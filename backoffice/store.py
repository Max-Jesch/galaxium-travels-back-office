"""Loads the back office from plain files and layers local writes on top.

Nothing is cached: every call reads `content/` and `state/` again, so editing a
content file or resetting the state shows up without a restart, and the MCP
process and the web view always agree.
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONTENT = ROOT / "content"
COMPANY = CONTENT / "company"
GENERATED = CONTENT / "generated"  # routine volume written by scripts/build.py
PACKS = CONTENT / "packs"
SNAPSHOT = ROOT / "snapshot"
STATE = Path(os.environ.get("BACKOFFICE_STATE", ROOT / "state"))
DEFAULT_PACK = "quote-mismatch"
SYSTEMS = ["mail", "chat", "calendar", "tickets", "crm", "drive"]
TEXT_SUFFIXES = {".md", ".txt", ".csv", ".yaml", ".yml", ".json"}


# ---------- small file helpers ----------

def _plain(value: Any) -> Any:
    """YAML turns 2026-11-04 into a date object. Keep dates as the text the author wrote."""
    if isinstance(value, dict):
        return {k: _plain(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_plain(v) for v in value]
    return str(value) if isinstance(value, (date, datetime)) else value


def _yaml(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default
    return _plain(yaml.safe_load(path.read_text(encoding="utf-8"))) or default


def _json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def _jsonl(path: Path) -> list[dict]:
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def _append_jsonl(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def _frontmatter(text: str) -> tuple[dict, str]:
    if text.startswith("---"):
        _, head, body = text.split("---", 2)
        return _plain(yaml.safe_load(head)) or {}, body.lstrip("\n")
    return {}, text


# ---------- packs, company, clock ----------

def list_packs() -> list[dict]:
    packs = []
    for folder in sorted(PACKS.iterdir()) if PACKS.exists() else []:
        meta = _yaml(folder / "pack.yaml")
        if meta:
            packs.append({"id": folder.name, **meta})
    return packs


def active_pack() -> str:
    marker = STATE / "active_pack.txt"
    name = marker.read_text().strip() if marker.exists() else os.environ.get("BACKOFFICE_PACK", DEFAULT_PACK)
    return name if name in ("all", "none") or (PACKS / name / "pack.yaml").exists() else DEFAULT_PACK


def set_active_pack(name: str) -> None:
    if name not in ("all", "none") and not (PACKS / name / "pack.yaml").exists():
        raise ValueError(f"Unknown scenario pack: {name}")
    STATE.mkdir(parents=True, exist_ok=True)
    (STATE / "active_pack.txt").write_text(name)


def _roots() -> list[tuple[str, Path]]:
    """Content folders in load order: the company, its generated routine, then the active pack(s)."""
    pack = active_pack()
    roots = [("company", COMPANY), ("generated", GENERATED)]
    if pack == "all":
        roots += [(p["id"], PACKS / p["id"]) for p in list_packs()]
    elif pack != "none":
        roots.append((pack, PACKS / pack))
    return roots


def pack_meta() -> dict:
    pack = active_pack()
    return {"id": pack, **(_yaml(PACKS / pack / "pack.yaml", {}) if pack not in ("all", "none") else {})}


def web_url() -> str:
    """Where the web view is served; `__main__` sets BACKOFFICE_URL when it starts on another port."""
    return os.environ.get("BACKOFFICE_URL", "http://localhost:8090").rstrip("/")


def company() -> dict:
    return _yaml(COMPANY / "company.yaml", {})


def today() -> str:
    """The story's date. Fixed, so that 'last week' means the same for everyone."""
    return str(pack_meta().get("today") or company().get("today"))


def story_now() -> str:
    return f"{today()} {datetime.now().strftime('%H:%M')}"


def people() -> list[dict]:
    return _yaml(COMPANY / "people.yaml", [])


def customers() -> list[dict]:
    return _yaml(COMPANY / "customers.yaml", [])


def persona() -> dict:
    wanted = pack_meta().get("persona") or company().get("persona")
    return next((p for p in people() if p["id"] == wanted), people()[0])


def find_person(key: str) -> dict | None:
    key = key.strip().lower()
    for p in people() + customers():
        if key in (str(p.get("id", "")).lower(), p["name"].lower(), p["email"].lower()):
            return p
    return None


# ---------- mail ----------

_ADDR = re.compile(r"<([^>]+)>")


def _emails(value: str) -> list[str]:
    out = []
    for part in (value or "").split(","):
        part = part.strip()
        if part:
            m = _ADDR.search(part)
            out.append((m.group(1) if m else part).lower())
    return out


def _parse_thread(path: Path, source: str) -> dict:
    meta, body = _frontmatter(path.read_text(encoding="utf-8"))
    messages = []
    for block in re.split(r"(?m)^\s*---\s*$", body):
        block = block.strip("\n")
        if not block.strip():
            continue
        head, _, text = block.partition("\n\n")
        headers = {}
        for line in head.splitlines():
            key, _, value = line.partition(":")
            headers[key.strip().lower()] = value.strip()
        messages.append({
            "from": headers.get("from", ""),
            "to": headers.get("to", ""),
            "cc": headers.get("cc", ""),
            "date": headers.get("date", ""),
            "body": text.strip(),
        })
    return {
        "id": meta.get("id", path.stem),
        "subject": meta.get("subject", path.stem),
        "labels": meta.get("labels", []),
        "attachments": meta.get("attachments", []),
        "flagged": bool(meta.get("flagged", False)),
        "source": source,
        "messages": messages,
    }


def mail_threads() -> list[dict]:
    threads: dict[str, dict] = {}
    for source, root in _roots():
        for path in sorted((root / "mail").glob("*.md")):
            t = _parse_thread(path, source)
            threads[t["id"]] = t
    for rec in _jsonl(STATE / "mail.jsonl"):
        t = threads.setdefault(rec["thread_id"], {
            "id": rec["thread_id"], "subject": rec["subject"], "labels": ["sent"],
            "attachments": [], "flagged": False, "source": "agent", "messages": [],
        })
        t["messages"].append({**rec["message"], "written": True})
    for t in threads.values():
        t["last_date"] = max((m["date"] for m in t["messages"]), default="")
        t["participants"] = sorted({e for m in t["messages"] for e in _emails(m["from"]) + _emails(m["to"]) + _emails(m["cc"])})
    return sorted(threads.values(), key=lambda t: t["last_date"], reverse=True)


def mailbox_addresses(person: dict) -> list[str]:
    """A person's own address, the shared mailboxes they read, and the lists they are on."""
    domain = company().get("domain", "galaxium.example")
    lists = [f"all@{domain}", f"{person.get('department', '').lower()}@{domain}"] if person.get("department") else []
    return list(dict.fromkeys([person["email"].lower()] + [a.lower() for a in person.get("also_reads", [])] + lists))


def send_mail(to: str, subject: str, body: str, cc: str = "", thread_id: str = "") -> dict:
    me = persona()
    threads = {t["id"]: t for t in mail_threads()}
    if thread_id and thread_id not in threads:
        raise ValueError(f"No mail thread with id {thread_id}")
    if not thread_id:
        thread_id = f"sent-{len(_jsonl(STATE / 'mail.jsonl')) + 1:03d}"
    else:
        subject = threads[thread_id]["subject"]
    message = {"from": f"{me['name']} <{me['email']}>", "to": to, "cc": cc, "date": story_now(), "body": body}
    _append_jsonl(STATE / "mail.jsonl", {"thread_id": thread_id, "subject": subject, "message": message})
    return {"thread_id": thread_id, "subject": subject, **message}


# ---------- chat ----------
#
# A channel is a folder of markdown files. Inside a file:
#   ## 2026-10-26                 starts a new day (or give `date:` in the frontmatter)
#   [10:02] Name: text            a message; following plain lines continue it
#       [10:05] Name: text        indented: a reply in the thread of the message above
#       + 👍 3, 🎉 1               indented, starting with +: reactions to the message above

_CHAT_LINE = re.compile(r"^(\s*)\[(\d{2}:\d{2})\]\s+([^:]+):\s?(.*)$")
_DAY_LINE = re.compile(r"^##\s+(\d{4}-\d{2}-\d{2})\s*$")
_REACTIONS = re.compile(r"^\s+\+\s+(.*)$")


def channels(everyone: bool = False) -> list[dict]:
    """Channels the person at the desk can see: every public one and their own direct messages."""
    known = {c["name"]: {"type": "channel", "purpose": "", **c} for c in _yaml(COMPANY / "channels.yaml", [])}
    for _, root in _roots():
        for folder in sorted((root / "chat").glob("*/")):
            known.setdefault(folder.name, {"name": folder.name, "type": "channel", "purpose": ""})
    me = persona()["id"]
    return [c for c in known.values() if everyone or c["type"] != "dm" or me in c.get("members", [])]


def _parse_chat(path: Path, source: str) -> list[dict]:
    meta, body = _frontmatter(path.read_text(encoding="utf-8"))
    day = str(meta.get("date", path.stem[:10]))
    out: list[dict] = []
    last: dict | None = None  # the message a continuation line belongs to
    for line in body.splitlines():
        if m := _DAY_LINE.match(line):
            day, last = m.group(1), None
        elif m := _CHAT_LINE.match(line):
            msg = {"channel": path.parent.name, "date": day, "time": m.group(2),
                   "author": m.group(3).strip(), "text": m.group(4), "source": source}
            if m.group(1) and out:
                out[-1].setdefault("replies", []).append(msg)
            else:
                out.append(msg)
            last = msg
        elif (m := _REACTIONS.match(line)) and out:
            out[-1]["reactions"] = [dict(zip(("emoji", "count"), (r.split() + ["1"])[:2])) for r in m.group(1).split(",") if r.strip()]
        elif line.strip() and last:
            last["text"] += "\n" + line.strip()
    return out


def chat_messages(channel: str | None = None) -> list[dict]:
    visible = {c["name"] for c in channels()}
    out = []
    for source, root in _roots():
        for path in sorted((root / "chat").glob("*/*.md")):
            if path.parent.name in visible and (not channel or path.parent.name == channel):
                out += _parse_chat(path, source)
    for rec in _jsonl(STATE / "chat.jsonl"):
        if rec["channel"] in visible and (not channel or rec["channel"] == channel):
            out.append({**rec, "source": "agent", "written": True})
    return sorted(out, key=lambda m: (m["date"], m["time"]))


def post_chat(channel: str, text: str) -> dict:
    channel = channel.lstrip("#")
    names = [c["name"] for c in channels()]
    if channel not in names:
        raise ValueError(f"No channel named {channel}. Channels: {', '.join(names)}")
    date, time = story_now().split(" ")
    rec = {"channel": channel, "date": date, "time": time, "author": persona()["name"], "text": text}
    _append_jsonl(STATE / "chat.jsonl", rec)
    return rec


# ---------- calendar ----------

def events() -> list[dict]:
    """Events from `<root>/calendar/*.yaml` (each file a list) plus the ones written at the desk."""
    out: dict[str, dict] = {}
    for source, root in _roots():
        for path in sorted((root / "calendar").glob("*.yaml")):
            for item in _yaml(path, []):
                out[str(item["id"])] = {"attendees": [], "location": "", "notes": "", **item, "source": source}
    for eid, item in _json(STATE / "calendar.json", {}).items():
        out[eid] = {**item, "written": True}
    return sorted(out.values(), key=lambda e: e["start"])


def create_event(fields: dict) -> dict:
    saved = _json(STATE / "calendar.json", {})
    event = {"id": f"EV-A{len(saved) + 1:03d}", "organizer": persona()["name"], "source": "agent", **fields}
    saved[event["id"]] = event
    _write_json(STATE / "calendar.json", saved)
    return event


# ---------- records: tickets and CRM ----------

def _records(kind: str, folder: str) -> dict[str, dict]:
    """Records from `<root>/<folder>/<kind>.yaml` (a list) and `<root>/<folder>/<kind>/*.yaml` (one each)."""
    out: dict[str, dict] = {}
    for source, root in _roots():
        items = list(_yaml(root / folder / f"{kind}.yaml", []))
        items += [_yaml(p) for p in sorted((root / folder / kind).glob("*.yaml"))]
        for item in items:
            out[str(item["id"])] = {**item, "source": source}
    for rid, item in _json(STATE / f"{kind}.json", {}).items():
        out[rid] = {**item, "written": True}
    return out


def _save_record(kind: str, record: dict) -> None:
    saved = _json(STATE / f"{kind}.json", {})
    saved[str(record["id"])] = {k: v for k, v in record.items() if k not in ("source", "written")}
    _write_json(STATE / f"{kind}.json", saved)


def tickets() -> dict[str, dict]:
    return _records("tickets", ".")


def accounts() -> dict[str, dict]:
    return _records("accounts", "crm")


def deals() -> dict[str, dict]:
    return _records("deals", "crm")


def create_ticket(fields: dict) -> dict:
    numbers = [int(t.split("-")[1]) for t in tickets() if t.startswith("GT-")]
    ticket = {"id": f"GT-{max(numbers, default=1000) + 1}", "status": "open", "created": story_now(),
              "created_by": persona()["name"], "comments": [], **fields}
    _save_record("tickets", ticket)
    return ticket


def update_record(kind: str, record_id: str, changes: dict, note: str = "") -> dict:
    records = _records(kind, "." if kind == "tickets" else "crm")
    if record_id not in records:
        raise ValueError(f"No {kind[:-1]} with id {record_id}")
    record = records[record_id]
    record.update({k: v for k, v in changes.items() if v not in (None, "")})
    if note:
        key = "comments" if kind == "tickets" else "notes"
        record[key] = list(record.get(key, [])) + [{"author": persona()["name"], "date": story_now(), "text": note}]
    record["updated"] = story_now()
    _save_record(kind, record)
    return record


# ---------- drive ----------

def _drive_roots() -> list[tuple[str, Path]]:
    return [(s, r / "drive") for s, r in _roots()] + [("agent", STATE / "drive")]


_DATED = re.compile(r"(20\d{2})-(0[1-9]|1[0-2])(?:-(0[1-9]|[12]\d|3[01]))?(?!\d)")


def _drive_meta(path: Path, rel: str, source: str, index: dict, owners: dict) -> dict:
    """Owner and last change of a file: from its frontmatter, the folder's `_index.yaml`, or a stable guess."""
    if source == "agent":
        return {"owner": persona()["name"], "modified": today()}
    meta = dict(index.get(rel, {}))
    if path.suffix == ".md" and not meta:
        with path.open(encoding="utf-8") as f:
            if f.readline().startswith("---"):
                head = f.read(600).split("---", 1)[0]
                meta = {k: v for k, v in (_plain(yaml.safe_load(head)) or {}).items() if k in ("owner", "modified")}
    if "modified" not in meta:
        day = date.fromisoformat(str(company().get("today")))
        if m := _DATED.search(path.name):  # a dated file was last touched shortly after its date
            named = date(int(m.group(1)), int(m.group(2)), int(m.group(3) or 28))
            meta["modified"] = str(min(day, named + timedelta(days=2)))
        else:
            meta["modified"] = str(day - timedelta(days=int(hashlib.md5(rel.encode()).hexdigest(), 16) % 240 + 6))
    meta.setdefault("owner", owners.get(rel.split("/")[0], ""))
    return meta


def drive_files() -> list[dict]:
    files: dict[str, dict] = {}
    owners = company().get("drive_owners", {})
    for source, root in _drive_roots():
        index = _yaml(root / "_index.yaml", {})
        for path in sorted(root.rglob("*")) if root.exists() else []:
            if path.is_file() and not path.name.startswith((".", "~", "_")):
                rel = path.relative_to(root).as_posix()
                files[rel] = {"path": rel, "size": path.stat().st_size, "source": source, "local_path": str(path),
                              **_drive_meta(path, rel, source, index, owners),
                              **({"written": True} if source == "agent" else {})}
    return sorted(files.values(), key=lambda f: f["path"])


def drive_preview(path: str) -> dict:
    """What the web view shows for a file: sheets as rows, everything else as text."""
    entry = next((f for f in drive_files() if f["path"] == path.strip("/")), None)
    if not entry:
        raise ValueError(f"No file at {path} in the drive")
    local = Path(entry["local_path"])
    public = {k: v for k, v in entry.items() if k != "local_path"}
    if local.suffix.lower() in (".xlsx", ".csv"):
        if local.suffix.lower() == ".csv":
            with local.open(encoding="utf-8", newline="") as f:
                sheets = [{"name": local.stem, "rows": [r for _, r in zip(range(301), csv.reader(f))]}]
        else:
            from openpyxl import load_workbook
            sheets = [{"name": ws.title, "total": ws.max_row,
                       "rows": [["" if v is None else v for v in row] for row in ws.iter_rows(max_row=301, values_only=True)]}
                      for ws in load_workbook(local, data_only=True, read_only=True).worksheets]
        return {**public, "kind": "sheet", "sheets": sheets}
    text = drive_text(path)["text"]
    if local.suffix.lower() == ".md":
        text = _frontmatter(text)[1]
    return {**public, "kind": "markdown" if local.suffix.lower() in (".md", ".pptx") else "text", "text": text,
            "rendered": local.suffix.lower() in RENDERED_SUFFIXES and bool(_OFFICE or _stored_pdf(local))}


RENDERED_SUFFIXES = {".pptx", ".docx"}
_OFFICE = shutil.which("soffice") or next((c for c in ("/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/libreoffice") if Path(c).exists()), None)
_RENDER_LOCK = threading.Lock()


def _stored_pdf(local: Path) -> Path | None:
    try:
        root, rel = local.relative_to(CONTENT).as_posix().split("/drive/", 1)
    except ValueError:
        return None
    pdf = CONTENT / "previews" / root / (rel + ".pdf")
    return pdf if pdf.exists() else None


def drive_pdf(path: str) -> Path | None:
    """The file as it looks when opened. The pages of the files in the repo are stored under
    content/previews by scripts/previews.py. Anything else is rendered by LibreOffice, if it
    is installed, once per version of the file, and kept in the temp folder."""
    entry = next((f for f in drive_files() if f["path"] == path.strip("/")), None)
    if not entry or Path(entry["local_path"]).suffix.lower() not in RENDERED_SUFFIXES:
        return None
    local = Path(entry["local_path"])
    if stored := _stored_pdf(local):
        return stored
    if not _OFFICE:
        return None
    stat = local.stat()
    cache = Path(tempfile.gettempdir()) / "galaxium-backoffice-previews"
    target = cache / (hashlib.sha1(f"{local}|{stat.st_size}|{stat.st_mtime_ns}".encode()).hexdigest() + ".pdf")
    with _RENDER_LOCK:  # one conversion at a time: LibreOffice shares a profile
        if not target.exists():
            work = cache / "work"
            shutil.rmtree(work, ignore_errors=True)
            work.mkdir(parents=True)
            try:
                subprocess.run([_OFFICE, "--headless", f"-env:UserInstallation={(cache / 'profile').as_uri()}", "--convert-to", "pdf",
                                "--outdir", str(work), str(local)], capture_output=True, timeout=90, check=False)
            except (OSError, subprocess.TimeoutExpired):
                return None
            made = work / (local.stem + ".pdf")
            if not made.exists():
                return None
            made.replace(target)
    return target


def drive_text(path: str, max_chars: int = 60000) -> dict:
    entry = next((f for f in drive_files() if f["path"] == path.strip("/")), None)
    if not entry:
        raise ValueError(f"No file at {path} in the drive")
    local = Path(entry["local_path"])
    suffix = local.suffix.lower()
    if suffix in TEXT_SUFFIXES:
        text = local.read_text(encoding="utf-8")
    elif suffix == ".xlsx":
        from openpyxl import load_workbook
        parts = []
        for ws in load_workbook(local, data_only=True).worksheets:
            parts.append(f"## Sheet: {ws.title} ({ws.max_row} rows)")
            for row in ws.iter_rows(values_only=True):
                parts.append(",".join("" if v is None else str(v) for v in row))
        text = "\n".join(parts)
    elif suffix == ".docx":
        from docx import Document
        parts = []
        for block in Document(str(local)).iter_inner_content():  # paragraphs and tables in the order they appear
            if hasattr(block, "rows"):
                parts += [" | ".join(dict.fromkeys(c.text.strip() for c in row.cells)) for row in block.rows]
            elif block.text.strip():
                parts.append(block.text)
        text = "\n".join(parts)
    elif suffix == ".pptx":
        from pptx import Presentation
        parts = []
        for i, slide in enumerate(Presentation(str(local)).slides, 1):
            parts.append(f"## Slide {i} (layout: {slide.slide_layout.name})")
            for s in slide.shapes:
                if s.has_text_frame and s.text_frame.text.strip():
                    parts.append(s.text_frame.text)
                elif getattr(s, "has_table", False) and s.has_table:
                    rows = [[c.text.replace("\n", " ") for c in row.cells] for row in s.table.rows]
                    parts += ["", "| " + " | ".join(rows[0]) + " |", "|" + " --- |" * len(rows[0])] + ["| " + " | ".join(r) + " |" for r in rows[1:]] + [""]
                elif getattr(s, "has_chart", False) and s.has_chart:
                    plot = s.chart.plots[0]
                    cats = [str(c) for c in plot.categories]
                    parts += ["", "| Chart | " + " | ".join(cats) + " |", "|" + " --- |" * (len(cats) + 1)]
                    parts += ["| " + series.name + " | " + " | ".join(f"{v:,.0f}" if isinstance(v, (int, float)) else "" for v in series.values) + " |" for series in plot.series] + [""]
        text = "\n".join(parts)
    else:
        text = "(binary file, open it from local_path)"
    return {**entry, "text": text[:max_chars], "truncated": len(text) > max_chars}


def drive_write(path: str, content: str) -> dict:
    rel = Path(path.strip("/"))
    if rel.is_absolute() or ".." in rel.parts or rel.suffix.lower() not in TEXT_SUFFIXES:
        raise ValueError("Path must stay inside the drive and end in a text suffix such as .md or .csv")
    target = STATE / "drive" / rel
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8")
    return {"path": rel.as_posix(), "size": target.stat().st_size, "local_path": str(target)}


# ---------- booking snapshot ----------

def snapshot(table: str) -> list[dict]:
    with (SNAPSHOT / f"{table}.csv").open(encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f))
    for row in rows:
        for key, value in row.items():
            if value.isdigit():
                row[key] = int(value)
    return rows


# ---------- permissions, audit, reset ----------

def permissions() -> dict[str, bool]:
    base = {s: True for s in SYSTEMS}
    base.update((_yaml(ROOT / "permissions.yaml", {}) or {}).get("write", {}))
    base.update(_json(STATE / "permissions.json", {}))
    return base


def set_permission(system: str, allowed: bool) -> None:
    if system not in SYSTEMS:
        raise ValueError(f"Unknown system: {system}")
    saved = _json(STATE / "permissions.json", {})
    saved[system] = allowed
    _write_json(STATE / "permissions.json", saved)


def audit(tool: str, system: str, args: dict, outcome: str, result: str) -> None:
    _append_jsonl(STATE / "audit.jsonl", {
        "time": datetime.now().isoformat(timespec="seconds"), "actor": persona()["name"],
        "tool": tool, "system": system, "args": args, "outcome": outcome, "result": result,
    })


def audit_log() -> list[dict]:
    return _jsonl(STATE / "audit.jsonl")


def reset() -> None:
    """Throw away everything the agent wrote. The active pack is kept."""
    pack = active_pack()
    if STATE.exists():
        shutil.rmtree(STATE)
    set_active_pack(pack)


def version() -> str:
    """Changes whenever content or state changes, so the web view knows when to reload."""
    stamp = 0.0
    count = 0
    for folder in (CONTENT, STATE):
        for path in folder.rglob("*") if folder.exists() else []:
            count += 1
            stamp = max(stamp, path.stat().st_mtime)
    return f"{active_pack()}-{count}-{stamp}"


def ground_truth() -> str:
    path = PACKS / active_pack() / "ground-truth.md"
    return path.read_text(encoding="utf-8") if path.exists() else ""
