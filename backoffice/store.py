"""Loads the back office from plain files and layers local writes on top.

Nothing is cached: every call reads `content/` and `state/` again, so editing a
content file or resetting the state shows up without a restart, and the MCP
process and the web view always agree.
"""

from __future__ import annotations

import csv
import json
import os
import re
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONTENT = ROOT / "content"
COMPANY = CONTENT / "company"
PACKS = CONTENT / "packs"
SNAPSHOT = ROOT / "snapshot"
STATE = Path(os.environ.get("BACKOFFICE_STATE", ROOT / "state"))
DEFAULT_PACK = "quote-mismatch"
SYSTEMS = ["mail", "chat", "tickets", "crm", "drive"]
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
    """Content folders in load order: the company first, then the active pack(s)."""
    pack = active_pack()
    roots = [("company", COMPANY)]
    if pack == "all":
        roots += [(p["id"], PACKS / p["id"]) for p in list_packs()]
    elif pack != "none":
        roots.append((pack, PACKS / pack))
    return roots


def pack_meta() -> dict:
    pack = active_pack()
    return {"id": pack, **(_yaml(PACKS / pack / "pack.yaml", {}) if pack not in ("all", "none") else {})}


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
            "source": "agent", "messages": [],
        })
        t["messages"].append({**rec["message"], "written": True})
    for t in threads.values():
        t["last_date"] = max((m["date"] for m in t["messages"]), default="")
        t["participants"] = sorted({e for m in t["messages"] for e in _emails(m["from"]) + _emails(m["to"]) + _emails(m["cc"])})
    return sorted(threads.values(), key=lambda t: t["last_date"], reverse=True)


def mailbox_addresses(person: dict) -> list[str]:
    return [person["email"].lower()] + [a.lower() for a in person.get("also_reads", [])]


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

_CHAT_LINE = re.compile(r"^\[(\d{2}:\d{2})\]\s+([^:]+):\s?(.*)$")


def channels() -> list[dict]:
    known = {c["name"]: c for c in _yaml(COMPANY / "channels.yaml", [])}
    for _, root in _roots():
        for folder in sorted((root / "chat").glob("*/")):
            known.setdefault(folder.name, {"name": folder.name, "purpose": ""})
    return list(known.values())


def chat_messages(channel: str | None = None) -> list[dict]:
    out = []
    for source, root in _roots():
        for path in sorted((root / "chat").glob("*/*.md")):
            if channel and path.parent.name != channel:
                continue
            meta, body = _frontmatter(path.read_text(encoding="utf-8"))
            date = str(meta.get("date", path.stem[:10]))
            for line in body.splitlines():
                m = _CHAT_LINE.match(line)
                if m:
                    out.append({"channel": path.parent.name, "date": date, "time": m.group(1),
                                "author": m.group(2).strip(), "text": m.group(3), "source": source})
                elif line.strip() and out:
                    out[-1]["text"] += "\n" + line
    for rec in _jsonl(STATE / "chat.jsonl"):
        if not channel or rec["channel"] == channel:
            out.append({**rec, "source": "agent", "written": True})
    return sorted(out, key=lambda m: (m["date"], m["time"]))


def post_chat(channel: str, text: str) -> dict:
    channel = channel.lstrip("#")
    if channel not in {c["name"] for c in channels()}:
        raise ValueError(f"No channel named {channel}. Channels: {', '.join(c['name'] for c in channels())}")
    date, time = story_now().split(" ")
    rec = {"channel": channel, "date": date, "time": time, "author": persona()["name"], "text": text}
    _append_jsonl(STATE / "chat.jsonl", rec)
    return rec


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
    numbers = [int(t.split("-")[1]) for t in tickets()]
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


def drive_files() -> list[dict]:
    files: dict[str, dict] = {}
    for source, root in _drive_roots():
        for path in sorted(root.rglob("*")) if root.exists() else []:
            if path.is_file() and not path.name.startswith((".", "~")):
                rel = path.relative_to(root).as_posix()
                files[rel] = {"path": rel, "size": path.stat().st_size, "source": source,
                              "local_path": str(path), **({"written": True} if source == "agent" else {})}
    return sorted(files.values(), key=lambda f: f["path"])


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
        doc = Document(str(local))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            parts += [" | ".join(c.text for c in row.cells) for row in table.rows]
        text = "\n".join(parts)
    elif suffix == ".pptx":
        from pptx import Presentation
        parts = []
        for i, slide in enumerate(Presentation(str(local)).slides, 1):
            parts.append(f"## Slide {i} (layout: {slide.slide_layout.name})")
            parts += [s.text_frame.text for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
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


def ground_truth() -> str:
    path = PACKS / active_pack() / "ground-truth.md"
    return path.read_text(encoding="utf-8") if path.exists() else ""
