"""The web view: one page that shows what is in the back office and what the agent changed."""

from __future__ import annotations

from pathlib import Path

from starlette.requests import Request
from starlette.responses import FileResponse, HTMLResponse, JSONResponse, Response

from . import store
from .server import TOOLS, bookings_summary, mcp

PAGE = Path(__file__).parent / "static" / "index.html"


@mcp.custom_route("/", methods=["GET"])
async def index(request: Request) -> Response:
    return HTMLResponse(PAGE.read_text(encoding="utf-8"))


@mcp.custom_route("/api/state", methods=["GET"])
async def state(request: Request) -> Response:
    perms = store.permissions()
    return JSONResponse({
        "company": store.company(),
        "today": store.today(),
        "persona": store.persona(),
        "pack": store.pack_meta(),
        "packs": store.list_packs(),
        "people": store.people(),
        "mail": store.mail_threads(),
        "channels": store.channels(),
        "chat": store.chat_messages(),
        "tickets": sorted(store.tickets().values(), key=lambda t: t["id"], reverse=True),
        "accounts": list(store.accounts().values()),
        "deals": list(store.deals().values()),
        "drive": store.drive_files(),
        "bookings": {
            "by_month": bookings_summary("month")["items"],
            "by_route": bookings_summary("route")["items"],
            "by_class": bookings_summary("seat_class")["items"],
            "quotes": len(store.snapshot("quotes")),
        },
        "audit": list(reversed(store.audit_log())),
        "permissions": perms,
        "tools": [{**t, "allowed": t["kind"] == "read" or perms.get(t["system"], True)} for t in TOOLS],
    })


@mcp.custom_route("/api/ground-truth", methods=["GET"])
async def ground_truth(request: Request) -> Response:
    return JSONResponse({"text": store.ground_truth()})


@mcp.custom_route("/api/reset", methods=["POST"])
async def reset(request: Request) -> Response:
    store.reset()
    return JSONResponse({"ok": True})


@mcp.custom_route("/api/pack", methods=["POST"])
async def pack(request: Request) -> Response:
    store.set_active_pack((await request.json())["pack"])
    return JSONResponse({"ok": True})


@mcp.custom_route("/api/permission", methods=["POST"])
async def permission(request: Request) -> Response:
    body = await request.json()
    store.set_permission(body["system"], bool(body["allowed"]))
    return JSONResponse({"ok": True})


@mcp.custom_route("/drive/{path:path}", methods=["GET"])
async def drive_file(request: Request) -> Response:
    entry = next((f for f in store.drive_files() if f["path"] == request.path_params["path"]), None)
    if not entry:
        return JSONResponse({"error": "not found"}, status_code=404)
    return FileResponse(entry["local_path"], filename=Path(entry["path"]).name)
