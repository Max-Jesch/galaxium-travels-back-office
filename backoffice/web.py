"""The web view: one page that shows what is in the back office and what the agent changed."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from starlette.concurrency import run_in_threadpool
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
    me = store.persona()
    bookings = store.snapshot("bookings")
    taken = Counter(b["flight_id"] for b in bookings if b["status"] != "cancelled")
    users = {u["user_id"]: u["name"] for u in store.snapshot("users")}
    flights = [{**f, "booked": taken[f["flight_id"]],
                "seats": f["economy_seats"] + f["business_seats"] + f["galaxium_seats"]} for f in store.snapshot("flights")]
    return JSONResponse({
        "version": store.version(),
        "my_addresses": store.mailbox_addresses(me),
        "events": store.events(),
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
        "drive": [{k: v for k, v in f.items() if k != "local_path"} for f in store.drive_files()],
        "customers": store.customers(),
        "bookings": {
            "by_month": bookings_summary("month")["items"],
            "by_route": bookings_summary("route")["items"],
            "by_class": bookings_summary("seat_class")["items"],
            "quotes": len(store.snapshot("quotes")),
            "total": len(bookings),
            "flights": flights,
            "recent": [{**b, "customer": users.get(b["user_id"], "")} for b in bookings[-60:]][::-1],
        },
        "audit": list(reversed(store.audit_log())),
        "permissions": perms,
        "tools": [{**t, "allowed": t["kind"] == "read" or perms.get(t["system"], True)} for t in TOOLS],
    })


@mcp.custom_route("/api/version", methods=["GET"])
async def version(request: Request) -> Response:
    return JSONResponse({"version": store.version()})


@mcp.custom_route("/api/drive-preview/{path:path}", methods=["GET"])
async def drive_preview(request: Request) -> Response:
    try:
        body = json.dumps(store.drive_preview(request.path_params["path"]), default=str)
        return Response(body, media_type="application/json")
    except ValueError as exc:
        return JSONResponse({"error": str(exc)}, status_code=404)


@mcp.custom_route("/api/drive-pdf/{path:path}", methods=["GET"])
async def drive_pdf(request: Request) -> Response:
    """A deck or document as pages, for the preview. Rendering takes a few seconds the first time."""
    pdf = await run_in_threadpool(store.drive_pdf, request.path_params["path"])
    if not pdf:
        return JSONResponse({"error": "no rendering available"}, status_code=404)
    return FileResponse(pdf, media_type="application/pdf", content_disposition_type="inline")


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
