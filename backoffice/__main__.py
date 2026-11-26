"""Start the back office.

    python -m backoffice                 web view and MCP on http://localhost:8090
    python -m backoffice --stdio         MCP over stdio, for clients that launch a command
    python -m backoffice reset           throw away everything the agent wrote
    python -m backoffice packs           list the scenario packs
"""

from __future__ import annotations

import argparse
import os
import sys
import threading

from . import store


def main() -> None:
    parser = argparse.ArgumentParser(prog="backoffice", description="Galaxium Travels back office")
    parser.add_argument("command", nargs="?", default="serve", choices=["serve", "reset", "packs"])
    parser.add_argument("--pack", help="scenario pack to load (see 'packs'), or 'all' or 'none'")
    parser.add_argument("--port", type=int, default=8090)
    parser.add_argument("--stdio", action="store_true", help="serve MCP over stdio instead of HTTP")
    args = parser.parse_args()

    if args.pack:
        store.set_active_pack(args.pack)

    if args.command == "packs":
        for p in store.list_packs():
            mark = "*" if p["id"] == store.active_pack() else " "
            print(f"{mark} {p['id']:<16} {p['title']}")
        return
    if args.command == "reset":
        store.reset()
        print("Back office reset to its starting state.")
        return

    from . import web  # noqa: F401  (registers the web view routes)
    from .server import mcp

    os.environ.setdefault("BACKOFFICE_URL", f"http://localhost:{args.port}")
    if args.stdio:
        mcp.run()
        return

    mcp.settings.host = "127.0.0.1"
    mcp.settings.port = args.port
    pack = store.pack_meta()
    print(
        f"\n  Galaxium back office\n"
        f"  Web view   http://localhost:{args.port}\n"
        f"  MCP        http://localhost:{args.port}/mcp\n"
        f"  Scenario   {pack.get('title', pack['id'])}\n"
        f"  Desk       {store.persona()['name']}, {store.persona()['role']}\n",
        file=sys.stderr,
    )
    threading.Thread(target=store.warm_drive_text, daemon=True).start()
    mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
