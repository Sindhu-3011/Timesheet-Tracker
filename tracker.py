#!/usr/bin/env python3
"""
tracker.py — Timesheet Tracker entry point.

Starts the HTTP server.  All application logic lives under app/.

Usage:
    pip install -r requirements.txt
    python tracker.py
"""
import logging
from http.server import HTTPServer

# Bootstrap DB + start email scheduler (imported as side-effects)
from app.database.migrations import init_db
from app.services.email import EMAIL_SCHEDULER_STARTED  # noqa: F401
from app.routes.handler import Handler
from app.config import HOST, PORT, DEBUG, DB_FILE, UPLOAD_DIR


def main() -> None:
    init_db()
    addr = (HOST, PORT)
    logging.info("Timesheet Tracker starting on http://%s:%s", HOST, PORT)
    logging.info("Debug=%s | DB=%s | Uploads=%s", DEBUG, DB_FILE, UPLOAD_DIR)
    server = HTTPServer(addr, Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logging.info("Server stopped.")


if __name__ == "__main__":
    main()
