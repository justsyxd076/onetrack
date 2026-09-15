#!/usr/bin/env python3
"""
OneTrack Keep-Alive Script
Prevents Oracle Cloud from reclaiming your Always Free VM.
Runs as a cron job every 6 hours.

Setup (on the VM):
    crontab -e
    Add: 0 */6 * * * /home/ubuntu/onetrack/venv/bin/python /home/ubuntu/onetrack/deploy/keepalive.py
"""

import urllib.request
import datetime

APP_URL = "http://localhost:5000"
LOG_FILE = "/home/ubuntu/onetrack/logs/keepalive.log"


def keepalive():
    """Ping the app to prevent VM reclaim."""
    try:
        req = urllib.request.Request(APP_URL)
        urllib.request.urlopen(req, timeout=10)
        status = "OK"
    except Exception as e:
        status = f"ERROR: {e}"

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] Keep-alive ping: {status}\n"

    try:
        with open(LOG_FILE, "a") as f:
            f.write(log_entry)
    except FileNotFoundError:
        import os
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "w") as f:
            f.write(log_entry)


if __name__ == "__main__":
    keepalive()
