#!/usr/bin/env python3
"""
OneTrack Auto-Restart Watchdog
Keeps OneTrack running. If it crashes, it restarts automatically.

Usage:
    python watchdog.py          # Start watching (run this instead of app.py)
    python watchdog.py --install # Install as Windows service (auto-start)
"""

import subprocess
import sys
import time
import os
import logging
from datetime import datetime

# Configuration
APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_SCRIPT = os.path.join(APP_DIR, "app.py")
PYTHON = sys.executable
CHECK_INTERVAL = 5  # seconds between checks
MAX_RESTARTS = 10  # max restarts before giving up
RESTART_DELAY = 10  # seconds to wait before restarting

# Setup logging
LOG_FILE = os.path.join(APP_DIR, "watchdog.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def is_app_running(process):
    """Check if the app process is still running."""
    if process is None:
        return False
    return process.poll() is None


def start_app():
    """Start the OneTrack app."""
    logger.info("Starting OneTrack...")
    try:
        process = subprocess.Popen(
            [PYTHON, APP_SCRIPT],
            cwd=APP_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        logger.info(f"OneTrack started with PID {process.pid}")
        return process
    except Exception as e:
        logger.error(f"Failed to start OneTrack: {e}")
        return None


def main():
    """Main watchdog loop."""
    logger.info("=" * 50)
    logger.info("OneTrack Watchdog Started")
    logger.info(f"App: {APP_SCRIPT}")
    logger.info(f"Python: {PYTHON}")
    logger.info("=" * 50)

    restart_count = 0
    process = None

    while restart_count < MAX_RESTARTS:
        # Start the app if not running
        if not is_app_running(process):
            if restart_count > 0:
                logger.info(f"Waiting {RESTART_DELAY} seconds before restart...")
                time.sleep(RESTART_DELAY)

            process = start_app()
            if process is None:
                logger.error("Failed to start app, waiting 30 seconds...")
                time.sleep(30)
                continue

            restart_count += 1
            logger.info(f"Restart count: {restart_count}/{MAX_RESTARTS}")

        # Check if app is still running
        time.sleep(CHECK_INTERVAL)

        if not is_app_running(process):
            logger.warning("OneTrack crashed! Restarting...")
            # Read any error output
            try:
                _, stderr = process.communicate(timeout=5)
                if stderr:
                    logger.error(f"Error output: {stderr.decode()[:500]}")
            except:
                pass

    logger.error(f"Max restarts ({MAX_RESTARTS}) reached. Watchdog stopping.")
    logger.error("Please check the logs and fix the issue.")
    logger.error("Then restart the watchdog.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--install":
        print("To install as Windows service, use setup_autostart.bat")
        print("This watchdog runs in the background and restarts OneTrack if it crashes.")
        print()
        print("To run the watchdog:")
        print("    python watchdog.py")
        print()
        print("The watchdog will:")
        print("    - Start OneTrack automatically")
        print("    - Restart it if it crashes")
        print("    - Log all activity to watchdog.log")
    else:
        main()
