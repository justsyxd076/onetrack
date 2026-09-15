#!/usr/bin/env python3
"""
OneTrack Auto-Sync
Run this on any PC to auto-commit and push changes to GitHub.

Usage:
    python auto_sync.py              # Start watching for changes
    python auto_sync.py --once       # Commit and push once, then exit
"""

import argparse
import hashlib
import os
import subprocess
import sys
import time

APP_DIR = os.path.dirname(os.path.abspath(__file__))
WATCH_FILES = ["app.py", "requirements.txt", ".env"]
WATCH_DIRS = ["static", "deploy"]
CHECK_INTERVAL = 5  # seconds


def get_file_hash(filepath):
    """Get hash of file content."""
    try:
        with open(filepath, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except (FileNotFoundError, PermissionError):
        return None


def get_all_hashes():
    """Get hashes of all watched files."""
    hashes = {}

    for f in WATCH_FILES:
        path = os.path.join(APP_DIR, f)
        hashes[f] = get_file_hash(path)

    for d in WATCH_DIRS:
        dir_path = os.path.join(APP_DIR, d)
        if os.path.exists(dir_path):
            for root, dirs, files in os.walk(dir_path):
                for file in files:
                    if not file.endswith((".pyc", ".pyo")):
                        rel_path = os.path.relpath(
                            os.path.join(root, file), APP_DIR
                        )
                        hashes[rel_path] = get_file_hash(
                            os.path.join(root, file)
                        )

    return hashes


def git_commit_and_push():
    """Commit and push changes to GitHub."""
    try:
        # Check if there are changes
        result = subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=APP_DIR,
            capture_output=True,
            text=True,
        )

        if not result.stdout.strip():
            return False

        # Count changed files
        changed_files = len(result.stdout.strip().split("\n"))

        # Commit
        subprocess.run(
            ["git", "add", "-A"],
            cwd=APP_DIR,
            check=True,
        )

        commit_msg = f"Auto-sync: {changed_files} file(s) updated at {time.strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(
            ["git", "commit", "-m", commit_msg],
            cwd=APP_DIR,
            check=True,
        )

        # Push
        subprocess.run(
            ["git", "push", "origin", "master"],
            cwd=APP_DIR,
            check=True,
        )

        print(f"[{time.strftime('%H:%M:%S')}] Pushed {changed_files} file(s)")
        return True

    except subprocess.CalledProcessError as e:
        print(f"Git error: {e}")
        return False


def watch():
    """Watch for file changes and auto-commit."""
    print("=== OneTrack Auto-Sync ===")
    print(f"Watching: {APP_DIR}")
    print(f"Files: {', '.join(WATCH_FILES)}")
    print(f"Dirs: {', '.join(WATCH_DIRS)}")
    print(f"Check interval: {CHECK_INTERVAL}s")
    print("Press Ctrl+C to stop\n")

    last_hashes = get_all_hashes()

    try:
        while True:
            time.sleep(CHECK_INTERVAL)
            current_hashes = get_all_hashes()

            if current_hashes != last_hashes:
                print(f"[{time.strftime('%H:%M:%S')}] Changes detected...")
                if git_commit_and_push():
                    last_hashes = current_hashes
                else:
                    last_hashes = current_hashes

    except KeyboardInterrupt:
        print("\nAuto-sync stopped.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="OneTrack Auto-Sync")
    parser.add_argument("--once", action="store_true", help="Commit and push once")
    args = parser.parse_args()

    if args.once:
        git_commit_and_push()
    else:
        watch()
