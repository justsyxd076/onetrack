#!/usr/bin/env python3
"""
OneTrack Startup Script for Waifly
This script is executed by Waifly when the server starts.
"""

import os
import sys
import subprocess

# Install dependencies
print("Installing dependencies...")
subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)

# Initialize database
print("Initializing database...")
from database import init_db, seed_defaults, delete_old_responses
init_db()
seed_defaults()

# Delete old responses (older than 2 years)
deleted = delete_old_responses(years=2)
if deleted:
    print(f"Deleted {deleted} old responses (older than 2 years)")

# Start the app with gunicorn
print("Starting OneTrack...")
port = os.environ.get("PORT", "8000")
subprocess.run([
    sys.executable, "-m", "gunicorn",
    "--bind", f"0.0.0.0:{port}",
    "--workers", "1",
    "--timeout", "120",
    "app:app"
], check=True)
