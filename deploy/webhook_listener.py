#!/usr/bin/env python3
"""
GitHub Webhook Listener for OneTrack
Listens for GitHub push events and auto-deploys.

Setup:
1. Run this script on the VM
2. Go to GitHub repo → Settings → Webhooks → Add webhook
3. Payload URL: http://YOUR_VM_IP:9000
4. Content type: application/json
5. Secret: (copy from WEBHOOK_SECRET env var)
6. Events: Just the push event
"""

import hashlib
import hmac
import json
import os
import subprocess
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler

WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "onetrack-webhook-secret-change-me")
PORT = int(os.environ.get("WEBHOOK_PORT", "9000"))
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def verify_signature(payload_body, signature_header):
    """Verify GitHub webhook signature."""
    if not signature_header:
        return False
    expected = "sha256=" + hmac.new(
        WEBHOOK_SECRET.encode(), payload_body, hashlib.sha256
    ).hexdigest()
    return hmac.compare_digest(expected, signature_header)


def deploy():
    """Pull latest code and restart the app."""
    print(f"[{os.popen('date').read().strip()}] Deploying...")
    try:
        subprocess.run(["git", "pull"], cwd=APP_DIR, check=True)
        subprocess.run(
            ["sudo", "systemctl", "restart", "onetrack"], check=True
        )
        print("Deployment successful!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Deployment failed: {e}")
        return False


class WebhookHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        payload_body = self.rfile.read(content_length)
        signature = self.headers.get("X-Hub-Signature-256", "")

        if not verify_signature(payload_body, signature):
            self.send_response(403)
            self.end_headers()
            self.wfile.write(b"Invalid signature")
            return

        try:
            data = json.loads(payload_body)
            if data.get("ref") == "refs/heads/master":
                deploy()
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"Deployed successfully")
            else:
                self.send_response(200)
                self.end_headers()
                self.wfile.write(b"Ignored (not master branch)")
        except Exception as e:
            self.send_response(500)
            self.end_headers()
            self.wfile.write(str(e).encode())

    def log_message(self, format, *args):
        print(f"[Webhook] {args[0]}")


if __name__ == "__main__":
    server = HTTPServer(("0.0.0.0", PORT), WebhookHandler)
    print(f"Webhook listener running on port {PORT}")
    print(f"GitHub webhook URL: http://YOUR_VM_IP:{PORT}")
    server.serve_forever()
