#!/usr/bin/env python3
"""
OneTrack Deploy Script
Run this on your local PC to deploy to Oracle Cloud VM.

Usage:
    python deploy/deploy.py              # Deploy latest code
    python deploy/deploy.py --setup      # First-time setup only
    python deploy/deploy.py --status     # Check VM status
"""

import argparse
import os
import subprocess
import sys

# Configuration - UPDATE THESE
VM_IP = ""  # Will be prompted if empty
VM_USER = "ubuntu"
SSH_KEY = os.path.expanduser("~/.ssh/onetrack_key")
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def run(cmd, check=True, capture=False):
    """Run a shell command."""
    result = subprocess.run(
        cmd, shell=isinstance(cmd, str), capture_output=capture, text=True
    )
    if check and result.returncode != 0:
        print(f"Error: {cmd}")
        if capture:
            print(result.stderr)
        sys.exit(1)
    return result


def get_vm_ip():
    """Get VM IP from user or config."""
    global VM_IP
    if VM_IP:
        return VM_IP

    config_file = os.path.join(APP_DIR, ".vm_ip")
    if os.path.exists(config_file):
        with open(config_file) as f:
            VM_IP = f.read().strip()
            return VM_IP

    VM_IP = input("Enter Oracle Cloud VM IP address: ").strip()
    with open(config_file, "w") as f:
        f.write(VM_IP)
    return VM_IP


def ssh(cmd, check=True):
    """Run command on VM via SSH."""
    vm_ip = get_vm_ip()
    return run(
        f'ssh -i {SSH_KEY} -o StrictHostKeyChecking=no {VM_USER}@{vm_ip} "{cmd}"',
        check=check,
    )


def first_time_setup():
    """Run initial setup on the VM."""
    print("=== First Time Setup ===")
    vm_ip = get_vm_ip()

    print(f"1. Connecting to VM at {vm_ip}...")
    ssh("echo 'Connected!'")

    print("2. Uploading code...")
    run(f"scp -i {SSH_KEY} -r {APP_DIR}/* {VM_USER}@{vm_ip}:~/onetrack/")

    print("3. Running setup script...")
    ssh("bash ~/onetrack/deploy/setup_vm.sh")

    print("4. Setting up GitHub webhook...")
    print(f"   Go to: https://github.com/justsyxd076/onetrack/settings/hooks")
    print(f"   Payload URL: http://{vm_ip}:9000")
    print(f"   Content type: application/json")
    print(f"   Secret: onetrack-webhook-secret-change-me")
    print(f"   Events: Just the push event")

    print("\n=== Setup Complete ===")
    print(f"Your app is running at: http://{vm_ip}")


def deploy():
    """Deploy latest code to VM."""
    print("=== Deploying OneTrack ===")

    # 1. Push to GitHub first
    print("1. Pushing to GitHub...")
    run(f"cd {APP_DIR} && git add -A")
    run(f'cd {APP_DIR} && git commit -m "Deploy: $(date +%Y-%m-%d_%H-%M-%S)" || true')
    run(f"cd {APP_DIR} && git push origin master")

    # 2. SSH and pull on VM
    print("2. Pulling on VM...")
    ssh("cd ~/onetrack && git pull")

    # 3. Restart app
    print("3. Restarting app...")
    ssh("sudo systemctl restart onetrack")

    # 4. Verify
    print("4. Verifying...")
    result = ssh("curl -s -o /dev/null -w '%{http_code}' http://localhost", check=False)
    if result.stdout.strip() == "200":
        print("✅ Deploy successful! App is running.")
    else:
        print(f"⚠️  App returned status: {result.stdout.strip()}")
        print("   Check logs: sudo journalctl -u onetrack -f")


def check_status():
    """Check VM and app status."""
    print("=== OneTrack VM Status ===")
    vm_ip = get_vm_ip()

    print(f"\nVM IP: {vm_ip}")
    print(f"App URL: http://{vm_ip}")

    print("\n--- Service Status ---")
    ssh("sudo systemctl status onetrack --no-pager", check=False)

    print("\n--- Recent Logs ---")
    ssh("sudo journalctl -u onetrack --no-pager -n 20", check=False)

    print("\n--- Disk Usage ---")
    ssh("df -h /", check=False)

    print("\n--- Database Size ---")
    ssh("ls -lh ~/onetrack/onetrack.db", check=False)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="OneTrack Deployment Script")
    parser.add_argument("--setup", action="store_true", help="First-time setup")
    parser.add_argument("--status", action="store_true", help="Check VM status")
    args = parser.parse_args()

    if args.setup:
        first_time_setup()
    elif args.status:
        check_status()
    else:
        deploy()
