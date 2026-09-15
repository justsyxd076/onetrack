# OneTrack Deployment Guide

## Prerequisites
- Oracle Cloud Always Free account
- SSH key at `~/.ssh/onetrack_key`
- Gmail app password for SMTP

## Step 1: Create Oracle Cloud VM

1. Go to https://cloud.oracle.com/free
2. Sign up / Log in
3. Create Instance:
   - Name: `onetrack`
   - Image: Ubuntu 24.04 (ARM)
   - Shape: `VM.Standard.A1.Flex` → 2 OCPUs, 12GB RAM
   - Upload SSH key: `~/.ssh/onetrack_key.pub`
4. Open ports in Security List:
   - 22 (SSH)
   - 80 (HTTP)
   - 443 (HTTPS)
   - 5000 (Flask direct)
   - 9000 (Webhook)
5. Note the public IP address

## Step 2: First-Time Setup

```bash
# From your PC
python deploy/deploy.py --setup
```

This will:
- Connect to your VM via SSH
- Upload all code
- Install Python, pip, nginx
- Configure firewall
- Setup systemd (auto-restart)
- Setup unattended-upgrades (auto security updates)
- Setup keep-alive cron (prevents VM reclaim)

## Step 3: Configure Gmail SMTP

```bash
# SSH into VM
ssh -i ~/.ssh/onetrack_key ubuntu@YOUR_VM_IP

# Edit .env
nano ~/onetrack/.env

# Fill in:
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_EMAIL=onetrackpn@gmail.com
SMTP_PASSWORD=your-app-password

# Restart app
sudo systemctl restart onetrack
```

## Step 4: Setup GitHub Webhook (Auto-Deploy)

1. Go to: https://github.com/justsyxd076/onetrack/settings/hooks
2. Add webhook:
   - Payload URL: `http://YOUR_VM_IP:9000`
   - Content type: `application/json`
   - Secret: `onetrack-webhook-secret-change-me`
   - Events: Just the push event

## Step 5: Enable Auto-Sync (Any PC)

```bash
# Run on any PC you develop on
python auto_sync.py
```

This watches for file changes and auto-pushes to GitHub.
GitHub webhook auto-deploys to VM.

## Daily Usage

### Sales Team
- Open `http://YOUR_VM_IP` on phone
- Login with credentials
- Submit forms through PWA
- Data syncs automatically

### You (Admin)
- Edit code on any PC
- `auto_sync.py` pushes to GitHub
- GitHub webhook deploys to VM
- **Live in 3 seconds**

## Useful Commands

```bash
# Check app status
python deploy/deploy.py --status

# Deploy manually
python deploy/deploy.py

# SSH into VM
ssh -i ~/.ssh/onetrack_key ubuntu@YOUR_VM_IP

# View logs on VM
sudo journalctl -u onetrack -f

# Restart app on VM
sudo systemctl restart onetrack

# Update code on VM
cd ~/onetrack && git pull
```

## Architecture

```
Any PC → auto_sync.py → GitHub → Webhook → VM (auto-deploy)
                                                  │
                                                  ▼
                                            Flask app (port 80)
                                                  │
                                                  ▼
                                            onetrack.db (SQLite)
```

## Cost

| Service | Cost |
|---------|------|
| Oracle Cloud Always Free | $0 forever |
| Cloudflare Free | $0 forever |
| GitHub Free | $0 forever |
| **Total** | **$0 forever** |
