#!/bin/bash
# OneTrack VM Setup Script
# Run this ONCE on a fresh Oracle Cloud Ubuntu VM
# Usage: bash setup_vm.sh

set -e

echo "=== OneTrack VM Setup ==="
echo ""

# 1. Update system
echo "[1/7] Updating system..."
sudo apt update && sudo apt upgrade -y

# 2. Install Python 3, pip, git, nginx
echo "[2/7] Installing Python, pip, git, nginx..."
sudo apt install -y python3 python3-pip python3-venv git nginx

# 3. Configure unattended-upgrades (auto security updates)
echo "[3/7] Configuring auto security updates..."
sudo apt install -y unattended-upgrades
sudo dpkg-reconfigure -plow unattended-upgrades 2>/dev/null || true

# 4. Open firewall ports
echo "[4/7] Configuring firewall..."
sudo ufw allow 22/tcp    # SSH
sudo ufw allow 80/tcp    # HTTP
sudo ufw allow 443/tcp   # HTTPS
sudo ufw allow 5000/tcp  # Flask (direct access during testing)
sudo ufw --force enable

# 5. Create app directory
echo "[5/7] Creating app directory..."
mkdir -p ~/onetrack
mkdir -p ~/onetrack/static/uploads

# 6. Clone repo (if not already cloned)
echo "[6/7] Cloning repository..."
cd ~/onetrack
if [ ! -d ".git" ]; then
    git clone https://github.com/justsyxd076/onetrack.git .
fi

# 7. Setup Python environment
echo "[7/7] Setting up Python environment..."
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install gunicorn

# Create .env file (user needs to fill this in)
if [ ! -f ".env" ]; then
    cat > .env << 'EOF'
# OneTrack Environment Variables
# Fill in your Gmail SMTP credentials below

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_EMAIL=your-email@gmail.com
SMTP_PASSWORD=your-app-password
EOF
    echo ""
    echo "⚠️  IMPORTANT: Edit ~/onetrack/.env with your Gmail credentials!"
    echo "   nano ~/onetrack/.env"
fi

# Create .secret_key if not exists
if [ ! -f ".secret_key" ]; then
    python3 -c "import secrets; open('.secret_key', 'wb').write(secrets.token_bytes(64))"
    echo "✅ Secret key generated"
fi

# Setup systemd service
echo "Setting up systemd service..."
sudo cp ~/onetrack/deploy/onetrack.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable onetrack
sudo systemctl start onetrack

# Setup keep-alive cron (prevents Oracle Cloud from reclaiming VM)
echo "Setting up keep-alive cron..."
(crontab -l 2>/dev/null; echo "0 */6 * * * curl -s http://localhost:5000 > /dev/null 2>&1") | crontab -

# Setup GitHub webhook listener (auto-deploy on git push)
echo "Setting up GitHub webhook listener..."
sudo cp ~/onetrack/deploy/webhook.service /etc/systemd/system/
sudo systemctl enable webhook
sudo systemctl start webhook

echo ""
echo "=== Setup Complete ==="
echo ""
echo "Your app is running at: http://$(curl -s ifconfig.me)"
echo ""
echo "Next steps:"
echo "1. Edit .env with your Gmail credentials: nano ~/onetrack/.env"
echo "2. Restart the app: sudo systemctl restart onetrack"
echo "3. Test: curl http://localhost:5000"
echo ""
echo "Useful commands:"
echo "  sudo systemctl status onetrack    # Check status"
echo "  sudo systemctl restart onetrack   # Restart app"
echo "  sudo journalctl -u onetrack -f    # View logs"
echo "  cd ~/onetrack && git pull         # Update code"
