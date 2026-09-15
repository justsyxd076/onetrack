# OneTrack Setup Guide (Local PC)

## Quick Setup (5 minutes)

### Step 1: Copy Files
Copy the entire `OneTrack` folder to the other PC.

### Step 2: Install Python
1. Go to: https://www.python.org/downloads/
2. Download Python 3.11 or newer
3. **IMPORTANT:** Check "Add Python to PATH" during installation
4. Click "Install Now"

### Step 3: Run Setup
1. Open the OneTrack folder
2. Double-click `setup_onetrack.bat`
3. Wait for installation to complete
4. App will start automatically

### Step 4: Access from Phones
1. Make sure all phones are on the same WiFi as the PC
2. Open browser on phone
3. Go to: `http://PC-IP-ADDRESS:5000`
4. Login with credentials

---

## Finding PC's IP Address

### Method 1: Setup Script
The setup script shows your IP address automatically.

### Method 2: Command Prompt
1. Open Command Prompt (cmd)
2. Type: `ipconfig`
3. Look for "IPv4 Address" — something like `192.168.1.100`

### Method 3: Settings
1. Open Settings
2. Go to Network & Internet
3. Click WiFi or Ethernet
4. Look for "IPv4 address"

---

## Default Login Credentials

| Username | Password | Role |
|----------|----------|------|
| admin | admin123 | Admin |
| user | user123 | User |

**Change these after first login!**

---

## Keeping App Running

The app must run 24/7 for sales team to access it.

### Option 1: Leave PC On
- Don't shut down the PC
- App runs as long as PC is on

### Option 2: Run as Service (Advanced)
- Use NSSM to run as Windows service
- Auto-starts when PC boots

---

## Troubleshooting

### "Python is not installed"
- Install Python from python.org
- Check "Add Python to PATH"

### "Port 5000 already in use"
- Another app is using port 5000
- Close that app, or change port in app.py

### "Can't access from phone"
- Make sure phone is on same WiFi
- Check PC's firewall (allow port 5000)
- Use PC's IP address, not localhost

### "App stopped working"
- PC was shut down or restarted
- Run `python app.py` again

---

## Firewall Setup (If Needed)

If phones can't connect:

1. Open Windows Defender Firewall
2. Click "Advanced settings"
3. Click "Inbound Rules"
4. Click "New Rule"
5. Select "Port"
6. Enter: `5000`
7. Click "Allow the connection"
8. Finish

---

## Summary

| Item | Value |
|------|-------|
| **URL** | `http://PC-IP:5000` |
| **Admin login** | admin / admin123 |
| **User login** | user / user123 |
| **Cost** | $0 forever |
| **Internet needed** | No (local network) |
| **PC must stay on** | Yes |
