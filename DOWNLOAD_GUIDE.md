# How to Download OneTrack from GitHub

## Step 1: Open GitHub

1. Open any browser on the other PC
2. Go to: **https://github.com/justsyxd076/onetrack**
3. You'll see the OneTrack repository

---

## Step 2: Download the Code

### Method A: Download ZIP (Easiest)

1. Click the green **"< > Code"** button
2. Click **"Download ZIP"**
3. Save the file to your Desktop
4. Wait for download to complete

### Method B: Clone with Git (Advanced)

1. Open Command Prompt (cmd)
2. Type:
```
cd Desktop
git clone https://github.com/justsyxd076/onetrack.git
```
3. Wait for download to complete

---

## Step 3: Extract the Files (If Method A)

1. Find the downloaded file: `onetrack-main.zip`
2. Right-click on it
3. Click **"Extract All..."**
4. Click **"Extract"**
5. You'll get a folder: `onetrack-main`

---

## Step 4: Rename the Folder (Optional)

1. Right-click on `onetrack-main`
2. Click **"Rename"**
3. Type: `OneTrack`
4. Press Enter

---

## Step 5: Install Python

1. Go to: **https://www.python.org/downloads/**
2. Click **"Download Python 3.12.x"** (or latest version)
3. Run the installer
4. **IMPORTANT:** Check **"Add Python to PATH"** ☑️
5. Click **"Install Now"**
6. Wait for installation to complete
7. Click **"Close"**

---

## Step 6: Run Setup

1. Open the `OneTrack` folder
2. Find **`setup_onetrack.bat`**
3. Double-click it
4. Wait for installation to complete
5. Choose how to run:
   - **Option 1:** Normal mode
   - **Option 2:** Watchdog mode (auto-restart)
   - **Option 3:** Auto-start on boot (recommended)

---

## Step 7: Get PC's IP Address

The setup script shows your IP address. It looks like:
- `192.168.1.100`
- `172.20.10.13`
- `10.0.0.5`

**Write down this IP address.** You'll need it for phones.

---

## Step 8: Access from Phones

1. Make sure phone is on the same WiFi as PC
2. Open browser on phone
3. Type: `http://YOUR-PC-IP:5000`
   - Example: `http://192.168.1.100:5000`
4. Login with:
   - Username: `admin`
   - Password: `admin123`

---

## Step 9: Install as PWA on Phone

### Android (Samsung, etc.)
1. Open Chrome on phone
2. Go to: `http://YOUR-PC-IP:5000`
3. Tap menu (3 dots in top right)
4. Tap **"Add to Home screen"**
5. Tap **"Install"**
6. App icon appears on home screen

### iPhone
1. Open Safari on phone
2. Go to: `http://YOUR-PC-IP:5000`
3. Tap Share button (square with arrow at bottom)
4. Tap **"Add to Home Screen"**
5. Tap **"Add"**
6. App icon appears on home screen

---

## Quick Reference

| Step | Action |
|------|--------|
| 1 | Go to github.com/justsyxd076/onetrack |
| 2 | Download ZIP |
| 3 | Extract the ZIP file |
| 4 | Install Python (with PATH) |
| 5 | Run setup_onetrack.bat |
| 6 | Note the IP address |
| 7 | Access from phone: http://IP:5000 |
| 8 | Install as PWA |

---

## Troubleshooting

### "Can't download from GitHub"
- Make sure you have internet connection
- Try a different browser
- Check if GitHub is blocked

### "Python not found"
- Reinstall Python
- Check "Add Python to PATH"
- Restart Command Prompt

### "Can't access from phone"
- Make sure phone is on same WiFi
- Check PC's IP address
- Try disabling firewall temporarily

### "setup_onetrack.bat doesn't work"
- Right-click → "Run as administrator"
- Make sure Python is installed
- Check if antivirus is blocking
