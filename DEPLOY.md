# OneTrack Deployment Guide (Waifly - Free Forever)

## Prerequisites
- Waifly account (free, no credit card)
- GitHub account (already have)
- Gmail app password (already configured)

## Step 1: Create Waifly Account

1. Go to https://dash.waifly.com
2. Sign up with your email
3. **No credit card required** ✅

## Step 2: Create Server

1. From Servers tab, click "Create"
2. Choose "Python Egg"
3. Name: `onetrack`
4. Location: Paris (or nearest)
5. **Free plan applies automatically**: 300 MB RAM, 1 GB disk, 30% CPU

## Step 3: Upload Code

### Option A: Git (Recommended)
1. SSH into your Waifly server (credentials in panel)
2. Run:
```bash
git clone https://github.com/justsyxd076/onetrack.git .
```

### Option B: File Manager
1. Use the panel file manager
2. Upload all files from your local OneTrack folder

## Step 4: Create MySQL Database

1. In panel, go to Databases tab
2. Click "Create Database"
3. Note the credentials:
   - Host: (shown in panel)
   - Database: (shown in panel)
   - User: (shown in panel)
   - Password: (shown in panel)

## Step 5: Set Environment Variables

In the panel's Startup tab, add these environment variables:

```
DB_HOST=your-mysql-host
DB_USER=your-mysql-user
DB_PASSWORD=your-mysql-password
DB_NAME=your-mysql-database
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_EMAIL=onetrackpn@gmail.com
SMTP_PASSWORD=your-gmail-app-password
```

## Step 6: Set Startup Command

In the panel's Startup tab, set:
```
python start.py
```

## Step 7: Start and Test

1. Click "Start" in the panel
2. Check console for errors
3. Your app is live at: `http://your-server-name.waifly.com`

## Step 8: Keep Server Alive

Waifly suspends servers offline for 3+ days. To prevent this:
- The app runs continuously when started
- If it crashes, enable "Auto-restart" in the panel

---

## Daily Usage

### Sales Team
- Open `http://your-server-name.waifly.com` on phone
- Login with credentials
- Submit forms through PWA
- Data syncs automatically

### You (Admin)
- Edit code on any PC
- Push to GitHub
- Pull on Waifly server
- **Live in seconds**

---

## Useful Commands

```bash
# Check app status
curl http://localhost:8000

# View logs
# Use the panel's console

# Restart app
# Use the panel's restart button

# Update code
cd ~ && git pull
# Then restart via panel
```

---

## Architecture

```
Any PC → GitHub → Waifly Server (MySQL)
                       │
                       ▼
                 Flask app (port from env)
                       │
                       ▼
                 MySQL database (free)
```

---

## Cost

| Service | Cost |
|---------|------|
| Waifly Free | $0 forever |
| GitHub Free | $0 forever |
| Gmail SMTP | $0 forever |
| **Total** | **$0 forever** |

---

## Troubleshooting

### App won't start
- Check console for errors
- Verify environment variables are set
- Check MySQL credentials

### Can't connect to database
- Verify DB_HOST, DB_USER, DB_PASSWORD, DB_NAME
- Check MySQL is running in panel

### Gmail not sending
- Verify SMTP_EMAIL and SMTP_PASSWORD
- Check Gmail app password is correct

### Server suspended
- Reactivate at unsuspend.waifly.com
- Enable auto-restart in panel
