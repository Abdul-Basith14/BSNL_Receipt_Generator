# Deployment Guide - Render.com

## Complete Step-by-Step Guide to Deploy Cash Receipts Generator on Render.com

---

## Prerequisites

- GitHub account (free)
- Render.com account (free tier available)
- Git installed on your computer
- Your application code ready

---

## Step 1: Prepare Your Application

### 1.1 Ensure All Files Are Ready

Your project should have these files:

```
BSNL/
├── app.py
├── requirements.txt
├── Procfile
├── templates/
│   ├── index.html
│   └── success.html
├── static/
│   ├── style.css
│   └── script.js
├── uploads/
│   └── .gitkeep
└── output/
    └── .gitkeep
```

### 1.2 Verify requirements.txt

Ensure your `requirements.txt` contains:
```
Flask==3.0.0
openpyxl==3.1.2
Werkzeug==3.0.1
gunicorn==21.2.0
```

### 1.3 Update app.py for Production

Add this at the bottom of `app.py`:

```python
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
```

---

## Step 2: Set Up Git Repository

### 2.1 Initialize Git (if not already done)

Open PowerShell in your project directory:

```powershell
cd C:\01.Myuse\BSNL
git init
```

### 2.2 Create .gitignore (if not exists)

Your `.gitignore` should contain:
```
*.pyc
__pycache__/
uploads/*
!uploads/.gitkeep
output/*
!output/.gitkeep
*.xlsx
*.xls
!Dec*.xlsx
.env
venv/
.vscode/
.DS_Store
```

### 2.3 Commit Your Code

```powershell
git add .
git commit -m "Initial commit - Cash Receipts Generator"
```

---

## Step 3: Create GitHub Repository

### 3.1 Create New Repository on GitHub

1. Go to https://github.com
2. Click the **"+"** icon (top right) → **"New repository"**
3. Repository name: `cash-receipts-generator`
4. Description: `Government Cash Receipts Generator - Official Document Management`
5. Select **Public** or **Private** (both work with Render)
6. **DO NOT** initialize with README (you already have files)
7. Click **"Create repository"**

### 3.2 Push Your Code to GitHub

Copy the commands from GitHub (after creating repo) or use:

```powershell
git remote add origin https://github.com/YOUR_USERNAME/cash-receipts-generator.git
git branch -M main
git push -u origin main
```

**Replace `YOUR_USERNAME`** with your actual GitHub username.

---

## Step 4: Sign Up for Render.com

### 4.1 Create Account

1. Go to https://render.com
2. Click **"Get Started"** or **"Sign Up"**
3. Choose **"Sign up with GitHub"** (recommended)
4. Authorize Render to access your GitHub account
5. Complete profile setup

---

## Step 5: Deploy on Render

### 5.1 Create New Web Service

1. On Render Dashboard, click **"New +"** button
2. Select **"Web Service"**

### 5.2 Connect Your Repository

1. Click **"Connect account"** next to GitHub (if not already connected)
2. Find your repository: `cash-receipts-generator`
3. Click **"Connect"**

### 5.3 Configure Web Service

Fill in the following details:

**Basic Settings:**
- **Name:** `cash-receipts-generator` (this will be part of your URL)
- **Region:** Choose closest to your location (e.g., Singapore for India)
- **Branch:** `main`
- **Root Directory:** Leave empty (or `.` if needed)
- **Runtime:** `Python 3`

**Build & Deploy Settings:**
- **Build Command:** 
  ```
  pip install -r requirements.txt
  ```

- **Start Command:**
  ```
  gunicorn app:app
  ```

**Instance Type:**
- Select **"Free"** (0.1 GB RAM, enough for this app)
- Note: Free tier may sleep after inactivity

### 5.4 Environment Variables (Optional)

Scroll down to **"Environment Variables"** section.

Click **"Add Environment Variable"** and add:

| Key | Value |
|-----|-------|
| `PYTHON_VERSION` | `3.11.0` |
| `SECRET_KEY` | `your-secret-key-here-change-this` |

**Note:** Change the SECRET_KEY to a random string for security.

### 5.5 Advanced Settings (Optional)

- **Auto-Deploy:** Keep it **ON** (deploys automatically when you push to GitHub)
- **Health Check Path:** Leave empty or set to `/`

### 5.6 Create Web Service

1. Click **"Create Web Service"** button at the bottom
2. Wait for deployment (5-10 minutes for first deployment)

---

## Step 6: Monitor Deployment

### 6.1 Watch Build Logs

- Render will show live logs
- You'll see:
  ```
  Installing dependencies...
  Collecting Flask==3.0.0
  ...
  Build successful
  Starting service...
  ```

### 6.2 Deployment Complete

When you see:
```
==> Your service is live 🎉
```

Your application is deployed!

---

## Step 7: Access Your Application

### 7.1 Get Your URL

Your app will be available at:
```
https://cash-receipts-generator.onrender.com
```

Or:
```
https://YOUR-SERVICE-NAME.onrender.com
```

### 7.2 Test the Application

1. Click the URL in Render dashboard
2. Upload a test Excel file
3. Generate receipts
4. Download and verify

---

## Step 8: Custom Domain (Optional)

### 8.1 Add Custom Domain

1. In your Render service dashboard, go to **"Settings"**
2. Scroll to **"Custom Domain"**
3. Click **"Add Custom Domain"**
4. Enter your domain: `receipts.yourdomain.com`
5. Follow DNS configuration instructions

### 8.2 Configure DNS

Add a CNAME record in your domain registrar:
```
Type: CNAME
Name: receipts
Value: cash-receipts-generator.onrender.com
```

### 8.3 SSL Certificate

- Render automatically provides free SSL/HTTPS
- Certificate is issued within minutes
- Your app will be accessible via `https://`

---

## Step 9: Ongoing Maintenance

### 9.1 Update Your Application

To deploy updates:

```powershell
# Make changes to your code
git add .
git commit -m "Description of changes"
git push origin main
```

Render will automatically redeploy (if Auto-Deploy is ON).

### 9.2 View Logs

1. Go to Render Dashboard
2. Click your service
3. Click **"Logs"** tab
4. View real-time application logs

### 9.3 Check Service Status

- **Events** tab: Shows deployment history
- **Metrics** tab: Shows CPU, Memory, Request count
- **Shell** tab: Access command line (paid plans only)

---

## Step 10: Troubleshooting

### 10.1 Common Issues

**Issue: "Build Failed"**
- Check `requirements.txt` syntax
- Ensure Python version compatibility
- View build logs for specific errors

**Issue: "Application Error"**
- Check Start Command is correct: `gunicorn app:app`
- Verify app.py has no syntax errors
- Check logs for detailed error messages

**Issue: "502 Bad Gateway"**
- Application failed to start
- Check if app binds to `0.0.0.0` and uses `PORT` env variable
- View logs for startup errors

**Issue: "Free Instance Sleeping"**
- Free tier sleeps after 15 minutes of inactivity
- First request after sleep takes 30-60 seconds to wake up
- Upgrade to paid plan for always-on service

### 10.2 View Detailed Logs

```powershell
# Or view in Render dashboard under "Logs" tab
```

### 10.3 Restart Service

1. Go to service dashboard
2. Click **"Manual Deploy"** → **"Clear build cache & deploy"**
3. Or click **"Suspend"** then **"Resume"**

---

## Step 11: Security Best Practices

### 11.1 Update Secret Key

In Render Environment Variables:
```
SECRET_KEY=your-very-long-random-secret-key-here
```

Generate secure key:
```python
import secrets
print(secrets.token_hex(32))
```

### 11.2 Restrict File Upload Size

Already configured in app.py:
```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
```

### 11.3 Monitor Access

- Check logs regularly for suspicious activity
- Use Render's built-in DDoS protection
- Consider adding rate limiting for production

---

## Step 12: Upgrade Options (Optional)

### Free Tier Limitations:
- ✅ 750 hours/month
- ✅ Automatic SSL
- ✅ 0.1 GB RAM
- ❌ Sleeps after 15 min inactivity
- ❌ Limited to 100GB bandwidth/month

### Paid Tier Benefits ($7/month):
- ✅ Always-on (no sleeping)
- ✅ More RAM (512 MB - 16 GB)
- ✅ Priority support
- ✅ Background workers
- ✅ SSH access

---

## Quick Reference Commands

### Local Development:
```powershell
# Run locally
python app.py

# Install dependencies
pip install -r requirements.txt
```

### Git Commands:
```powershell
# Push updates
git add .
git commit -m "Update message"
git push origin main

# Check status
git status

# View history
git log --oneline
```

### Render Dashboard Links:
- Dashboard: https://dashboard.render.com
- Docs: https://render.com/docs
- Status: https://status.render.com

---

## Support & Resources

- **Render Documentation:** https://render.com/docs
- **Render Community:** https://community.render.com
- **Flask Documentation:** https://flask.palletsprojects.com
- **Support Email:** support@render.com

---

## Summary Checklist

- [ ] Code pushed to GitHub
- [ ] requirements.txt includes all dependencies
- [ ] Procfile configured (optional with gunicorn)
- [ ] Render account created
- [ ] Web service created and configured
- [ ] Build command set: `pip install -r requirements.txt`
- [ ] Start command set: `gunicorn app:app`
- [ ] Environment variables configured
- [ ] Deployment successful
- [ ] Application tested and working
- [ ] Custom domain configured (optional)
- [ ] SSL certificate active

---

**🎉 Congratulations! Your application is now live on Render.com**

Access your app at: `https://cash-receipts-generator.onrender.com`

---

*Last Updated: February 19, 2026*
*Version: 1.0*
