# PDF Print Production App

This guide is for first-time deployment. It shows exactly what to run and explains why.

## 0) What we are deploying
- App entrypoint: `app.py`
- Web server in production: `gunicorn`
- Host target: Render (using Docker)

Why Docker: your app depends on PDF libraries (`pikepdf`, `PyMuPDF`). Docker gives a consistent runtime so deployment is much less fragile.

## 1) Local prerequisites
Install these tools first:
- Python 3.11+
- Docker Desktop
- Git
- A GitHub account
- A Render account

Check tools:
```bash
python3 --version
docker --version
git --version
```

## 2) Run locally (non-Docker)
This confirms the app works before deployment.

```bash
cd "/Users/upw_aszejko/Documents/Print Q:A+/PrintPreso-duplicate"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open: `http://localhost:5001`

What this does:
- creates an isolated Python environment (`.venv`)
- installs dependencies from `requirements.txt`
- runs Flask app directly

## 3) Test the Docker image locally
This mirrors production behavior.

```bash
cd "/Users/upw_aszejko/Documents/Print Q:A+/PrintPreso-duplicate"
docker build -t print-pdf-app .
docker run --rm -p 10000:10000 -e PORT=10000 print-pdf-app
```

Open: `http://localhost:10000`

What this does:
- `docker build`: creates a container image from `Dockerfile`
- `docker run`: starts it and maps container port 10000 to your machine

## 4) Put code in GitHub
If this folder is already linked to the correct GitHub repo, skip to step 5.

```bash
cd "/Users/upw_aszejko/Documents/Print Q:A+/PrintPreso-duplicate"
git status
git add .
git commit -m "Add deployment setup (Docker + requirements + docs)"
```

Create a new GitHub repo in browser, then connect and push:

```bash
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
```

## 5) Deploy on Render
1. Open Render dashboard
2. Click **New** -> **Web Service**
3. Connect your GitHub repo
4. Select this repo
5. Render should detect `Dockerfile` automatically
6. Choose a service name
7. Click **Create Web Service**

What Render does:
- pulls your repo
- runs Docker build
- starts your app with `gunicorn`
- gives you a public URL

## 6) Verify deployment
- Open the Render URL
- Upload a PDF
- Process it
- Download output
- Open Flipbook preview

If it fails, open Render logs and check:
- build logs (dependency install)
- runtime logs (startup errors)

## 7) Update app later
Any time you change code:

```bash
cd "/Users/upw_aszejko/Documents/Print Q:A+/PrintPreso-duplicate"
git add .
git commit -m "Describe your change"
git push
```

Render auto-deploys after push.

## 8) Notes about file storage
This app writes temporary files in `/tmp`.
- That is fine for short-lived upload/process/download workflows.
- On many hosts, `/tmp` is not permanent across restarts.

If you later need saved job history, add persistent storage (disk or object storage).

## Files added for deployment
- `requirements.txt`: Python dependencies
- `Dockerfile`: production container build
- `.dockerignore`: excludes local/dev files from image
