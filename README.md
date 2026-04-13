# Resource Report Generator — Deployment Guide

## Folder structure

```
mis_app/
├── app.py              ← Streamlit app (main file)
├── requirements.txt    ← Python dependencies
├── Employee.xlsx       ← Permanent mapping file (Role + Location)
├── Project.xlsx        ← Permanent mapping file (Client)
└── README.md
```

---

## Step 1 — Create a free GitHub account
Go to https://github.com and sign up (if you don't have one).

---

## Step 2 — Create a new GitHub repository

1. Click the **+** icon (top right) → **New repository**
2. Name it: `mis-report-generator`
3. Set it to **Private** (so only you control it)
4. Click **Create repository**

---

## Step 3 — Upload your files to GitHub

In your new repo, click **Add file → Upload files** and upload:
- `app.py`
- `requirements.txt`
- `Employee.xlsx`
- `Project.xlsx`

Click **Commit changes**.

---

## Step 4 — Deploy on Streamlit Cloud (free)

1. Go to https://streamlit.io/cloud and sign in with your GitHub account
2. Click **New app**
3. Select your repository: `mis-report-generator`
4. Main file path: `app.py`
5. Click **Deploy**

Wait ~2 minutes. You'll get a public URL like:
```
https://your-name-mis-report-generator.streamlit.app
```

Share this link with your team. That's it! ✅

---

## How your team uses it

1. Open the link in any browser
2. Upload the weekly MIS Report (.xlsx)
3. Click **Generate Consolidated Report**
4. Click **Download Report**

---

## Updating Employee.xlsx or Project.xlsx

Whenever you add a new employee or project:

1. Go to your GitHub repo
2. Click on `Employee.xlsx` or `Project.xlsx`
3. Click the pencil ✏️ icon → **Upload** the new file
4. Commit the change

Streamlit will pick up the new file automatically within a minute.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Employee.xlsx not found" | Make sure the file is uploaded to GitHub repo root |
| App is slow to load first time | Streamlit free tier "sleeps" — just wait 30 seconds |
| Wrong client shown for a project | Update Project.xlsx on GitHub |
| New employee has no role/location | Update Employee.xlsx on GitHub |
