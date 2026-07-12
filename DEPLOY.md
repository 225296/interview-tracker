# Deploy the Interview Tracker to a public URL (free)

This puts the tool online at a permanent link like
`https://interview-tracker.streamlit.app` that your whole team can open from
any browser, anywhere. Free, always-on. ~15 minutes, one-time.

You'll do it once; after that, the team just uses the URL.

## What gets deployed

Only three files (no candidate data — that's blocked by `.gitignore`):

- `app.py`
- `interview_automation.py`
- `requirements.txt`

## Step 1 — Create a free GitHub account (skip if you have one)

1. Go to <https://github.com> → **Sign up**.

## Step 2 — Put the 3 files in a GitHub repository

Easiest, no-terminal way:

1. On GitHub, click **+** (top right) → **New repository**.
2. Name it e.g. `interview-tracker`, keep it **Public**, click **Create repository**.
3. On the new repo page, click **uploading an existing file**.
4. Drag in **`app.py`**, **`interview_automation.py`**, and
   **`requirements.txt`** from your `Desktop/Excel` folder.
5. Click **Commit changes**.

## Step 3 — Deploy on Streamlit Community Cloud

1. Go to <https://share.streamlit.io> → **Sign in with GitHub** (approve access).
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository:** `your-username/interview-tracker`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**. First build takes ~2 minutes.
5. You'll get a URL like `https://interview-tracker-xxxx.streamlit.app`.

## Step 4 — Share it

Send that URL to your team. Anyone with the link can:
- enter/adjust the HR names,
- pick the interview date,
- upload the candidate Excel,
- download the finished tracker.

## Updating the tool later

When I change the code, re-upload the changed file to the GitHub repo
(Step 2, method: **Add file → Upload files**). Streamlit auto-redeploys within
a minute. No other action needed.

## Notes

- **Data:** files are processed on Streamlit's cloud servers (in memory, not
  stored). If your org requires candidate data to stay in-house, use the
  "Company IT / internal server" option instead — ask me and I'll package it.
- **Sleeping:** a free app goes to sleep after long inactivity and wakes on the
  next visit (a few seconds). That's normal on the free tier.
- **Alternative host:** Hugging Face Spaces works too (create a Space →
  "Streamlit" SDK → upload the same 3 files). Ask me if you'd prefer that.
