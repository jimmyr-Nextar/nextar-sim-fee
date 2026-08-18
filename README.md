# Nextar SIM Fee Report

Monthly BYOD SIM Fee report generator for Nextar Telecom Group.

## What it does
- Uploads 4 monthly files and generates a formatted Excel report
- Filters to Sub Dealer & Preferred doors only (excludes Direct)
- Applies configurable fee tiers and return credits
- Handles Change of Ownership (COO) exceptions with a full audit tab
- Outputs a 5-tab Excel report (+ COO tab when applicable)

## Files required each month
1. Current month BYOD MCS Report
2. Previous month Door Info
3. Current month Door Info
4. Current month Remorse Returns

## Deployment

### Step 1 — Create GitHub repo
1. Go to github.com and sign in
2. Click **New repository**
3. Name it `nextar-sim-fee`
4. Set to **Private**
5. Click **Create repository**
6. Upload all files from this folder (app.py, requirements.txt, README.md)

### Step 2 — Deploy on Streamlit
1. Go to share.streamlit.io and sign in
2. Click **New app**
3. Select your `nextar-sim-fee` GitHub repo
4. Main file path: `app.py`
5. Click **Deploy**

Your app will be live at:
`https://nextar-sim-fee.streamlit.app`

## Fee Schedule
Editable in the sidebar before each run. Defaults:
- < $4.00 → $0
- $4.01 – $7.00 → $5
- $7.01 – $11.00 → $10
- $11.01+ → $15
- JKME doors → $5 (flat, always)

## Return Credits
- Phone → $20
- BTS → $9
