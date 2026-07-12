# Interview Tracker Automation

Automates the recruitment interview tracker: panel distribution, status/date
logic, and all data-validation dropdowns.

## Files

| File | What it is |
|------|-----------|
| `interview_automation.py` | The tool. Template generator + import processor. |
| `make_sample.py` | Generates `sample_candidates.xlsx` for a demo. |
| `Interview_Template.xlsx` | Blank format your team fills in and imports. |
| `Interview_Tracker_Output.xlsx` | Example finished tracker (from the sample). |

## How to use — the app (recommended)

**Double-click `Interview Tracker Tool.command`** in this folder.

1. Type your HR interviewer names (they become the panels; they're remembered
   for next time).
2. Click **Import Excel…** and pick your candidate file.
3. Click **Run ▶** — it saves `<yourfile>_Tracker.xlsx` next to your input and
   offers to open it. Use **Get blank template instead** for an empty format.

> First launch: if macOS blocks the `.command` file, right-click it → **Open**
> once to approve it. After that, double-click works normally.

## How to use — command line (optional)

**1. Get the blank template (share this with the team to fill in):**
```
python3 interview_automation.py --template
```

**2. Process an imported file:**
```
python3 interview_automation.py your_candidates.xlsx
```
Writes `Interview_Tracker_Output.xlsx`.

Input columns are matched by name (case-insensitive). At minimum provide
**Candidate Name** and **Primary Skill** — everything else is created/assigned.

## What it does

- **Dropdowns (data validation):**
  - Primary Skill: SAP, Oracle, Salesforce, Agentforce
  - Interview Status: Scheduled skill interview, Scheduled final interview,
    Final selected, Skill selected, Skill status pending, Final status pending
  - Skill Panel / Final Panel: Panel 1–6
  - Update / Disposition: Reject, Skill mismatch, OHCM, PHC, Moved to final,
    Moved to skill 2, Cand renege, EI renege, Tech issue
- **Skill-round distribution:** every candidate gets a Skill Panel, balanced so
  each panel carries an equal load (30 → 5 each) and each skill is spread across
  all panels.
- **Final-round distribution:** candidates flagged for finals get a Final Panel,
  one per panel (6 → 1 each).
- **Date logic:** Interview Date is meaningful only for the two *Scheduled*
  statuses. The sheet has an AutoFilter on row 1 — filter the Interview Status
  column to those two values to see only scheduled interviews and their dates.
- Frozen header row, colored headers, auto-fill highlight on assigned panels.

## Configuration

Edit the lists at the top of `interview_automation.py` to change dropdown
values, add skills, or change the panel count (currently 6). No other changes
needed — everything downstream reads from those lists.

## OneDrive sharing (manual step)

Fully-automated OneDrive sharing needs the Microsoft Graph API (an app
registration + org admin consent), which your IT would have to set up. Until
then, sharing is a quick manual step:

1. Upload `Interview_Tracker_Output.xlsx` to your OneDrive / SharePoint folder.
2. **Share → add team members → set permission to *Can edit*.**
3. For live co-editing, have everyone open it in Excel for the web.

If you want, this can be upgraded later to push the file and set edit
permissions automatically via Graph — say the word and we'll scope it.
