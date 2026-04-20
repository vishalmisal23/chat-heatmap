# Chat Heatmap Generator

Paste your raw Yellow.ai chat export CSV → get a fully formatted Excel heatmap in seconds.

No manual date fixing. No slot calculation. Everything is automatic.

---

## What it does

- Parses `MM/DD/YYYY hh:mm:ss AM/PM` timestamps automatically (no manual reformatting)
- Derives day of week and 30-min slot from each ticket's `initialized_time`
- Builds a normalised pivot (avg chats per slot per day across your date range)
- Calculates **Agents Needed** using your AHT and concurrency
- Outputs a formatted `.xlsx` with:
  - **⚙️ Config** — edit AHT, concurrency, roster size, shrinkage % here
  - **Half-Hour Heatmap** — colour-coded volume grid, agents needed, gap, status
  - **Shrinkage Planner** — effective headcount after unplanned + planned shrinkage

All calculated columns are **live Excel formulas** — change AHT or concurrency in the Config sheet and every number updates instantly.

---

## Quickstart

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run

```bash
python generate_heatmap.py your_chat_export.csv
```

Output: `chat_heatmap_output.xlsx` in the same folder.

### 3. Open in Excel or upload to Google Sheets

---

## Configuration

Edit the top of `generate_heatmap.py`:

```python
AHT_MINUTES  = 16       # Average Handle Time (minutes)
CONCURRENCY  = 3.0      # Simultaneous chats per agent
SLOT_MINUTES = 30       # Slot size (15 or 30)
ROSTERED     = 39       # Total agents on roster
UNPLN_PCT    = 0.10     # Unplanned shrinkage rate (10%)

SHIFTS = [
    ("S1  7AM–4PM",   7,  16, 13),
    ("S2  3PM–12AM", 15,  24, 10),
    ("S3  6PM–3AM",  18,  27, 10),
    ("S4 11PM–8AM",  23,  32,  6),
]
```

Or just change the yellow cells in the **⚙️ Config** sheet after generating — no need to re-run the script.

---

## Input format

The script expects a CSV export from Yellow.ai with at least an `initialized_time` column in `MM/DD/YYYY hh:mm:ss AM/PM` format.

It also tries `assigned_time`, `opened_time`, `Created time` as fallbacks.

---

## Timestamp formats supported

| Format | Example |
|--------|---------|
| `MM/DD/YYYY hh:mm:ss AM/PM` | `04/06/2026 12:00:03 AM` ✓ default |
| `MM/DD/YYYY HH:MM:SS` | `04/06/2026 00:00:03` ✓ |
| `YYYY-MM-DD HH:MM:SS` | `2026-04-06 00:00:03` ✓ |
| `DD/MM/YYYY HH:MM:SS` | `06/04/2026 00:00:03` ✓ |

---

## GitHub Actions (optional auto-run)

If you commit a new CSV to this repo, you can auto-generate the heatmap using the included workflow:

```
.github/workflows/generate.yml
```

It runs on every push, generates the Excel, and uploads it as a build artifact.
