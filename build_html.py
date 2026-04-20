"""HTML dashboard builder — called by generate_heatmap.py"""
import json

def build_html(pivot, all_slots, date_range, total_tickets,
               AHT_MINUTES, CONCURRENCY, SLOT_MINUTES, ROSTERED, SHIFTS,
               output_file="index.html"):

    DAYS_ORDER = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    DAYS_SHORT = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

    matrix = []
    for slot in all_slots:
        row = {
            "slot": slot,
            "days": [int(pivot.loc[slot, d]) if slot in pivot.index else 0 for d in DAYS_ORDER],
            "avg":  int(pivot.loc[slot, "Daily Avg"]) if slot in pivot.index else 0,
        }
        matrix.append(row)

    peak_slot = str(pivot["Daily Avg"].idxmax()) if len(pivot) else "--"
    peak_val  = int(pivot["Daily Avg"].max()) if len(pivot) else 0
    low_val   = int(pivot["Daily Avg"].min()) if len(pivot) else 0

    hourly_avg = {}
    for h in range(24):
        slots_in_hour = [s for s in all_slots if int(s.split(":")[0]) == h]
        vals = [int(pivot.loc[s,"Daily Avg"]) for s in slots_in_hour if s in pivot.index]
        hourly_avg[str(h)] = round(sum(vals), 1) if vals else 0

    day_totals = {}
    for i, d in enumerate(DAYS_ORDER):
        day_totals[DAYS_SHORT[i]] = int(pivot[d].sum()) if d in pivot.columns else 0

    shift_colors = ["#f59e0b","#00e5a0","#f97316","#0ea5e9"]
    shift_cards = ""
    for i, s in enumerate(SHIFTS):
        color = shift_colors[i % len(shift_colors)]
        pct   = round(s[3] / ROSTERED * 100)
        shift_cards += f"""
    <div class="shift-card shift-s{i+1}">
      <div class="shift-name">{s[0].strip()}</div>
      <div class="shift-time">{s[1]:02d}:00 &ndash; {s[2] % 24:02d}:00</div>
      <div class="shift-agents">{s[3]} <span class="shift-agents-label">agents</span></div>
      <div class="shift-bar"><div class="shift-bar-fill" style="width:{pct}%;background:{color}"></div></div>
    </div>"""

    shift_summary = " &middot; ".join(f"{s[0].strip()}({s[3]})" for s in SHIFTS)

    data_block = f"""
const MATRIX = {json.dumps(matrix)};
const HOURLY = {json.dumps(hourly_avg)};
const DAY_TOTALS = {json.dumps(day_totals)};
const DAYS_SHORT = {json.dumps(DAYS_SHORT)};
let AHT = {AHT_MINUTES}, CONC = {CONCURRENCY};
"""

    # Read the JS/CSS template
    import os
    template_path = os.path.join(os.path.dirname(__file__), "dashboard_template.html")
    with open(template_path, encoding="utf-8") as f:
        template = f.read()

    html = (template
        .replace("{{DATE_RANGE}}", date_range)
        .replace("{{TOTAL_TICKETS}}", f"{total_tickets:,}")
        .replace("{{PEAK_SLOT}}", peak_slot)
        .replace("{{PEAK_VAL}}", str(peak_val))
        .replace("{{LOW_VAL}}", str(low_val))
        .replace("{{ROSTERED}}", str(ROSTERED))
        .replace("{{SHIFT_CARDS}}", shift_cards)
        .replace("{{SHIFT_SUMMARY}}", shift_summary)
        .replace("{{AHT}}", str(AHT_MINUTES))
        .replace("{{CONC}}", f"{CONCURRENCY:.1f}")
        .replace("{{DATA_BLOCK}}", data_block)
    )

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  HTML dashboard saved to: {output_file}")
