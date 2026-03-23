#!/usr/bin/env python3
"""Generate an interactive HTML report for The Snack Choice territory analysis."""

import json
import os

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

# Territory coordinates (centroid approximations from postcodes)
COORDS = {
    "Park Royal": [51.527, -0.263],
    "Manor Royal": [51.130, -0.183],
    "Slough Trading Estate": [51.510, -0.605],
    "Brimsdown Industrial Estate": [51.660, -0.032],
    "DP World London Gateway": [51.502, 0.478],
    "Crossways Business Park": [51.440, 0.235],
    "Prologis Park West London": [51.495, -0.480],
    "Watford Business Park": [51.665, -0.400],
    "London Medway Commercial Park": [51.380, 0.480],
    "SEGRO Greenford Central": [51.530, -0.330],
    "Thames Road / Erith-Belvedere Corridor": [51.480, 0.160],
    "Beddington Trading Estate": [51.370, -0.140],
    "Stockley Park": [51.505, -0.460],
    "Croxley Park": [51.660, -0.430],
    "Willow Lane Trading Estate": [51.400, -0.170],
    "Beddington Lane Corridor": [51.373, -0.140],
    "Crayfields Park": [51.385, 0.105],
    "Kimpton Industrial Estate": [51.365, -0.190],
    "Charlton Riverside": [51.488, 0.040],
    "Purley Way Corridor": [51.370, -0.120],
}

TIER_COLOURS = {1: "#22c55e", 2: "#f59e0b", 3: "#9ca3af"}
TIER_LABELS = {1: "Tier 1 — Prime Target", 2: "Tier 2 — Secondary", 3: "Tier 3 — Monitor"}

DESERT_BADGE = {
    "FOOD DESERT": '<span style="background:#dc2626;color:white;padding:2px 8px;border-radius:4px;font-size:12px">DESERT</span>',
    "PARTIAL": '<span style="background:#f59e0b;color:white;padding:2px 8px;border-radius:4px;font-size:12px">PARTIAL</span>',
    "NOT": '<span style="background:#22c55e;color:white;padding:2px 8px;border-radius:4px;font-size:12px">SERVED</span>',
    "OK": '<span style="background:#6b7280;color:white;padding:2px 8px;border-radius:4px;font-size:12px">NOT SCOUTED</span>',
}


def get_desert_status(food_access):
    if "FOOD DESERT" in food_access:
        return "FOOD DESERT"
    if "PARTIAL" in food_access:
        return "PARTIAL"
    if "NOT" in food_access and "desert" in food_access.lower():
        return "NOT"
    return "OK"


def get_badge(food_access):
    status = get_desert_status(food_access)
    return DESERT_BADGE.get(status, DESERT_BADGE["OK"])


def main():
    territories = json.load(open(os.path.join(DATA_DIR, "territories.json")))
    territories.sort(key=lambda t: (-t["vending_score"], t["tier"]))

    # Build markers JS
    markers_js = []
    for t in territories:
        coord = COORDS.get(t["name"])
        if not coord:
            continue
        colour = TIER_COLOURS.get(t["tier"], "#9ca3af")
        popup = (
            f"<b>{t['name']}</b><br>"
            f"Score: {t['vending_score']}/10 | {TIER_LABELS.get(t['tier'], '')}<br>"
            f"Workers: {t.get('workers', '?')}<br>"
            f"Food: {t.get('food_access', '?')[:80]}"
        ).replace('"', '\\"').replace("'", "\\'").replace("\n", "")
        markers_js.append(
            f'  L.circleMarker([{coord[0]}, {coord[1]}], '
            f'{{radius: {8 + t["vending_score"]}, fillColor: "{colour}", '
            f'color: "#333", weight: 1, fillOpacity: 0.85}})'
            f'.addTo(map).bindPopup("{popup}");'
        )

    # Build table rows
    table_rows = []
    for t in territories:
        badge = get_badge(t.get("food_access", ""))
        table_rows.append(
            f"<tr>"
            f"<td><b>{t['name']}</b><br><small>{t.get('area', '')}</small></td>"
            f"<td style='text-align:center'><b>{t['vending_score']}</b>/10</td>"
            f"<td style='text-align:center'>{t['tier']}</td>"
            f"<td>{t.get('workers', '?')}</td>"
            f"<td>{badge}</td>"
            f"<td><small>{t.get('notes', '')[:100]}</small></td>"
            f"</tr>"
        )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Snack Choice — Territory Analysis</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; color: #1a1a1a; background: #f8f9fa; }}
  .header {{ background: #1e293b; color: white; padding: 32px 24px; }}
  .header h1 {{ font-size: 28px; margin-bottom: 8px; }}
  .header p {{ color: #94a3b8; font-size: 16px; }}
  .container {{ max-width: 1200px; margin: 0 auto; padding: 24px; }}
  #map {{ height: 500px; border-radius: 12px; margin-bottom: 32px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
  .legend {{ background: white; padding: 12px 16px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
  .legend-item {{ display: flex; align-items: center; margin: 4px 0; font-size: 14px; }}
  .legend-dot {{ width: 14px; height: 14px; border-radius: 50%; margin-right: 8px; border: 1px solid #333; }}
  .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 32px; }}
  .stat-card {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
  .stat-card .number {{ font-size: 32px; font-weight: 700; color: #1e293b; }}
  .stat-card .label {{ font-size: 14px; color: #64748b; margin-top: 4px; }}
  .section {{ margin-bottom: 32px; }}
  .section h2 {{ font-size: 22px; margin-bottom: 16px; color: #1e293b; }}
  table {{ width: 100%; border-collapse: collapse; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
  th {{ background: #1e293b; color: white; padding: 12px 16px; text-align: left; font-size: 13px; text-transform: uppercase; }}
  td {{ padding: 12px 16px; border-bottom: 1px solid #e2e8f0; font-size: 14px; }}
  tr:last-child td {{ border-bottom: none; }}
  tr:hover {{ background: #f1f5f9; }}
  .callout {{ background: #fef3c7; border-left: 4px solid #f59e0b; padding: 16px 20px; border-radius: 0 8px 8px 0; margin-bottom: 16px; font-size: 15px; }}
  .callout.red {{ background: #fee2e2; border-color: #dc2626; }}
  .callout.green {{ background: #dcfce7; border-color: #22c55e; }}
  .strategy {{ background: white; padding: 24px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
  .strategy li {{ margin: 8px 0; line-height: 1.6; }}
  .footer {{ text-align: center; padding: 32px; color: #94a3b8; font-size: 13px; }}
</style>
</head>
<body>

<div class="header">
  <h1>The Snack Choice — London Expansion</h1>
  <p>Territory Analysis &amp; Opportunity Map &bull; March 2026</p>
</div>

<div class="container">

  <div class="stats">
    <div class="stat-card">
      <div class="number">{len(territories)}</div>
      <div class="label">Territories Evaluated</div>
    </div>
    <div class="stat-card">
      <div class="number">{sum(1 for t in territories if t['tier'] == 1)}</div>
      <div class="label">Tier 1 Targets</div>
    </div>
    <div class="stat-card">
      <div class="number">{sum(1 for t in territories if 'DESERT' in t.get('food_access', ''))}</div>
      <div class="label">Confirmed Food Deserts</div>
    </div>
    <div class="stat-card">
      <div class="number">£17k</div>
      <div class="label">Annual Profit (5-machine cluster)</div>
    </div>
  </div>

  <div class="section">
    <h2>Opportunity Map</h2>
    <div id="map"></div>
  </div>

  <div class="section">
    <div class="callout red">
      <b>Key Insight:</b> Inner London has a corner shop every 200m — vending's "captive audience" value proposition fails.
      Industrial estates and leisure venues are where people are genuinely captive — either spatially (no walkable food) or temporally (stuck for 1-4 hours).
    </div>
    <div class="callout green">
      <b>The Play:</b> Croydon corridor — 16 pre-qualified leads within 30 min of base. Mix industrial sites (Beddington Lane),
      leisure venues (Oxygen, Kidspace), gyms, and car dealerships along one restock route. Free placement, 5-7 sites in phase 1.
    </div>
  </div>

  <div class="section">
    <h2>Territory Ranking</h2>
    <table>
      <thead>
        <tr><th>Estate</th><th>Score</th><th>Tier</th><th>Workers</th><th>Food Access</th><th>Notes</th></tr>
      </thead>
      <tbody>
        {"".join(table_rows)}
      </tbody>
    </table>
  </div>

  <div class="section">
    <h2>Strategy Summary</h2>
    <div class="strategy">
      <ul>
        <li><b>Croydon corridor first.</b> Romell is Croydon-based — prioritise by operator proximity, not market size. Beddington Lane (2mi) beats Park Royal (12mi).</li>
        <li><b>Two types of captivity.</b> Spatial (industrial estates — no walkable food) and temporal (gyms, leisure, dealerships — people stuck for hours). Mix both on one route.</li>
        <li><b>Cluster along one restock route.</b> 16 leads within 30 min of Croydon. 5 clustered machines = £44/hr effective rate.</li>
        <li><b>Free placement over site rent.</b> Breakeven drops from 16 vends/day to 4 vends/day. Always pitch free placement first.</li>
        <li><b>Hybrid format.</b> Vending for public/untrusted sites. Micro-markets for closed office/warehouse staff rooms. Match format to trust level.</li>
        <li><b>Product mix: stock Monster not Red Bull.</b> Red Bull = 30% margin. Monster = 59%. Same category, double the profit.</li>
      </ul>
    </div>
  </div>

</div>

<div class="footer">
  The Snack Choice &bull; Ruben &amp; Romell &bull; Data: Google Places API, Companies House, supplier websites &bull; Generated March 2026
</div>

<script>
var map = L.map('map').setView([51.45, -0.05], 10);
L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
  maxZoom: 18,
  attribution: '&copy; OpenStreetMap contributors'
}}).addTo(map);

{chr(10).join(markers_js)}

var legend = L.control({{position: 'bottomright'}});
legend.onAdd = function(map) {{
  var div = L.DomUtil.create('div', 'legend');
  div.innerHTML = '<b>Territory Tiers</b><br>' +
    '<div class="legend-item"><div class="legend-dot" style="background:#22c55e"></div>Tier 1 — Prime Target</div>' +
    '<div class="legend-item"><div class="legend-dot" style="background:#f59e0b"></div>Tier 2 — Secondary</div>' +
    '<div class="legend-item"><div class="legend-dot" style="background:#9ca3af"></div>Tier 3 — Monitor</div>' +
    '<br><small>Marker size = vending score</small>';
  return div;
}};
legend.addTo(map);
</script>

</body>
</html>"""

    output = os.path.join(os.path.dirname(__file__), "TheSnackChoice_Report.html")
    with open(output, "w") as f:
        f.write(html)
    print(f"Generated: {output}")


if __name__ == "__main__":
    main()
