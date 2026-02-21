import os
import urllib.parse
import json
from datetime import datetime, date

import pandas as pd
import folium
from folium import FeatureGroup
from folium.plugins import Search
from openpyxl import load_workbook


# ------------------------
# Config
# ------------------------
EXCEL_FILE = "co_courses.xlsx"
OUTPUT_HTML = "index.html"
OUTPUT_JSON = "courses.json"

TYPE_COLORS = {
    "Public": "#2ecc71",
    "Private": "#3498db",
    "Semi-Private": "#9b59b6",
    "Resort": "#f39c12",
}

DOT_RADIUS = 6
DOT_WEIGHT = 2


# ------------------------
# Helpers
# ------------------------
def normalize_type(x: str) -> str:
    if not isinstance(x, str) or not x.strip():
        return "Public"
    low = x.strip().lower()
    if low in ["semi private", "semi-private", "semi"]:
        return "Semi-Private"
    if low == "public":
        return "Public"
    if low == "private":
        return "Private"
    if low == "resort":
        return "Resort"
    return "Public"


def is_blank(v) -> bool:
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def fmt_date(v) -> str:
    try:
        if pd.isna(v):
            return "—"
    except Exception:
        pass

    if v is None:
        return "—"

    if isinstance(v, pd.Timestamp):
        if pd.isna(v):
            return "—"
        v = v.to_pydatetime()

    if isinstance(v, (datetime, date)):
        try:
            return v.strftime("%-m/%-d/%Y")
        except Exception:
            return "—"

    s = str(v).strip()
    if not s:
        return "—"

    try:
        parsed = pd.to_datetime(s, errors="coerce")
        if pd.isna(parsed):
            return s
        return parsed.to_pydatetime().strftime("%-m/%-d/%Y")
    except Exception:
        return s


def clean_text(s) -> str:
    if is_blank(s):
        return ""
    return str(s).strip()


def build_maps_links(address: str):
    q = urllib.parse.quote(address)
    google = f"https://www.google.com/maps/search/?api=1&query={q}"
    apple = f"https://maps.apple.com/?q={q}"
    return apple, google


def safe_js_str(s: str) -> str:
    return (
        str(s)
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
        .replace("\r", " ")
    )


def extract_reel_links_xlsx(path: str, course_col_name="Course", reel_col_name="Reel"):
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if isinstance(v, str):
            headers[v.strip()] = c

    if course_col_name not in headers or reel_col_name not in headers:
        return {}

    course_col = headers[course_col_name]
    reel_col = headers[reel_col_name]

    out = {}
    for r in range(2, ws.max_row + 1):
        course = ws.cell(row=r, column=course_col).value
        if not isinstance(course, str) or not course.strip():
            continue
        course_name = course.strip()

        cell = ws.cell(row=r, column=reel_col)
        url = ""

        if cell.hyperlink and cell.hyperlink.target:
            url = str(cell.hyperlink.target).strip()

        if not url:
            v = cell.value
            if isinstance(v, str) and v.strip().lower().startswith("http"):
                url = v.strip()

        out[course_name] = url

    return out


# ------------------------
# Load + validate data
# ------------------------
df = pd.read_excel(EXCEL_FILE)
df = df.rename(columns=lambda c: c.strip())

required_cols = ["Course", "Address", "City", "Type", "Region", "Lat", "Long"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing columns: {missing}. Required: {required_cols}")

has_first_played = "1st Played" in df.columns
has_order = "Order" in df.columns
has_reel = "Reel" in df.columns

df["Type"] = df["Type"].apply(normalize_type)
df["Lat"] = pd.to_numeric(df["Lat"], errors="coerce")
df["Long"] = pd.to_numeric(df["Long"], errors="coerce")
df = df.dropna(subset=["Lat", "Long"]).copy()

reel_links = extract_reel_links_xlsx(EXCEL_FILE) if has_reel else {}


# ------------------------
# Build map
# ------------------------
m = folium.Map(
    location=[39.0, -105.55],
    zoom_start=7,
    control_scale=True,
    tiles="OpenStreetMap",
)

MAP_JS_NAME = m.get_name()

type_groups = {}
type_group_js = {}
for t in TYPE_COLORS.keys():
    g = FeatureGroup(name=t, show=True, control=False)
    g.add_to(m)
    type_groups[t] = g
    type_group_js[t] = g.get_name()

markers_meta = []
courses_export = []

for _, r in df.iterrows():
    course = clean_text(r["Course"])
    city = clean_text(r["City"])
    ctype = normalize_type(r["Type"])
    region = clean_text(r["Region"])
    address = clean_text(r["Address"])

    first_played_val = fmt_date(r["1st Played"]) if has_first_played else "—"
    played_by_cam = first_played_val != "—"

    order_val = "—"
    if has_order and not is_blank(r["Order"]):
        try:
            order_val = str(int(r["Order"]))
        except Exception:
            ov = clean_text(r["Order"])
            order_val = ov if ov else "—"

    reel_url = ""
    if has_reel:
        reel_url = reel_links.get(course, "")
        if not reel_url:
            v = r.get("Reel", "")
            if isinstance(v, str) and v.strip().lower().startswith("http"):
                reel_url = v.strip()

    has_video = bool(reel_url)

    apple_maps, google_maps = build_maps_links(address if address else f"{course}, {city}, CO")

    if has_video:
        video_html = f"""
        <a href="{reel_url}" target="_blank" rel="noopener" style="display:block;text-decoration:none;">
          <div style="width:100%;padding:9px 10px;border-radius:12px;border:1px solid rgba(0,0,0,0.18);
                      text-align:center;font-weight:800;color:#0b6aa2;background:white;">IG Reel</div>
        </a>
        """
    else:
        if played_by_cam:
            video_html = """
            <div style="width:100%;padding:9px 10px;border-radius:12px;border:1px solid rgba(0,0,0,0.12);
                        text-align:center;font-weight:800;color:rgba(0,0,0,0.45);background:rgba(0,0,0,0.03);">No video yet</div>
            """
        else:
            video_html = """
            <div style="width:100%;padding:9px 10px;border-radius:12px;border:1px solid rgba(0,0,0,0.12);
                        text-align:center;font-weight:800;color:rgba(0,0,0,0.35);background:rgba(0,0,0,0.02);">Not played yet</div>
            """

    color = TYPE_COLORS.get(ctype, TYPE_COLORS["Public"])

    popup_html = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial;">
      <div style="font-weight:900;font-size:20px;line-height:1.1;margin-bottom:4px;">{course}</div>
      <div style="font-size:14px;opacity:0.75;margin-bottom:10px;">{city} · {region}</div>

      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:70px;opacity:0.55;">Type</div>
        <div style="display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;
                    background:rgba(0,0,0,0.04);font-weight:800;">
          <span style="width:10px;height:10px;border-radius:3px;background:{color};display:inline-block;"></span>
          <span>{ctype}</span>
        </div>
      </div>

      <div style="display:flex;gap:10px;margin-bottom:10px;">
        <div style="width:70px;opacity:0.55;">Address</div>
        <div style="flex:1;font-weight:650;">{address}</div>
      </div>

      <div style="display:flex;gap:10px;margin:10px 0 6px 0;">
        <a href="{apple_maps}" target="_blank" rel="noopener" style="flex:1;text-decoration:none;">
          <div style="padding:10px 12px;border-radius:14px;border:1px solid rgba(0,0,0,0.18);
                      text-align:center;font-weight:900;color:#0b6aa2;background:white;">Open in Maps</div>
        </a>
        <a href="{google_maps}" target="_blank" rel="noopener" style="flex:1;text-decoration:none;">
          <div style="padding:10px 12px;border-radius:14px;border:1px solid rgba(0,0,0,0.18);
                      text-align:center;font-weight:900;color:#0b6aa2;background:white;">Google Maps</div>
        </a>
      </div>

      <div style="height:1px;background:rgba(0,0,0,0.12);margin:12px 0;"></div>

      <div style="font-weight:900;font-size:18px;margin-bottom:8px;">Cam’s Every Course Journey</div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
        <div style="border:1px solid rgba(0,0,0,0.12);border-radius:14px;padding:10px 12px;">
          <div style="font-weight:800;opacity:0.6;margin-bottom:6px;">Course #</div>
          <div style="font-weight:950;font-size:20px;line-height:1.1;">{order_val}</div>
        </div>

        <div style="border:1px solid rgba(0,0,0,0.12);border-radius:14px;padding:10px 12px;
                    display:flex;flex-direction:column;gap:8px;">
          <div>
            <div style="font-weight:800;opacity:0.6;margin-bottom:6px;">First Played</div>
            <div style="font-weight:950;font-size:18px;line-height:1.1;">{first_played_val}</div>
          </div>
          {video_html}
        </div>
      </div>
    </div>
    """

    marker = folium.CircleMarker(
        location=[float(r["Lat"]), float(r["Long"])],
        radius=DOT_RADIUS,
        weight=DOT_WEIGHT,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.9,
        popup=folium.Popup(popup_html, max_width=520),
    )
    marker.add_to(type_groups.get(ctype, type_groups["Public"]))

    markers_meta.append(
        {"js": marker.get_name(), "type": ctype, "played": played_by_cam, "video": has_video}
    )

    courses_export.append(
        {
            "name": course,
            "city": city,
            "region": region,
            "type": ctype,
            "address": address,
            "lat": float(r["Lat"]),
            "lng": float(r["Long"]),
            "played": bool(played_by_cam),
            "first_played": first_played_val,
            "order": None if order_val == "—" else order_val,
            "video_url": reel_url,
            "has_video": bool(has_video),
            "apple_maps": apple_maps,
            "google_maps": google_maps,
        }
    )


features = [
    {
        "type": "Feature",
        "properties": {"Course": clean_text(r["Course"])},
        "geometry": {"type": "Point", "coordinates": [float(r["Long"]), float(r["Lat"])]},
    }
    for _, r in df.iterrows()
]

search_layer = folium.GeoJson(
    {"type": "FeatureCollection", "features": features},
    name="__search_index__",
    show=False,
    control=False,
    marker=folium.CircleMarker(radius=0, opacity=0, fill_opacity=0),
    style_function=lambda x: {"opacity": 0, "fillOpacity": 0},
).add_to(m)

Search(
    layer=search_layer,
    search_label="Course",
    placeholder="Search a course name…",
    collapsed=False,
    position="topright",
    geom_type="Point",
    marker=False,
).add_to(m)


filter_rows = "\n".join(
    [
        f"""
        <label style="display:flex;align-items:center;gap:10px;margin:7px 0;cursor:pointer;">
          <input type="checkbox" class="type-toggle" data-layer="{t}" checked>
          <span style="width:14px;height:14px;background:{TYPE_COLORS[t]};display:inline-block;border:1px solid rgba(0,0,0,0.25);"></span>
          <span>{t}</span>
        </label>
        """
        for t in TYPE_COLORS.keys()
    ]
)

custom_css = """
<style>
  .leaflet-control-search {
    z-index: 9999 !important;
    box-shadow: 0 6px 20px rgba(0,0,0,0.12) !important;
    border-radius: 12px !important;
    background: rgba(255,255,255,0.92) !important;
    border: 1px solid rgba(0,0,0,0.18) !important;
  }
  .leaflet-control-search .search-input {
    width: 230px !important;
    border-radius: 12px !important;
  }
</style>
"""

ui_html = f"""
{custom_css}

<div style="
  position:fixed; bottom:24px; right:18px; z-index:9998;
  background:rgba(255,255,255,0.92);
  border:1px solid rgba(0,0,0,0.2);
  border-radius:12px;
  padding:12px 14px;
  box-shadow:0 6px 20px rgba(0,0,0,0.12);
  font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial;
  font-size:13px;
  width:230px;
">
  <div style="font-weight:900;font-size:14px;margin-bottom:8px;">Course Filters</div>

  <div style="margin-bottom:10px;">
    <label style="display:flex;align-items:center;gap:10px;margin:7px 0;cursor:pointer;">
      <input id="playedOnly" type="checkbox">
      <span style="font-weight:800;">Played by Cam</span>
    </label>

    <label style="display:flex;align-items:center;gap:10px;margin:7px 0;cursor:pointer;">
      <input id="videoOnly" type="checkbox">
      <span style="font-weight:800;">Has video review</span>
    </label>
  </div>

  <div style="height:1px;background:rgba(0,0,0,0.12);margin:10px 0;"></div>

  {filter_rows}

  <div style="display:flex;gap:10px;margin-top:10px;">
    <button id="filterAll" style="flex:1;padding:8px 10px;border-radius:10px;border:1px solid rgba(0,0,0,0.2);background:white;cursor:pointer;">All</button>
    <button id="filterNone" style="flex:1;padding:8px 10px;border-radius:10px;border:1px solid rgba(0,0,0,0.2);background:white;cursor:pointer;">None</button>
  </div>
</div>
"""

type_layers_js = ",\n".join([f'"{t}": window["{type_group_js[t]}"]' for t in TYPE_COLORS.keys()])

markers_js_list = ",\n".join(
    [
        f'{{m: window["{mm["js"]}"], type:"{safe_js_str(mm["type"])}", played:{str(mm["played"]).lower()}, video:{str(mm["video"]).lower()}}}'
        for mm in markers_meta
    ]
)

ui_js = f"""
document.addEventListener("DOMContentLoaded", function() {{
  const mapObj = window.{MAP_JS_NAME};
  if (!mapObj) {{
    console.warn("Map object not found");
    return;
  }}

  const typeLayers = {{
    {type_layers_js}
  }};

  const markers = [
    {markers_js_list}
  ];

  function getTypeState() {{
    const state = {{}};
    document.querySelectorAll(".type-toggle").forEach(cb => {{
      state[cb.dataset.layer] = cb.checked;
    }});
    return state;
  }}

  function applyFilters() {{
    const typeState = getTypeState();
    const playedOnly = !!document.getElementById("playedOnly")?.checked;
    const videoOnly = !!document.getElementById("videoOnly")?.checked;

    Object.keys(typeLayers).forEach(t => {{
      const g = typeLayers[t];
      if (!g) return;

      const wantType = !!typeState[t];
      if (wantType) {{
        if (!mapObj.hasLayer(g)) mapObj.addLayer(g);
      }} else {{
        if (mapObj.hasLayer(g)) mapObj.removeLayer(g);
      }}
    }});

    markers.forEach(obj => {{
      if (!obj.m) return;

      const wantType = !!typeState[obj.type];
      let ok = wantType;

      if (playedOnly) ok = ok && obj.played;
      if (videoOnly) ok = ok && obj.video;

      if (ok) {{
        if (!mapObj.hasLayer(obj.m)) mapObj.addLayer(obj.m);
      }} else {{
        if (mapObj.hasLayer(obj.m)) mapObj.removeLayer(obj.m);
      }}
    }});
  }}

  document.querySelectorAll(".type-toggle").forEach(cb => cb.addEventListener("change", applyFilters));
  document.getElementById("playedOnly")?.addEventListener("change", applyFilters);
  document.getElementById("videoOnly")?.addEventListener("change", applyFilters);

  document.getElementById("filterAll")?.addEventListener("click", () => {{
    document.querySelectorAll(".type-toggle").forEach(cb => cb.checked = true);
    applyFilters();
  }});

  document.getElementById("filterNone")?.addEventListener("click", () => {{
    document.querySelectorAll(".type-toggle").forEach(cb => cb.checked = false);
    applyFilters();
  }});

  applyFilters();
}});
"""

m.get_root().html.add_child(folium.Element(ui_html))
m.get_root().script.add_child(folium.Element(ui_js))

base_dir = os.path.dirname(os.path.abspath(__file__))
out_html_path = os.path.join(base_dir, OUTPUT_HTML)
out_json_path = os.path.join(base_dir, OUTPUT_JSON)

m.save(out_html_path)

with open(out_json_path, "w", encoding="utf-8") as f:
    json.dump(courses_export, f, ensure_ascii=False, indent=2)

print("Map created!")
print(f"Exported {len(courses_export)} rows to {out_json_path}")
