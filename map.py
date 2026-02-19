import pandas as pd
import folium
from folium import FeatureGroup
from folium.plugins import Search
from urllib.parse import quote_plus
import openpyxl

EXCEL_FILE = "co_courses.xlsx"
OUTPUT_HTML = "colorado_golf_courses_map.html"

TYPE_COLORS = {
    "Public": "#2ecc71",
    "Private": "#3498db",
    "Semi-Private": "#9b59b6",
    "Resort": "#f39c12",
    "Military": "#16a085",
    "Other/Unknown": "#7f8c8d",
}

# Only show these in the filter UI
FILTER_TYPES = ["Public", "Private", "Semi-Private", "Resort"]


def normalize_type(x: str) -> str:
    if not isinstance(x, str) or not x.strip():
        return "Other/Unknown"
    t = x.strip().lower()
    if "semi" in t:
        return "Semi-Private"
    if t == "public":
        return "Public"
    if t == "private":
        return "Private"
    if t == "resort":
        return "Resort"
    if any(k in t for k in ["military", "usaf", "army", "navy", "air force", "space force", "marines"]):
        return "Military"
    return "Other/Unknown"


def fmt_date(v) -> str:
    if v is None:
        return "—"
    try:
        if pd.isna(v):
            return "—"
    except Exception:
        pass
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return "—"
    return dt.strftime("%-m/%-d/%Y")


def load_reel_hyperlinks(path: str):
    """
    Reads the Excel workbook directly so we can grab REAL hyperlink targets
    from the 'Reel' column (since pandas won't preserve hyperlinks).
    Returns a list aligned to data rows (row 2 = index 0).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    reel_col = None
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if isinstance(val, str) and val.strip().lower() == "reel":
            reel_col = c
            break

    if reel_col is None:
        return []

    urls = []
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=reel_col)
        url = None
        if cell.hyperlink and cell.hyperlink.target:
            url = cell.hyperlink.target
        else:
            # fallback if the cell contains a raw URL text
            if isinstance(cell.value, str) and cell.value.strip().startswith(("http://", "https://")):
                url = cell.value.strip()
        urls.append(url)

    return urls


# -------------------------
# Load and prep data
# -------------------------
df = pd.read_excel(EXCEL_FILE)
df = df.rename(columns=lambda c: c.strip())

required_cols = ["Course", "Address", "City", "Type", "Region", "Lat", "Long"]
for c in required_cols:
    if c not in df.columns:
        raise ValueError(f"Missing column: {c}. Need columns: {required_cols}")

has_order = "Order" in df.columns
has_first_played = "1st Played" in df.columns
has_reel = "Reel" in df.columns

df["Type"] = df["Type"].apply(normalize_type)
df["Lat"] = pd.to_numeric(df["Lat"], errors="coerce")
df["Long"] = pd.to_numeric(df["Long"], errors="coerce")

if has_first_played:
    df["1st Played"] = pd.to_datetime(df["1st Played"], errors="coerce")

df = df.dropna(subset=["Lat", "Long"]).reset_index(drop=True)

# Attach reel URLs from hyperlinks
df["Reel_URL"] = None
if has_reel:
    reel_urls = load_reel_hyperlinks(EXCEL_FILE)
    if len(reel_urls) >= len(df):
        df["Reel_URL"] = reel_urls[: len(df)]
    else:
        df["Reel_URL"] = (reel_urls + [None] * (len(df) - len(reel_urls)))[: len(df)]


# -------------------------
# Build map (street only)
# -------------------------
m = folium.Map(
    location=[39.0, -105.55],
    zoom_start=7,
    tiles="OpenStreetMap",
    control_scale=True,
)
MAP_JS_NAME = m.get_name()

# Popup CSS (cleaner layout, better fitting numbers)
popup_css = """
<style>
  .leaflet-popup-content-wrapper { border-radius: 14px !important; }
  .leaflet-popup-content { margin: 10px 12px !important; width: auto !important; }

  .cg-wrap { width: 420px; max-width: 92vw; font-family:-apple-system,Segoe UI,Roboto,Arial; }
  .cg-top { display:flex; align-items:flex-start; justify-content:space-between; gap:10px; }
  .cg-title { font-weight:950; font-size:18px; line-height:1.15; margin:0; color:#111; }
  .cg-sub { font-size:13px; color:#60646b; margin:6px 0 10px 0; }

  .cg-pill { display:inline-flex; align-items:center; gap:7px; padding:5px 10px;
             border-radius:999px; background:rgba(0,0,0,0.06);
             font-size:12px; font-weight:950; white-space:nowrap; }
  .cg-dot { width:9px; height:9px; border-radius:3px; border:1px solid rgba(0,0,0,0.18); display:inline-block; }

  .cg-address { font-size:13px; color:#111; margin:0 0 10px 0; }

  .cg-actions { display:flex; gap:10px; margin:10px 0 8px 0; }
  .cg-btn { flex:1; display:block; text-align:center; padding:9px 10px;
            border-radius:12px; border:1px solid rgba(0,0,0,0.18);
            background:white; font-size:13px; font-weight:950;
            text-decoration:none; color:#0b4f6c; }
  .cg-btn:hover { background:rgba(0,0,0,0.05); }

  .cg-journey { margin-top:10px; padding-top:10px; border-top:1px solid rgba(0,0,0,0.10); }
  .cg-journey-title { font-size:13px; font-weight:950; color:#222; margin:0 0 8px 0; }

  .cg-grid { display:grid; grid-template-columns: 1fr 1fr; gap:10px; }
  .cg-box { border:1px solid rgba(0,0,0,0.10); border-radius:12px; padding:10px; }
  .cg-k { font-size:11px; color:#6b7280; font-weight:900; margin:0 0 5px 0; }
  .cg-v { font-size:16px; color:#111; font-weight:950; margin:0; line-height:1.1; }

  .cg-mini-btn { margin-top:8px; display:block; text-align:center; padding:7px 8px;
                 border-radius:10px; border:1px solid rgba(0,0,0,0.14);
                 font-size:12px; font-weight:950; text-decoration:none; color:#0b4f6c; }
  .cg-muted { margin-top:8px; font-size:12px; font-weight:900; color:#8a8f98; text-align:center; }
</style>
"""
m.get_root().html.add_child(folium.Element(popup_css))

# -------------------------
# Layer groups by type
# -------------------------
type_groups = {}
type_group_js = {}
for t in TYPE_COLORS.keys():
    g = FeatureGroup(name=t, show=True, control=False)
    g.add_to(m)
    type_groups[t] = g
    type_group_js[t] = g.get_name()

# -------------------------
# Add markers + popups
# -------------------------
for _, r in df.iterrows():
    course = str(r["Course"]).strip()
    city = str(r["City"]).strip()
    region = str(r["Region"]).strip()
    ctype = normalize_type(r["Type"])
    address = str(r["Address"]).strip()

    lat = float(r["Lat"])
    lon = float(r["Long"])
    color = TYPE_COLORS.get(ctype, TYPE_COLORS["Other/Unknown"])

    # Google Maps link uses address (not lat/long)
    full_address = f"{address}, {city}, CO"
    maps_url = f"https://www.google.com/maps/search/?api=1&query={quote_plus(full_address)}"

    # Stats
    order_val = "—"
    if has_order:
        ov = r.get("Order")
        if ov is not None and not pd.isna(ov):
            try:
                order_val = str(int(ov))
            except Exception:
                order_val = str(ov)

    first_played_val = fmt_date(r.get("1st Played")) if has_first_played else "—"

    reel_url = r.get("Reel_URL")
    has_reel_btn = isinstance(reel_url, str) and reel_url.strip().startswith(("http://", "https://"))

    if first_played_val != "—":
        if has_reel_btn:
            reel_html = f'<a class="cg-mini-btn" href="{reel_url}" target="_blank" rel="noopener">IG Reel</a>'
        else:
            reel_html = '<div class="cg-muted">No video yet</div>'
    else:
        reel_html = '<div class="cg-muted">Not played</div>'

    popup_html = f"""
    <div class="cg-wrap">
      <div class="cg-top">
        <div>
          <div class="cg-title">{course}</div>
          <div class="cg-sub">{city} · {region}</div>
        </div>
        <div class="cg-pill">
          <span class="cg-dot" style="background:{color};"></span>
          {ctype}
        </div>
      </div>

      <div class="cg-address">{address}</div>

      <div class="cg-actions">
        <a class="cg-btn" href="{maps_url}" target="_blank" rel="noopener">Open in Maps</a>
      </div>

      <div class="cg-journey">
        <div class="cg-journey-title">Cam’s Every Course Journey</div>
        <div class="cg-grid">
          <div class="cg-box">
            <div class="cg-k">Course #</div>
            <div class="cg-v">{order_val}</div>
          </div>
          <div class="cg-box">
            <div class="cg-k">First Played</div>
            <div class="cg-v">{first_played_val}</div>
            {reel_html}
          </div>
        </div>
      </div>
    </div>
    """

    folium.CircleMarker(
        location=[lat, lon],
        radius=6,
        weight=2,
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.9,
        popup=folium.Popup(popup_html, max_width=540),
    ).add_to(type_groups.get(ctype, type_groups["Other/Unknown"]))

# -------------------------
# Search index
# -------------------------
features = []
for _, r in df.iterrows():
    features.append(
        {
            "type": "Feature",
            "properties": {"Course": str(r["Course"]).strip()},
            "geometry": {"type": "Point", "coordinates": [float(r["Long"]), float(r["Lat"])]},
        }
    )

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

# -------------------------
# Filter UI (bottom right)
# -------------------------
filter_rows = "\n".join(
    [
        f"""
        <label style="display:flex;align-items:center;gap:10px;margin:7px 0;cursor:pointer;">
          <input type="checkbox" class="type-toggle" data-layer="{t}" checked>
          <span style="width:14px;height:14px;background:{TYPE_COLORS[t]};display:inline-block;border:1px solid rgba(0,0,0,0.25);"></span>
          <span>{t}</span>
        </label>
        """
        for t in FILTER_TYPES
    ]
)

type_layer_map_js = ",\n".join([f'"{t}": {type_group_js[t]}' for t in FILTER_TYPES])

ui_html = f"""
<div style="
  position:fixed; bottom:24px; right:18px; z-index:9999;
  background:rgba(255,255,255,0.95);
  border:1px solid rgba(0,0,0,0.2);
  border-radius:12px;
  padding:12px 14px;
  box-shadow:0 6px 20px rgba(0,0,0,0.12);
  font-family:-apple-system,Segoe UI,Roboto,Arial;
  font-size:13px;
  width:230px;
">
  <div style="font-weight:950;font-size:14px;margin-bottom:8px;">Course Filters</div>
  {filter_rows}
  <div style="display:flex;gap:10px;margin-top:10px;">
    <button id="filterAll" style="flex:1;padding:8px 10px;border-radius:10px;border:1px solid rgba(0,0,0,0.2);background:white;cursor:pointer;font-weight:900;">All</button>
    <button id="filterNone" style="flex:1;padding:8px 10px;border-radius:10px;border:1px solid rgba(0,0,0,0.2);background:white;cursor:pointer;font-weight:900;">None</button>
  </div>
</div>
"""
ui_js = f"""
window.addEventListener("load", function() {{
  const mapObj = window.{MAP_JS_NAME};
  if (!mapObj) return;

  const typeLayers = {{
    {type_layer_map_js}
  }};

  function setTypeVisible(name, visible) {{
    const layer = typeLayers[name];
    if (!layer) return;
    if (visible) {{
      if (!mapObj.hasLayer(layer)) mapObj.addLayer(layer);
    }} else {{
      if (mapObj.hasLayer(layer)) mapObj.removeLayer(layer);
    }}
  }}

  document.querySelectorAll(".type-toggle").forEach(cb => {{
    cb.addEventListener("change", () => {{
      setTypeVisible(cb.dataset.layer, cb.checked);
    }});
  }});

  document.getElementById("filterAll").addEventListener("click", () => {{
    document.querySelectorAll(".type-toggle").forEach(cb => {{
      cb.checked = true;
      setTypeVisible(cb.dataset.layer, true);
    }});
  }});

  document.getElementById("filterNone").addEventListener("click", () => {{
    document.querySelectorAll(".type-toggle").forEach(cb => {{
      cb.checked = false;
      setTypeVisible(cb.dataset.layer, false);
    }});
  }});
}});
"""
m.get_root().html.add_child(folium.Element(ui_html))
m.get_root().script.add_child(folium.Element(ui_js))

# -------------------------
# Logo overlay (top left)
# -------------------------
logo_filename = "EveryCourse LogoStacked.PNG"
logo_html = f"""
<div style="
  position: fixed;
  top: 18px;
  left: 18px;
  z-index: 9999;
  background: rgba(255,255,255,0.92);
  padding: 10px 12px;
  border-radius: 14px;
  box-shadow: 0 6px 20px rgba(0,0,0,0.15);
">
  <img src="{logo_filename}" style="height:52px; display:block;">
</div>
"""
m.get_root().html.add_child(folium.Element(logo_html))

# Save
m.save(OUTPUT_HTML)
print("Map created!")