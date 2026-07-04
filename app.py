from __future__ import annotations
import os
import io
import uuid
import copy
import json
import math
import threading
import traceback
import re
import time
from pathlib import Path
import requests

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree
import anthropic

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# In-memory job registry
# {job_id: {"status": ..., "message": ..., "progress": ..., "plan": ..., "pptx_path": ..., "output": ...}}
jobs: dict = {}
jobs_lock = threading.Lock()


# ---------------------------------------------------------------------------
# PPTX helpers
# ---------------------------------------------------------------------------

def clone_slide(prs: Presentation, source_index: int = 0):
    """Clone the slide at source_index and append a copy to the presentation."""
    source = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source.slide_layout)

    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree):
        sp_tree.remove(child)

    for child in source.shapes._spTree:
        sp_tree.append(copy.deepcopy(child))

    for rel in source.part.rels.values():
        if "image" in rel.reltype:
            try:
                new_slide.part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                pass

    return new_slide


def replace_text_in_slide(slide, replacements: dict, ordered: dict = None):
    """
    Replace placeholder tokens in every text frame on a slide.

    replacements  – {old: new} for unique tokens
    ordered       – {old: [val1, val2, val3]} for tokens that appear
                    multiple times; replaced in document order (top→bottom)
    """
    order_counts = {k: 0 for k in (ordered or {})}

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full_text = "".join(run.text for run in para.runs)
            if not full_text.strip():
                continue

            modified = full_text
            changed = False

            for placeholder, value in replacements.items():
                if placeholder in modified:
                    modified = modified.replace(
                        placeholder, str(value) if value is not None else ""
                    )
                    changed = True

            for placeholder, values in (ordered or {}).items():
                if placeholder in modified:
                    idx = order_counts[placeholder]
                    if idx < len(values):
                        modified = modified.replace(
                            placeholder,
                            str(values[idx]) if values[idx] is not None else "",
                        )
                        order_counts[placeholder] += 1
                        changed = True

            if changed and para.runs:
                para.runs[0].text = modified
                for run in para.runs[1:]:
                    run.text = ""


# ---------------------------------------------------------------------------
# Real landmark lookup — Google Maps (preferred) with OSM fallback
# ---------------------------------------------------------------------------

OSM_HEADERS = {"User-Agent": "Skyscale-OOH-Generator/1.0 (contact@skyscale.com)"}

def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(dlon / 2) ** 2)
    return R * 2 * math.asin(math.sqrt(a))


def _get_landmarks_google(location: str, city: str, n: int = 3):
    api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "")
    if not api_key:
        return None
    try:
        geo = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params={"address": f"{location}, {city}", "key": api_key},
            timeout=8,
        ).json()
        if geo.get("status") != "OK":
            return None
        loc = geo["results"][0]["geometry"]["location"]
        lat, lng = loc["lat"], loc["lng"]

        places = requests.get(
            "https://maps.googleapis.com/maps/api/place/nearbysearch/json",
            params={
                "location": f"{lat},{lng}",
                "radius": 5000,
                "type": "point_of_interest",
                "key": api_key,
            },
            timeout=10,
        ).json()

        results = []
        seen: set = set()
        for p in places.get("results", []):
            name = p.get("name", "").strip()
            if not name or name in seen:
                continue
            seen.add(name)
            p_lat = p["geometry"]["location"]["lat"]
            p_lng = p["geometry"]["location"]["lng"]
            dist = _haversine_km(lat, lng, p_lat, p_lng)
            results.append((dist, name))

        results.sort(key=lambda x: x[0])
        filtered = [(d, name) for d, name in results if d <= 5.0][:n]
        return [
            f"{name} – {round(d, 1) if d >= 0.1 else 0.1}km"
            for d, name in filtered
        ] or None

    except Exception:
        return None


def _get_landmarks_osm(location: str, city: str, n: int = 3):
    try:
        geo_resp = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": f"{location}, {city}", "format": "json", "limit": 1},
            headers=OSM_HEADERS,
            timeout=8,
        )
        geo_data = geo_resp.json()
        if not geo_data:
            return None
        lat = float(geo_data[0]["lat"])
        lon = float(geo_data[0]["lon"])
        time.sleep(1.1)

        overpass_query = f"""
[out:json][timeout:12];
(
  node["name"]["tourism"](around:5000,{lat},{lon});
  node["name"]["amenity"~"^(restaurant|cafe|hotel|bank|museum|theatre|cinema|hospital|university|library|historic)$"](around:5000,{lat},{lon});
  node["name"]["historic"](around:5000,{lat},{lon});
  node["name"]["shop"~"^(mall|department_store|supermarket)$"](around:5000,{lat},{lon});
);
out center 20;
"""
        elements = requests.post(
            "https://overpass-api.de/api/interpreter",
            data={"data": overpass_query},
            headers=OSM_HEADERS,
            timeout=15,
        ).json().get("elements", [])

        seen: set = set()
        ranked: list = []
        for el in elements:
            name = el.get("tags", {}).get("name", "").strip()
            if not name or name in seen:
                continue
            seen.add(name)
            el_lat = el.get("lat") or el.get("center", {}).get("lat", lat)
            el_lon = el.get("lon") or el.get("center", {}).get("lon", lon)
            ranked.append((_haversine_km(lat, lon, float(el_lat), float(el_lon)), name))

        ranked.sort(key=lambda x: x[0])
        results = []
        for dist, name in ranked:
            if dist > 5.0:
                break
            km = round(dist, 1) if dist >= 0.1 else 0.1
            results.append(f"{name} – {km}km")
            if len(results) == n:
                break

        return results if len(results) >= n else None

    except Exception:
        return None


def get_real_landmarks(location: str, city: str, n: int = 3):
    result = _get_landmarks_google(location, city, n)
    if result and len(result) >= n:
        return result
    return _get_landmarks_osm(location, city, n)


# ---------------------------------------------------------------------------
# Google Maps Static screenshot
# Adjust these constants to match your template layout (inches from top-left)
# ---------------------------------------------------------------------------

MAP_IMG_LEFT   = Inches(6.60)   # horizontal offset from slide left edge
MAP_IMG_TOP    = Inches(1.20)   # vertical offset from slide top edge
MAP_IMG_WIDTH  = Inches(3.10)   # image width
MAP_IMG_HEIGHT = Inches(2.10)   # image height


def get_map_image_bytes(location: str, city: str, zoom: int = 16) -> bytes | None:
    """
    Return a map image for the location.
    Uses Google Maps Static API when GOOGLE_MAPS_API_KEY is set,
    otherwise geocodes via Nominatim and fetches from staticmap.openstreetmap.de
    — no extra dependencies needed.
    """
    UA = "OOHProposalGenerator/1.0 (skyscalemedia.com)"
    api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "")

    if api_key:
        try:
            geo = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params={"address": f"{location}, {city}", "key": api_key},
                timeout=8,
            ).json()
            if geo.get("status") != "OK" or not geo.get("results"):
                return None
            loc = geo["results"][0]["geometry"]["location"]
            lat, lng = loc["lat"], loc["lng"]
            resp = requests.get(
                "https://maps.googleapis.com/maps/api/staticmap",
                params={
                    "center": f"{lat},{lng}",
                    "zoom": zoom,
                    "size": "600x400",
                    "scale": 2,
                    "maptype": "roadmap",
                    "markers": f"color:red|size:mid|{lat},{lng}",
                    "key": api_key,
                },
                timeout=12,
            )
            ct = resp.headers.get("content-type", "")
            if resp.status_code == 200 and ct.startswith("image"):
                return resp.content
        except Exception:
            pass
        return None

    # ── Free fallback: Nominatim geocode → stitch OSM tiles with Pillow ──
    # Respect Nominatim's 1 req/sec rate limit
    time.sleep(1.1)
    try:
        geo = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": f"{location}, {city}", "format": "json", "limit": 1},
            headers={"User-Agent": UA},
            timeout=8,
        ).json()
        if not geo:
            print(f"[MAP] Nominatim found nothing for {location!r}, {city!r}")
            return None
        lat, lng = float(geo[0]["lat"]), float(geo[0]["lon"])
        print(f"[MAP] geocoded {location!r} → lat={lat:.4f} lng={lng:.4f}")
    except Exception as e:
        print(f"[MAP] Nominatim failed for {location!r}: {e}")
        return None

    # Stitch a 3×3 grid of OSM tiles into one image
    try:
        from PIL import Image, ImageDraw

        z = min(zoom, 15)
        n = 2 ** z
        lat_rad = math.radians(lat)
        # Fractional tile position of the exact location
        fx = (lng + 180) / 360 * n
        fy = (1 - math.log(math.tan(lat_rad) + 1 / math.cos(lat_rad)) / math.pi) / 2 * n
        tx = int(fx)
        ty = int(fy)

        TILE = 256
        canvas = Image.new("RGB", (TILE * 3, TILE * 3), (220, 220, 215))
        ok = 0
        for row in range(3):
            for col in range(3):
                url = f"https://tile.openstreetmap.org/{z}/{tx - 1 + col}/{ty - 1 + row}.png"
                try:
                    r = requests.get(url, headers={"User-Agent": UA}, timeout=8)
                    if r.status_code == 200 and "image" in r.headers.get("content-type", ""):
                        tile = Image.open(io.BytesIO(r.content)).convert("RGB")
                        canvas.paste(tile, (col * TILE, row * TILE))
                        ok += 1
                except Exception:
                    pass

        print(f"[MAP] stitched {ok}/9 tiles for {location!r}")

        # Pin at the exact sub-pixel location within the 3×3 canvas
        pr = 13
        cx = int((fx - (tx - 1)) * TILE)
        cy = int((fy - (ty - 1)) * TILE)
        cx = max(pr + 3, min(canvas.width  - pr - 3, cx))
        cy = max(pr + 3, min(canvas.height - pr - 3, cy))

        draw = ImageDraw.Draw(canvas)
        # Soft shadow (solid dark grey offset)
        draw.ellipse([cx - pr + 3, cy - pr + 3, cx + pr + 3, cy + pr + 3],
                     fill=(100, 100, 100))
        # Red pin with white border
        draw.ellipse([cx - pr, cy - pr, cx + pr, cy + pr],
                     fill=(220, 30, 30), outline="white", width=3)

        buf = io.BytesIO()
        canvas.save(buf, format="PNG")
        return buf.getvalue()
    except Exception as e:
        print(f"[MAP] tile stitch failed for {location!r}: {e}")
        return None


def _cleanup_map_images(job_id: str):
    """Delete all map image files for a job."""
    with jobs_lock:
        map_paths = jobs.get(job_id, {}).get("map_paths", {})
    for path_str in map_paths.values():
        try:
            Path(path_str).unlink(missing_ok=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# AI content generation
# ---------------------------------------------------------------------------

def generate_site_content(site: dict, client: anthropic.Anthropic) -> dict:
    site_name = site.get("Site Name", "")
    location  = site.get("Location", "")
    market    = site.get("Market", "")
    fmt       = site.get("Format", "")
    size      = site.get("Size", "")
    is_mobile = str(location).strip().lower() == "various"

    lookup_address = market if is_mobile else location
    real_landmarks: list | None = get_real_landmarks(lookup_address, market)

    if real_landmarks:
        landmark_instruction = (
            "Real nearby landmarks have already been sourced from a map service. "
            "For landmark_1/2/3 return exactly these strings unchanged:\n"
            + "\n".join(f"  {i+1}. {l}" for i, l in enumerate(real_landmarks))
        )
        # Use json.dumps to safely escape any special characters in landmark strings
        landmark_format = (
            f'"landmark_1": {json.dumps(real_landmarks[0])},\n'
            f'  "landmark_2": {json.dumps(real_landmarks[1])},\n'
            f'  "landmark_3": {json.dumps(real_landmarks[2])}'
        )
    else:
        landmark_instruction = (
            "Real map lookup was unavailable. Use your knowledge of this city to name "
            "3 specific, well-known nearby landmarks within 5km. "
            'Format each as "Landmark Name – 0.Xkm" (max 5km).'
        )
        landmark_format = (
            '"landmark_1": "Landmark Name – 0.Xkm",\n'
            '  "landmark_2": "Landmark Name – 0.Xkm",\n'
            '  "landmark_3": "Landmark Name – 0.Xkm"'
        )

    prompt = f"""You are writing punchy, professional copy for an OOH (Out-of-Home) advertising proposal.

Site details:
- Name: {site_name}
- Location / Address: {location}
- City / Market: {market}
- Format: {fmt}
- Size: {size}

Return ONLY valid JSON (no markdown fences, no extra text) with exactly these keys:

{{
  "tagline": "<4–7 word punchy advertising tagline for this site>",
  "location_desc": "<2–3 sentences describing where the site is and what surrounds it>",
  "visibility_desc": "<2–3 sentences about viewing angles, physical size, and sightlines>",
  "audience_desc": "<2–3 sentences about who passes by and approximate daily volume>",
  {landmark_format}
}}

{landmark_instruction}"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    text = raw.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # Extract the outermost {...} block and retry
        m = re.search(r'\{[\s\S]*\}', text)
        if m:
            try:
                return json.loads(m.group())
            except json.JSONDecodeError:
                pass
        raise ValueError(f"AI returned invalid JSON: {text[:200]}")


# ---------------------------------------------------------------------------
# Helper: build replacement maps from a site plan dict
# ---------------------------------------------------------------------------

def _build_replacements(site: dict) -> tuple[dict, dict]:
    """Return (replacements, ordered) dicts for replace_text_in_slide."""
    lm = [site.get("landmark_1", ""), site.get("landmark_2", ""), site.get("landmark_3", "")]

    replacements = {
        # xyz-format template tokens
        "Site Name":       site.get("site_name", ""),
        "Headline":        site.get("tagline", ""),
        "Size: xyz":       f"Size: {site.get('size', '')}",
        "Format: xyz":     f"Format: {site.get('format', '')}",
        "Location: xyz":   f"Location: {site.get('location', '')}",
        "Frequency: xyz":  f"Frequency: {site.get('frequency', '')}",
        "Units: xyz":      f"Units: {site.get('units', '')}",
        "Traffic: xyz":    f"Traffic: {site.get('traffic', '')}",
        # {TOKEN} style
        "{SITE_NAME}":      site.get("site_name", ""),
        "{TAGLINE}":        site.get("tagline", ""),
        "{LOCATION_DESC}":  site.get("location_desc", ""),
        "{VISIBILITY_DESC}": site.get("visibility_desc", ""),
        "{AUDIENCE_DESC}":  site.get("audience_desc", ""),
        "{SIZE}":           site.get("size", ""),
        "{LOCATION}":       site.get("location", ""),
        "{UNITS}":          site.get("units", ""),
        "{FORMAT}":         site.get("format", ""),
        "{FREQUENCY}":      site.get("frequency", ""),
        "{TRAFFIC}":        site.get("traffic", ""),
        "{LANDMARK_1}":     lm[0],
        "{LANDMARK_2}":     lm[1],
        "{LANDMARK_3}":     lm[2],
        "{MARKET}":         site.get("market", ""),
    }

    ordered = {
        "Text": [
            site.get("location_desc", ""),
            site.get("visibility_desc", ""),
            site.get("audience_desc", ""),
        ],
        "Xyz - 0.5km":     lm,
        "Xyz – 0.5km": lm,
        "Xyz –0.5km":  lm,
        "Xyz -0.5km":      lm,
    }

    return replacements, ordered


# ---------------------------------------------------------------------------
# PPTX build worker (shared by both plan-based and legacy one-shot flows)
# ---------------------------------------------------------------------------

def build_pptx_from_plan(job_id: str, pptx_path: Path, plan: list):
    """Background job: build PPTX from a pre-computed plan list."""
    def update(status: str, message: str, progress: int = 0):
        with jobs_lock:
            jobs[job_id]["status"] = status
            jobs[job_id]["message"] = message
            jobs[job_id]["progress"] = progress

    try:
        update("building", "Loading template…", 5)
        prs = Presentation(str(pptx_path))
        if not prs.slides:
            raise ValueError("The PowerPoint template has no slides.")

        template_slide  = prs.slides[0]
        template_layout = template_slide.slide_layout
        template_spTree = copy.deepcopy(template_slide.shapes._spTree)
        template_spTree_xml = etree.tostring(template_spTree, encoding="unicode")

        # Grab server-side map paths (not sent to client)
        with jobs_lock:
            map_paths = dict(jobs.get(job_id, {}).get("map_paths", {}))

        total = len(plan)

        for idx, site in enumerate(plan):
            pct = 5 + int((idx / total) * 90)
            update("building", f"Building slide {idx + 1}/{total}: {site.get('site_name', '')}…", pct)

            replacements, ordered = _build_replacements(site)

            if idx == 0:
                slide = prs.slides[0]
            else:
                slide = prs.slides.add_slide(template_layout)

                rId_map = {}
                for rel_id, rel in template_slide.part.rels.items():
                    if "image" in rel.reltype or "media" in rel.reltype:
                        try:
                            new_rId = slide.part.relate_to(rel.target_part, rel.reltype)
                            if new_rId != rel_id:
                                rId_map[rel_id] = new_rId
                        except Exception:
                            pass

                xml = template_spTree_xml
                for old_id, new_id in rId_map.items():
                    xml = xml.replace(f'r:embed="{old_id}"', f'r:embed="{new_id}"')
                    xml = xml.replace(f'r:link="{old_id}"',  f'r:link="{new_id}"')

                new_tree = slide.shapes._spTree
                for child in list(new_tree):
                    new_tree.remove(child)
                for child in etree.fromstring(xml):
                    new_tree.append(copy.deepcopy(child))

            replace_text_in_slide(slide, replacements, ordered)

            # Add Google Maps screenshot if available
            map_path_str = map_paths.get(idx)
            if map_path_str:
                map_file = Path(map_path_str)
                if map_file.exists():
                    try:
                        slide.shapes.add_picture(
                            str(map_file),
                            MAP_IMG_LEFT, MAP_IMG_TOP,
                            MAP_IMG_WIDTH, MAP_IMG_HEIGHT,
                        )
                        print(f"[PPTX] map added to slide {idx}")
                    except Exception as pic_err:
                        print(f"[WARN] add_picture failed for slide {idx}: {pic_err}")
                else:
                    print(f"[WARN] map file missing on disk for slide {idx}: {map_path_str}")
            else:
                print(f"[PPTX] no map path for slide {idx} — skipping")

        update("building", "Saving output file…", 96)
        output_filename = f"OOH_Proposal_{job_id[:8]}.pptx"
        output_path = OUTPUT_FOLDER / output_filename
        prs.save(str(output_path))

        with jobs_lock:
            jobs[job_id]["status"]   = "done"
            jobs[job_id]["message"]  = f"Done! {total} slide(s) generated."
            jobs[job_id]["progress"] = 100
            jobs[job_id]["output"]   = output_filename

    except Exception as exc:
        with jobs_lock:
            jobs[job_id]["status"]   = "error"
            jobs[job_id]["message"]  = f"Error: {exc}"
            jobs[job_id]["progress"] = 0
        print(traceback.format_exc())

    finally:
        try:
            pptx_path.unlink(missing_ok=True)
        except Exception:
            pass
        _cleanup_map_images(job_id)


# ---------------------------------------------------------------------------
# Plan generation worker
# ---------------------------------------------------------------------------

def generate_plan_job(job_id: str, excel_path: Path):
    """Background job: read Excel + AI/landmarks, produce a content plan."""
    def update(status: str, message: str, progress: int = 0):
        with jobs_lock:
            jobs[job_id]["status"]   = status
            jobs[job_id]["message"]  = message
            jobs[job_id]["progress"] = progress

    try:
        update("planning", "Reading Excel file…", 5)
        df = pd.read_excel(excel_path, engine="openpyxl")
        df.columns = [c.strip() for c in df.columns]

        if "Market" in df.columns:
            df["Market"] = df["Market"].ffill()

        if "Site Name" not in df.columns:
            raise ValueError("Excel file must have a 'Site Name' column.")

        df = df[df["Site Name"].notna() & (df["Site Name"].astype(str).str.strip() != "")]
        df = df.reset_index(drop=True)

        if df.empty:
            raise ValueError("No valid site rows found in the Excel file.")

        total = len(df)
        update("planning", f"Found {total} site(s). Connecting to AI…", 10)

        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable is not set.")
        client = anthropic.Anthropic(api_key=api_key)

        plan = []
        for idx, row in df.iterrows():
            pct = 10 + int(((idx + 1) / total) * 85)
            site_name = str(row.get("Site Name", "")).strip()
            update("planning", f"Researching site {idx + 1}/{total}: {site_name}…", pct)

            location  = str(row.get("Location", "")).strip()
            market    = str(row.get("Market", "")).strip()
            is_mobile = location.lower() == "various"

            spot_dur = str(row.get("Spot Duration", "")).strip()
            sov_loop = str(row.get("SOV/Loop", "")).strip()
            if spot_dur.lower() in ("", "nan", "n/a", "na"):
                frequency = sov_loop
            else:
                frequency = f"{spot_dur} {sov_loop}".strip()

            raw_impacts = row.get("Impacts", "")
            try:
                traffic = f"{int(float(str(raw_impacts).replace(',', ''))):,}"
            except (ValueError, TypeError):
                traffic = str(raw_impacts).strip()

            try:
                ai = generate_site_content(row.to_dict(), client)
            except Exception as ai_err:
                print(f"[WARN] AI failed for site {site_name!r}: {ai_err}")
                ai = {
                    "tagline":         "",
                    "location_desc":   "",
                    "visibility_desc": "",
                    "audience_desc":   "",
                    "landmark_1":      "",
                    "landmark_2":      "",
                    "landmark_3":      "",
                }

            # Fetch map screenshot
            map_address = market if is_mobile else location
            map_zoom    = 11 if is_mobile else 16
            print(f"[MAP] fetching for site {idx} '{site_name}': address={map_address!r}, city={market!r}, zoom={map_zoom}")
            map_bytes   = get_map_image_bytes(map_address, market, zoom=map_zoom)
            has_map     = False
            if map_bytes:
                map_file = UPLOAD_FOLDER / f"{job_id}_map_{idx}.png"
                map_file.write_bytes(map_bytes)
                with jobs_lock:
                    jobs[job_id]["map_paths"][idx] = str(map_file)
                has_map = True
                print(f"[MAP] saved {len(map_bytes)} bytes for site {idx} '{site_name}'")
            else:
                print(f"[MAP] no map returned for site {idx} '{site_name}'")

            plan.append({
                "site_name":       site_name,
                "market":          market,
                "location":        location,
                "format":          str(row.get("Format", "")).strip(),
                "size":            str(row.get("Size", "")).strip(),
                "units":           str(row.get("Units/Faces", "")).strip(),
                "frequency":       frequency,
                "traffic":         traffic,
                "tagline":         ai.get("tagline", ""),
                "location_desc":   ai.get("location_desc", ""),
                "visibility_desc": ai.get("visibility_desc", ""),
                "audience_desc":   ai.get("audience_desc", ""),
                "landmark_1":      ai.get("landmark_1", ""),
                "landmark_2":      ai.get("landmark_2", ""),
                "landmark_3":      ai.get("landmark_3", ""),
                "has_map":         has_map,
            })

        with jobs_lock:
            jobs[job_id]["status"]   = "plan_ready"
            jobs[job_id]["message"]  = f"Plan ready — {total} site(s). Review and edit, then build."
            jobs[job_id]["progress"] = 100
            jobs[job_id]["plan"]     = plan

    except Exception as exc:
        with jobs_lock:
            jobs[job_id]["status"]   = "error"
            jobs[job_id]["message"]  = f"Error: {exc}"
            jobs[job_id]["progress"] = 0
        print(traceback.format_exc())

    finally:
        try:
            excel_path.unlink(missing_ok=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Legacy one-shot worker (kept for /api/generate backward compat)
# ---------------------------------------------------------------------------

def process_job(job_id: str, excel_path: Path, pptx_path: Path):
    """Upload + AI + build in one shot (legacy endpoint)."""
    def update(status: str, message: str, progress: int = 0):
        with jobs_lock:
            jobs[job_id]["status"]   = status
            jobs[job_id]["message"]  = message
            jobs[job_id]["progress"] = progress

    try:
        update("processing", "Reading Excel file…", 5)
        df = pd.read_excel(excel_path, engine="openpyxl")
        df.columns = [c.strip() for c in df.columns]

        if "Market" in df.columns:
            df["Market"] = df["Market"].ffill()

        if "Site Name" not in df.columns:
            raise ValueError("Excel file must have a 'Site Name' column.")
        df = df[df["Site Name"].notna() & (df["Site Name"].astype(str).str.strip() != "")]
        df = df.reset_index(drop=True)

        if df.empty:
            raise ValueError("No valid site rows found in the Excel file.")

        total = len(df)
        update("processing", f"Found {total} site(s). Loading template…", 10)

        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable is not set.")
        client = anthropic.Anthropic(api_key=api_key)

        plan = []
        for idx, row in df.iterrows():
            pct = 10 + int(((idx + 1) / total) * 80)
            site_name = str(row.get("Site Name", "")).strip()
            update("processing", f"Processing {idx + 1}/{total}: {site_name}…", pct)

            spot_dur = str(row.get("Spot Duration", "")).strip()
            sov_loop = str(row.get("SOV/Loop", "")).strip()
            frequency = sov_loop if spot_dur.lower() in ("", "nan", "n/a", "na") else f"{spot_dur} {sov_loop}".strip()

            raw_impacts = row.get("Impacts", "")
            try:
                traffic = f"{int(float(str(raw_impacts).replace(',', ''))):,}"
            except (ValueError, TypeError):
                traffic = str(raw_impacts).strip()

            ai = generate_site_content(row.to_dict(), client)

            plan.append({
                "site_name":       site_name,
                "market":          str(row.get("Market", "")).strip(),
                "location":        str(row.get("Location", "")).strip(),
                "format":          str(row.get("Format", "")).strip(),
                "size":            str(row.get("Size", "")).strip(),
                "units":           str(row.get("Units/Faces", "")).strip(),
                "frequency":       frequency,
                "traffic":         traffic,
                "tagline":         ai.get("tagline", ""),
                "location_desc":   ai.get("location_desc", ""),
                "visibility_desc": ai.get("visibility_desc", ""),
                "audience_desc":   ai.get("audience_desc", ""),
                "landmark_1":      ai.get("landmark_1", ""),
                "landmark_2":      ai.get("landmark_2", ""),
                "landmark_3":      ai.get("landmark_3", ""),
            })

        excel_path.unlink(missing_ok=True)
        build_pptx_from_plan(job_id, pptx_path, plan)

    except Exception as exc:
        with jobs_lock:
            jobs[job_id]["status"]   = "error"
            jobs[job_id]["message"]  = f"Error: {exc}"
            jobs[job_id]["progress"] = 0
        print(traceback.format_exc())
        try:
            excel_path.unlink(missing_ok=True)
            pptx_path.unlink(missing_ok=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


# ── Pro flow ────────────────────────────────────────────────────────────────

@app.route("/api/plan", methods=["POST"])
def create_plan():
    """Step 1: Upload files, generate content plan (AI + landmarks)."""
    if "excel" not in request.files or "template" not in request.files:
        return jsonify({"error": "Both 'excel' and 'template' files are required."}), 400

    excel_file    = request.files["excel"]
    template_file = request.files["template"]

    if not excel_file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Excel file must be .xlsx or .xls"}), 400
    if not template_file.filename.endswith(".pptx"):
        return jsonify({"error": "Template file must be .pptx"}), 400

    job_id     = uuid.uuid4().hex
    excel_path = UPLOAD_FOLDER / f"{job_id}_data.xlsx"
    pptx_path  = UPLOAD_FOLDER / f"{job_id}_template.pptx"
    excel_file.save(str(excel_path))
    template_file.save(str(pptx_path))

    with jobs_lock:
        jobs[job_id] = {
            "status":    "planning",
            "message":   "Starting…",
            "progress":  0,
            "plan":      None,
            "map_paths": {},        # {site_idx: str(path)} — server-side only
            "pptx_path": str(pptx_path),
            "output":    None,
        }

    threading.Thread(target=generate_plan_job, args=(job_id, excel_path), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/api/build", methods=["POST"])
def build():
    """Step 2: Submit (possibly edited) plan → build PPTX."""
    data = request.get_json(force=True, silent=True) or {}
    job_id = data.get("job_id")
    plan   = data.get("plan")

    if not job_id or not plan:
        return jsonify({"error": "job_id and plan are required"}), 400

    with jobs_lock:
        job = jobs.get(job_id)

    if not job:
        return jsonify({"error": "Job not found — session may have expired."}), 404

    pptx_path = Path(job.get("pptx_path", ""))
    if not pptx_path.exists():
        return jsonify({"error": "Template file not found. Please re-upload and start again."}), 404

    with jobs_lock:
        jobs[job_id]["status"]   = "building"
        jobs[job_id]["message"]  = "Starting build…"
        jobs[job_id]["progress"] = 0
        jobs[job_id]["output"]   = None

    threading.Thread(target=build_pptx_from_plan, args=(job_id, pptx_path, plan), daemon=True).start()
    return jsonify({"job_id": job_id})


# ── Shared polling + download ────────────────────────────────────────────────

@app.route("/api/status/<job_id>")
def status(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    # Only return fields safe for the frontend (omit pptx_path)
    return jsonify({
        "status":   job["status"],
        "message":  job["message"],
        "progress": job["progress"],
        "plan":     job.get("plan"),
        "output":   job.get("output"),
    })


@app.route("/api/download/<job_id>")
def download(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "File not ready"}), 404

    output_path = OUTPUT_FOLDER / job["output"]
    if not output_path.exists():
        return jsonify({"error": "Output file missing"}), 404

    response = send_file(
        str(output_path),
        as_attachment=True,
        download_name=job["output"],
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    @response.call_on_close
    def _cleanup():
        try:
            output_path.unlink(missing_ok=True)
        except Exception:
            pass
        with jobs_lock:
            jobs.pop(job_id, None)

    return response


# ── Map image preview (served to the plan review UI) ────────────────────────

@app.route("/api/map/<job_id>/<int:site_idx>")
def serve_map(job_id: str, site_idx: int):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return "", 404
    map_path_str = job.get("map_paths", {}).get(site_idx)
    if not map_path_str:
        return "", 404
    p = Path(map_path_str)
    if not p.exists():
        return "", 404
    return send_file(str(p), mimetype="image/png")


# ── Legacy one-shot endpoint (backward compat) ───────────────────────────────

@app.route("/api/generate", methods=["POST"])
def generate():
    if "excel" not in request.files or "template" not in request.files:
        return jsonify({"error": "Both 'excel' and 'template' files are required."}), 400

    excel_file    = request.files["excel"]
    template_file = request.files["template"]

    if not excel_file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Excel file must be .xlsx or .xls"}), 400
    if not template_file.filename.endswith(".pptx"):
        return jsonify({"error": "Template file must be .pptx"}), 400

    job_id     = uuid.uuid4().hex
    excel_path = UPLOAD_FOLDER / f"{job_id}_data.xlsx"
    pptx_path  = UPLOAD_FOLDER / f"{job_id}_template.pptx"
    excel_file.save(str(excel_path))
    template_file.save(str(pptx_path))

    with jobs_lock:
        jobs[job_id] = {
            "status":    "queued",
            "message":   "Queued…",
            "progress":  0,
            "plan":      None,
            "pptx_path": str(pptx_path),
            "output":    None,
        }

    threading.Thread(target=process_job, args=(job_id, excel_path, pptx_path), daemon=True).start()
    return jsonify({"job_id": job_id})


# ---------------------------------------------------------------------------
# CN Print Plan Filler
# ---------------------------------------------------------------------------
# Fills the bundled Conde Nast print plan template (assets/) from up to three
# publisher rate cards (India / UK / US). All extraction and template surgery
# lives in cn_print_plan.py; see that module for the section geometry.

from cn_print_plan import fill_cn_print_plan, CNPlanError

CN_TEMPLATE_PATH = BASE_DIR / "assets" / "CN_Print_Plan_Template_EMPTY.xlsx"


@app.route('/print-plan')
def print_plan_page():
    return render_template('print_plan.html')


@app.route('/fill-cn-plan', methods=['POST'])
def fill_cn_plan():
    cards = {}
    for mkt in ('india', 'uk', 'us'):
        f = request.files.get(mkt)
        if f and f.filename:
            cards[mkt] = f.read()
    if not cards:
        return 'Upload at least one rate card (India, UK or US)', 400

    tpl = request.files.get('template')
    template_source = io.BytesIO(tpl.read()) if tpl and tpl.filename else CN_TEMPLATE_PATH
    header = {k: (request.form.get(k) or '').strip()
              for k in ('client', 'campaign', 'agency', 'contact')}

    try:
        out = fill_cn_print_plan(template_source, cards, header)
    except CNPlanError as e:
        return str(e), 400
    except Exception:
        print(traceback.format_exc())
        return 'Failed to fill the plan — check the server logs', 500

    campaign = re.sub(r'[^A-Za-z0-9 _\-]', '', header.get('campaign', '')).strip()
    fname = (campaign.replace(' ', '_') + '_Print_Plan.xlsx') if campaign else 'CN_Print_Plan_Filled.xlsx'
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------------------------------------------------------------------------
# Etihad OOH Proposal Builder
# ---------------------------------------------------------------------------
# Fills the Etihad-branded template from the vendor-supplied Excel media plan
# (one row per site: State/Market/Location/Contacts Per Day/Media/Type/Units/
# W/H/SOV %/Lead Time/Language/Screen Resolution/Screen Loop). The template's
# data slide (shape IDs 138-146) is cloned once per site, and a divider slide
# (shape IDs 169-171) is cloned once per market/city, in first-seen order.
# Since the Excel has no vendor photos, the "Pictures" placeholder is dropped
# and "Map" is filled with a geocoded Google Maps screenshot (same lookup the
# main OOH generator uses) — fetching one map per site is network-bound, so
# the build runs as a background job polled via the existing /api/status and
# /api/download endpoints rather than a single synchronous request.

_ETIHAD_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_ETIHAD_NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_ETIHAD_REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
_ETIHAD_LINK_H = 324000   # EMU height of the "View on Google Maps" text box
_ETIHAD_GAP    = 36000    # EMU gap between map image and hyperlink text box

# Header text -> internal field name. Matched case-insensitively against the
# Excel's header row, wherever that row happens to land (there's a Date/
# Client/Campaign block above it).
_ETIHAD_HEADER_KEYWORDS = {
    'sno':        ['s. no.', 's.no.', 's no', 'sno'],
    'state':      ['state'],
    'market':     ['market'],
    'location':   ['location'],
    'audience':   ['contacts per day'],
    'media':      ['media'],
    'type':       ['type'],
    'units':      ['units'],
    'w':          ['w'],
    'h':          ['h'],
    'sov':        ['sov %', 'sov%'],
    'lead_time':  ['lead time'],
    'language':   ['language'],
    'resolution': ['screen resolution (in pixels)', 'screen resolution'],
    'loop':       ['screen loop'],
}
# These columns get merged across multi-row site clusters in the source sheet
# (e.g. two screens at the same address) — blank cells inherit the last value.
_ETIHAD_FORWARD_FILL_FIELDS = ('state', 'market', 'location', 'audience', 'sov')


def _etihad_find_header_row(rows: list) -> int:
    for i, row in enumerate(rows[:20]):
        vals = {str(v).strip().lower() for v in row if v is not None and str(v).strip()}
        if 'location' in vals and 'market' in vals:
            return i
    raise ValueError("Could not find the header row (expected 'Location' and 'Market' columns) in the Excel file.")


def _etihad_parse_excel_sites(excel_bytes: bytes) -> list:
    """Parse the vendor media-plan Excel into a flat list of site dicts, one per data row."""
    import openpyxl as opx
    wb = opx.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    raw = [[c.value for c in row] for row in ws.iter_rows()]
    if not raw:
        raise ValueError("The Excel file is empty.")

    hdr_idx = _etihad_find_header_row(raw)
    headers = [str(h).strip().lower() if h is not None else '' for h in raw[hdr_idx]]

    field_col = {}
    for field, keywords in _ETIHAD_HEADER_KEYWORDS.items():
        for ci, h in enumerate(headers):
            if h in keywords:
                field_col[field] = ci
                break

    if 'location' not in field_col or 'market' not in field_col:
        raise ValueError("Could not find 'Location' and 'Market' columns in the Excel file.")

    def cell(row, field):
        ci = field_col.get(field)
        if ci is None or ci >= len(row) or row[ci] is None:
            return ''
        return str(row[ci]).strip()

    sites = []
    last = {f: '' for f in _ETIHAD_FORWARD_FILL_FIELDS}
    sno_counter = 0

    for row in raw[hdr_idx + 1:]:
        if not cell(row, 'sno'):
            continue  # blank separator row or a "<City> Total" subtotal row — neither has a serial number

        values = {}
        for field in _ETIHAD_FORWARD_FILL_FIELDS:
            v = cell(row, field)
            values[field] = v or last[field]
            last[field] = values[field]

        sno_counter += 1
        sites.append({
            'sno':        sno_counter,
            'state':      values['state'],
            'market':     values['market'] or 'Other',
            'location':   values['location'],
            'audience':   values['audience'],
            'media':      cell(row, 'media'),
            'type':       cell(row, 'type'),
            'units':      cell(row, 'units') or '1',
            'w':          cell(row, 'w'),
            'h':          cell(row, 'h'),
            'sov':        values['sov'],
            'lead_time':  cell(row, 'lead_time'),
            'language':   cell(row, 'language') or 'English',
            'resolution': cell(row, 'resolution'),
            'loop':       cell(row, 'loop'),
        })

    if not sites:
        raise ValueError("No site rows found under the header row.")
    return sites


def _etihad_format_site_fields(site: dict) -> dict:
    """Derive the display strings that go on the Etihad slide from raw Excel fields."""
    audience_digits = site['audience'].replace(',', '')
    try:
        audience_fmt = f"{int(float(audience_digits)):,} daily contacts" if audience_digits else ''
    except ValueError:
        audience_fmt = site['audience']

    size_fmt = f"{site['w']} x {site['h']} ft" if site['w'] and site['h'] else ''

    sov_fmt = ''
    if site['sov']:
        try:
            sov_fmt = f"{float(site['sov']):.1f}% SOV"
        except ValueError:
            sov_fmt = site['sov']

    spot_duration = ''
    m = re.search(r'(\d+)\s*sec', site['loop'], re.IGNORECASE)
    if m:
        spot_duration = m.group(1) + ' seconds'

    visibility = ', '.join(p for p in (site['media'], site['type'], site['resolution']) if p)

    return {
        'title':         f"{site['market'].upper()} — SITE {site['sno']:02d}",
        'location':      site['location'][:300],
        'visibility':    visibility[:150],
        'audience':      audience_fmt,
        'size':          size_fmt,
        'units':         site['units'],
        'spot_duration': spot_duration,
        'format':        site['type'] or site['media'],
        'sov':           sov_fmt,
        'language':      site['language'],
        'deadline':      site['lead_time'],
    }


def _etihad_clone_slide(prs: Presentation, idx: int):
    tpl = prs.slides[idx]
    new = prs.slides.add_slide(tpl.slide_layout)
    sp = new.shapes._spTree
    sp.clear()
    for el in tpl.shapes._spTree:
        sp.append(copy.deepcopy(el))
    return new


def _etihad_make_run(text: str, sz: int = 900, color: str = '1A1A1A', bold: bool = False):
    rPr = etree.Element(f'{{{_ETIHAD_NS_A}}}rPr', attrib={'lang': 'en-US', 'dirty': '0'})
    if bold:
        rPr.set('b', '1')
    rPr.set('sz', str(sz))
    sf = etree.SubElement(rPr, f'{{{_ETIHAD_NS_A}}}solidFill')
    sc = etree.SubElement(sf, f'{{{_ETIHAD_NS_A}}}srgbClr')
    sc.set('val', color)
    lt = etree.SubElement(rPr, f'{{{_ETIHAD_NS_A}}}latin')
    lt.set('typeface', 'Calibri')
    r = etree.Element(f'{{{_ETIHAD_NS_A}}}r')
    r.insert(0, rPr)
    t = etree.SubElement(r, f'{{{_ETIHAD_NS_A}}}t')
    t.text = text
    return r


def _etihad_clear_runs(p_elem):
    for r in list(p_elem.findall(f'{{{_ETIHAD_NS_A}}}r')):
        p_elem.remove(r)
    for b in list(p_elem.findall(f'{{{_ETIHAD_NS_A}}}br')):
        p_elem.remove(b)


def _etihad_set_para(tf, idx: int, text: str, sz: int = 900, color: str = '1A1A1A', bold: bool = False):
    """Replace paragraph idx content with a single run."""
    try:
        p = tf.paragraphs[idx]._p
    except IndexError:
        return
    _etihad_clear_runs(p)
    if text:
        p.append(_etihad_make_run(text, sz=sz, color=color, bold=bold))


def _etihad_append_value(tf, idx: int, value: str, sz: int = 900):
    """Append a dark value run to paragraph idx, keeping existing label runs."""
    if not value or not str(value).strip():
        return
    try:
        p = tf.paragraphs[idx]._p
    except IndexError:
        return
    p.append(_etihad_make_run('  ' + value, sz=sz, color='1A1A1A', bold=False))


def _etihad_fill_data_slide(slide, f: dict):
    """Fill all text fields on a cloned Etihad data slide from formatted site fields."""
    for sh in slide.shapes:
        sid = sh.shape_id
        try:
            tf = sh.text_frame
        except Exception:
            continue

        if sid == 138:    # Site title
            _etihad_set_para(tf, 0, f['title'], sz=2000, color='898989', bold=True)
        elif sid == 139:  # Location / Visibility / Audience
            _etihad_set_para(tf, 2, f['location'], sz=950)
            _etihad_set_para(tf, 5, f['visibility'], sz=950)
            _etihad_set_para(tf, 7, f['audience'], sz=950)
        elif sid == 140:  # Size / Period / Nb Units / Spot duration
            _etihad_append_value(tf, 0, f['size'], sz=950)
            _etihad_append_value(tf, 4, f['units'], sz=950)
            _etihad_append_value(tf, 6, f['spot_duration'], sz=950)
        elif sid == 141:  # Format / SOV / Traffic / Specs
            _etihad_append_value(tf, 0, f['format'], sz=950)
            _etihad_append_value(tf, 2, f['sov'], sz=950)
            _etihad_append_value(tf, 4, f['audience'], sz=950)
            _etihad_append_value(tf, 5, f['size'], sz=950)
        elif sid == 142:  # Nearby location (no data source — left blank) / Lead Language / Material Deadline
            _etihad_append_value(tf, 3, f['language'], sz=950)
            _etihad_append_value(tf, 5, f['deadline'], sz=950)


def _etihad_insert_map_with_link(slide, map_stream, ML, MT, MW, MH, maps_url: str):
    """Insert map image, leaving room below for a 'View on Google Maps' hyperlink."""
    ih = MH - _ETIHAD_LINK_H - _ETIHAD_GAP
    slide.shapes.add_picture(map_stream, ML, MT, MW, ih)

    tb = slide.shapes.add_textbox(ML, MT + ih + _ETIHAD_GAP, MW, _ETIHAD_LINK_H)
    tf2 = tb.text_frame
    tf2.word_wrap = False
    p = tf2.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = '📍 View on Google Maps'
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x3B, 0x7D, 0xC4)
    run.font.underline = True

    if maps_url:
        rId = slide.part.relate_to(maps_url, _ETIHAD_REL_HYPERLINK, is_external=True)
        rPr = run._r.get_or_add_rPr()
        hl = etree.SubElement(rPr, f'{{{_ETIHAD_NS_A}}}hlinkClick')
        hl.set(f'{{{_ETIHAD_NS_R}}}id', rId)


def _etihad_insert_map_only(slide, map_bytes: bytes | None, maps_url: str):
    """Drop the 'Pictures' placeholder (no vendor photo available) and fill 'Map' with a geocoded screenshot."""
    pics_sh = map_sh = None
    ML = MT = MW = MH = None

    for sh in slide.shapes:
        try:
            txt = sh.text_frame.text.strip()
        except Exception:
            continue
        if txt == 'Pictures':
            pics_sh = sh
        elif txt in ('Map', 'Map '):
            map_sh = sh
            ML, MT, MW, MH = sh.left, sh.top, sh.width, sh.height

    spt = slide.shapes._spTree
    if pics_sh is not None:
        spt.remove(pics_sh._element)
    if map_sh is not None:
        spt.remove(map_sh._element)
        if map_bytes:
            _etihad_insert_map_with_link(slide, io.BytesIO(map_bytes), ML, MT, MW, MH, maps_url)


def _etihad_maps_search_url(location: str, market: str) -> str:
    from urllib.parse import quote
    return f"https://www.google.com/maps/search/?api=1&query={quote(f'{location}, {market}')}"


def _etihad_build_deck(prs: Presentation, sites: list, map_lookup: dict):
    """Clone the data template slide once per site and a divider slide once per market."""
    markets_in_order = list(dict.fromkeys(s['market'] for s in sites))

    for market in markets_in_order:
        divider = _etihad_clone_slide(prs, 2)
        for sh in divider.shapes:
            if sh.shape_id in (169, 170, 171):
                try:
                    _etihad_set_para(sh.text_frame, 0, market.upper(), sz=3600, color='1A1A1A', bold=True)
                    break
                except Exception:
                    pass

        for site in (s for s in sites if s['market'] == market):
            slide = _etihad_clone_slide(prs, 1)
            _etihad_fill_data_slide(slide, _etihad_format_site_fields(site))
            map_bytes, maps_url = map_lookup.get(site['sno'], (None, ''))
            _etihad_insert_map_only(slide, map_bytes, maps_url)

    # Remove the original template slides 1 & 2, keep the title slide (0)
    sll = prs.slides._sldIdLst
    pp = prs.part
    for idx in [2, 1]:
        elem = list(sll)[idx]
        rId = elem.get(f'{{{_ETIHAD_NS_R}}}id')
        sll.remove(elem)
        if rId:
            try:
                pp._rels.pop(rId)
            except Exception:
                pass


def _etihad_build_job(job_id: str, template_bytes: bytes, sites: list):
    """Background job: fetch a map screenshot per site, then assemble the deck."""
    def update(status, message, progress=0):
        with jobs_lock:
            jobs[job_id]['status'] = status
            jobs[job_id]['message'] = message
            jobs[job_id]['progress'] = progress

    try:
        total = len(sites)
        map_lookup = {}
        for idx, site in enumerate(sites):
            pct = int((idx / total) * 90)
            update('building', f"Fetching map {idx + 1}/{total}: {site['location'][:60]}…", pct)
            map_bytes = None
            try:
                map_bytes = get_map_image_bytes(site['location'], site['market'], zoom=16)
                if not map_bytes:
                    # The full location text (e.g. "Set of 2 - Bandra Kurla Complex - Entrance")
                    # is often too descriptive for the geocoder — fall back to a city-level map
                    # rather than leaving the placeholder blank.
                    map_bytes = get_map_image_bytes(site['market'], site['market'], zoom=12)
            except Exception as e:
                print(f"[ETIHAD MAP] failed for site {site['sno']} ({site['location']!r}): {e}")
            map_lookup[site['sno']] = (map_bytes, _etihad_maps_search_url(site['location'], site['market']))

        update('building', 'Assembling slides…', 92)
        prs = Presentation(io.BytesIO(template_bytes))
        _etihad_build_deck(prs, sites, map_lookup)

        update('building', 'Saving file…', 97)
        output_filename = f"Etihad_OOH_Proposal_{job_id[:8]}.pptx"
        output_path = OUTPUT_FOLDER / output_filename
        prs.save(str(output_path))

        with jobs_lock:
            jobs[job_id]['status']   = 'done'
            jobs[job_id]['message']  = f"Done! {total} site slide(s) generated."
            jobs[job_id]['progress'] = 100
            jobs[job_id]['output']   = output_filename

    except Exception as exc:
        with jobs_lock:
            jobs[job_id]['status']   = 'error'
            jobs[job_id]['message']  = f"Error: {exc}"
            jobs[job_id]['progress'] = 0
        print(traceback.format_exc())


@app.route('/etihad-ooh')
def etihad_ooh_page():
    return render_template('etihad_ooh.html')


@app.route('/build-etihad-proposal', methods=['POST'])
def build_etihad_proposal_route():
    """Step 1: upload the vendor Excel plan + Etihad template, parse the Excel, start the build job."""
    excel_file    = request.files.get('excel')
    template_file = request.files.get('template')

    if not excel_file or not excel_file.filename:
        return jsonify({'error': 'Excel site-plan file is required.'}), 400
    if not template_file or not template_file.filename:
        return jsonify({'error': 'Etihad template file is required.'}), 400
    if not excel_file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Site plan must be .xlsx or .xls'}), 400
    if not template_file.filename.lower().endswith('.pptx'):
        return jsonify({'error': 'Etihad template must be a .pptx'}), 400

    try:
        sites = _etihad_parse_excel_sites(excel_file.read())
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    template_bytes = template_file.read()
    job_id = uuid.uuid4().hex
    with jobs_lock:
        jobs[job_id] = {
            'status':   'building',
            'message':  f"Found {len(sites)} site(s). Starting…",
            'progress': 0,
            'plan':     None,
            'output':   None,
        }

    threading.Thread(target=_etihad_build_job, args=(job_id, template_bytes, sites), daemon=True).start()
    return jsonify({'job_id': job_id, 'site_count': len(sites)})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
