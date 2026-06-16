#!/usr/bin/env python3
"""
Threads & Sew — Fabric Catalogue
Render.com deployment — runs 24/7
"""

import os, re, json, zipfile, shutil, time, threading, mimetypes, requests
from pathlib import Path
from flask import Flask, jsonify, send_file, abort
from flask_cors import CORS
import openpyxl

app = Flask(__name__)
CORS(app)

# ── CONFIG — set these in Render Environment Variables ──────────
DROPBOX_EXCEL_LINK  = os.environ.get('DROPBOX_EXCEL_LINK',  '')
DROPBOX_IMAGES_LINK = os.environ.get('DROPBOX_IMAGES_LINK', '')

BASE        = Path(__file__).parent
XLSX_PATH   = BASE / 'data' / 'Raw_Data.xlsx'
IMAGES_DIR  = BASE / 'data' / 'images'
HTML_PATH   = BASE / 'Fabric_Catalogue.html'

FABRIC_DATA = []
IMAGE_MAP   = {}

# ── Dropbox URL fixer ────────────────────────────────────────────
def dropbox_direct(url):
    url = url.strip()
    url = re.sub(r'[?&]dl=\d', lambda m: m.group(0)[:-1] + '1', url)
    if 'dl=' not in url:
        url += ('&' if '?' in url else '?') + 'dl=1'
    url = re.sub(r'[?&]st=[^&]+', '', url)
    url = url.replace('www.dropbox.com', 'dl.dropboxusercontent.com')
    return url

# ── Prefix map ───────────────────────────────────────────────────
PREFIX_MAP = [
    ('SCC', ('Self Check Cotton',  '100% Cotton', 'Check'  )),
    ('SCP', ('Self Plain Cotton',  '100% Cotton', 'Plain'  )),
    ('SCS', ('Self Stripe Cotton', '100% Cotton', 'Stripe' )),
    ('SWC', ('Self Check Woven',   '100% Cotton', 'Check'  )),
    ('SWP', ('Self Plain Woven',   '100% Cotton', 'Plain'  )),
    ('SWS', ('Self Stripe Woven',  '100% Cotton', 'Stripe' )),
    ('PR',  ('Printed 100% Cotton','100% Cotton', 'Printed')),
    ('A',   ('Stripe CVC',         'CVC',         'Stripe' )),
    ('B',   ('Check CVC',          'CVC',         'Check'  )),
    ('C',   ('100% Cotton Stripe', '100% Cotton', 'Stripe' )),
    ('D',   ('100% Cotton Plain',  '100% Cotton', 'Plain'  )),
    ('E',   ('100% Cotton Check',  '100% Cotton', 'Check'  )),
    ('F',   ('Plain CVC',          'CVC',         'Plain'  )),
]

def derive_meta(design):
    code = str(design).upper().strip()
    for prefix, vals in PREFIX_MAP:
        if code.startswith(prefix):
            return vals
    return ('Other', 'Other', '')

def to_float(v):
    try: return round(float(v or 0), 2)
    except: return 0.0

# ── Download Excel ───────────────────────────────────────────────
def download_excel():
    global FABRIC_DATA
    if not DROPBOX_EXCEL_LINK:
        print("⚠️  DROPBOX_EXCEL_LINK not set")
        return
    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    print("Downloading Excel from Dropbox...")
    r = requests.get(dropbox_direct(DROPBOX_EXCEL_LINK), timeout=60)
    r.raise_for_status()
    XLSX_PATH.write_bytes(r.content)
    print(f"✅ Excel downloaded: {len(r.content):,} bytes")
    FABRIC_DATA = read_fabric_data()
    match_images()
    print(f"✅ Loaded {len(FABRIC_DATA):,} records")

# ── Download Images ──────────────────────────────────────────────
def download_images():
    global IMAGE_MAP
    if not DROPBOX_IMAGES_LINK:
        print("⚠️  DROPBOX_IMAGES_LINK not set — no images")
        return
    IMAGES_ZIP = BASE / 'data' / 'images.zip'
    IMAGES_DIR.parent.mkdir(parents=True, exist_ok=True)
    print("Downloading images from Dropbox...")
    for attempt in range(3):
        try:
            r = requests.get(DROPBOX_IMAGES_LINK, timeout=300, stream=True)
            r.raise_for_status()
            total = 0
            with open(IMAGES_ZIP, 'wb') as f:
                for chunk in r.iter_content(chunk_size=65536):
                    f.write(chunk)
                    total += len(chunk)
            print(f"✅ Images downloaded: {total//1024//1024} MB")
            break
        except Exception as e:
            print(f"Attempt {attempt+1} failed: {e}")
            if attempt < 2: time.sleep(5)
            else: return

    if IMAGES_DIR.exists(): shutil.rmtree(IMAGES_DIR)
    IMAGES_DIR.mkdir()
    with zipfile.ZipFile(IMAGES_ZIP, 'r') as z:
        z.extractall(IMAGES_DIR)
    os.remove(IMAGES_ZIP)

    img_exts = {'.jpg','.jpeg','.png','.webp','.gif'}
    all_images = [f for f in IMAGES_DIR.rglob('*') if f.suffix.lower() in img_exts]
    IMAGE_MAP = {}
    for img in all_images:
        stem = img.stem
        while True:
            s2, e2 = os.path.splitext(stem)
            if e2.lower() in img_exts: stem = s2
            else: break
        key = stem.upper().strip()
        if key not in IMAGE_MAP:
            IMAGE_MAP[key] = str(img.relative_to(IMAGES_DIR))
    print(f"✅ Image map: {len(IMAGE_MAP):,} images")
    match_images()

# ── Read Excel ───────────────────────────────────────────────────
def read_fabric_data():
    if not XLSX_PATH.exists(): return []
    wb = openpyxl.load_workbook(str(XLSX_PATH), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return []
    headers = [str(h).strip() if h is not None else '' for h in rows[0]]
    col = {h.lower(): i for i, h in enumerate(headers)}
    def get(row, name, default=''):
        idx = col.get(name.lower())
        if idx is None or idx >= len(row): return default
        v = row[idx]
        return v if v is not None else default
    seen, records = set(), []
    for row in rows[1:]:
        raw_design = str(get(row, 'Design', '')).strip()
        if not raw_design or raw_design.lower() in ('none', 'nan'): continue
        key = raw_design.upper()
        if key in seen: continue
        seen.add(key)
        category, content, pattern = derive_meta(raw_design)
        records.append({
            'design': raw_design,
            'style': str(get(row, 'Style', '') or '').strip(),
            'color': str(get(row, 'Color', '') or '').strip(),
            'salePrice': to_float(get(row, 'Sale_Price')),
            'actualStock': to_float(get(row, 'ActualStock')),
            'availableStock': to_float(get(row, 'AvailableStock')),
            'season': str(get(row, 'Season', '') or '').strip(),
            'brand': str(get(row, 'BrandName', '') or '').strip(),
            'vendor': str(get(row, 'Vendor_Name', '') or '').strip(),
            'category': category, 'content': content, 'pattern': pattern,
            'imagePath': '',
        })
    return records

# ── Match images to fabric records ───────────────────────────────
def match_images():
    matched = 0
    for record in FABRIC_DATA:
        key = record['design'].upper().strip()
        if key in IMAGE_MAP:
            record['imagePath'] = '/img/' + IMAGE_MAP[key]
            matched += 1
        else:
            base_key = re.sub(r'-\d+$', '', key)
            if base_key in IMAGE_MAP:
                record['imagePath'] = '/img/' + IMAGE_MAP[base_key]
                matched += 1
            else:
                record['imagePath'] = ''
    print(f"✅ Matched {matched:,} of {len(FABRIC_DATA):,} fabrics with images")

# ── Auto-refresh every 6 hours ───────────────────────────────────
def auto_refresh():
    while True:
        time.sleep(6 * 60 * 60)  # 6 hours
        print("🔄 Auto-refreshing data from Dropbox...")
        try:
            download_excel()
        except Exception as e:
            print(f"Auto-refresh failed: {e}")

# ── Routes ───────────────────────────────────────────────────────
@app.route('/')
@app.route('/index.html')
def index():
    return send_file(str(HTML_PATH), mimetype='text/html')

@app.route('/api/data')
def api_data():
    return jsonify(FABRIC_DATA)

@app.route('/api/refresh')
def api_refresh():
    try:
        download_excel()
        return jsonify({'status': 'ok', 'count': len(FABRIC_DATA)})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/img/<path:imgpath>')
def serve_image(imgpath):
    candidate = (IMAGES_DIR / imgpath).resolve()
    if str(candidate).startswith(str(IMAGES_DIR.resolve())) and candidate.exists():
        mime = mimetypes.guess_type(str(candidate))[0] or 'application/octet-stream'
        return send_file(str(candidate), mimetype=mime)
    abort(404)

# ── Startup ──────────────────────────────────────────────────────
def startup():
    print("🚀 Starting Threads & Sew Catalogue...")
    download_excel()
    download_images()
    threading.Thread(target=auto_refresh, daemon=True).start()
    print("✅ Ready!")

if __name__ == '__main__':
    startup()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
else:
    # Called by gunicorn
    startup()
