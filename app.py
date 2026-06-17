import os, re, json, threading, requests
from pathlib import Path
from flask import Flask, jsonify, send_file, Response
from flask_cors import CORS
import openpyxl

app = Flask(__name__)
CORS(app)

DROPBOX_EXCEL_LINK    = os.environ.get('DROPBOX_EXCEL_LINK', '')
DROPBOX_ACCESS_TOKEN  = os.environ.get('DROPBOX_ACCESS_TOKEN', '')
DROPBOX_IMAGES_FOLDER = os.environ.get('DROPBOX_IMAGES_FOLDER', '/A (Stripe CVC) 20-May-26')

BASE      = Path(__file__).parent
XLSX_PATH = BASE / 'Raw_Data.xlsx'
HTML_PATH = BASE / 'Fabric_Catalogue.html'

FABRIC_DATA = []
IMAGE_MAP   = {}
READY       = False

def dropbox_direct(url):
    url = url.strip()
    url = re.sub(r'[?&]dl=\d', lambda m: m.group(0)[:-1] + '1', url)
    if 'dl=' not in url:
        url += ('&' if '?' in url else '?') + 'dl=1'
    url = re.sub(r'[?&]st=[^&]+', '', url)
    url = re.sub(r'[?&]e=\d+', '', url)
    url = url.replace('www.dropbox.com', 'dl.dropboxusercontent.com')
    return url

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
        if code.startswith(prefix): return vals
    return ('Other', 'Other', '')

def to_float(v):
    try: return round(float(v or 0), 2)
    except: return 0.0

def dropbox_list_all_images():
    if not DROPBOX_ACCESS_TOKEN:
        print("No Dropbox access token set")
        return {}

    headers = {
        'Authorization': f'Bearer {DROPBOX_ACCESS_TOKEN}',
        'Content-Type': 'application/json',
    }
    img_exts = {'.jpg', '.jpeg', '.png', '.webp', '.gif'}
    image_map = {}

    url = 'https://api.dropboxapi.com/2/files/list_folder'
    payload = {
        'path': DROPBOX_IMAGES_FOLDER,
        'recursive': True,
        'limit': 2000,
    }

    try:
        while True:
            r = requests.post(url, headers=headers, json=payload, timeout=60)
            r.raise_for_status()
            data = r.json()
            for entry in data.get('entries', []):
                if entry.get('.tag') != 'file':
                    continue
                name = entry['name']
                p = Path(name)
                if p.suffix.lower() not in img_exts:
                    continue
                stem = p.stem
                while True:
                    s2, e2 = os.path.splitext(stem)
                    if e2.lower() in img_exts: stem = s2
                    else: break
                key = stem.upper().strip()
                if key not in image_map:
                    image_map[key] = entry['path_lower']
            if data.get('has_more'):
                url = 'https://api.dropboxapi.com/2/files/list_folder/continue'
                payload = {'cursor': data['cursor']}
            else:
                break
        print(f"Dropbox API: found {len(image_map)} images")
    except Exception as e:
        print(f"Dropbox list error: {e}")
    return image_map

def read_fabric_data(image_map):
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
        img_path = ''
        if key in image_map:
            img_path = '/api/img/' + key
        else:
            base_key = re.sub(r'-\d+$', '', key)
            if base_key in image_map:
                img_path = '/api/img/' + base_key
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
            'imagePath': img_path,
        })
    return records

def background_startup():
    global FABRIC_DATA, IMAGE_MAP, READY
    try:
        print("Downloading Excel...")
        r = requests.get(dropbox_direct(DROPBOX_EXCEL_LINK), timeout=60)
        r.raise_for_status()
        XLSX_PATH.write_bytes(r.content)
        print(f"Excel: {len(r.content):,} bytes")

        print("Listing images from Dropbox API...")
        IMAGE_MAP = dropbox_list_all_images()

        FABRIC_DATA = read_fabric_data(IMAGE_MAP)
        matched = sum(1 for r in FABRIC_DATA if r['imagePath'])
        print(f"Ready! {len(FABRIC_DATA)} records, {matched} with images")
    except Exception as e:
        print(f"Startup error: {e}")
    finally:
        READY = True

threading.Thread(target=background_startup, daemon=True).start()

@app.route('/')
@app.route('/index.html')
def index():
    return send_file(str(HTML_PATH), mimetype='text/html')

@app.route('/api/data')
def api_data():
    return jsonify(FABRIC_DATA)

@app.route('/api/img/<key>')
def serve_img(key):
    path = IMAGE_MAP.get(key.upper())
    if not path:
        return '', 404
    try:
        headers = {
            'Authorization': f'Bearer {DROPBOX_ACCESS_TOKEN}',
            'Dropbox-API-Arg': json.dumps({'path': path}),
        }
        r = requests.post(
            'https://content.dropboxapi.com/2/files/download',
            headers=headers, timeout=30
        )
        r.raise_for_status()
        ext = Path(path).suffix.lower()
        mime = {'.jpg':'image/jpeg','.jpeg':'image/jpeg','.png':'image/png','.webp':'image/webp','.gif':'image/gif'}.get(ext,'image/jpeg')
        return Response(r.content, mimetype=mime, headers={'Cache-Control': 'public, max-age=86400'})
    except Exception as e:
        print(f"Image fetch error for {key}: {e}")
        return '', 500

@app.route('/api/status')
def api_status():
    matched = sum(1 for r in FABRIC_DATA if r['imagePath'])
    return jsonify({'ready': READY, 'records': len(FABRIC_DATA), 'images_total': len(IMAGE_MAP), 'images_matched': matched})

@app.route('/api/refresh')
def api_refresh():
    global READY
    READY = False
    threading.Thread(target=background_startup, daemon=True).start()
    return jsonify({'status': 'refreshing'})
