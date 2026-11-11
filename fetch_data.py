# -*- coding: utf-8 -*-
"""從 Apps Script 網址下載 Excel 到 data/，並寫出 config.json。環境變數：APPS_SCRIPT_URL、APPS_SCRIPT_TOKEN"""
import base64, json, os, sys, time, urllib.parse, urllib.request

URL   = os.environ['APPS_SCRIPT_URL']
TOKEN = os.environ['APPS_SCRIPT_TOKEN']
OUT   = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
FILES = {   # key: (存檔名, 必要)
    'reorder':  ('追加單_2026.xlsx', True),
    'purchase': ('採購總表_2026.xlsx', True),
    'transfer': ('調撥出庫明細表_2026.xlsx', True),
    'preorder': ('預購清單_2026.xlsx', False),
    'store':    ('門市資料表.xlsx', True),
}

def call(**params):
    q = urllib.parse.urlencode({'token': TOKEN, **params})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(f"{URL}?{q}", timeout=180) as r:
                d = json.loads(r.read().decode('utf-8'))
            if 'error' in d: raise RuntimeError(d['error'])
            return d
        except Exception as e:
            if attempt == 3: raise
            print(f"  重試 {attempt+1}：{e}"); time.sleep(5 * (attempt + 1))

os.makedirs(OUT, exist_ok=True)
for key, (fname, required) in FILES.items():
    try:
        meta = call(file=key)
        buf = b''.join(base64.b64decode(call(file=key, part=i)['data']) for i in range(meta['parts']))
        if len(buf) != meta['size']: raise RuntimeError(f"大小不符 {len(buf)} != {meta['size']}")
        with open(os.path.join(OUT, fname), 'wb') as f: f.write(buf)
        print(f"OK {key}: {meta['name']} ({meta['size']/1e6:.1f}MB, 更新 {meta['updated']})")
    except Exception as e:
        if required: print(f"ERROR {key}: {e}"); sys.exit(1)
        print(f"略過 {key}: {e}")

with open(os.path.join(os.path.dirname(OUT), 'config.json'), 'w', encoding='utf-8') as f:
    json.dump({'data_dir': OUT, 'store_file': os.path.join(OUT, '門市資料表.xlsx'),
               'store_sheet': '門市資料表_底檔'}, f, ensure_ascii=False)
