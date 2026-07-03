"""
追單到貨追蹤系統 - 資料產生器
使用方式：python generate.py
輸出：index.html

資料夾結構（與此腳本同目錄）：
  追加單_2026.xlsx / 採購總表_2026.xlsx / 調撥單_2026.xlsx
  門市基本資料.xlsx / 預購清單_2026.xlsx（可選）
  generate.py / template.html → index.html
"""
import glob, io, json, os, re, struct, sys, zlib, zipfile
from datetime import datetime, timedelta
import pandas as pd

def read_xlsx(path, sheet_name=0, header=0):
    """Read xlsx that may be missing its ZIP central directory."""
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header)
    except Exception:
        pass
    with open(path, 'rb') as f:
        data = f.read()
    positions = [m.start() for m in re.finditer(b'PK\x03\x04', data)]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for i, pos in enumerate(positions):
            fname_len = struct.unpack_from('<H', data, pos+26)[0]
            extra_len = struct.unpack_from('<H', data, pos+28)[0]
            fname = data[pos+30:pos+30+fname_len].decode('utf-8', errors='replace')
            data_start = pos + 30 + fname_len + extra_len
            data_end = positions[i+1] if i+1 < len(positions) else len(data)
            if data_end <= data_start:
                continue
            compressed = data[data_start:data_end]
            dec = zlib.decompressobj(-15)
            try:
                raw = dec.decompress(compressed) + dec.flush()
            except zlib.error:
                raw = zlib.decompressobj(-15).decompress(compressed)
            if not raw:
                continue
            if 'sheet' in fname and fname.endswith('.xml') and not raw.rstrip().endswith(b'</worksheet>'):
                last_row = raw.rfind(b'</row>')
                if last_row > 0:
                    raw = raw[:last_row+6] + b'</sheetData></worksheet>'
            zout.writestr(fname, raw)
    buf.seek(0)
    return pd.read_excel(buf, sheet_name=sheet_name, header=header)

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
OUTPUT_HTML   = os.path.join(BASE_DIR, 'index.html')
TEMPLATE_FILE = os.path.join(BASE_DIR, 'template.html')

DAYS_VENDOR       = 30
DAYS_WH           = 10
DAYS_STORE        = 3
ALERT_OVERDUE_DAY = 40

OUTLET_STORES = {'AS51','AS63','AT08'}
SPECIAL_CATS  = {'泳裝','鞋子','內搭小可愛','內睡衣褲'}
BRAND_ORDER   = ['AS','OL','CA','PLUS','IP','泳裝','鞋子']
SIZE_ORDER    = ['XS','S','S+','M','M+','L','L+','XL','2XL','2L','3XL','3L','F','FF','ONE SIZE']

def find_file(pattern, required=True):
    matches = sorted(glob.glob(os.path.join(BASE_DIR, pattern)), reverse=True)
    if not matches:
        if required: raise FileNotFoundError(f"找不到 {pattern}，請確認資料夾：{BASE_DIR}")
        return None
    return matches[0]

def parse_date(v):
    if not v: return None
    s = re.sub(r'\D','',str(v))[:8]
    if len(s)<8: return None
    try: return datetime.strptime(s,'%Y%m%d')
    except: return None

def add_days(d,n): return (d+timedelta(days=n)) if d else None
def fmt_date(d): return d.strftime('%Y/%m/%d') if d else ''
def ymd(d): return d.strftime('%Y%m%d') if d else ''
def week_label(d): return f"W{int(d.strftime('%W'))}" if d else ''
def size_rank(s):
    try: return SIZE_ORDER.index(str(s).strip())
    except: return 99

print("📂 尋找資料檔案…")
file_reorder  = find_file('追加單_*.xlsx')
file_purchase = find_file('採購總表_*.xlsx')
file_transfer = find_file('調撥單_*.xlsx')
file_store    = find_file('門市基本資料*.xlsx')
file_preorder = find_file('預購清單_*.xlsx', required=False)

for label, f in [('追加單',file_reorder),('採購總表',file_purchase),('調撥單',file_transfer),
                  ('門市資料',file_store),('預購清單',file_preorder or '（無）')]:
    print(f"  {label}：{os.path.basename(f) if f and f!='（無）' else '（無）'}")

print("\n📊 讀取資料…")
df_r = read_xlsx(file_reorder)
df_p = read_xlsx(file_purchase)
df_t = read_xlsx(file_transfer)
df_s_raw = read_xlsx(file_store, sheet_name='店數', header=None)
df_pre = read_xlsx(file_preorder) if file_preorder else None
print(f"  追加單 {len(df_r)} / 採購總表 {len(df_p)} / 調撥單 {len(df_t)}")

# ── 預購款號 ──
preorder_styles = set()
if df_pre is not None:
    preorder_styles = set(df_pre['款號'].dropna().astype(str).str.strip())

# ── 門市對照 + 轄區負責人 ──
hdr = df_s_raw.iloc[0].tolist()
df_s = df_s_raw.iloc[1:].copy(); df_s.columns = hdr
df_s = df_s[df_s['倉庫編號'].notna() & (df_s['區域'] != '已歇業')]
store_map = {}
for _, row in df_s.iterrows():
    code = str(row['倉庫編號']).strip()
    sell = str(row['販售店']).strip() if pd.notna(row['販售店']) else ''
    mgr  = str(row['轄區負責人']).strip() if pd.notna(row['轄區負責人']) else '其他'
    store_map[code] = {
        'name':       str(row['倉庫名稱']).strip() if pd.notna(row['倉庫名稱']) else code,
        'is_outlet':  code in OUTLET_STORES or sell == 'OUTLET',
        'manager':    mgr,
        'district':   str(row['轄區']).strip() if pd.notna(row['轄區']) else '',
    }

# 轄區負責人順序（依門市數量排序）
from collections import Counter
mgr_counts = Counter(v['manager'] for v in store_map.values())
MGR_ORDER = [m for m,_ in mgr_counts.most_common()]

# ── 採購索引 ──
print("🔗 建立採購索引…")
pu_idx = {}
for _, row in df_p[df_p['通路']=='門市'].iterrows():
    rid = str(row['預購單號']).strip() if pd.notna(row['預購單號']) else ''
    bc  = str(row['商品條碼']).strip() if pd.notna(row['商品條碼']) else ''
    if not rid or not bc: continue
    key = f"{rid}|{bc}"
    if key not in pu_idx: pu_idx[key] = {'purchase_date':None,'qty':0,'received':0}
    rec = pu_idx[key]
    dv = row['採購日期']
    if pd.notna(dv):
        ds = re.sub(r'\D','',str(dv))[:8]
        if not rec['purchase_date'] or ds < rec['purchase_date']: rec['purchase_date'] = ds
    rec['qty']      += int(row['數量'])      if pd.notna(row['數量'])      else 0
    rec['received'] += int(row['已驗收數量']) if pd.notna(row['已驗收數量']) else 0

# ── 調撥索引 ──
print("🔗 建立調撥索引…")
tr_idx = {}
for _, row in df_t.iterrows():
    bc = str(row['商品條碼']).strip() if pd.notna(row['商品條碼']) else ''
    if not bc: continue
    if bc not in tr_idx: tr_idx[bc] = []
    dv = row['調出日期']
    ds = re.sub(r'\D','',str(dv))[:8] if pd.notna(dv) else ''
    tr_idx[bc].append({
        'transfer_date': ds,
        'store_code':    str(row['調入倉庫']).strip()   if pd.notna(row['調入倉庫'])   else '',
        'store_name':    str(row['倉庫名稱.1']).strip() if pd.notna(row['倉庫名稱.1']) else '',
        'qty_out': int(row['調出數量']) if pd.notna(row['調出數量']) else 0,
        'qty_in':  int(row['調入數量']) if pd.notna(row['調入數量']) else 0,
    })

# ── 主處理 ──
print("⚙️  資料串接中…")
today = (datetime.utcnow() + __import__('datetime').timedelta(hours=8)).replace(hour=0,minute=0,second=0,microsecond=0)
result = []
has_preorder_col = '預購款' in df_r.columns

for _, row in df_r.iterrows():
    twa = row['TW_A(門市)頭單']
    if pd.isna(twa) or int(twa)<=0: continue

    barcode  = str(row['商品條碼']).strip()  if pd.notna(row['商品條碼'])  else ''
    order_id = str(row['小白單編號']).strip() if pd.notna(row['小白單編號']) else ''
    ptype    = str(row['採購類別']).strip()   if pd.notna(row['採購類別'])  else ''
    is_outlet= (ptype == 'Outlet追加單')

    brand_raw= str(row['品牌']).strip()         if pd.notna(row['品牌'])         else ''
    category = str(row['商品大分類']).strip()   if pd.notna(row['商品大分類'])   else ''
    brand    = category if category in SPECIAL_CATS else ('PLUS' if brand_raw=='A PLUS' else brand_raw)

    name      = str(row['商品名稱']).strip()  if pd.notna(row['商品名稱']) else ''
    is_special= bool(re.search(r'【.*?】', name))
    style_no  = str(row['款號']).strip()      if pd.notna(row['款號'])     else ''
    size      = str(row['尺寸']).strip()      if pd.notna(row['尺寸'])     else ''

    if has_preorder_col:
        pv = row['預購款']
        is_preorder = pd.notna(pv) and str(pv).strip() not in ('','nan')
    else:
        is_preorder = style_no in preorder_styles

    odate_raw      = row['單據日期']
    order_date_str = re.sub(r'\D','',str(odate_raw))[:8] if pd.notna(odate_raw) else ''
    order_date     = parse_date(order_date_str)

    p_key        = f"{order_id}|{barcode}"
    p_rec        = pu_idx.get(p_key)
    has_purchase = p_rec is not None
    purchase_date= parse_date(p_rec['purchase_date']) if has_purchase else None

    all_tr = tr_idx.get(barcode, [])
    if has_purchase and purchase_date:
        cutoff = ymd(purchase_date)
        rel_tr = [t for t in all_tr if t['transfer_date']>=cutoff]
    else:
        cutoff_d = add_days(order_date, DAYS_VENDOR-5)
        cutoff   = ymd(cutoff_d) if cutoff_d else ''
        rel_tr   = [t for t in all_tr if cutoff and t['transfer_date']>=cutoff]

    t_recv    = [t for t in rel_tr if t['qty_in']>0]
    t_transit = [t for t in rel_tr if t['qty_in']==0]

    if t_recv:
        status='closed';       slabel='已結案'
    elif t_transit:
        status='in_transit';   slabel='倉庫已調撥在途中'
    elif has_purchase:
        status='at_warehouse'; slabel='廠商已交貨待發貨'
    else:
        status='pending';      slabel='等待廠商交貨'

    if status=='in_transit' and t_transit:
        latest    = max(t_transit, key=lambda x:x['transfer_date'])
        eta       = add_days(parse_date(latest['transfer_date']), DAYS_STORE)
        eta_basis = f'調出日+{DAYS_STORE}天'; eta_conf='high'
    elif has_purchase and purchase_date:
        eta       = add_days(purchase_date, DAYS_WH)
        eta_basis = f'交貨日+{DAYS_WH}天'; eta_conf='mid'
    else:
        vendor_days = 45 if category == '泳裝' else DAYS_VENDOR
        eta       = add_days(order_date, vendor_days)
        eta_basis = f'下單日+{vendor_days}天'; eta_conf='low'

    overdue_days = 0; is_alert = False
    if status=='pending' and order_date:
        diff = (today - order_date).days
        if diff > ALERT_OVERDUE_DAY:
            is_alert=True; overdue_days=diff-ALERT_OVERDUE_DAY

    stores_src = t_recv if t_recv else t_transit
    stores = [{'code':t['store_code'],'name':t['store_name'],
               'qty_in':t['qty_in'],'qty_out':t['qty_out'],'date':t['transfer_date']}
              for t in stores_src]

    result.append({
        'img':           str(row['商品內部圖片']).strip() if pd.notna(row['商品內部圖片']) else '',
        'brand':         brand,
        'is_outlet':     is_outlet,
        'order_id':      order_id,
        'style_no':      style_no,
        'item_code':     str(row['商品編號']).strip() if pd.notna(row['商品編號']) else '',
        'name':          name,
        'is_special':    is_special,
        'is_preorder':   is_preorder,
        'size':          size,
        'size_rank':     size_rank(size),
        'color_name':    str(row['顏色名稱']).strip() if pd.notna(row['顏色名稱']) else '',
        'color_code':    str(row['顏色']).strip()     if pd.notna(row['顏色'])     else '',
        'barcode':       barcode,
        'order_date':    fmt_date(order_date),
        'order_date_raw': order_date_str,
        'order_qty':     int(twa),
        'purchase_date': fmt_date(purchase_date),
        'purchase_qty':  p_rec['qty'] if has_purchase else 0,
        'eta':           fmt_date(eta),
        'eta_raw':       ymd(eta),
        'eta_week':      week_label(eta),
        'eta_basis':     eta_basis,
        'eta_conf':      eta_conf,
        'status':        status,
        'status_label':  slabel,
        'stores':        stores,
        'category':      category,
        'is_alert':      is_alert,
        'overdue_days':  overdue_days,
    })

counts = {}
for r in result: counts[r['status']] = counts.get(r['status'],0)+1
alert_count    = sum(1 for r in result if r['is_alert'])
preorder_count = sum(1 for r in result if r['is_preorder'])

print(f"  串接完成：{len(result)} 筆  警示：{alert_count}  預購：{preorder_count}")
print(f"  狀態：{counts}")

# 產出 HTML
print("\n📝 產出 HTML…")
with open(TEMPLATE_FILE,'r',encoding='utf-8') as f: template = f.read()

payload = json.dumps({
    'data':          result,
    'generated':     (datetime.utcnow() + __import__('datetime').timedelta(hours=8)).strftime('%Y/%m/%d %H:%M'),
    'counts':        counts,
    'alert_count':   alert_count,
    'preorder_count':preorder_count,
    'store_map':     store_map,
    'mgr_order':     MGR_ORDER,
}, ensure_ascii=False, separators=(',',':'))

with open(OUTPUT_HTML,'w',encoding='utf-8') as f:
    f.write(template.replace('/*__DATA_PLACEHOLDER__*/null', payload))
print(f"✅ 完成！總筆數：{len(result)} / 警示：{alert_count} / 預購：{preorder_count}")
