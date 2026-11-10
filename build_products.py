# -*- coding: utf-8 -*-
"""
韋德機車精品 商城 — 蝦皮匯出檔 → products.json / desc.json

用法：
  python build_products.py <放蝦皮匯出檔的資料夾>
資料夾內放：
  mass_update_basic_info_*.xlsx   （商品標題及描述）
  mass_update_media_info_*.xlsx   （媒體資訊）
  mass_update_sales_info_*.zip 或解壓後的多個 xlsx（價格及庫存）
輸出到目前資料夾：products.json、desc.json
"""
import sys, os, re, io, json, glob, zipfile
import pandas as pd

SRC = sys.argv[1] if len(sys.argv) > 1 else '.'
OUT = sys.argv[2] if len(sys.argv) > 2 else '.'

# ---------- 讀檔（蝦皮檔案的 sheetView 格式有錯，先修掉）----------
def fixed_bytes(raw):
    zi = zipfile.ZipFile(io.BytesIO(raw)); buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            d = zi.read(it.filename)
            if it.filename.startswith('xl/worksheets/') and it.filename.endswith('.xml'):
                d = re.sub(rb'<sheetViews>.*?</sheetViews>', b'', d, flags=re.S)
            zo.writestr(it, d)
    buf.seek(0); return buf

def read_sheet(raw):
    df = pd.read_excel(fixed_bytes(raw), header=None, dtype=str)
    df.columns = df.iloc[2].tolist()
    return df.iloc[6:].reset_index(drop=True)

def find(pat):
    r = sorted(glob.glob(os.path.join(SRC, pat)))
    return r

basic_f = find('*basic_info*.xlsx'); media_f = find('*media_info*.xlsx')
if not basic_f or not media_f: sys.exit('找不到 basic_info 或 media_info 檔案')
basic = read_sheet(open(basic_f[-1], 'rb').read())
media = read_sheet(open(media_f[-1], 'rb').read())

sales_parts = []
for z in find('*sales_info*.zip'):
    with zipfile.ZipFile(z) as zf:
        for n in sorted(zf.namelist()):
            if n.endswith('.xlsx'): sales_parts.append(read_sheet(zf.read(n)))
for x in find('*sales_info*.xlsx'):
    sales_parts.append(read_sheet(open(x, 'rb').read()))
if not sales_parts: sys.exit('找不到 sales_info 檔案')
sales = pd.concat(sales_parts).drop_duplicates(subset=['商品選項ID', '商品ID'])

# ---------- 分類：依 config.json 的 categories，用商品標題關鍵字分 ----------
CFG_PATH = sys.argv[3] if len(sys.argv) > 3 else os.path.join(OUT, 'config.json')
CATS = json.load(open(CFG_PATH, encoding='utf-8'))['categories']
_brand = [c for c in CATS if not c.get('g')]
_generic = [c for c in reversed(CATS) if c.get('g') and c['k']]
OTHER = next((c['n'] for c in CATS if not c['k']), '其他')
def cat_of(name):
    u = ' ' + (name or '').upper() + ' '
    for c in _brand:                      # 標題直接寫「XX部品」就以它為準
        if c['n'].endswith('部品') and c['n'].replace(' ', '').upper() in u.replace(' ', ''): return c['n']
    for c in _brand + _generic:
        if any(k.upper() in u for k in c['k']): return c['n']
    return OTHER

# ---------- 描述整理：拿掉蝦皮通用說明，圖片另外列 ----------
IMG_RE = re.compile(r'https?://s-cf-tw\.shopeesz\.com/file/([\w\-]+)')
DROP = ['營業時間', '公休', '聯絡電話', '02-2998', '0229987231', '02-29987231', '0975', '歡迎來到',
        '請先詢問是否有貨', '下標前', '下單資訊', '七日', '組裝或使用', '全程錄影', '賣場數量',
        '＿＿', '__', '注意事項', '確定有無需要', '蝦皮', '週一到週五', '店家地址', '新莊區', '大漢橋', 'IG：', 'FB：', 'IG:', 'FB:', '官方', '聯絡方式', '頭前庄', '評價']
def clean_desc(text):
    text = text if isinstance(text, str) else ''
    imgs = IMG_RE.findall(text)
    lines = []
    for ln in IMG_RE.sub('', text).split('\n'):
        s = ln.strip().lstrip('●•').strip()
        if not s:
            if lines and lines[-1] != '': lines.append('')
            continue
        if any(k in s for k in DROP): continue
        if re.fullmatch(r'[☎📱\d\-\s]+', s): continue
        s = s.replace('聊聊', 'LINE')
        lines.append(s)
    while lines and lines[0] == '': lines.pop(0)
    while lines and lines[-1] == '': lines.pop()
    return '\n'.join(lines), imgs

def img_id(u):
    m = IMG_RE.search(u) if isinstance(u, str) else None; return m.group(1) if m else None

# ---------- 組合 ----------
media = media.set_index('商品ID'); basic = basic.set_index('商品ID')
sales['價格'] = pd.to_numeric(sales['價格'], errors='coerce')
sales['庫存'] = pd.to_numeric(sales['庫存'], errors='coerce').fillna(0).astype(int)

products, descs = [], {}
order = list(media.index)  # 保持蝦皮匯出順序（新上架在前）
grp = {k: g for k, g in sales.groupby('商品ID', sort=False)}
for pid in order:
    if pid not in grp: continue
    mr = media.loc[pid]; g = grp[pid]
    name = re.sub(r'^\s*韋德機車精品\s*', '', str(mr['商品名稱'])).strip()
    imgs = [img_id(mr.get(c)) for c in ['主商品圖片'] + [f'商品圖片 {i}' for i in range(1, 9)]]
    imgs = [x for x in imgs if x]
    # 選項圖片（第一層規格）
    oimg = {}
    for i in range(1, 51):
        on, oi = mr.get(f'選項 {i} 的名稱'), mr.get(f'選項 {i} 的圖片')
        if isinstance(on, str) and img_id(oi): oimg[on.strip()] = img_id(oi)
    tier1 = mr.get('規格名稱 1'); tier1 = tier1.strip() if isinstance(tier1, str) else ''
    vars_ = []
    for _, r in g.iterrows():
        if pd.isna(r['價格']): continue
        vn = r['商品規格名稱']; vn = vn.strip() if isinstance(vn, str) else ''
        vars_.append([vn, int(r['價格']), int(r['庫存']), str(r['商品選項ID'])])
    if not vars_: continue
    prices = [v[1] for v in vars_]
    p = {'i': str(pid), 'n': name, 'c': cat_of(name), 'im': imgs,
         'p': min(prices), 'P': max(prices), 'v': vars_}
    if tier1 and tier1.lower() != 'variation': p['t'] = tier1
    if oimg: p['oi'] = oimg
    if sum(v[2] for v in vars_) == 0: p['so'] = 1
    products.append(p)
    d, dimgs = clean_desc(basic.loc[pid]['商品描述'] if pid in basic.index else '')
    if d or dimgs: descs[str(pid)] = [d, dimgs]

cats = [{'k': c['n'], 'n': c['n'], 'cnt': sum(1 for p in products if p['c'] == c['n'])} for c in CATS]
cats = [c for c in cats if c['cnt']]
from datetime import datetime, timezone, timedelta
out = {'updated': datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M'),
       'img': 'https://s-cf-tw.shopeesz.com/file/', 'cats': cats, 'items': products}
os.makedirs(OUT, exist_ok=True)
json.dump(out, open(os.path.join(OUT, 'products.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
json.dump(descs, open(os.path.join(OUT, 'desc.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print(f'完成：{len(products)} 個商品、{sum(len(p["v"]) for p in products)} 個規格')
for c in cats: print(f"  {c['n']}: {c['cnt']}")
