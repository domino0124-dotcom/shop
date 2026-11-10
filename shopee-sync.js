/* 蝦皮匯出檔 → products.json / desc.json（在瀏覽器裡轉換，不需要其他程式）
 * 用法：ShopeeSync.build(files, config) → { products, desc, stats }
 * files：蝦皮「批次更新」下載的 basic_info.xlsx、media_info.xlsx、sales_info.zip（或解壓後的 xlsx）
 */
(function (root) {
  'use strict';

  /* ---------- 迷你 zip 讀取（用瀏覽器內建的解壓縮）---------- */
  async function unzip(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error('不是有效的 zip / xlsx 檔');
    const n = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const out = {};
    for (let k = 0; k < n; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nl));
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      const data = u8.subarray(start, start + csize);
      out[name] = { method, data };
      p += 46 + nl + el + cl;
    }
    return {
      names: Object.keys(out),
      async get(name) {
        const f = out[name]; if (!f) return null;
        if (f.method === 0) return f.data;
        const ds = new DecompressionStream('deflate-raw');
        const res = await new Response(new Blob([f.data]).stream().pipeThrough(ds)).arrayBuffer();
        return new Uint8Array(res);
      }
    };
  }

  /* ---------- 讀 xlsx 第一個工作表 → { 列號: [欄值...] } ---------- */
  const colIdx = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  async function readXlsx(buf) {
    const z = await unzip(buf), dec = new TextDecoder(), dp = new DOMParser();
    const ssRaw = await z.get('xl/sharedStrings.xml');
    const ss = [];
    if (ssRaw) {
      const doc = dp.parseFromString(dec.decode(ssRaw), 'application/xml');
      for (const si of doc.getElementsByTagName('si')) {
        let t = ''; for (const x of si.getElementsByTagName('t')) { if (x.parentNode.nodeName !== 'rPh') t += x.textContent; }
        ss.push(t);
      }
    }
    const sheetName = z.names.filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];
    const doc = dp.parseFromString(dec.decode(await z.get(sheetName)), 'application/xml');
    const rows = {};
    for (const r of doc.getElementsByTagName('row')) {
      const rn = +r.getAttribute('r'), arr = [];
      for (const c of r.getElementsByTagName('c')) {
        const t = c.getAttribute('t'); let v;
        if (t === 's') { const x = c.getElementsByTagName('v')[0]; v = x ? ss[+x.textContent] : ''; }
        else if (t === 'inlineStr') { v = [...c.getElementsByTagName('t')].map(x => x.textContent).join(''); }
        else { const x = c.getElementsByTagName('v')[0]; v = x ? x.textContent : ''; }
        arr[colIdx(c.getAttribute('r'))] = v;
      }
      rows[rn] = arr;
    }
    return rows;
  }
  // 蝦皮格式：第 3 列是中文欄名，第 7 列起是資料
  function table(rows) {
    const head = (rows[3] || []).map(x => (x || '').trim());
    const out = [];
    Object.keys(rows).map(Number).filter(n => n >= 7).sort((a, b) => a - b).forEach(n => {
      const o = {}; head.forEach((h, k) => { if (h) o[h] = rows[n][k]; }); if (o['商品ID']) out.push(o);
    });
    return out;
  }

  /* ---------- 和 build_products.py 相同的規則 ---------- */
  const IMG_RE = /https?:\/\/s-cf-tw\.shopeesz\.com\/file\/([\w\-]+)/g;
  const DROP = ['營業時間', '公休', '聯絡電話', '02-2998', '0229987231', '02-29987231', '0975', '歡迎來到',
    '請先詢問是否有貨', '下標前', '下單資訊', '七日', '組裝或使用', '全程錄影', '賣場數量',
    '＿＿', '__', '注意事項', '確定有無需要', '蝦皮', '週一到週五', '店家地址', '新莊區', '大漢橋', 'IG：', 'FB：', 'IG:', 'FB:', '官方', '聯絡方式', '頭前庄', '評價'];
  const isStr = v => typeof v === 'string' && v !== '';
  const imgId = u => { if (!isStr(u)) return null; const m = /https?:\/\/s-cf-tw\.shopeesz\.com\/file\/([\w\-]+)/.exec(u); return m ? m[1] : null; };
  function cleanDesc(text) {
    text = isStr(text) ? text : '';
    const imgs = [...text.matchAll(IMG_RE)].map(m => m[1]);
    const lines = [];
    for (const ln of text.replace(IMG_RE, '').split('\n')) {
      const s0 = ln.trim().replace(/^[●•]+/, '').trim();
      if (!s0) { if (lines.length && lines[lines.length - 1] !== '') lines.push(''); continue; }
      if (DROP.some(k => s0.includes(k))) continue;
      if (/^[☎📱\d\-\s]+$/u.test(s0)) continue;
      lines.push(s0.split('聊聊').join('LINE'));
    }
    while (lines.length && lines[0] === '') lines.shift();
    while (lines.length && lines[lines.length - 1] === '') lines.pop();
    return [lines.join('\n'), imgs];
  }
  function catOf(cats) {
    const brand = cats.filter(c => !c.g), generic = cats.filter(c => c.g && c.k.length).reverse();
    const other = ([...cats].reverse().find(c => !c.k.length) || { n: '其他' }).n;
    // 英文字母開頭／結尾的關鍵字要整個字對上（避免 RAK 對到 BRAKE）
    const hit = (u, k) => { k = k.toUpperCase(); let i = u.indexOf(k);
      while (i >= 0) { const a = u[i - 1] || ' ', b = u[i + k.length] || ' ';
        if (!(/[A-Z]/.test(k[0]) && /[A-Z]/.test(a)) && !(/[A-Z]/.test(k[k.length - 1]) && /[A-Z]/.test(b))) return true;
        i = u.indexOf(k, i + 1); } return false; };
    return name => { const u = ' ' + (name || '').toUpperCase() + ' ', u2 = u.replace(/ /g, '');
      for (const c of brand) if (c.n.endsWith('部品') && u2.includes(c.n.replace(/ /g, '').toUpperCase())) return c.n;
      for (const c of brand.concat(generic)) if (c.k.some(k => hit(u, k))) return c.n; return other; };
  }

  async function build(files, cfg, onStep) {
    const step = onStep || (() => {});
    let basic, media; const sales = [];
    for (const f of files) {
      const buf = await f.arrayBuffer(), nm = f.name.toLowerCase();
      if (nm.includes('basic_info') && nm.endsWith('.xlsx')) { step('讀取 商品標題及描述'); basic = table(await readXlsx(buf)); }
      else if (nm.includes('media_info') && nm.endsWith('.xlsx')) { step('讀取 媒體資訊'); media = table(await readXlsx(buf)); }
      else if (nm.includes('sales_info') && nm.endsWith('.zip')) {
        step('讀取 價格及庫存');
        const z = await unzip(buf);
        for (const n of z.names.filter(n => n.toLowerCase().endsWith('.xlsx')).sort()) sales.push(...table(await readXlsx((await z.get(n)).slice().buffer)));
      } else if (nm.includes('sales_info') && nm.endsWith('.xlsx')) { step('讀取 價格及庫存'); sales.push(...table(await readXlsx(buf))); }
    }
    const miss = [!basic && '商品標題及描述 (basic_info)', !media && '媒體資訊 (media_info)', !sales.length && '價格及庫存 (sales_info)'].filter(Boolean);
    if (miss.length) throw new Error('缺少檔案：' + miss.join('、'));
    step('整理商品');
    const cat = catOf(cfg.categories);
    const bmap = {}; basic.forEach(r => bmap[r['商品ID']] = r);
    const grp = {}, seen = new Set();
    sales.forEach(r => { const key = r['商品ID'] + '|' + r['商品選項ID']; if (seen.has(key)) return; seen.add(key); (grp[r['商品ID']] = grp[r['商品ID']] || []).push(r); });
    const products = [], descs = {};
    for (const mr of media) {
      const pid = mr['商品ID'], g = grp[pid]; if (!g) continue;
      const name = String(mr['商品名稱'] || '').replace(/^\s*韋德機車精品\s*/, '').trim();
      const imgs = ['主商品圖片', ...Array.from({ length: 8 }, (_, i) => '商品圖片 ' + (i + 1))].map(c => imgId(mr[c])).filter(Boolean);
      const oimg = {};
      for (let i = 1; i <= 50; i++) { const on = mr['選項 ' + i + ' 的名稱'], oi = imgId(mr['選項 ' + i + ' 的圖片']); if (isStr(on) && oi) oimg[on.trim()] = oi; }
      let t1 = mr['規格名稱 1']; t1 = isStr(t1) ? t1.trim() : '';
      const vars = [];
      g.forEach(r => {
        const price = parseFloat(r['價格']); if (isNaN(price)) return;
        const stock = parseInt(parseFloat(r['庫存']), 10) || 0;
        vars.push([isStr(r['商品規格名稱']) ? r['商品規格名稱'].trim() : '', Math.trunc(price), stock, String(r['商品選項ID'])]);
      });
      if (!vars.length) continue;
      const prices = vars.map(v => v[1]);
      const p = { i: String(pid), n: name, c: cat(name), im: imgs, p: Math.min(...prices), P: Math.max(...prices), v: vars };
      if (t1 && t1.toLowerCase() !== 'variation') p.t = t1;
      if (Object.keys(oimg).length) p.oi = oimg;
      if (vars.every(v => v[2] === 0)) p.so = 1;
      products.push(p);
      const [d, dimgs] = cleanDesc(bmap[pid] ? bmap[pid]['商品描述'] : '');
      if (d || dimgs.length) descs[String(pid)] = [d, dimgs];
    }
    const cats = cfg.categories.map(c => ({ k: c.n, n: c.n, cnt: products.filter(p => p.c === c.n).length })).filter(c => c.cnt);
    const now = new Date(Date.now() + 8 * 3600e3).toISOString().slice(0, 16).replace('T', ' ');
    const out = { updated: now, img: 'https://s-cf-tw.shopeesz.com/file/', cats, items: products };
    return { products: out, desc: descs, stats: { items: products.length, vars: products.reduce((a, p) => a + p.v.length, 0) } };
  }

  root.ShopeeSync = { build, catOf };
})(window);
