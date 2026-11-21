// Chạy trong trang https://report.mwgroup.vn/ManagerDownload (javascript_tool).
// Lấy file Excel của dòng "Chi tiết yêu cầu xuất" mới nhất đã xuất xong, ĐỌC NGAY TRONG TRÌNH DUYỆT
// (không tải về máy, không bật hộp thoại Save as), trả về danh sách đơn Ahamove chưa giao hẹn giao hôm nay
// dạng "mã đơn|kho(N/B/T)|giờ hẹn (số serial Excel)|đã xuất(1/0)" + thời điểm xuất.
await (async () => {
  const row = [...document.querySelectorAll('tr')].find(t => t.textContent.includes('Chi tiết yêu cầu xuất'));
  if (!row) return { loi: 'Không thấy dòng Chi tiết yêu cầu xuất' };
  const txt = row.innerText;
  if (!/Đã xuất xong/.test(txt)) return { loi: 'Chưa xuất xong', dong: txt };
  const xuatLuc = (txt.match(/(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2})/) || [])[1];
  const sc = angular.element(row).scope();
  const ow = window.open; let u = null; window.open = x => { u = x; return null; };
  try { sc.ViewNow(sc.dataItem); } finally { window.open = ow; }
  if (!u) return { loi: 'Không lấy được link ParquetViewer' };
  const blobs = [];
  const f = document.createElement('iframe'); f.style.display = 'none';
  f.src = '/' + u.trim().replace(/(%20|\s)+$/, '') + '&download=xlsx&bg=1';
  const patch = () => { try { const w = f.contentWindow; if (w && !w.__p && w.URL) {
    const oc = w.URL.createObjectURL; w.URL.createObjectURL = function (b) { if (b && b.size) blobs.push(b); return oc.apply(this, arguments); };
    const ac = w.HTMLAnchorElement.prototype.click; w.HTMLAnchorElement.prototype.click = function () { if (/^blob:/.test(this.href)) return; return ac.apply(this, arguments); };
    const ad = w.HTMLAnchorElement.prototype.dispatchEvent; w.HTMLAnchorElement.prototype.dispatchEvent = function (e) { if (/^blob:/.test(this.href)) return true; return ad.apply(this, arguments); };
    w.__p = 1; } } catch (e) {} };
  const pi = setInterval(patch, 10); document.body.appendChild(f);
  for (let i = 0; i < 40 && !blobs.length; i++) await new Promise(r => setTimeout(r, 1000));
  clearInterval(pi); setTimeout(() => f.remove(), 500);
  if (!blobs.length) return { loi: 'Hết 40 giây chưa tạo được file Excel' };
  const buf = new Uint8Array(await blobs[0].arrayBuffer()); const dv = new DataView(buf.buffer);
  let eocd = buf.length - 22; while (dv.getUint32(eocd, true) !== 0x06054b50) eocd--;
  const n = dv.getUint16(eocd + 10, true); let off = dv.getUint32(eocd + 16, true); const files = {};
  for (let k = 0; k < n; k++) { const method = dv.getUint16(off + 10, true), csize = dv.getUint32(off + 20, true), nl = dv.getUint16(off + 28, true), el = dv.getUint16(off + 30, true), cl = dv.getUint16(off + 32, true), lho = dv.getUint32(off + 42, true);
    files[new TextDecoder().decode(buf.subarray(off + 46, off + 46 + nl))] = { method, csize, lho }; off += 46 + nl + el + cl; }
  const fe = files['xl/worksheets/sheet1.xml']; const l = fe.lho; const nl = dv.getUint16(l + 26, true), el = dv.getUint16(l + 28, true);
  const data = buf.subarray(l + 30 + nl + el, l + 30 + nl + el + fe.csize);
  const x = fe.method === 0 ? new TextDecoder().decode(data) : await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
  const rows = x.slice(x.indexOf('<sheetData>') + 11, x.indexOf('</sheetData>')).split('</row>');
  const dec = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const parse = r => { const o = {}; const re = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let m;
    while ((m = re.exec(r))) { const inner = m[3] || ''; let v; const t = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/);
      if (t) v = dec(t[1]); else { const vv = inner.match(/<v>([\s\S]*?)<\/v>/); v = vv ? vv[1] : null; if (v !== null && !/t="s(tr)?"/.test(m[2])) v = Number(v); }
      o[m[1]] = v; } return o; };
  const norm = v => String(v ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  const head = parse(rows[0]); const C = w => Object.keys(head).find(k => norm(head[k]) === norm(w));
  const c = { loai: C('Loại YCX'), don: C('Mã đơn hàng'), xuat: C('Trạng thái xuất'), giao: C('Trạng thái giao hàng'), huy: C('Trạng thái hủy') || 'P',
    ht: C('Hình thức giao hàng'), hen: C('Thời gian hẹn giao'), kho: C('Siêu thị xuất'), tra: C('Tình trạng nhập trả của sản phẩm đổi với sản phẩm chính') || 'AS' };
  const O = {};
  for (let k = 1; k < rows.length; k++) { if (!rows[k].includes('<c')) continue; const r = parse(rows[k]); const don = String(r[c.don] ?? '').trim(); if (!don) continue;
    if (norm(r[c.huy]) === 'đã hủy' || norm(r[c.tra]) === 'đã trả' || /trả hàng/i.test(String(r[c.loai] ?? ''))) continue;
    const o = O[don] || (O[don] = { kho: String(r[c.kho] ?? ''), hen: r[c.hen], ht: String(r[c.ht] ?? ''), xuat: true, giao: true });
    o.xuat = o.xuat && norm(r[c.xuat]) === 'đã xuất'; o.giao = o.giao && norm(r[c.giao]) === 'đã giao'; }
  const [dd, mm, yy] = xuatLuc.slice(0, 10).split('/').map(Number);
  const today = Math.round((Date.UTC(yy, mm - 1, dd) - Date.UTC(1899, 11, 30)) / 86400000);
  const code = k => /tân bình/i.test(k) ? 'N' : /bình chánh/i.test(k) ? 'B' : /thủ đức/i.test(k) ? 'T' : '?';
  const out = []; const khos = new Set();
  for (const [d, o] of Object.entries(O)) { if (typeof o.hen !== 'number') continue; khos.add(o.kho);
    if (/ahamove/i.test(o.ht) && !o.giao && Math.floor(o.hen) === today) out.push([d, code(o.kho), o.hen, o.xuat ? 1 : 0].join('|')); }
  return { xuatLuc, soDong: rows.length - 2, soDon: Object.keys(O).length, khos: [...khos], soDong_Ahamove_homnay: out.length, rows: out.join('\n') };
})()
