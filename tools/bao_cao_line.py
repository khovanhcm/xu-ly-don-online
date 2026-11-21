"""Tạo tin nhắn LINE từ file "Chi tiết yêu cầu xuất_YYYYMMDD_HHMM.xlsx".

Quy tắc giống dashboard Xử lý đơn Online theo kho:
- Bỏ dòng "Đã hủy" (Trạng thái hủy) và "Đã trả" (Tình trạng nhập trả...), bỏ phiếu "Yêu cầu xuất trả hàng mua".
- Đếm theo mã đơn.
- Hạn xuất: Ahamove = giờ hẹn - 1 giờ, GHTK = giờ hẹn - 72 giờ, Giao tại siêu thị = giờ hẹn.
- Trễ xuất = chưa xuất, đã qua hạn xuất. Trễ giao = đã qua giờ hẹn, chưa giao (GHTK không tính).
- Cần xử lý = chưa xuất, hạn xuất đến hết hôm nay (gồm cả trễ xuất).
Dùng: python3 bao_cao_line.py <file.xlsx>  -> in tin nhắn ra stdout.
"""
import sys, re, datetime as D
import openpyxl

LEAD = {'Ahamove': 1, 'GHTK': 72, 'Siêu thị': 0}
KHO_ORDER = ['Tân Bình', 'Bình Chánh', 'Thủ Đức']
MAX_LIST = 10
MAX_HAN = 20  # số đơn tối đa liệt kê ở mục Hạn giao mỗi kho
SP_LEN = 45   # cắt tên sản phẩm cho gọn tin

norm = lambda s: re.sub(r'\s+', ' ', str(s if s is not None else '').strip().lower())


def ht_of(s):
    s = str(s or '')
    if re.search('ahamove', s, re.I): return 'Ahamove'
    if re.search('ghtk', s, re.I): return 'GHTK'
    return 'Siêu thị' if re.search('siêu thị', s, re.I) else (s or 'Khác')


def kho_short(k):
    k = str(k or '')
    s = re.sub(r'(?i)^.*kho online\s*', '', k).strip()
    s = re.sub(r'(?i)^tt\s+', '', s)
    return s or k or '(Không rõ kho)'


def to_dt(v):
    if v is None or v == '': return None
    if isinstance(v, D.datetime): return v.replace(microsecond=0)
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%Y-%m-%d', '%d/%m/%Y'):
        try: return D.datetime.strptime(s[:19], f)
        except ValueError: pass
    return None


def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        head = None
        for _ in range(15):
            r = next(it, None)
            if r is None: break
            if any(norm(c) == 'mã đơn hàng' for c in r): head = [norm(c) for c in r]; break
        if head: break
    if not head: raise SystemExit('Không thấy cột "Mã đơn hàng".')
    col = lambda n: head.index(norm(n)) if norm(n) in head else -1
    c = dict(don=col('Mã đơn hàng'), kho=col('Siêu thị xuất'), hen=col('Thời gian hẹn giao'), xuat=col('Trạng thái xuất'),
             giao=col('Trạng thái giao hàng'), ht=col('Hình thức giao hàng'), loai=col('Loại YCX'),
             sp=col('Tên sản phẩm'), sl=col('Số lượng'), gia=col('Giá bán'),
             huy=col('Trạng thái hủy') if col('Trạng thái hủy') >= 0 else 15,
             tra=col('Tình trạng nhập trả của sản phẩm đổi với sản phẩm chính') if col('Tình trạng nhập trả của sản phẩm đổi với sản phẩm chính') >= 0 else 44)
    g = lambda r, i: r[i] if 0 <= i < len(r) else None
    orders = {}
    for r in it:
        don = str(g(r, c['don']) or '').strip()
        if not don: continue
        if norm(g(r, c['huy'])) == 'đã hủy' or norm(g(r, c['tra'])) == 'đã trả': continue
        if re.search('trả hàng', str(g(r, c['loai']) or ''), re.I): continue
        o = orders.setdefault(don, dict(don=don, kho=kho_short(g(r, c['kho'])), hen=to_dt(g(r, c['hen'])), ht=ht_of(g(r, c['ht'])),
                                        xuat=True, giao=True, sp=[]))
        # một đơn nhiều dòng: còn dòng nào chưa xuất / chưa giao thì cả đơn tính là chưa
        o['xuat'] = o['xuat'] and norm(g(r, c['xuat'])) == 'đã xuất'
        o['giao'] = o['giao'] and norm(g(r, c['giao'])) == 'đã giao'
        ten = re.sub(r'\s+', ' ', str(g(r, c['sp']) or '')).strip()
        if ten:
            try: sl = int(float(g(r, c['sl']) or 1))
            except (TypeError, ValueError): sl = 1
            try: gia = float(g(r, c['gia']) or 0)
            except (TypeError, ValueError): gia = 0
            o['sp'].append((ten, sl, gia))
    return list(orders.values())


def snap_of(path):
    m = re.search(r'(20\d{2})(\d{2})(\d{2})[_\- ]?(\d{2})(\d{2})', path)
    return D.datetime(*map(int, m.groups())) if m else D.datetime.now().replace(second=0, microsecond=0)


def late_txt(h):
    """h = số giờ (float). Dưới 1 giờ hiện phút."""
    m = int(round(h * 60))
    if m < 60: return f'{m}p'
    hh = m // 60
    return f'{hh // 24}n{hh % 24}g' if hh >= 24 else f'{hh}g{m % 60:02d}p'


VOUCHER = re.compile(r'(?i)^(pmh|phiếu mua hàng|phiếu quà tặng|voucher)')


def sp_txt(o):
    """Sản phẩm chính của đơn (giá cao nhất, bỏ phiếu mua hàng) + số SP khác."""
    sp = [x for x in o['sp'] if not VOUCHER.search(x[0])] or o['sp']
    if not sp: return ''
    sp = sorted(sp, key=lambda x: -x[2])
    ten, sl, _ = sp[0]
    if len(ten) > SP_LEN: ten = ten[:SP_LEN - 1].rstrip() + '…'
    t = f' · {ten} x{sl}'
    return t + (f' (+{len(sp) - 1} SP)' if len(sp) > 1 else '')


def build(path, partners=('Ahamove',)):
    snap = snap_of(path)
    end_day = snap.replace(hour=23, minute=59, second=59)
    O = [o for o in load(path) if o['hen']]
    for o in O:
        o['dl'] = o['hen'] - D.timedelta(hours=LEAD.get(o['ht'], 0))
    khos = sorted({o['kho'] for o in O}, key=lambda k: next((i for i, s in enumerate(KHO_ORDER) if s.lower() in k.lower()), 9))
    NAME = {'Ahamove': 'Đơn Hoả Tốc (Ahamove)', 'GHTK': 'Đơn GHTK'}
    L = [f'📦 BÁO CÁO ĐƠN ONLINE · {snap:%H:%M %d/%m/%Y}']
    for k in khos:
        L += ['', f'🏬 Kho {k}']
        for p in partners:
            d = [o for o in O if o['kho'] == k and o['ht'] == p and not o['giao']]
            han = sorted([o for o in d if o['hen'] <= end_day and o['hen'] >= snap], key=lambda o: o['hen'])
            # chỉ tính đơn hẹn giao trong ngày N; hẹn N-1 trở về trước coi là lỗi dữ liệu, bỏ qua
            tre = sorted([o for o in d if o['hen'].date() == snap.date() and o['hen'] < snap], key=lambda o: o['hen'])
            canh = sorted([o for o in d if o['hen'].date() == snap.date() and o['hen'] >= snap and not o['xuat'] and o['dl'] < snap], key=lambda o: o['hen'])
            L.append(NAME.get(p, p))
            L.append(f'• Hạn giao: {len(han)} đơn hẹn giao còn lại trong hôm nay')
            for o in han[:MAX_HAN]:
                L.append(f'   - {o["don"]} · hẹn {o["hen"]:%H:%M} · {"đã xuất" if o["xuat"] else "chưa xuất"}{sp_txt(o)}')
            if len(han) > MAX_HAN: L.append(f'   … và {len(han) - MAX_HAN} đơn khác')
            L.append(f'• Cảnh báo trễ: {len(canh)} đơn' + (' (chưa xuất, đã quá hạn xuất)' if canh else ''))
            L.append(f'• Đơn trễ: {len(tre)} đơn')
            if tre:
                L.append('• Danh sách đơn trễ:')
                for o in tre[:MAX_LIST]:
                    L.append(f'   - {o["don"]} · hẹn {o["hen"]:%H:%M %d/%m} · trễ {late_txt((snap - o["hen"]).total_seconds() / 3600)} · {"đã xuất" if o["xuat"] else "chưa xuất"}{sp_txt(o)}')
                if len(tre) > MAX_LIST: L.append(f'   … và {len(tre) - MAX_LIST} đơn khác')
            if canh:
                L.append('• Đơn cảnh báo trễ:')
                for o in canh[:MAX_LIST]:
                    L.append(f'   - {o["don"]} · hẹn {o["hen"]:%H:%M %d/%m} · hạn xuất {o["dl"]:%H:%M}{sp_txt(o)}')
                if len(canh) > MAX_LIST: L.append(f'   … và {len(canh) - MAX_LIST} đơn khác')
    return '\n'.join(L)


if __name__ == '__main__':
    print(build(sys.argv[1]))
