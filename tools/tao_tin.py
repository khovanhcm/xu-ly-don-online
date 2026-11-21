"""Dựng tin nhắn LINE từ danh sách đơn do tools/lay_file_excel.js trả về.
Dùng: python3 tao_tin.py "<xuatLuc dd/mm/yyyy HH:MM:SS>" <rows.txt>   (rows: mã|N/B/T|serial|1/0)"""
import sys, os, datetime as D, tempfile, subprocess, openpyxl
K = {'N': 'ĐMX_HCM - kho Online TT Tân Bình', 'B': 'HCM - Kho Online Bình Chánh', 'T': 'HCM - Kho Online Thủ Đức 1'}
t = D.datetime.strptime(sys.argv[1].strip(), '%d/%m/%Y %H:%M:%S')
wb = openpyxl.Workbook(); ws = wb.active
ws.append(['Loại YCX', 'Mã đơn hàng', 'Trạng thái xuất', 'Trạng thái giao hàng', 'Trạng thái hủy', 'Hình thức giao hàng',
           'Thời gian hẹn giao', 'Siêu thị xuất', 'Tình trạng nhập trả của sản phẩm đổi với sản phẩm chính'])
for k in 'NBT':  # giữ đủ 3 kho trong tin (dòng đã giao, không ảnh hưởng số liệu)
    ws.append(['', '_' + k, 'Đã xuất', 'Đã giao', 'Chưa hủy', 'GHTK', t - D.timedelta(days=1), K[k], ''])
n = 0
for l in open(sys.argv[2], encoding='utf-8'):
    l = l.strip()
    if not l: continue
    d, k, h, x = l.split('|'); n += 1
    ws.append(['Yêu cầu xuất bán hàng Online', d, 'Đã xuất' if x == '1' else 'Chưa xuất', 'Chưa giao', 'Chưa hủy',
               'Giao hàng qua đối tác Ahamove', D.datetime(1899, 12, 30) + D.timedelta(seconds=round(float(h) * 86400)), K[k], ''])
p = os.path.join(tempfile.gettempdir(), f'Chi tiết yêu cầu xuất_{t:%Y%m%d_%H%M}.xlsx'); wb.save(p)
sys.stdout.write(subprocess.check_output([sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bao_cao_line.py'), p], text=True))
