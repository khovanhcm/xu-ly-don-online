# Ghi chú dự án cho Claude – Xử lý đơn Online theo kho

Đọc file này trước khi sửa dashboard.

## Bối cảnh
- Người dùng: quản lý kho vận HCM. Trả lời tiếng Việt, ngắn gọn, thẳng vào việc.
- Dashboard một file `index.html`, chạy trên GitHub Pages: https://khovanhcm.github.io/xu-ly-don-online/
- Bản tương ứng trong Claude là artifact "Xử lý đơn Online theo kho" (có nhúng file mẫu 18/11/2026 09:07). Bản GitHub không nhúng dữ liệu.
- Mỗi lần sửa: cập nhật cả artifact và `index.html` trên repo, giữ hai bản giống nhau (khác nhau duy nhất ở dữ liệu nhúng và cách tải file).

## Dữ liệu đầu vào
- File Excel "Chi tiết yêu cầu xuất" xuất từ hệ thống, tên dạng `Chi_tiết_yêu_cầu_xuất_YYYYMMDD_HHMM.xlsx`. Thời điểm dữ liệu lấy từ tên file.
- Cột dùng: Mã đơn hàng, Siêu thị xuất, Thời gian hẹn giao, Hình thức giao hàng, Trạng thái xuất / giao hàng / hủy / duyệt, Loại YCX, Mã sản phẩm, Tên sản phẩm, Số lượng.
- Không đọc thông tin khách hàng (tên, SĐT, địa chỉ, email).
- 3 kho: ĐMX_HCM - kho Online TT Tân Bình, HCM - Kho Online Bình Chánh, HCM - Kho Online Thủ Đức 1.

## Đã chốt với người dùng
- **Lọc trước khi chạy:** bỏ hẳn dòng có "Đã hủy" ở cột P (Trạng thái hủy) và "Đã trả" ở cột AS (Tình trạng nhập trả của sản phẩm đổi với sản phẩm chính). Không có công tắc tính lại đơn hủy. Thẻ "Dòng đã loại khi nhập" hiện số dòng đã bỏ.
- Mặc định đếm theo **đơn** (mã đơn duy nhất), không tính phiếu "Yêu cầu xuất trả hàng mua" (có công tắc bật lại). Bảng Pivot cũ của họ đếm theo dòng SP và tính cả đơn hủy.
- Mã nguồn: sửa `template.html` rồi chạy `build.py` để ra cả 2 bản (lưu ngoài repo); repo chỉ giữ bản GitHub.
- Cần xử lý = chưa xuất, chưa hủy. Đang giao = đã xuất, chưa giao. Trễ hẹn = giờ hẹn trước thời điểm dữ liệu, chưa giao, chưa hủy.
- Hạn xuất kho: **Ahamove trước giờ hẹn 1 giờ**, **GHTK trước giờ hẹn 72 giờ** (cố định, không có ô chỉnh).
- Báo cáo trễ Ahamove: 4 mức (Trễ giao, Trễ xuất, Cần xuất hôm nay, Trong hạn). Báo cáo trễ GHTK: **không có mức Trễ giao** (3 mức).
- Bảng "Lịch xử lý Ahamove theo khung giờ": mỗi kho một con số (tổng), cột Tổng và Chưa xuất; không dùng dạng "x / y".
- Các file Excel xuất: Cần xử lý trong ngày, Trễ hẹn khách, Trễ đơn Ahamove (có sheet khung giờ), Trễ đơn GHTK.
- Giao diện: nền chàm đậm, thẻ màu oải hương, nút chàm, font Lexend (theo mẫu báo cáo NCC của họ).

## Việc còn mở
- Chưa có.

## Báo cáo LINE tự động (từ 20/11/2026)
- 2 scheduled task trên máy của người dùng (Windows, app Claude desktop): 8g25 và 12g55/14g55/16g55 (để tin đến khoảng 8g30, 13g, 15g, 17g). Gửi nhóm LINE "Kho Online Miền HCM".
- Luồng: report.mwgroup.vn/home/dashboard/77 (đăng nhập QR DMXWork, phiên 12 giờ) → xuất Excel 3 kho 15698/36703/36813, Tìm theo = Kho xuất, ngày tạo N-3..N → ManagerDownload "Tải Excel" → C:\Users\Dell\Downloads → `tools/bao_cao_line.py` → dán vào LINE desktop bằng computer use.
- Mẫu tin (người dùng chốt): mỗi kho, Đơn Hoả Tốc (Ahamove): Hạn giao, Cảnh báo trễ, Đơn trễ, Danh sách đơn trễ.
