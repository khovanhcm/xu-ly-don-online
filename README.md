# Xử lý đơn Online theo kho

Dashboard theo dõi hạn giao và xử lý đơn Online của 3 kho (TT Tân Bình, Bình Chánh, Thủ Đức 1), chạy từ file Excel **"Chi tiết yêu cầu xuất"** xuất từ hệ thống.

## Cách dùng hằng ngày
1. Mở trang dashboard (link GitHub Pages của repo).
2. Bấm **Chọn file Excel** hoặc kéo thả file `Chi_tiết_yêu_cầu_xuất_YYYYMMDD_HHMM.xlsx` vào trang.
3. Xem các khối: Hạn giao theo ngày, Báo cáo trễ Ahamove / GHTK, Tổng quan.
4. Tải Excel: *Cần xử lý trong ngày*, *Trễ hẹn khách*, *Trễ đơn Ahamove*, *Trễ đơn GHTK*.

File Excel chỉ được đọc trong trình duyệt của người dùng, **không tải lên GitHub hay máy chủ nào**.
Giữ nguyên cách đặt tên file `..._YYYYMMDD_HHMM.xlsx` để trang lấy đúng thời điểm dữ liệu khi tính trễ hạn.

## Quy tắc tính
- Mặc định đếm theo **đơn hàng**, bỏ đơn đã hủy và phiếu "Yêu cầu xuất trả hàng mua" (có thể bật lại trên trang).
- **Cần xử lý** = chưa xuất kho, chưa hủy. **Đang giao** = đã xuất, chưa giao.
- **Trễ hẹn** = giờ hẹn giao trước thời điểm dữ liệu, chưa giao, chưa hủy.
- **Hạn xuất kho**: Ahamove trước giờ hẹn **1 giờ**, GHTK trước giờ hẹn **72 giờ**.
  Sửa trong `index.html`, dòng `lead: {Ahamove: 1, GHTK: 72}`.

## Đưa lên GitHub Pages (một lần)
1. Tạo repo mới trên GitHub (ví dụ `xu-ly-don-online`).
2. Tải lên 2 file `index.html` và `README.md` (Add file → Upload files → Commit).
3. Vào **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `(root)` → Save.
4. Sau 1–2 phút, trang chạy tại `https://<tên-tài-khoản>.github.io/<tên-repo>/`.

Lưu ý: GitHub Pages miễn phí yêu cầu repo **public** (ai có link đều mở được trang). Trang không chứa dữ liệu đơn hàng nào; dữ liệu chỉ nằm trong file Excel người dùng tự chọn trên máy mình.

## Thư viện dùng (tải từ cdnjs)
- d3 7.9.0 (biểu đồ)
- SheetJS xlsx 0.18.5 (đọc và xuất Excel)
