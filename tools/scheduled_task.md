# Scheduled task "Báo cáo đơn Online LINE"

Lưu lại để dựng lại task khi đổi máy chạy tự động.

## Lịch (giờ Việt Nam) – tin vào nhóm LINE "Kho Online Miền HCM"
| Tin đến | Chạy lúc | Task |
|---|---|---|
| 8g30 | 8g25 | "Báo cáo đơn Online LINE 8g30" – cron `CRON_TZ=Asia/Ho_Chi_Minh 25 8 * * *` |
| 11g, 13g, 15g, 17g, 19g | 10g55, 12g55, 14g55, 16g55, 18g55 | "Báo cáo đơn Online LINE 11g-13g-15g-17g-19g" – cron `CRON_TZ=Asia/Ho_Chi_Minh 55 10,12,14,16,18 * * *` |

Hai task vì 8g25 khác phút với các khung còn lại (một lịch chỉ đặt được một phút). Cả hai dùng chung nội dung bên dưới.

## Máy chạy cần có
- Windows, app Claude desktop đăng nhập đúng tài khoản, bật Computer use, máy không ngủ.
- LINE desktop đã đăng nhập, có nhóm "Kho Online Miền HCM".
- Trình duyệt trong app Claude đã đăng nhập report.mwgroup.vn (quét QR DMXWork, phiên ~12 giờ) – người dùng tự lo.
- Không cần cấp thư mục nào (không tải file về máy).

## Chuyển sang máy mới
Trên máy mới, mở app Claude, trong một cuộc trò chuyện Cowork nối với máy đó, nhờ Claude: tạo lại 2 task theo file này (requires_local_device = true) và xoá 2 task cũ gắn máy cũ. Hoặc trong cài đặt từng task, bật "Require this computer" trên máy mới.

## Nội dung task (prompt)
```
Gửi báo cáo đơn Online (Ahamove) theo kho vào nhóm LINE "Kho Online Miền HCM". Trả lời bằng tiếng Việt. Chạy hết các bước, không hỏi lại, trừ khi bị chặn. Không cần quyền thư mục Downloads (không tải file về máy).

BƯỚC 0 – Chuẩn bị: git clone --depth 1 https://github.com/khovanhcm/xu-ly-don-online /tmp/xldo ; đọc /tmp/xldo/tools/kho.txt lấy danh sách MÃ KHO (mỗi dòng một mã, bỏ dòng bắt đầu bằng #). Đọc /tmp/xldo/tools/lay_file_excel.js (dùng ở bước 1.4).

BƯỚC 1 – Xuất và đọc dữ liệu (trình duyệt trong app Claude, tools mcp__remote-devices__Claude_Browser__*; nạp hết bằng ToolSearch query "mcp__remote-devices__Claude_Browser__", và đọc skill built-in-browser trước):
1. preview_start https://report.mwgroup.vn/home/dashboard/77 . Nếu bị chuyển tới authkey.dmx.vn (màn hình quét QR) = phiên đăng nhập đã hết hạn: TUYỆT ĐỐI KHÔNG gửi gì vào nhóm LINE. Dừng ngay, gọi PushNotification "⚠️ Báo cáo đơn Online HH:MM chưa gửi: report.mwgroup.vn cần đăng nhập lại (quét QR DMXWork trên máy tính)", câu trả lời cuối cũng bắt đầu bằng đúng dòng đó.
2. javascript_tool (đợi trang tải ~4 giây trong cùng lệnh):
   - window.alert = m => (window.__alerts=(window.__alerts||[])).push(String(m)).
   - Từ ngày = N-3, Đến ngày = N+1 (giờ Việt Nam): 2 kendoDatePicker đầu trang: [...document.querySelectorAll('input')].map(e=>$(e).data('kendoDatePicker')).filter(Boolean); dp.value(new Date(y,m,d)); dp.trigger('change').
   - "Tìm theo (Kho)": trong các kendoComboBox/kendoDropDownList, cái có dữ liệu chứa "Kho xuất" → w.value('1'); w.trigger('change').
3. Chọn kho: chụp screenshot, CLICK THẬT (computer left_click) vào ô "… Siêu thị được chọn" dưới nhãn Kho (xác định toạ độ từ screenshot; khung 800x600 thường ở khoảng x=90,y=246). Rồi javascript_tool: chờ 1,5 giây; $('.jstree').filter(':visible').jstree(true).open_all(); chờ 2 giây; với mỗi MÃ KHO tìm .jstree-anchor đang hiển thị có text bắt đầu "<mã> - ", nếu chưa có class jstree-clicked thì .click(); rồi click nút <button> hiển thị có text "Xong". Mã nào không thấy → báo lỗi, không gửi. Kiểm tra s = angular.element(document.getElementById('btnExportData')).scope(); s.ListCondition: Kho = đúng các MÃ KHO, "Tìm theo (Kho)" = "1", Từ/Đến ngày đúng. Sai thì làm lại bước 2–3 một lần, vẫn sai thì báo lỗi.
   Sau đó s.$apply(()=>s.ExportData('EXCEL')); chờ tới khi s.continueDownload = true (tối đa 30 giây); ghi lại __alerts nếu có.
4. navigate https://report.mwgroup.vn/ManagerDownload; cứ ~15 giây location.reload() rồi đọc dòng <tr> đầu tiên chứa "Chi tiết yêu cầu xuất" (phải là giờ vừa xuất), đến khi báo "Đã xuất xong" (tối đa 5 phút). KHÔNG bấm "Tải Excel". Chạy javascript_tool với NGUYÊN VĂN nội dung file tools/lay_file_excel.js (đọc file ở bước 0, dán toàn bộ). Kết quả: {xuatLuc, soDong, soDon, khos, soDong_Ahamove_homnay, rows} hoặc {loi}. Có loi → báo lỗi, không gửi. Kiểm tra xuatLuc đúng là lần xuất vừa rồi và khos có đủ 3 kho.

BƯỚC 2 – Tính báo cáo: ghi chuỗi rows (mỗi dòng một đơn, giữ nguyên) vào /tmp/rows.txt (có thể rỗng), rồi python3 /tmp/xldo/tools/tao_tin.py "<xuatLuc>" /tmp/rows.txt > /tmp/msg.txt . Đếm số dòng /tmp/rows.txt phải bằng soDong_Ahamove_homnay. Nội dung /tmp/msg.txt là tin nhắn, gửi nguyên văn, không sửa.

BƯỚC 3 – Gửi LINE (computer use, tools mcp__remote-devices__computer_*; đọc skill computer-use trước):
1. computer_resolve_access ["LINE"] (clipboardWrite true) rồi computer_request_access với TẤT CẢ các entry LINE trả về (truyền nguyên willHide, screenshotFiltering). Nếu không thấy LINE, dùng computer_list_apps query "LINE" để tìm đúng tên/đường dẫn line.exe rồi resolve lại.
2. computer_open_application với entry LINE (ưu tiên entry có bundleId kết thúc bằng line.exe, không phải linelauncher.exe), chụp màn hình. Trong cửa sổ danh sách chat của LINE, tìm "Kho Online Miền HCM" (ô tìm kiếm trên cùng hoặc danh sách chat), double-click mở đúng nhóm. Zoom kiểm tra tiêu đề khung chat là "Kho Online Miền HCM" (số trong ngoặc là số thành viên); không đúng thì KHÔNG gửi, báo lỗi.
3. computer_write_clipboard nội dung tin, click ô "Enter a message" của khung chat nhóm, ctrl+v, chụp kiểm tra, nhấn Return. Zoom xác nhận tin đã lên. Nếu cửa sổ LINE bị che/thu nhỏ: computer_open_application LINE lại, click vào phần cửa sổ LINE đang thấy; nếu vẫn không được, xin quyền "File Explorer" (resolve rồi request) và bấm biểu tượng LINE ở thanh tác vụ. Nếu có hộp thoại "Save as" của app Claude đang mở thì bỏ qua (không bấm Save). Cuối cùng computer_release_lock.

Kết thúc: trả lời 1–2 dòng: đã gửi lúc mấy giờ, tổng số đơn trễ / cảnh báo trễ (không cần PushNotification khi thành công). Nếu bất kỳ bước nào lỗi: KHÔNG gửi gì vào nhóm, gọi PushNotification với dòng mở đầu "⚠️ Báo cáo đơn Online HH:MM chưa gửi:" + bước lỗi và lý do.
```
