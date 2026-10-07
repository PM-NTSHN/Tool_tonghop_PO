# Tool tổng hợp PO – NTS Hanoi Corp.

**`Tool tổng hợp PO.html`** là một file HTML duy nhất, mở trực tiếp bằng trình duyệt (Chrome/Edge/Firefox), **chạy offline 100%**, không gửi dữ liệu đi đâu. Có thể gửi file này cho người khác dùng.

## Cách dùng
1. Mở `Tool tổng hợp PO.html` → kéo-thả (hoặc chọn) file Excel **raw** xuất từ Base Service (*Service Tickets – Yêu cầu đặt hàng*). Tool cũng nhận lại file `Data_sạch` đã chuẩn hóa.
2. Tool tự convert sang template chuẩn **Data_sạch** (30 cột) theo đúng logic `tools/convert_service_tickets.py` → nút **Tải Data_sạch** để tải file template (`<tên file>_CLEAN.xlsx`).
3. Lọc / phân tích: tìm nhanh có gợi ý, lọc nhanh, bộ lọc đa chọn (Hãng, Sale, Khách hàng, Đại lý/Reseller, Nhóm KH, PM, Trạng thái, Bước hiện tại, Tình trạng HĐ, BB thanh lý, Luồng nhập hàng Trực tiếp/Exwork, Công ty ký HĐ, Cờ nhắc, Giao hàng, License, SLA…), lọc ngày (Ngày tạo, Ngày giao hàng, License start/end, Cập nhật cuối) bằng chọn nhanh / lưới tháng-quý / lịch. Bấm vào biểu đồ, bảng, KPI để lọc.
4. **Xuất Excel** → `Tổng hợp PO_Hãng_Tháng M.YYYY.xlsx` (Tổng quan, Chi tiết PO, Theo Hãng/Sale/Đại lý/Nhóm KH, Hãng × Tháng, Cờ nhắc, Tuân thủ HĐ-BB, License renew, Data_sạch).

## Cờ nhắc (mặc định, chỉnh trong ⚙ Thiết lập)
Tính trên phiếu **Đang xử lý**, theo số ngày từ *Cập nhật lần cuối* đến *Ngày đối chiếu* (mặc định = lần cập nhật mới nhất trong file; có thể chọn Hôm nay / tự chọn):
Bình thường ≤ 7 ngày · Mức 1: 8–14 · Mức 2: 15–30 · Mức 3 (tồn đọng): > 30 ngày. Kèm: trễ hạn giao, sắp đến hạn giao (≤ 7 ngày), license hết hạn ≤ 90 ngày.

## Phát triển
- Mã nguồn: `src/` (`app.html`, `app.css`, `app.js`, `converter.js` – port JS của script Python), thư viện nhúng trong `vendor/` (xlsx-js-style, Chart.js).
- Đóng gói: `python3 build.py` → tạo lại `Tool tổng hợp PO.html`.
- Kiểm tra convert bằng Node: `node tools/convert_node.js RAW.xlsx OUT.xlsx` (cần `npm i xlsx-js-style`). Đã đối chiếu với script Python: khớp 100% từng ô trên dữ liệu thật và dữ liệu thử nghiệm biến dạng.
