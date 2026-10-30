# Đối chiếu Sticker / Nhãn

Web tĩnh (GitHub Pages) để so **file tác nghiệp nhãn (.xlsx)** với **proof sticker** chủ hàng / nhà cung cấp nhãn gửi về (.pdf) trước khi sản xuất.

Mọi xử lý chạy trên trình duyệt — không file nào được gửi lên server. Thư viện (SheetJS, pdf.js, Tesseract.js + dữ liệu OCR tiếng Anh) nằm sẵn trong `lib/`, không phụ thuộc CDN.

## Cách dùng
1. Chọn file tác nghiệp Excel.
2. Chọn 1 hoặc nhiều file PDF proof.
3. Bấm **Bắt đầu đối chiếu** → xem kết quả, bấm **Xuất báo cáo Excel** nếu cần gửi lại.

## Web tự kiểm tra
| Nhóm | Hạng mục |
|---|---|
| Thông tin chung | Mã nhãn (LB…) khớp sheet tác nghiệp · Kích thước nhãn · Danh sách PO trên proof so với VPONO trong tác nghiệp (PO lạ / PO bị thiếu) |
| Từng nhãn | Style-color có trong đơn · RN# / CA# (mã cố định) · ID# = FACTORYID · Dịch size (EN/FR/ES) và size China theo bảng size · Đủ / thừa size so với đơn của style-color · Tháng SX (nếu nhập) · "made in Vietnam" |
| Thành phần | Tên sợi và tỷ lệ % (khi nhập thành phần mong đợi ở Thiết lập, hoặc cột *Care content* trong tác nghiệp có dữ liệu) |
| Kiểm tra mắt | Biểu tượng giặt, chữ Trung Quốc, bố cục — hiển thị ảnh nhãn cạnh hình mẫu trong tác nghiệp |

## Cấu trúc tác nghiệp được nhận
- Ô `Mùa:`, `Style:`, `Item:`, `Thị trường:` và dòng `Kích thước nhãn: 20*88mm`.
- Bảng đơn hàng có tiêu đề `VPONO` (+ `CUSTSTYLE`/`MASTSTYLE`, `CUSTCOLORCODE`/`MASTCOLORCODE`, `CUSTSIZEDESC`/`MASTSIZEDESC`, `FACTORYID`, `Care content`).
- Bảng Style-color có `TECHPACKNO` + `Color code` / `VSD color code` để nhận màu tương đương (Z81QO = 81QO = QCZ).

## Lưu ý
Proof của ITL thường đã chuyển chữ thành nét (không còn text) nên web dùng OCR; mất khoảng 30–90 giây mỗi trang. Mã đọc nhầm kiểu Q↔0 được tự khớp với mã trong tác nghiệp và ghi chú lại.
