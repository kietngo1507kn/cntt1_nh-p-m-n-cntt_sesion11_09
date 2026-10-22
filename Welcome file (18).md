**_Phần 1: Phân tích bài toán (I/O)_**

**1.** Xác định rõ: Lookup_value là gì, Table_array là vùng nào , Col_index_num cần gì cho từng cột D, E, F. Trình bày dưới dạng bảng 4 cột.
**Cột cần điền (DanhSachBaoCao)**

**Lookup_value (Giá trị tra cứu)**

**Table_array (Vùng tra cứu)**

**Col_index_num (Cột lấy dữ liệu)**

**Range_lookup (Kiểu tra cứu)**

**Cột D:** Tên Tác Giả

Mã TG tại ô `C2`

`'DanhMucTacGia'!$A$2:$E$100` _(hoặc toàn cột `'DanhMucTacGia'!$A:$E`)_

**2** (Cột B: Tên Tác Giả)

`0` (hoặc `FALSE` - tra cứu chính xác)

**Cột E:** Phòng trình bày

Mã TG tại ô `C2`

`'DanhMucTacGia'!$A$2:$E$100`

**4** (Cột D: Phòng)

`0` (hoặc `FALSE`)

**Cột F:** Thù lao

Mã TG tại ô `C2`

`'DanhMucTacGia'!$A$2:$E$100`

**5** (Cột E: Thù lao)

`0` (hoặc `FALSE`)
**2.** Đề xuất giải pháp: Mô tả ngắn gọn chiến lược xử lý 3 bẫy dữ liệu nêu trên.
-   **Xử lý Bẫy 1 (Khoảng trắng thừa):** Dùng hàm `TRIM(C2)` làm `Lookup_value`. Hàm này sẽ cắt bỏ toàn bộ dấu cách thừa ở đầu/cuối của chuỗi mã tác giả (biến `'TG001 '` thành `'TG001'`).
    
-   **Xử lý Bẫy 2 (Mã chưa có trong danh mục):** Lồng toàn bộ hàm `VLOOKUP` vào trong hàm `IFERROR(...)`. Cú pháp:  `IFERROR(VLOOKUP(...), "Chưa cập nhật")`. Khi VLOOKUP không tìm thấy mã và trả về lỗi `#N/A`,  `IFERROR` sẽ bắt lỗi này và hiển thị văn bản thân thiện `"Chưa cập nhật"`.
    
-   **Xử lý Bẫy 3 (Bảng tra cứu bị trôi):** Cố định tuyệt đối vùng tra cứu bằng dấu `$` (`'DanhMucTacGia'!$A$2:$E$100` hoặc dùng địa chỉ cột `'DanhMucTacGia'!$A:$E`) để khi sao chép công thức xuống các dòng dưới, tọa độ bảng tra giữ nguyên.
**3.** Thiết kế các bước thực hiện: Gạch đầu dòng tuần tự từng bước trong Excel.
-   **Bước 1: Chuẩn bị tại ô D2 (Sheet `DanhSachBaoCao`)**
    
    Chọn ô `D2`, nhập công thức xử lý lồng ghép (`IFERROR` + `VLOOKUP` + `TRIM`) để lấy Tên tác giả.
    
-   **Bước 2: Thiết lập cho ô E2 và F2**
    
    Sao chép công thức từ ô `D2` sang ô `E2` và `F2`, sau đó sửa lại tham số `Col_index_num` tương ứng (thành `4` cho Phòng và `5` cho Thù lao).
    
-   **Bước 3: Sao chép công thức xuống toàn bộ 80 dòng (D2:F81)**
    
    Bôi đen vùng `D2:F2`, đưa con trỏ chuột vào góc dưới bên phải (AutoFill handle) đến khi xuất hiện dấu cộng đen `+`, nhấp kép chuột (Double-click) hoặc kéo xuống hết dòng 81.
    
-   **Bước 4: Định dạng dữ liệu (Formatting)**
    
    Bôi đen cột F (`F2:F81`), nhấn `Ctrl + 1` chọn định dạng **Currency** hoặc **Accounting** (đơn vị VNĐ) để hiển thị thù lao dạng phân cách hàng nghìn (ví dụ:  `500,000`).
**_Phần 2: Triển khai & Chống lỗi_**

**4.** Viết công thức hoàn chỉnh cho ô D2 (Tên Tác Giả) có xử lý đầy đủ 3 bẫy: TRIM lookup_value + cố định vùng bằng $ + IFERROR trả 'Chưa cập nhật'.
Excel

```
=IFERROR(VLOOKUP(TRIM(C2), 'DanhMucTacGia'!$A$2:$E$100, 2, FALSE), "Chưa cập nhật")
```
**5.** Viết công thức cho ô E2 (Phòng) và F2 (Thù lao). Giải thích tại sao chỉ cần thay Col_index_num chứ không cần viết lại toàn bộ cú pháp.
Excel

```
=IFERROR(VLOOKUP(TRIM(C2), 'DanhMucTacGia'!$A$2:$E$100, 4, FALSE), "Chưa cập nhật")
```
Excel

```
=IFERROR(VLOOKUP(TRIM(C2), 'DanhMucTacGia'!$A$2:$E$100, 5, FALSE), "Chưa cập nhật
```
1.  Cùng một đối tượng cần tra cứu (`TRIM(C2)` - Mã tác giả).
    
2.  Cùng một bảng chứa dữ liệu nguồn (`'DanhMucTacGia'!$A$2:$E$100`).
    
3.  Cùng cơ chế tra chính xác và bắt lỗi (`IFERROR` + `FALSE`).
    

	Điểm khác biệt duy nhất là vị trí cột dữ liệu cần rút trích ra khỏi bảng `A:E`: Cột Tên là cột 2, cột Phòng là cột 4, cột Thù lao là cột 5. Do đó chỉ cần đổi tham số `Col_index_num` tương ứng.
**6.** Giải thích hiện tượng: Nếu bỏ $ trong Table_array, khi kéo công thức D2 xuống D5, vùng tra cứu sẽ thay đổi thành gì? Cho ví dụ cụ thể với địa chỉ ô.
Khi bạn bỏ dấu `$` trong `Table_array` (viết dạng địa chỉ tương đối:  `'DanhMucTacGia'!A2:E100`), cơ chế tự động điều chỉnh địa chỉ của Excel sẽ khiến vùng tra cứu bị lệch xuống tương ứng với từng dòng bạn copy:

-   Tại ô **D2**, vùng tra cứu là:  `'DanhMucTacGia'!A2:E100` (đúng).
    
-   Khi kéo công thức xuống ô **D3**, vùng tra cứu bị trôi thành:  `'DanhMucTacGia'!A3:E101`.
    
-   Khi kéo xuống ô **D4**, vùng tra cứu biến thành:  `'DanhMucTacGia'!A4:E102`.
    
-   Khi kéo xuống ô **D5**, vùng tra cứu biến thành:  `'DanhMucTacGia'!A5:E103`.
    

**Hậu quả thực tế:**

-   Nếu bài báo cáo ở dòng 5 (ô `C5`) mang mã tác giả là **TG001** (nằm ở ô `A2` của sheet `DanhMucTacGia`), Excel chỉ thực hiện tìm kiếm từ `A5` trở đi.
    
-   Do ô `A2` đã bị loại khỏi vùng tra cứu `A5:E103`, hàm `VLOOKUP` sẽ không tìm thấy **TG001** và báo lỗi `#N/A` (hoặc trả về `"Chưa cập nhật"` nếu có `IFERROR`), dẫn đến sai lệch kết quả nghiêm trọng mặc dù tác giả đó hoàn toàn có trong danh mục.
