# BÀI LUYỆN TẬP NGÀY 3 – NHẬP MÔN CNTT

## Bài 1. Thao tác với tài liệu Markdown

### 1.1. Tạo file README.md

Tạo file `README.md` tại thư mục gốc `trung-tam-tieng-anh` với nội dung sau:

# Trung Tâm Tiếng Anh TTA

## 1. Giới thiệu

**Trung tâm tiếng Anh TTA** cung cấp thông tin về các khóa học, học viên và học phí. Dự án hướng đến việc *quản lý dữ liệu rõ ràng, khoa học và dễ sử dụng*.

## 2. Cấu trúc thư mục

| Thư mục    | Nội dung                                   |
| ---------- | ------------------------------------------ |
| `archive/` | Lưu trữ tài liệu cũ                        |
| `docs/`    | Chứa tài liệu dự án                        |
| `logs/`    | Lưu nhật ký hoạt động                      |
| `media/`   | Chứa hình ảnh và tài nguyên đa phương tiện |
| `data/`    | Chứa bảng tính học viên và học phí         |

## 3. Quy trình đăng ký học

1. Học viên tìm hiểu các khóa học tại TTA.
2. Học viên chọn khóa học phù hợp.
3. Nhân viên ghi nhận thông tin đăng ký.
4. Trung tâm xác nhận thời gian học và học phí.
5. Học viên hoàn thành thủ tục thanh toán.

## 4. Lệnh kiểm tra dự án

```powershell
Get-Location
Get-ChildItem -Force
Get-ChildItem .\data
```

## 5. Công việc thực hiện

* [x] Tạo cấu trúc tài liệu README.md.
* [ ] Hoàn thiện công thức tính học phí trong Excel.
* [ ] Kiểm tra dữ liệu đăng ký và lưu kết quả.

## 6. Lưu ý bảo mật

> Không tải tệp `.env` lên GitHub vì tệp này có thể chứa thông tin bí mật hoặc cấu hình nhạy cảm.

---

## Bài 2. Thực hành Excel

### 2.1. Tra cứu tên khóa học bằng VLOOKUP

**Yêu cầu:** Sheet `DangKy`, lấy tên khóa học từ sheet `KhoaHoc` cho 15 lượt đăng ký.

Tại ô `E2` của sheet `DangKy`, nhập công thức:

```excel
=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,2,FALSE)
```

**Giải thích:**

* `C2`: Mã khóa học cần tra cứu.
* `KhoaHoc!$A$2:$C$8`: Vùng dữ liệu tra cứu.
* `2`: Trả về dữ liệu ở cột thứ hai của vùng tra cứu.
* `FALSE`: Tra cứu khớp chính xác.

Sau đó kéo công thức từ `E2` xuống `E16`.

### 2.2. Tra cứu học phí bằng VLOOKUP

Tại ô `F2` của sheet `DangKy`, nhập công thức:

```excel
=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,3,FALSE)
```

**Giải thích:**

* `C2`: Mã khóa học cần tra cứu.
* `KhoaHoc!$A$2:$C$8`: Vùng dữ liệu chứa mã khóa, tên khóa và học phí.
* `3`: Trả về học phí ở cột thứ ba.
* `FALSE`: Yêu cầu khớp chính xác.

Kéo công thức từ `F2` xuống `F16` để hoàn thiện 15 lượt đăng ký.

### 2.3. Tính tổng học phí

Tại ô `G2` của sheet `DangKy`, nhập công thức:

```excel
=D2*F2
```

Trong đó:

* `D2`: Số tháng đăng ký.
* `F2`: Học phí mỗi tháng.
* `G2`: Tổng học phí cần thanh toán.

Kéo công thức từ `G2` xuống `G16`.

**Lưu ý:** Không nhập tay tên khóa học, học phí hoặc tổng học phí thay cho công thức.

### 2.4. Tra cứu học phí tại sheet KhoaHoc

Tại ô `F2` của sheet `KhoaHoc`, nhập công thức:

```excel
=VLOOKUP(E2,$A$2:$C$8,3,FALSE)
```

**Giải thích:**

* `E2`: Mã khóa học cần tra cứu, ban đầu là `IE02`.
* `$A$2:$C$8`: Vùng dữ liệu tra cứu được cố định bằng dấu `$`.
* `3`: Lấy học phí ở cột thứ ba.
* `FALSE`: Tra cứu chính xác.

Thử thay đổi mã khóa học tại `E2` để kiểm tra kết quả. Sau đó đặt lại giá trị `IE02`.

### 2.5. Sử dụng AutoFilter

Thực hiện trên sheet `DangKy`:

1. Chọn toàn bộ bảng dữ liệu, bao gồm hàng tiêu đề và các dòng từ 2 đến 16.
2. Chọn **Data → Filter**.
3. Nhấn mũi tên lọc ở cột `tong_hoc_phi`.
4. Chọn **Number Filters → Greater Than Or Equal To**.
5. Nhập giá trị `6000000`.
6. Nhấn **OK**.
7. Đếm số dòng dữ liệu còn hiển thị, không tính hàng tiêu đề.
8. Lưu file Excel.

**Kết quả lọc:** Có `...` dòng đăng ký có tổng học phí lớn hơn hoặc bằng 6.000.000 đồng.

*Lưu ý: Thay dấu `...` bằng số dòng thực tế sau khi lọc dữ liệu trong Excel.*

### 2.6. Giải thích tác dụng của dấu `$` trong Excel

Dấu `$` được sử dụng để cố định địa chỉ ô hoặc vùng dữ liệu khi sao chép công thức trong Excel.

Ví dụ:

```excel
=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,2,FALSE)
```

Trong công thức trên, vùng `KhoaHoc!$A$2:$C$8` được cố định để không thay đổi khi kéo công thức xuống các dòng bên dưới.

Nếu bỏ dấu `$`, vùng tra cứu `KhoaHoc!A2:C8` có thể tự động dịch chuyển thành `KhoaHoc!A3:C9`, `KhoaHoc!A4:C10` khi sao chép công thức xuống.

Điều này có thể khiến công thức tra cứu sai vùng dữ liệu và trả về kết quả không chính xác hoặc lỗi `#N/A`.

**Kết luận:** Cần sử dụng dấu `$` để cố định vùng tra cứu khi sao chép công thức cho nhiều dòng.

---

## Bài 3. Kiểm tra và hoàn thiện bài làm

### 3.1. Kiểm tra cấu trúc thư mục

Dự án sau khi hoàn thành có cấu trúc cơ bản như sau:

```text
trung-tam-tieng-anh/
├── .env
├── README.md
├── bai_lam_ngay3.md
├── archive/
├── docs/
├── logs/
├── media/
└── data/
    ├── hoc_vien_khoa_3.xlsx
    └── hoc_phi_thang_10.xlsx
```

### 3.2. Danh sách kiểm tra trước khi nộp

* [ ] Tạo file `README.md` tại thư mục gốc dự án.
* [ ] README có tiêu đề, các heading, bảng, danh sách đánh số và code block.
* [ ] README có task list và blockquote nhắc không tải `.env` lên GitHub.
* [ ] Sao chép file mẫu `bai_lam_ngay3.md` vào thư mục gốc.
* [ ] Hoàn thiện công thức VLOOKUP tại các ô `E2`, `F2` của sheet `DangKy`.
* [ ] Hoàn thiện công thức tính tổng học phí tại `G2:G16`.
* [ ] Hoàn thiện công thức tra học phí tại ô `F2` của sheet `KhoaHoc`.
* [ ] Bật AutoFilter và lọc tổng học phí từ 6.000.000 đồng trở lên.
* [ ] Ghi số dòng thực tế còn hiển thị vào bài làm.
* [ ] Lưu file Excel sau khi hoàn thành.
* [ ] Kiểm tra lại toàn bộ nội dung trước khi nộp.
