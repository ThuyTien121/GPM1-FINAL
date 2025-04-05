# Cách Chạy Phần Mềm Phân Tích Tài Chính (Windows)

## LƯU Ý QUAN TRỌNG VỀ MÔI TRƯỜNG ẢO

**QUAN TRỌNG**: Khi tải project từ GitHub hoặc từ máy tính khác về, bạn PHẢI xóa thư mục môi trường ảo cũ (thư mục `venv`) và tạo môi trường ảo mới trên máy của bạn. Điều này là bắt buộc vì môi trường ảo chứa các đường dẫn tuyệt đối và cấu hình đặc thù cho máy tính đã tạo ra nó.

### Nhận biết và xóa môi trường ảo cũ:
1. Môi trường ảo thường nằm trong thư mục `venv` hoặc `.venv` trong thư mục dự án
2. Xóa thư mục này bằng cách:
   ```
   rmdir /s /q venv
   ```

## Tạo Môi Trường Ảo Mới

1. Mở Command Prompt (hoặc PowerShell) với quyền quản trị viên
2. Điều hướng tới thư mục dự án:
   ```
   cd đường-dẫn-tới-thư-mục-dự-án
   ```

3. Tạo môi trường ảo mới:
   ```
   python -m venv venv
   ```

4. Kích hoạt môi trường ảo:
   ```
   venv\Scripts\activate
   ```
   (Command Prompt sẽ hiển thị prefix `(venv)` ở đầu dòng khi kích hoạt thành công)

## Cài Đặt Các Thư Viện Cần Thiết

Sau khi đã tạo và kích hoạt môi trường ảo, cài đặt các thư viện:

1. Nâng cấp pip (công cụ cài đặt):
   ```
   python -m pip install --upgrade pip
   ```

2. Cài đặt các thư viện chính:
   ```
   pip install flask pandas numpy matplotlib seaborn
   ```

3. Cài đặt thư viện đọc Excel:
   ```
   pip install openpyxl xlrd
   ```

4. Cài đặt thư viện xuất báo cáo PDF:
   ```
   pip install weasyprint
   ```
   (Đảm bảo bạn đã cài đặt GTK3 như hướng dẫn trước đó)

5. Các thư viện khác:
   ```
   pip install plotly
   ```

6. Hoặc cài đặt tất cả từ file requirements.txt nếu có:
   ```
   pip install -r requirements.txt
   ```

## Cấu Trúc Thư Mục

Đảm bảo cấu trúc thư mục dự án như sau:
```
financial_analysis_app/
├── app.py
├── requirements.txt
├── static/
│   ├── css/
│   ├── js/
│   └── img/
├── templates/
│   ├── base.html
│   ├── index.html
│   ├── sector_analysis.html
│   └── ...
├── utils/
│   └── financial_calculations.py
└── data/
    ├── Average_by_Code.csv
    ├── Average_by_Sector.csv
    ├── BCDKT.csv
    └── ...
```

**Lưu ý**: Thư mục `venv` sẽ được tạo mới bởi các bước phía trên và không nên được sao chép từ máy khác.

## Chạy Phần Mềm

1. Trong Command Prompt với môi trường ảo đã kích hoạt (có hiển thị `(venv)`):
   ```
   python app.py
   ```

2. Nếu không có lỗi, bạn sẽ thấy thông báo:
   ```
   * Running on http://127.0.0.1:5000/ (Press CTRL+C to quit)
   ```

3. Mở trình duyệt web và truy cập địa chỉ:
   ```
   http://localhost:5000
   ```

## Lỗi Thường Gặp và Cách Khắc Phục

### Lỗi khi tạo môi trường ảo
Nếu gặp lỗi "Access is denied" khi tạo môi trường ảo:
- Đảm bảo bạn đã chạy Command Prompt với quyền Admin
- Thử tạo môi trường ảo ở thư mục khác mà bạn có quyền ghi

### Lỗi thiếu thư viện
Nếu gặp lỗi `ModuleNotFoundError: No module named 'xxx'`:
- Đảm bảo môi trường ảo đã được kích hoạt (hiển thị `(venv)`)
- Cài đặt thư viện còn thiếu: `pip install xxx`

### Lỗi không tìm thấy file dữ liệu
- Kiểm tra xem các file CSV có đúng tên và nằm trong thư mục `data/`
- Nếu báo lỗi encoding, hãy đảm bảo các file CSV là UTF-8

### Lỗi khi xuất PDF
Nếu chức năng xuất PDF không hoạt động:
- Kiểm tra đã cài đặt GTK3 đúng cách (xem hướng dẫn trước)
- Kiểm tra WeasyPrint đã được cài đặt: `pip show weasyprint`
- Thử khởi động lại máy tính để cập nhật biến môi trường PATH

## Nhắc Nhở Cuối Cùng

1. **KHÔNG BAO GIỜ** sao chép thư mục môi trường ảo (`venv`) từ máy tính khác
2. **LUÔN** tạo môi trường ảo mới trên mỗi máy tính
3. **LUÔN** cài đặt lại tất cả thư viện trong môi trường ảo mới
4. Đảm bảo cài đặt GTK3 đúng cách nếu muốn sử dụng chức năng xuất PDF

Làm theo các bước trên sẽ giúp bạn tránh nhiều lỗi phổ biến khi chạy phần mềm Python trên các máy tính khác nhau!