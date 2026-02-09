# CÔNG CỤ CHUYỂN ĐỔI PDF SANG EXCEL BẰNG AI

## 📋 Tổng quan

Công cụ này giúp bạn chuyển đổi file PDF chứa bảng dữ liệu thành file Excel, sử dụng AI (Claude) để nhận diện và trích xuất dữ liệu từ mỗi trang.

### Workflow:
```
PDF → Tách từng trang → AI OCR → Excel → Ghép lại 1 file
```

## 🔧 Cài đặt

### 1. Cài đặt Python packages:

```bash
pip install pypdf pdf2image pillow requests openpyxl pandas --break-system-packages
```

### 2. Cài đặt Poppler (để convert PDF sang ảnh):

**Ubuntu/Debian:**
```bash
sudo apt-get install poppler-utils
```

**macOS:**
```bash
brew install poppler
```

**Windows:**
- Tải Poppler: https://github.com/oschwartz10612/poppler-windows/releases
- Giải nén và thêm vào PATH

## 🚀 Cách sử dụng

### Chạy với file PDF:

```bash
python pdf_to_excel_ai.py input.pdf
```

### Quy trình chi tiết:

1. **Bước 1 - Tách PDF**: Tự động tách file PDF thành từng trang riêng lẻ
   
2. **Bước 2 - AI OCR**: 
   - Chuyển mỗi trang PDF thành ảnh
   - Gọi Claude API để phân tích bảng dữ liệu
   - Lưu kết quả thành file Excel riêng
   - Sau mỗi trang, bạn có thể kiểm tra và quyết định tiếp tục

3. **Bước 3 - Ghép Excel**: Gộp tất cả các sheet thành 1 file Excel duy nhất

## 📁 Cấu trúc thư mục output

```
output/
├── temp/
│   ├── pages/          # Các trang PDF đã tách
│   │   ├── page_001.pdf
│   │   ├── page_002.pdf
│   │   └── ...
│   ├── excel_sheets/   # Các file Excel từng trang
│   │   ├── page_001.xlsx
│   │   ├── page_002.xlsx
│   │   └── ...
│   └── page_*.png      # Ảnh tạm của từng trang
└── merged_excel_YYYYMMDD_HHMMSS.xlsx  # File Excel cuối cùng
```

## ⚙️ Cấu hình API

**LÀM SAO ĐỂ CHẠY ĐƯỢC?**

Script này cần API key để hoạt động. Hiện tại API key KHÔNG được bao gồm trong code vì lý do bảo mật.

### Cách thêm API key:

Mở file `.env.example` và thêm API key vào, sau đó rename thành `.env`

```
DEEPSEEK_API_KEY=""
GEMINI_API_KEY=""
CLAUDE_API_KEY=""
```

### Lấy API key:
1. Truy cập: https://console.anthropic.com/
2. Đăng nhập/Đăng ký
3. Tạo API key mới
4. Copy và paste vào code

## 💡 Ví dụ sử dụng

```bash
# Chuyển đổi file bảng giá đất
python pdf_to_excel_ai.py NQ100-pl-1.pdf

# Kết quả sẽ có trong thư mục output/
```

## 🔍 Kiểm tra từng bước

Sau khi chạy script, bạn có thể:

1. **Kiểm tra tách trang**: Xem thư mục `output/temp/pages/`
2. **Kiểm tra Excel từng trang**: Xem thư mục `output/temp/excel_sheets/`
3. **Xem file cuối cùng**: File `merged_excel_*.xlsx` trong thư mục `output/`

## ⚠️ Lưu ý quan trọng

1. **Chi phí API**: Mỗi lần gọi Claude API có thể tốn tiền. Với PDF 24 trang như của bạn, ước tính ~$0.5-1 USD

2. **Chất lượng ảnh**: Ảnh càng rõ nét, kết quả OCR càng tốt (DPI mặc định: 300)

3. **Định dạng bảng phức tạp**: Với bảng có nhiều cột và merged cells, kết quả có thể cần chỉnh sửa thủ công

4. **Rate limit**: API có giới hạn số request/phút, nếu PDF quá dài cần điều chỉnh

## 🛠️ Tùy chỉnh

### Thay đổi DPI của ảnh (chất lượng):

Trong file `pdf_to_excel_ai.py`, tìm dòng:
```python
images = convert_from_path(page_pdf, dpi=300)
```

Tăng lên `dpi=600` để có chất lượng cao hơn (nhưng tốn nhiều thời gian hơn)

### Thay đổi AI prompt:

Chỉnh sửa phần `text` trong hàm `_call_claude_api()` để AI hiểu đúng cấu trúc bảng của bạn

## 🐛 Xử lý lỗi thường gặp

### Lỗi: "poppler not found"
→ Chưa cài đặt poppler-utils (xem mục Cài đặt)

### Lỗi: "API key not found" 
→ Chưa thêm API key vào code (xem mục Cấu hình API)

### Lỗi: "JSONDecodeError"
→ AI trả về không đúng format JSON, có thể do ảnh quá mờ hoặc bảng quá phức tạp

## 📞 Hỗ trợ

Nếu gặp vấn đề, hãy kiểm tra:
1. File PDF có mở được không?
2. Đã cài đủ dependencies chưa?
3. API key có hợp lệ không?
4. Thư mục output/ có quyền ghi không?
# pdf_to_excel_ai
