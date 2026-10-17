# SPX Converter Pro

Web upload RAR/ZIP/PDF/JPG/PNG/WEBP -> Excel mã vận đơn SPXVN.

## Chạy local trên Windows
1. Cài Python 3.11+, Tesseract OCR và `unar` (đảm bảo có trong PATH). Hoặc dùng Docker để tự cài các phụ thuộc.
2. `pip install -r requirements.txt`
3. `uvicorn app:app --reload`
4. Mở http://127.0.0.1:8000

## Chạy Docker / triển khai Render
- Đưa toàn bộ thư mục lên GitHub.
- Render > New > Web Service > kết nối repo.
- Chọn Runtime **Docker**, triển khai và mở URL Render cung cấp.
- Hoặc chạy: `docker build -t spx-converter .` rồi `docker run -p 10000:10000 spx-converter`.

## Ghi chú
- Chỉ trích xuất mã dạng `SPXVN...`; OCR có thể bỏ sót mã hoặc đọc nhầm ký tự.
- Excel có sheet Danh sách mã SPX, Thống kê, Cảnh báo.
- Giới hạn 30 file tải lên, 60MB/file, 1500 file trong mỗi archive, 250MB sau giải nén, PDF tối đa 150 trang.
- Tệp tạm tự xóa sau mỗi lần xử lý; không lưu lịch sử.
- Không dùng dịch vụ miễn phí để xử lý khối lượng cực lớn; tăng tài nguyên nếu cần.
