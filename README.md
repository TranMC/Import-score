# Ứng dụng Quản lý Điểm Học Sinh

Ứng dụng giúp giáo viên dễ dàng quản lý và thống kê điểm số học sinh, xuất báo cáo, và phân tích kết quả.

## Tính năng chính

- Nhập điểm từ file Excel
- Tìm kiếm học sinh nhanh chóng
- Tính điểm tự động từ số câu đúng
- Xuất báo cáo thống kê dạng PDF
- Biểu đồ phân phối điểm số
- Thống kê tỷ lệ đạt/không đạt
- Sao lưu và khôi phục dữ liệu có mã hóa

## Cài đặt

### Sử dụng file EXE (Windows)

1. Tải file EXE mới nhất từ [trang Releases](https://github.com/TranMC/Import-score/releases)
2. Chạy file để sử dụng ứng dụng

### Phát triển (Development)

```bash
# Cài đặt dependencies
pip install -r requirements.txt

# Chạy ứng dụng
python import_score.py

# Build ứng dụng
python build_optimized.py
```

## Tự động triển khai (CI/CD)

Dự án sử dụng GitHub Actions để tự động:
- Kiểm tra lỗi cú pháp
- Build ứng dụng
- Tạo phiên bản và changelog
- Phát hành bản mới

## Phiên bản

Xem lịch sử phiên bản và thay đổi trong [CHANGELOG.md](CHANGELOG.md)
