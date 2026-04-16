# Giải đáp thắc mắc

# Giải đáp thắc mắc

## 1. UI Cài đặt ở đâu?

✅ **Đã có UI Settings!** 

### Cách mở:
```
Menu: ⚙️ Cài đặt > 📊 Đọc Excel
```

### Tabs có sẵn:

#### 📍 **Tab "Phát hiện Header"**
- ✅ Bật/tắt tự động phát hiện header
- ✅ Số dòng quét tối đa (10-100 dòng)
- ✅ Hỗ trợ header nhiều cấp (merged cells)
- ✅ Ký tự ngăn cách (mặc định: "_")
- ✅ Số cột tối thiểu phải khớp
- ✅ Validate dòng data

#### 🔗 **Tab "Merged Cells"**
- ✅ Bật/tắt xử lý merged cells
- ✅ Forward fill merged cells
- ✅ Kết hợp với sub-header

#### ⚡ **Tab "Performance"**
- ✅ Bật/tắt caching
- ✅ Số file Excel cache (1-10)
- ✅ Cache kết quả tìm kiếm
- ✅ Cache thống kê
- ✅ Nút xóa cache

### Hoặc edit file JSON:
Nếu muốn edit trực tiếp:
```
C:\Users\[Username]\AppData\Local\StudentScoreImport\app_config.json
```

## 2. Tại sao không tìm thấy cột "ĐĐGck"?

### Nguyên nhân:
Hàm tìm kiếm cột chưa được cập nhật đầy đủ để sử dụng patterns mới.

### Đã fix:
- ✅ Cập nhật `find_matching_column()` để dùng config patterns
- ✅ Thêm variations: "ĐĐGck", "Điểm CK", "Điểm cuối kỳ", "CK"
- ✅ Hỗ trợ tìm kiếm linh hoạt hơn

### Cách test:
1. Đóng app
2. Mở lại app
3. Load file Excel
4. Kiểm tra xem cột ĐĐGck có được nhận diện không

## 3. Phân biệt "Mã đề" và "Mã định danh"

### Mã định danh Bộ GD&ĐT (Student ID)
- Ví dụ: `0142976685`, `1181872257`
- Là mã số học sinh do Bộ GD&ĐT cấp
- Dùng để định danh học sinh
- **Không dùng trong app này**

### Mã đề (Exam Code)
- Ví dụ: `701`, `702`, `703`, `604`
- Là mã đề thi trắc nghiệm
- Dùng để tính điểm theo đáp án đúng
- **Cần có cột này nếu dùng tính điểm trắc nghiệm**

### Trong file Excel của bạn:
```
✅ Có: Mã định danh Bộ GD&ĐT (0142976685...)
❌ Không có: Mã đề (701, 702...)
```

### Config hiện tại:
```json
{
  "columns": {
    "exam_code": "Mã đề"  // Đang tìm cột này nhưng không có trong Excel
  }
}
```

### Giải pháp:

**Option 1**: Nếu không cần tính điểm trắc nghiệm
```json
{
  "columns": {
    "name": "Họ và tên",
    "exam_code": "",        // Để trống
    "score": "ĐĐGck"
  }
}
```

**Option 2**: Nếu cần mã đề, thêm cột "Mã đề" vào Excel
```
| STT | Họ và tên     | Mã đề | ĐĐGck |
|-----|---------------|-------|-------|
| 1   | Nguyễn Bình An| 701   | 5     |
| 2   | Nguyễn Thái An| 702   | 5     |
```

## 4. Các cột được nhận diện tự động

App hiện có thể tự động nhận diện các variations:

### Tên học sinh:
- ✅ "Họ và tên"
- ✅ "Tên học sinh"
- ✅ "Họ tên"
- ✅ "Tên"
- ✅ "Học sinh"

### Điểm cuối kỳ:
- ✅ "ĐĐGck"
- ✅ "ĐĐGcuối kỳ"
- ✅ "Điểm CK"
- ✅ "Điểm cuối kỳ"
- ✅ "CK"

### Điểm giữa kỳ:
- ✅ "ĐĐGgk"
- ✅ "Điểm GK"
- ✅ "Điểm giữa kỳ"
- ✅ "GK"

### Điểm thường xuyên:
- ✅ "ĐĐGtx"
- ✅ "Điểm TX"
- ✅ "Điểm thường xuyên"
- ✅ "TX"

### Điểm trung bình:
- ✅ "ĐTBmhkI"
- ✅ "ĐTB HK1"
- ✅ "Điểm TB HK1"
- ✅ "Điểm trung bình HK1"

## 5. Test lại

### Bước 1: Check config
```json
{
  "columns": {
    "name": "Họ và tên",
    "exam_code": "",      // Để trống nếu không cần
    "score": "ĐĐGck"     // Hoặc cột điểm bạn muốn nhập
  }
}
```

### Bước 2: Restart app
1. Đóng app hoàn toàn
2. Mở lại

### Bước 3: Load file Excel
File > Mở file Excel > Chọn file

### Bước 4: Kiểm tra
- ✅ Có hiển thị danh sách học sinh?
- ✅ Cột "Họ và tên" có đúng?
- ✅ Cột "ĐĐGck" có được nhận diện?
- ✅ Không còn lỗi "Thiếu cột dữ liệu"?

## 6. Nếu vẫn lỗi

### Debug:
1. Check file Excel có đúng format không
2. Check tên cột có chính xác không (không có khoảng trắng thừa)
3. Check header có ở dòng 7-8 như trong ảnh không

### Logs:
App sẽ print ra console:
```
Đang tìm cột: ĐĐGck
Các cột tìm thấy: [list of columns]
Column matched: ĐĐGck
```

### Report:
Nếu vẫn lỗi, gửi cho tôi:
1. Screenshot lỗi
2. File Excel (hoặc ảnh header)
3. Nội dung config hiện tại

---

**Tóm tắt:**
- ✅ Đã fix hàm tìm cột để nhận diện tốt hơn
- ⏳ UI Settings sẽ được thêm sau
- ℹ️ "Mã đề" ≠ "Mã định danh" - có thể để trống nếu không cần tính điểm trắc nghiệm
