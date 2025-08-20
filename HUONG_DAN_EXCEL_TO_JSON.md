# Hướng dẫn chuyển đổi Excel sang JSON

## Mô tả
Công cụ này chuyển đổi file Excel chứa thông tin lịch học thành file JSON theo định dạng chuẩn.

## Format Excel đầu vào

### Cấu trúc cột:
1. **Cột đầu tiên**: Tên lớp và phòng học (cách nhau bằng xuống dòng)
2. **Các cột tiếp theo**: Cặp "Thứ" và "Tiết" (lặp lại cho nhiều buổi học)

### Ví dụ format Excel:

| Lớp | Thứ | Tiết | Thứ | Tiết |
|-----|-----|------|-----|------|
| 10 AĐ VẬT LÝ TC 1<br/>A703 | Thứ 4 | 1-2 | Thứ 6 | 3-4 |
| 10 LÝ CHUYÊN<br/>A702 | Thứ 2 | 3-5 | | |

### Quy tắc format:
- **Tên lớp và phòng**: Viết trong cùng 1 ô, cách nhau bằng xuống dòng
- **Thứ**: "Thứ 2", "Thứ 3", ..., "Thứ 6"
- **Tiết**: 
  - Tiết đơn: "1", "2", "3"...
  - Khoảng tiết: "1-2", "3-5", "6-9"...
- **Ô trống**: Để trống nếu không có lịch học

## Cách sử dụng

### 1. Chuyển đổi file cụ thể:
```python
from excel_to_json import excel_to_json

# Chuyển đổi 1 file Excel
result = excel_to_json('lich_hoc.xlsx', 'output.json')
```

### 2. Chuyển đổi tất cả file Excel:
```bash
python excel_to_json.py
```

## Format JSON đầu ra

```json
{
  "classes": [
    {
      "name": "10 AĐ VẬT LÝ TC 1",
      "schedule": [
        {
          "room": "A703",
          "day": 2,        // Thứ 4 = 2 (Thứ 2=0, Thứ 3=1, ...)
          "period": 0      // Tiết 1 = 0 (Tiết 1=0, Tiết 2=1, ...)
        },
        {
          "room": "A703", 
          "day": 2,
          "period": 1      // Tiết 2 = 1
        }
      ]
    }
  ]
}
```

## Chuyển đổi số:
- **Thứ**: Thứ 2→0, Thứ 3→1, Thứ 4→2, Thứ 5→3, Thứ 6→4
- **Tiết**: Tiết 1→0, Tiết 2→1, Tiết 3→2, ...

## Lưu ý
- File Excel phải có định dạng .xlsx hoặc .xls
- Tên phòng PHẢI được viết trong cùng ô với tên lớp
- Không tạo trường "students" trong JSON
- Code tự động xử lý khoảng tiết (VD: "3-5" → [2,3,4])

## Yêu cầu hệ thống
```bash
pip install pandas openpyxl
```
