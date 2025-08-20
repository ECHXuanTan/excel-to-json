# Hướng dẫn sử dụng công cụ xử lý lịch học

## Mô tả
Công cụ này đọc các file JSON chứa thông tin lịch học và tạo file Excel tổng hợp.

## Cấu trúc file JSON
File JSON cần có cấu trúc như sau:
```json
{
  "classes": [
    {
      "name": "Tên lớp",
      "students": ["mã_sv_1", "mã_sv_2", ...],  // Tùy chọn
      "schedule": [
        {
          "room": "Tên phòng",
          "day": 0,     // 0-4 tương ứng Thứ 2-6
          "period": 0   // 0-9 tương ứng Tiết 1-10
        }
      ]
    }
  ]
}
```

## Cách sử dụng
1. Đặt tất cả file JSON vào cùng thư mục với file `process_schedule.py`
2. Chạy lệnh: `python process_schedule.py`
3. File Excel `lich_hoc_tong_hop.xlsx` sẽ được tạo

## Kết quả
- Mỗi file JSON sẽ trở thành 1 sheet trong Excel
- Thông tin hiển thị:
  - Tên lớp
  - Thứ, Tiết (được gộp nếu liên tiếp trong cùng ngày)  
  - Phòng học
  - Sỉ số (số học sinh hoặc "N/A" nếu không có)

## Yêu cầu hệ thống
- Python 3.x
- pandas: `pip install pandas`
- openpyxl: `pip install openpyxl`
