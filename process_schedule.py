import json
import pandas as pd
import os
import glob
from collections import defaultdict


def convert_day_to_text(day):
    """Chuyển đổi số ngày thành tên thứ"""
    day_map = {
        0: "Thứ 2",
        1: "Thứ 3", 
        2: "Thứ 4",
        3: "Thứ 5",
        4: "Thứ 6"
    }
    return day_map.get(day, f"Ngày {day}")


def convert_period_to_text(period):
    """Chuyển đổi số tiết thành tên tiết"""
    return f"Tiết {period + 1}"


def process_schedule(schedule):
    """
    Xử lý lịch học để gộp các tiết liên tiếp trong cùng ngày
    Input: Danh sách các schedule items
    Output: Danh sách các chuỗi lịch học đã được gộp
    """
    # Nhóm theo ngày và phòng
    day_groups = defaultdict(list)
    room_info = {}
    
    for item in schedule:
        day = item['day']
        period = item['period']
        room = item['room']
        
        day_groups[day].append(period)
        room_info[day] = room
    
    # Sắp xếp và gộp các tiết liên tiếp
    processed_schedule = []
    
    for day in sorted(day_groups.keys()):
        periods = sorted(day_groups[day])
        room = room_info[day]
        
        # Gộp các tiết liên tiếp
        groups = []
        if periods:
            current_group = [periods[0]]
            
            for i in range(1, len(periods)):
                if periods[i] == periods[i-1] + 1:  # Tiết liên tiếp
                    current_group.append(periods[i])
                else:  # Tiết không liên tiếp, bắt đầu nhóm mới
                    groups.append(current_group)
                    current_group = [periods[i]]
            
            groups.append(current_group)
        
        # Tạo chuỗi mô tả cho mỗi nhóm tiết
        day_text = convert_day_to_text(day)
        
        for group in groups:
            if len(group) == 1:
                period_text = convert_period_to_text(group[0])
            else:
                start_period = convert_period_to_text(group[0])
                end_period = convert_period_to_text(group[-1])
                period_text = f"{start_period} - {end_period}"
            
            processed_schedule.append({
                'day': day_text,
                'period': period_text,
                'room': room
            })
    
    return processed_schedule


def process_json_file(file_path):
    """Xử lý một file JSON và trả về DataFrame"""
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    all_rows = []
    max_schedule_items = 0
    
    # Tìm số lượng lịch học tối đa để tạo đủ cột
    for class_info in data['classes']:
        schedule = class_info['schedule']
        processed_schedule = process_schedule(schedule)
        max_schedule_items = max(max_schedule_items, len(processed_schedule))
    
    # Tạo tên cột động
    columns = ['Tên lớp']
    
    # Thêm cột cho từng lịch học
    for i in range(max_schedule_items):
        if i == 0:
            columns.extend(['Thứ', 'Tiết', 'Phòng'])
        else:
            columns.extend([f'Thứ {i+1}', f'Tiết {i+1}', f'Phòng {i+1}'])
    
    columns.append('Sỉ số')
    
    for class_info in data['classes']:
        name = class_info['name']
        # Kiểm tra xem có trường students không
        students_count = len(class_info.get('students', [])) if 'students' in class_info else "N/A"
        schedule = class_info['schedule']
        
        # Xử lý lịch học
        processed_schedule = process_schedule(schedule)
        
        # Tạo dòng dữ liệu cho Excel
        row = [name]  # Tên lớp
        
        # Thêm thông tin lịch học
        for i in range(max_schedule_items):
            if i < len(processed_schedule):
                schedule_item = processed_schedule[i]
                row.extend([
                    schedule_item['day'],
                    schedule_item['period'],
                    schedule_item['room']
                ])
            else:
                # Thêm ô trống nếu không có lịch học
                row.extend(['', '', ''])
        
        row.append(students_count)  # Sỉ số
        all_rows.append(row)
    
    # Tạo DataFrame với tên cột đã định nghĩa
    df = pd.DataFrame(all_rows, columns=columns)
    return df


def create_excel_from_json_files():
    """Đọc tất cả file JSON và tạo file Excel"""
    # Tìm tất cả file JSON trong thư mục hiện tại
    json_files = glob.glob("*.json")
    
    if not json_files:
        print("Không tìm thấy file JSON nào trong thư mục!")
        return
    
    # Tạo Excel writer
    output_file = "lich_hoc_tong_hop.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for json_file in json_files:
            print(f"Đang xử lý file: {json_file}")
            
            try:
                # Xử lý file JSON
                df = process_json_file(json_file)
                
                # Tên sheet là tên file không có extension
                sheet_name = os.path.splitext(json_file)[0]
                
                # Ghi vào Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Định dạng cột
                worksheet = writer.sheets[sheet_name]
                
                # Tự động điều chỉnh độ rộng cột
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                print(f"Đã xử lý xong file {json_file} -> sheet {sheet_name}")
                
            except Exception as e:
                print(f"Lỗi khi xử lý file {json_file}: {str(e)}")
    
    print(f"\nĐã tạo file Excel: {output_file}")
    print(f"Tổng cộng đã xử lý {len(json_files)} file JSON")


if __name__ == "__main__":
    create_excel_from_json_files()
