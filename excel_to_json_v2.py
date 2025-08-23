import pandas as pd
import json
import re
import os


def parse_day_text(day_text):
    """Chuyển đổi text thứ thành số (Thứ 2=0, Thứ 3=1, ...)"""
    if pd.isna(day_text) or not day_text:
        return None
    
    day_text = str(day_text).strip()
    if day_text == "" or day_text == "nan":
        return None
    
    day_map = {
        "Thứ 2": 0,
        "Thứ 3": 1, 
        "Thứ 4": 2,
        "Thứ 5": 3,
        "Thứ 6": 4
    }
    
    return day_map.get(day_text, None)


def parse_period_text(period_text):
    """Chuyển đổi text tiết thành danh sách số (1-2 -> [0,1], 3-5 -> [2,3,4])"""
    if pd.isna(period_text) or not period_text:
        return []
    
    period_text = str(period_text).strip()
    if period_text == "" or period_text == "nan":
        return []
    
    # Xử lý khoảng tiết (ví dụ: 1-2, 3-5, 6-9)
    if '-' in period_text:
        parts = period_text.split('-')
        if len(parts) == 2:
            try:
                start = int(parts[0]) - 1  # Chuyển từ 1-based sang 0-based
                end = int(parts[1]) - 1
                return list(range(start, end + 1))
            except ValueError:
                return []
    else:
        # Xử lý tiết đơn lẻ
        try:
            return [int(float(period_text)) - 1]  # Chuyển từ 1-based sang 0-based, xử lý cả float
        except ValueError:
            return []
    
    return []


def parse_class_name(class_text):
    """Tách tên lớp từ text và format theo yêu cầu (thêm GDTC, viết thường)"""
    if pd.isna(class_text) or not class_text:
        return None
    
    class_text = str(class_text).strip()
    if class_text == "" or class_text == "nan":
        return None
    
    # Tách theo xuống dòng - chỉ lấy dòng đầu tiên làm tên lớp
    lines = class_text.split('\n')
    
    if len(lines) >= 1:
        class_name = lines[0].strip()
    elif len(lines) == 1:
        # Nếu chỉ có 1 dòng, thử tách bằng regex để lấy phần tên lớp
        # Tìm pattern: tên lớp + mã phòng (A703, B505, etc.)
        match = re.match(r'^(.+?)\s*([A-Z]\d+)$', class_text)
        if match:
            class_name = match.group(1).strip()
        else:
            class_name = class_text
    else:
        return None
    
    if not class_name:
        return None
    
    # Xử lý format tên lớp
    # Nếu đã có "GDTC" ở đầu thì bỏ đi để tránh trùng lặp
    if class_name.startswith("GDTC "):
        class_name = class_name[5:]  # Bỏ "GDTC " ở đầu
    
    # Tách thành các phần để xử lý
    parts = class_name.split()
    if len(parts) >= 2:
        # Xử lý từng phần
        formatted_parts = []
        for i, part in enumerate(parts):
            if i == 0:  # Phần đầu tiên (số lớp)
                formatted_parts.append(part)
            else:
                # Kiểm tra nếu có cụm -LN (giữ nguyên viết hoa)
                if '-LN' in part:
                    # Tách phần trước và sau -LN
                    ln_parts = part.split('-LN')
                    if len(ln_parts) == 2:
                        # Phần trước -LN viết thường, phần -LN giữ nguyên
                        before_ln = ln_parts[0]
                        after_ln = ln_parts[1]
                        if before_ln.upper() == before_ln and len(before_ln) > 1:
                            before_ln = before_ln.capitalize()
                        formatted_part = f"{before_ln}-LN{after_ln}"
                        formatted_parts.append(formatted_part)
                    else:
                        formatted_parts.append(part)
                else:
                    # Các phần khác: viết thường chữ cái đầu, giữ nguyên phần còn lại
                    if part.upper() == part and len(part) > 1:  # Nếu toàn bộ là chữ hoa
                        formatted_parts.append(part.capitalize())
                    else:
                        formatted_parts.append(part)
        
        formatted_name = " ".join(formatted_parts)
    else:
        formatted_name = class_name
    
    # Thêm "GDTC" vào đầu
    return f"GDTC {formatted_name}"


def excel_to_json(excel_file, output_file=None):
    """Chuyển đổi file Excel thành JSON"""
    
    # Đọc file Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"Lỗi khi đọc file Excel: {e}")
        return None
    
    classes = []
    
    # Duyệt qua từng dòng
    for index, row in df.iterrows():
        # Lấy tên lớp từ cột đầu tiên
        first_col = row.iloc[0]  # Cột đầu tiên
        class_name = parse_class_name(first_col)
        
        if not class_name:
            continue
        
        schedule = []
        
        # Duyệt qua các cặp cột Thứ-Tiết
        col_index = 1  # Bắt đầu từ cột thứ 2
        while col_index < len(row) - 1:  # Đảm bảo còn ít nhất 2 cột
            day_text = row.iloc[col_index] if col_index < len(row) else None
            period_text = row.iloc[col_index + 1] if col_index + 1 < len(row) else None
            
            # Parse thứ và tiết
            day_num = parse_day_text(day_text)
            periods = parse_period_text(period_text)
            
            # Tạo schedule entries cho từng tiết
            if day_num is not None and periods:
                for period in periods:
                    schedule.append({
                        "room": "Sân trường",
                        "day": day_num,
                        "period": period
                    })
            
            # Chuyển sang cặp cột tiếp theo
            col_index += 2
        
        # Thêm class vào danh sách nếu có schedule
        if class_name and schedule:
            classes.append({
                "name": class_name,
                "originalClassId": "",
                "schedule": schedule
            })
    
    # Tạo JSON output
    json_data = {
        "classes": classes
    }
    
    # Ghi ra file nếu có output_file
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"Đã tạo file JSON: {output_file}")
    
    return json_data


def convert_all_excel_files():
    """Chuyển đổi tất cả file Excel trong thư mục hiện tại"""
    excel_files = []
    
    # Tìm tất cả file Excel
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print("Không tìm thấy file Excel nào trong thư mục!")
        return
    
    for excel_file in excel_files:
        print(f"Đang chuyển đổi: {excel_file}")
        
        # Tạo tên file JSON output
        base_name = os.path.splitext(excel_file)[0]
        json_file = f"{base_name}_converted_v2.json"
        
        try:
            result = excel_to_json(excel_file, json_file)
            if result:
                print(f"Đã chuyển đổi thành công: {excel_file} -> {json_file}")
                print(f"Số lớp được xử lý: {len(result['classes'])}")
            else:
                print(f"Lỗi khi chuyển đổi file: {excel_file}")
        except Exception as e:
            print(f"Lỗi khi xử lý file {excel_file}: {e}")
        
        print("-" * 50)


if __name__ == "__main__":
    print("=== CÔNG CỤ CHUYỂN ĐỔI EXCEL SANG JSON (VERSION 2) ===\n")
    convert_all_excel_files()
