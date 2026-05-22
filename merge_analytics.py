import os
import glob
import re
import pandas as pd

# 硬编码绝对路径
root_folder = '/Users/yiiing/Desktop/Work/Excel-Auto-Fetch/Students self-paced learning'
output_file = '/Users/yiiing/Desktop/Work/Excel-Auto-Fetch/Final_All_Weeks_Analytics.xlsx'

all_students_data = {}
student_name_to_key = {}
column_order = {}

print("==================================================")
print("🚀 超级明朗版脚本启动（全量 Header 监控模式）...")
print("📍 目标路径:", root_folder)
print("==================================================")

if not os.path.exists(root_folder):
    print("❌ 路径不存在，请检查！")
    exit()

all_folders = sorted(
    [f for f in os.listdir(root_folder) if os.path.isdir(os.path.join(root_folder, f)) and f.startswith('Week')],
    key=lambda x: int(''.join(filter(str.isdigit, x))) if any(char.isdigit() for char in x) else 0
)
print(f"扫描到根目录下的所有项: {all_folders}\n")


def normalize_text(value):
    if pd.isna(value):
        return ''
    return re.sub(r'\s+', '', str(value)).strip().lower()


def normalize_student_id(value):
    if pd.isna(value):
        return None
    candidate = str(value).strip()
    if candidate == '' or candidate.lower() == 'not started':
        return None
    return candidate


def is_blank_value(value):
    if pd.isna(value):
        return True
    text = str(value).strip().lower()
    if text == '' or text == 'not started':
        return True
    try:
        return float(text) == 0
    except Exception:
        return False


def should_be_blank_in_output(value, col_name):
    """判断值在输出中是否应该显示为空（如viewed列的no）"""
    if is_blank_value(value):
        return True
    text = str(value).strip().lower()
    if text == 'no' and 'viewed' in col_name.lower():
        return True
    return False


def find_target_sheets(xl):
    """返回所有匹配的sheet列表，优先返回analytics per student，如果没有则返回analytics per learning step"""
    student_sheets = [sheet for sheet in xl.sheet_names if 'analytics per student' in sheet.lower()]
    if student_sheets:
        return student_sheets
    
    learning_sheets = [sheet for sheet in xl.sheet_names if 'analytics per learning step' in sheet.lower()]
    if learning_sheets:
        return learning_sheets
    
    return []



def choose_student_key(student_id, student_name):
    name_key = normalize_text(student_name)
    if student_id:
        id_key = f'ID::{student_id}'
        if name_key and name_key in student_name_to_key:
            existing_key = student_name_to_key[name_key]
            if existing_key != id_key:
                if existing_key.startswith('NAME::') and existing_key in all_students_data:
                    existing_data = all_students_data.pop(existing_key)
                    if id_key not in all_students_data:
                        all_students_data[id_key] = existing_data
                    else:
                        all_students_data[id_key].update(existing_data)
                    for key_name, mapped_key in list(student_name_to_key.items()):
                        if mapped_key == existing_key:
                            student_name_to_key[key_name] = id_key
        if id_key not in all_students_data:
            all_students_data[id_key] = {
                'Student ID': student_id,
                'Name': student_name
            }
        else:
            if all_students_data[id_key].get('Name', '') == '' and student_name:
                all_students_data[id_key]['Name'] = student_name
            if all_students_data[id_key].get('Student ID', '') == '' and student_id:
                all_students_data[id_key]['Student ID'] = student_id
        if name_key:
            student_name_to_key[name_key] = id_key
        return id_key

    if name_key:
        if name_key in student_name_to_key:
            return student_name_to_key[name_key]
        new_key = f'NAME::{name_key}'
        all_students_data[new_key] = {
            'Student ID': '',
            'Name': student_name
        }
        student_name_to_key[name_key] = new_key
        return new_key

    new_key = f'NOID::{len(all_students_data) + 1}'
    all_students_data[new_key] = {
        'Student ID': '',
        'Name': ''
    }
    return new_key


def update_score(key, col_name, value):
    blank = is_blank_value(value)
    current = all_students_data[key].get(col_name, '')
    if current != '' and current != 0 and blank:
        return
    if blank:
        all_students_data[key][col_name] = ''
    else:
        all_students_data[key][col_name] = value


for week_folder in all_folders:
    week_path = os.path.join(root_folder, week_folder)
    if not os.path.isdir(week_path):
        continue

    print(f"\n==================================================")
    print(f"📂 正在全力处理文件夹: {week_folder}")
    print(f"==================================================")

    excel_files = sorted(glob.glob(os.path.join(week_path, "*.xlsx")))
    print(f"📄 该 Week 目录下系统找到的 Excel 文件总数: {len(excel_files)} 个")

    for file_path in excel_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        if file_name.startswith('~$'):
            print(f"   ⏭️  自动跳过 Excel 临时锁文件: {file_name}")
            continue

        print(f"\n   🔍 正在剖析文件: [{file_name}.xlsx]")

        try:
            xl = pd.ExcelFile(file_path)
            target_sheets = find_target_sheets(xl)
            if not target_sheets:
                print(f"      ⚠️  该文件找不到任何匹配的 Sheet 页！它包含的 Sheet 有: {xl.sheet_names}")
                continue
            
            # 处理所有找到的sheet
            for target_sheet in target_sheets:
                if 'analytics per learning step' in target_sheet.lower() and 'analytics per student' not in target_sheet.lower():
                    print(f"      ℹ️  未找到 Student 页，已自动切换至 '{target_sheet}' 页")
                df = pd.read_excel(file_path, sheet_name=target_sheet)
                
                df.columns = [str(col).strip() for col in df.columns]
                print(f"      📋 本表【全部原始 Header】为:\n      👉 {list(df.columns)}")

                id_col = None
                name_col = None
                for original_col in df.columns:
                    col_lower = original_col.lower()
                    if ('student id' in col_lower or 'sourcedid' in col_lower or 'user id' in col_lower or col_lower == 'id') and id_col is None:
                        id_col = original_col
                    if 'name' in col_lower and 'username' not in col_lower and 'email' not in col_lower and name_col is None:
                        name_col = original_col

                if not name_col:
                    print(f"      ❌ [严重错误] 无法在这个 Header 列表里定位到学生姓名列！")
                    continue
                if not id_col:
                    print(f"      ⚠️  未找到 Student ID 列，后续会只用姓名进行匹配")

                exclude_keywords = ['student']
                data_cols = [col for col in df.columns if col != id_col and col != name_col and not any(kw in col.lower() for kw in exclude_keywords)]
                print(f"      📊 筛选出的【待汇总数据列】为: {data_cols}")

                week_digits = ''.join(filter(str.isdigit, week_folder))
                week_short = f'W{week_digits}' if week_digits else week_folder
                q_prefix = f"{week_short} - {file_name}"

                file_grouped_rows = {}
                for _, row in df.iterrows():
                    student_name = row[name_col]
                    if pd.isna(student_name) or str(student_name).strip() == '':
                        continue
                    student_name = str(student_name).strip()
                    raw_student_id = row[id_col] if id_col else None
                    student_id = normalize_student_id(raw_student_id) if id_col else None
                    name_key = normalize_text(student_name)

                    existing_key = student_name_to_key.get(name_key)
                    if student_id:
                        student_key = choose_student_key(student_id, student_name)
                        if existing_key and existing_key != student_key and existing_key in file_grouped_rows:
                            old_group = file_grouped_rows.pop(existing_key)
                            if student_key not in file_grouped_rows:
                                file_grouped_rows[student_key] = {
                                    'Student ID': student_id or old_group['Student ID'],
                                    'Name': student_name or old_group['Name'],
                                    'row': {}
                                }
                            target_group = file_grouped_rows[student_key]
                            for col, old_value in old_group['row'].items():
                                existing_value = target_group['row'].get(col, '')
                                if existing_value != '' and not is_blank_value(existing_value) and is_blank_value(old_value):
                                    continue
                                target_group['row'][col] = old_value
                    elif existing_key:
                        student_key = existing_key
                    else:
                        student_key = choose_student_key(None, student_name)

                    if student_key not in file_grouped_rows:
                        file_grouped_rows[student_key] = {
                            'Student ID': student_id or '',
                            'Name': student_name,
                            'row': {}
                        }

                    grouped = file_grouped_rows[student_key]
                    if grouped['Student ID'] == '' and student_id:
                        grouped['Student ID'] = student_id
                    if grouped['Name'] == '' and student_name:
                        grouped['Name'] = student_name

                    for col in data_cols:
                        value = row[col]
                        existing_value = grouped['row'].get(col, '')
                        if existing_value != '' and not is_blank_value(existing_value) and is_blank_value(value):
                            continue
                        grouped['row'][col] = value

                for student_key, grouped in file_grouped_rows.items():
                    if student_key not in all_students_data:
                        all_students_data[student_key] = {
                            'Student ID': grouped['Student ID'],
                            'Name': grouped['Name']
                        }
                    else:
                        if all_students_data[student_key].get('Student ID', '') == '' and grouped['Student ID']:
                            all_students_data[student_key]['Student ID'] = grouped['Student ID']
                        if all_students_data[student_key].get('Name', '') == '' and grouped['Name']:
                            all_students_data[student_key]['Name'] = grouped['Name']

                    for col, value in grouped['row'].items():
                        new_col_name = f"{q_prefix} - {col}"
                        if new_col_name not in column_order:
                            column_order[new_col_name] = len(column_order)
                        update_score(student_key, new_col_name, value)
        except Exception as e:
            print(f"      ❌ 读取失败。错误原因: {e}")
            continue

if all_students_data:
    final_df = pd.DataFrame.from_dict(all_students_data, orient='index')
    if 'Student ID' not in final_df.columns:
        final_df['Student ID'] = ''
    if 'Name' not in final_df.columns:
        final_df['Name'] = ''

    data_cols_only = [c for c in final_df.columns if c not in ['Student ID', 'Name']]

    def get_week_num(col_name):
        try:
            part = col_name.split('-')[0].strip()
            num = ''.join(filter(str.isdigit, part))
            return int(num) if num else 999
        except Exception:
            return 999

    sorted_data_cols = sorted(data_cols_only, key=lambda c: (get_week_num(c), column_order.get(c, 9999)))
    ordered_cols = ['Student ID', 'Name'] + sorted_data_cols
    final_df = final_df.reindex(columns=ordered_cols)
    
    for col in final_df.columns:
        if col not in ['Student ID', 'Name']:
            final_df[col] = final_df[col].apply(lambda x: '' if should_be_blank_in_output(x, col) else x)
    
    final_df.to_excel(output_file, index=False)
    print(f"\n==================================================")
    print(f"🎉 跑完了！全量监控后的总表已生成至:\n{output_file}")
    print(f"==================================================")
else:
    print("\n❌ 结束：依然未能提取到任何有效的学生数据。")
