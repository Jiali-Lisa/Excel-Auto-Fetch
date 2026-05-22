import os
import glob
import pandas as pd
import numpy as np

# 硬编码绝对路径
root_folder = '/Users/yiiing/Desktop/Work/Students self-paced learning'
output_file = '/Users/yiiing/Desktop/Work/Final_All_Weeks_Analytics.xlsx'

all_students_data = {}  

print("==================================================")
print("🚀 超级明朗版脚本启动（全量 Header 监控模式）...")
print("📍 目标路径:", root_folder)
print("==================================================")

if not os.path.exists(root_folder):
    print("❌ 路径不存在，请检查！")
    exit()

all_folders = sorted(os.listdir(root_folder))
print(f"扫描到根目录下的所有项: {all_folders}\n")

for week_folder in all_folders:
    week_path = os.path.join(root_folder, week_folder)
    
    if not os.path.isdir(week_path) or not week_folder.startswith('Week'):
        continue
        
    print(f"\n==================================================")
    print(f"📂 正在全力处理文件夹: {week_folder}")
    print(f"==================================================")
    
    excel_files = glob.glob(os.path.join(week_path, "*.xlsx"))
    print(f"📄 该 Week 目录下系统找到的 Excel 文件总数: {len(excel_files)} 个")
    
    for file_path in excel_files:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        if file_name.startswith("~$"):
            print(f"   ⏭️  自动跳过 Excel 临时锁文件: {file_name}")
            continue
            
        print(f"\n   🔍 正在剖析文件: [{file_name}.xlsx]")
        
        try:
            xl = pd.ExcelFile(file_path)
            
            # 双模态 Sheet 名字匹配
            target_sheet = None
            for sheet in xl.sheet_names:
                if "analytics per student" in sheet.lower():
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                for sheet in xl.sheet_names:
                    if "analytics per learning step" in sheet.lower():
                        target_sheet = sheet
                        print(f"      ℹ️  未找到 Student 页，已自动切换至 '{sheet}' 页")
                        break
            
            if not target_sheet:
                print(f"      ⚠️  [警告] 该文件找不到任何匹配的 Sheet 页！它包含的 Sheet 有: {xl.sheet_names}")
                continue
                
            df = pd.read_excel(file_path, sheet_name=target_sheet)
        except Exception as e:
            print(f"      ❌ 读取失败。错误原因: {e}")
            continue
        
        # 清洗并拉取原始列名
        df.columns = [str(col).strip() for col in df.columns]
        
        # =================【用户核心需求：在这里强行把所有 Header 吐出来】=================
        print(f"      📋 本表【全部原始 Header】为:\n      👉 {list(df.columns)}")
        # ==============================================================================
        
        # 精准子字符串匹配学生定位
        id_col = None
        name_col = None
        for original_col in df.columns:
            col_lower = original_col.lower()
            if 'student id' in col_lower or 'sourcedid' in col_lower or 'user id' in col_lower:
                id_col = original_col
            elif 'name' in col_lower and 'username' not in col_lower and 'email' not in col_lower:
                name_col = original_col

        if not id_col or not name_col:
            print(f"      ❌ [严重错误] 无法在这个 Header 列表里定位到学生 ID 或姓名列！")
            continue
            
        # 排除干扰列，留下真正的指标列
        exclude_keywords = [
            'student'
        ]
        data_cols = [col for col in df.columns if not any(kw in col.lower() for kw in exclude_keywords)]
        print(f"      📊 筛选出的【核心数据指标列】为: {data_cols}")
        
        # 提取周数
        week_digits = ''.join(filter(str.isdigit, week_folder))
        week_short = f"W{week_digits}" if week_digits else week_folder
        
        # 使用【周数 + 文件名】作为唯一前缀，防止多周同编号相互覆盖
        q_prefix = f"{week_short} - {file_name}"
        
        # 开始抓取数据
        for _, row in df.iterrows():
            student_id = row[id_col]
            student_name = row[name_col]
            
            if pd.isna(student_id):
                continue
                
            student_id = str(student_id).strip()
            student_scores = row[data_cols]
            
            # 检查是否有非 0 参与
            has_participation = False
            for val in student_scores:
                try:
                    if pd.notna(val) and str(val).strip() != '' and float(val) != 0:
                        has_participation = True
                        break
                except ValueError:
                    if pd.notna(val) and str(val).strip() != '':
                        has_participation = True
                        break
                    
            if has_participation:
                if student_id not in all_students_data:
                    all_students_data[student_id] = {
                        'Student ID': student_id,
                        'Name': student_name
                    }
                
                for col in data_cols:
                    new_col_name = f"{q_prefix} - {col}"
                    val = row[col]

                    try:
                        # 把格子的内容转成小写字符串，用来防范 'Not started' 或 'not started'
                        val_str = str(val).strip().lower()
                        
                        if pd.isna(val) or val_str == '' or val_str == 'not started' or float(val) == 0:
                            all_students_data[student_id][new_col_name] = ""
                        else:
                            all_students_data[student_id][new_col_name] = val
                    except ValueError:
                        # 如果转 float 报错了（说明是文本），再次双重保险检查是不是 'not started'
                        if str(val).strip().lower() == 'not started':
                            all_students_data[student_id][new_col_name] = ""
                        else:
                            all_students_data[student_id][new_col_name] = val

# 导出最终结果
if all_students_data:
    final_df = pd.DataFrame.from_dict(all_students_data, orient='index')
    cols = ['Student ID', 'Name'] + [c for c in final_df.columns if c not in ['Student ID', 'Name']]
    final_df = final_df[cols]
    
    final_df.to_excel(output_file, index=False)
    print(f"\n==================================================")
    print(f"🎉 跑完了！全量监控后的总表已生成至:\n{output_file}")
    print(f"==================================================")
else:
    print("\n❌ 结束：依然未能提取到任何有效的学生数据。")