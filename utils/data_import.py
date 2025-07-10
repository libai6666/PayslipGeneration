"""
数据导入模块
用于导入员工数据和导出模板
"""

import os
import csv
from openpyxl import Workbook, load_workbook
from datetime import datetime


def import_employee_data(file_path):
    """
    导入员工数据 - 不依赖pandas，支持中文表头，自动检测表头行
    
    参数:
        file_path (str): 数据文件路径
    
    返回:
        list: 员工数据字典列表
    """
    # 获取文件扩展名
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    employees = []
    
    # 列名映射（中文 -> 英文）
    column_map = {
        '姓名': 'name',
        '月份': 'month',
        '基本工资': 'base_salary',
        '应出勤天数': 'required_days',
        '实际出勤天数': 'actual_days',
        '夜班补助': 'night_shift',
        '高温补贴': 'high_temp',
        '迟到罚款': 'late_fine',
        '其他': 'others'
    }
    
    # 反向映射字典，用于错误提示（英文 -> 中文）
    reverse_map = {v: k for k, v in column_map.items()}
    
    # 必要的中文列
    required_cn_columns = ['姓名', '基本工资', '应出勤天数', '实际出勤天数']
    
    try:
        if ext in ['.xlsx', '.xls']:
            # 使用openpyxl读取Excel文件
            wb = load_workbook(filename=file_path, read_only=True, data_only=True)
            ws = wb.active
            
            # 尝试检测表头行（第一行或第二行）
            header_row = 1  # 默认假设表头在第一行
            data_start_row = 2  # 默认数据从第二行开始
            
            # 从第一行获取可能的表头
            headers1 = []
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=1, column=col).value
                if cell_value is not None:
                    headers1.append(str(cell_value).strip())
                else:
                    headers1.append(None)
            
            # 检查第一行是否包含必要的列
            found_headers1 = [h for h in headers1 if h is not None]
            missing_count1 = 0
            for required_cn in required_cn_columns:
                if not any(required_cn in h or h == required_cn for h in found_headers1):
                    missing_count1 += 1
            
            # 如果第一行缺少必要的列，尝试第二行
            if missing_count1 > 0 and ws.max_row > 1:
                headers2 = []
                for col in range(1, ws.max_column + 1):
                    cell_value = ws.cell(row=2, column=col).value
                    if cell_value is not None:
                        headers2.append(str(cell_value).strip())
                    else:
                        headers2.append(None)
                
                # 检查第二行是否包含必要的列
                found_headers2 = [h for h in headers2 if h is not None]
                missing_count2 = 0
                for required_cn in required_cn_columns:
                    if not any(required_cn in h or h == required_cn for h in found_headers2):
                        missing_count2 += 1
                
                # 如果第二行比第一行更像是表头，使用第二行
                if missing_count2 < missing_count1:
                    print("检测到表头可能在第二行，使用第二行作为表头")
                    headers = headers2
                    header_row = 2
                    data_start_row = 3
                else:
                    headers = headers1
            else:
                headers = headers1
            
            print(f"使用第{header_row}行作为表头: {headers}")
            
            # 映射列名
            mapped_indices = {}
            for i, header in enumerate(headers):
                if header is None:
                    continue
                    
                header_lower = header.lower()
                
                # 1. 首先尝试精确匹配
                if header in column_map:
                    mapped_indices[column_map[header]] = i
                    continue
                
                # 2. 然后尝试不区分大小写的精确匹配
                for cn, en in column_map.items():
                    if cn.lower() == header_lower:
                        mapped_indices[en] = i
                        break
                        
                # 3. 最后尝试包含关系匹配
                if not any(cn.lower() == header_lower for cn in column_map):
                    for cn, en in column_map.items():
                        if cn.lower() in header_lower or header_lower in cn.lower():
                            mapped_indices[en] = i
                            break
            
            print(f"列映射结果: {mapped_indices}")
            
            # 检查必要的列是否存在
            required_columns = ['name', 'base_salary', 'required_days', 'actual_days']
            missing_columns = [col for col in required_columns if col not in mapped_indices]
            if missing_columns:
                # 准备中文名称的错误消息
                missing_cn = [reverse_map[col] for col in missing_columns]
                raise ValueError(f"输入数据缺少必要的列：{', '.join(missing_cn)}")
            
            # 读取数据行
            for row in range(data_start_row, ws.max_row + 1):
                employee = {}
                
                # 获取映射后的值
                for en, i in mapped_indices.items():
                    cell_value = ws.cell(row=row, column=i+1).value
                    if cell_value is not None:
                        employee[en] = cell_value
                
                # 如果有姓名，添加到员工列表
                if employee.get('name'):
                    # 确保数值字段类型正确
                    for field in ['base_salary', 'required_days', 'actual_days', 
                                 'night_shift', 'high_temp', 'late_fine', 'others']:
                        if field in employee:
                            try:
                                if field in ['required_days', 'actual_days']:
                                    employee[field] = int(float(str(employee[field]).replace(',', '')))
                                else:
                                    employee[field] = float(str(employee[field]).replace(',', ''))
                            except (ValueError, TypeError):
                                employee[field] = 0
                        else:
                            employee[field] = 0
                    
                    # 处理月份字段
                    if 'month' in employee:
                        try:
                            employee['month'] = int(float(str(employee['month']).replace(',', '')))
                        except (ValueError, TypeError):
                            employee['month'] = datetime.now().month
                    else:
                        employee['month'] = datetime.now().month
                    
                    employees.append(employee)
            
            wb.close()
            
        elif ext == '.csv':
            # 使用csv模块读取CSV文件
            with open(file_path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
                
                if not rows:
                    raise ValueError("CSV文件为空")
                
                # 检测表头行
                header_row = 0  # 默认表头在第一行
                data_start_row = 1  # 默认数据从第二行开始
                
                # 检查第一行是否包含必要的列
                headers1 = [h.strip() for h in rows[0] if h]
                missing_count1 = 0
                for required_cn in required_cn_columns:
                    if not any(required_cn in h or h == required_cn for h in headers1):
                        missing_count1 += 1
                
                # 如果第一行缺少必要的列，尝试第二行
                if missing_count1 > 0 and len(rows) > 1:
                    headers2 = [h.strip() for h in rows[1] if h]
                    missing_count2 = 0
                    for required_cn in required_cn_columns:
                        if not any(required_cn in h or h == required_cn for h in headers2):
                            missing_count2 += 1
                    
                    # 如果第二行比第一行更像是表头，使用第二行
                    if missing_count2 < missing_count1:
                        print("检测到表头可能在第二行，使用第二行作为表头")
                        headers = headers2
                        header_row = 1
                        data_start_row = 2
                    else:
                        headers = headers1
                else:
                    headers = headers1
                
                print(f"使用第{header_row+1}行作为表头: {headers}")
                
                # 映射列名
                mapped_indices = {}
                for i, header in enumerate(headers):
                    if not header:
                        continue
                        
                    header_lower = header.lower()
                    
                    # 1. 首先尝试精确匹配
                    if header in column_map:
                        mapped_indices[column_map[header]] = i
                        continue
                    
                    # 2. 然后尝试不区分大小写的精确匹配
                    for cn, en in column_map.items():
                        if cn.lower() == header_lower:
                            mapped_indices[en] = i
                            break
                            
                    # 3. 最后尝试包含关系匹配
                    if not any(cn.lower() == header_lower for cn in column_map):
                        for cn, en in column_map.items():
                            if cn.lower() in header_lower or header_lower in cn.lower():
                                mapped_indices[en] = i
                                break
                
                print(f"列映射结果: {mapped_indices}")
                
                # 检查必要的列是否存在
                required_columns = ['name', 'base_salary', 'required_days', 'actual_days']
                missing_columns = [col for col in required_columns if col not in mapped_indices]
                if missing_columns:
                    # 准备中文名称的错误消息
                    missing_cn = [reverse_map[col] for col in missing_columns]
                    raise ValueError(f"输入数据缺少必要的列：{', '.join(missing_cn)}")
                
                # 读取数据行
                for row in rows[data_start_row:]:
                    employee = {}
                    
                    # 获取映射后的值
                    for en, i in mapped_indices.items():
                        if i < len(row):
                            cell_value = row[i].strip()
                            if cell_value:
                                employee[en] = cell_value
                    
                    # 如果有姓名，添加到员工列表
                    if employee.get('name'):
                        # 确保数值字段类型正确
                        for field in ['base_salary', 'required_days', 'actual_days', 
                                     'night_shift', 'high_temp', 'late_fine', 'others']:
                            if field in employee:
                                try:
                                    if field in ['required_days', 'actual_days']:
                                        employee[field] = int(float(str(employee[field]).replace(',', '')))
                                    else:
                                        employee[field] = float(str(employee[field]).replace(',', ''))
                                except (ValueError, TypeError):
                                    employee[field] = 0
                            else:
                                employee[field] = 0
                        
                        # 处理月份字段
                        if 'month' in employee:
                            try:
                                employee['month'] = int(float(str(employee['month']).replace(',', '')))
                            except (ValueError, TypeError):
                                employee['month'] = datetime.now().month
                        else:
                            employee['month'] = datetime.now().month
                        
                        employees.append(employee)
        else:
            raise ValueError(f"不支持的文件类型：{ext}")
        
        if not employees:
            raise ValueError("没有找到有效的员工数据")
        
        return employees
    
    except Exception as e:
        print(f"导入数据时出错：{str(e)}")
        raise ValueError(f"导入数据时出错：{str(e)}")


def export_template(file_path):
    """
    导出数据导入模板 - 不依赖pandas
    
    参数:
        file_path (str): 模板文件保存路径
    """
    try:
        # 创建工作簿
        wb = Workbook()
        ws = wb.active
        ws.title = "员工数据"
        
        # 添加表头 - 增加月份列
        headers = ["姓名", "月份", "基本工资", "应出勤天数", "实际出勤天数", 
                  "夜班补助", "高温补贴", "迟到罚款", "其他"]
        
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col).value = header
        
        # 添加示例数据
        ws.cell(row=2, column=1).value = "示例：张三"
        ws.cell(row=2, column=2).value = datetime.now().month  # 当前月份作为示例
        ws.cell(row=2, column=3).value = 5000
        ws.cell(row=2, column=4).value = 30
        ws.cell(row=2, column=5).value = 28
        ws.cell(row=2, column=6).value = 200
        ws.cell(row=2, column=7).value = 100
        ws.cell(row=2, column=8).value = -50
        ws.cell(row=2, column=9).value = 0
        
        # 设置列宽
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[chr(64 + col)].width = 15
        
        # 保存文件
        wb.save(file_path)
        return True
    except Exception as e:
        print(f"导出模板时出错：{str(e)}")
        raise Exception(f"导出模板时出错：{str(e)}")