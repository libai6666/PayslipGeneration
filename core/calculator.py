"""
工资条计算逻辑模块
包含缺勤扣款和实发工资的计算函数
"""


def calculate_absence_deduction(base_salary, required_days, actual_days):
    """
    计算缺勤扣款
    
    参数:
        base_salary (float): 基本工资
        required_days (int): 应出勤天数
        actual_days (int): 实际出勤天数
    
    返回:
        float: 缺勤扣款金额 (负值表示扣款)
    """
    if required_days <= 0:
        return 0.0
    
    if actual_days >= required_days:
        return 0.0
    
    deduction = -(base_salary / required_days) * (required_days - actual_days)
    return round(deduction, 2)  # 保留两位小数


def calculate_net_salary(base_salary, absence_deduction, night_shift, high_temp, late_fine, others):
    """
    计算实发工资
    
    参数:
        base_salary (float): 基本工资
        absence_deduction (float): 缺勤扣款
        night_shift (float): 夜班补助
        high_temp (float): 高温补贴
        late_fine (float): 迟到罚款
        others (float): 其他
    
    返回:
        float: 实发工资金额
    """
    net_salary = base_salary + absence_deduction + night_shift + high_temp + late_fine + others
    return round(net_salary, 2)  # 保留两位小数


def validate_input(text, default=0.0):
    """
    验证并转换用户输入
    
    参数:
        text (str): 输入文本
        default (float/int, optional): 默认值，如果输入无效则返回此值
    
    返回:
        float/int: 转换后的数值，如果无法转换则返回默认值
    """
    if not text or not text.strip():
        return default
    
    try:
        # 尝试将输入转换为浮点数或整数
        value = float(text)
        if isinstance(default, int) and value == int(value):
            return int(value)
        return value
    except (ValueError, TypeError):
        return default