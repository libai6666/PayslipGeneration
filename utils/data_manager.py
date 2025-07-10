"""
数据管理器模块
用于在单人模式和批量模式之间共享数据
"""

from datetime import datetime

class DataManager:
    """数据管理器单例类"""
    
    _instance = None
    
    @classmethod
    def get_instance(cls):
        """获取单例实例"""
        if cls._instance is None:
            cls._instance = DataManager()
        return cls._instance
    
    def __init__(self):
        """初始化数据管理器"""
        # 确保只有一个实例
        if DataManager._instance is not None:
            raise RuntimeError("尝试创建DataManager的第二个实例")
        
        # 当前月份（默认为当前月份）
        self.current_month = datetime.now().month
        
        # 单人模式数据
        self.single_mode_data = {
            'name': '',
            'base_salary': 0.0,
            'required_days': 0,
            'actual_days': 0,
            'night_shift': 0.0,
            'high_temp': 0.0,
            'late_fine': 0.0,
            'others': 0.0,
            'month': self.current_month
        }
        
        # 批量模式数据
        self.batch_mode_data = []
    
    def save_single_mode_data(self, data):
        """
        保存单人模式数据
        
        参数:
            data (dict): 单人模式数据字典
        """
        self.single_mode_data = data.copy()
        # 确保有月份字段
        if 'month' not in self.single_mode_data:
            self.single_mode_data['month'] = self.current_month
    
    def get_single_mode_data(self):
        """
        获取单人模式数据
        
        返回:
            dict: 单人模式数据字典的副本
        """
        return self.single_mode_data.copy()
    
    def save_batch_mode_data(self, data_list):
        """
        保存批量模式数据
        
        参数:
            data_list (list): 员工数据字典列表
        """
        # 确保每个数据项都有月份字段
        for item in data_list:
            if 'month' not in item:
                item['month'] = self.current_month
        
        self.batch_mode_data = [item.copy() for item in data_list]
    
    def get_batch_mode_data(self):
        """
        获取批量模式数据
        
        返回:
            list: 员工数据字典列表的副本
        """
        return [item.copy() for item in self.batch_mode_data]
    
    def set_current_month(self, month):
        """
        设置当前月份
        
        参数:
            month (int): 月份（1-12）
        """
        if 1 <= month <= 12:
            self.current_month = month
        else:
            raise ValueError("月份必须在1到12之间")
    
    def convert_single_to_batch(self):
        """
        将单人模式数据转换为批量模式数据
        
        返回:
            list: 包含单人模式数据的列表
        """
        # 如果单人模式数据中有名字，则添加到批量模式数据中
        if self.single_mode_data.get('name'):
            # 确保有月份字段
            if 'month' not in self.single_mode_data:
                self.single_mode_data['month'] = self.current_month
                
            # 如果批量模式数据为空，直接添加单人数据
            if not self.batch_mode_data:
                self.batch_mode_data = [self.single_mode_data.copy()]
            # 否则，检查是否已有同名数据，如有则更新，没有则添加
            else:
                name = self.single_mode_data['name']
                for i, item in enumerate(self.batch_mode_data):
                    if item.get('name') == name:
                        self.batch_mode_data[i] = self.single_mode_data.copy()
                        return self.get_batch_mode_data()
                # 没找到同名数据，添加新数据
                self.batch_mode_data.append(self.single_mode_data.copy())
        
        return self.get_batch_mode_data()
    
    def clear_all_data(self):
        """清除所有数据"""
        self.single_mode_data = {
            'name': '',
            'base_salary': 0.0,
            'required_days': 0,
            'actual_days': 0,
            'night_shift': 0.0,
            'high_temp': 0.0,
            'late_fine': 0.0,
            'others': 0.0,
            'month': self.current_month
        }
        self.batch_mode_data = [] 