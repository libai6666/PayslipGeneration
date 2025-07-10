"""
工资条生成器GUI界面
"""

import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QFormLayout, QLabel, QLineEdit, 
                           QPushButton, QMessageBox, QDesktopWidget, QSpinBox)
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QRegExpValidator, QIcon, QFont

# 导入自定义模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.calculator import calculate_absence_deduction, calculate_net_salary
from utils.excel import generate_excel
from utils.data_manager import DataManager


class PayslipGeneratorWindow(QMainWindow):
    """
    工资条生成器主窗口
    """
    
    def __init__(self):
        """初始化窗口和UI元素"""
        super().__init__()
        self.setWindowTitle("工资条生成器")
        self.setMinimumSize(600, 400)
        
        # 获取数据管理器实例
        self.data_manager = DataManager.get_instance()
        
        # 居中显示窗口
        self.center_window()
        
        # 设置中心部件和布局
        self.setup_ui()
        
        # 连接信号和槽
        self.connect_signals()
        
        # 加载保存的数据
        self.load_data()
    
    def center_window(self):
        """居中显示窗口"""
        frame_geometry = self.frameGeometry()
        center_point = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())
    
    def setup_ui(self):
        """设置UI布局和组件"""
        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # 标题标签
        title_label = QLabel("工资条生成器")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        main_layout.addWidget(title_label)
        
        # 月份选择
        month_layout = QHBoxLayout()
        month_layout.setSpacing(10)
        
        month_label = QLabel("月份:")
        self.month_spinbox = QSpinBox()
        self.month_spinbox.setRange(1, 12)
        self.month_spinbox.setValue(self.data_manager.current_month)
        self.month_spinbox.setFixedWidth(60)
        
        month_layout.addWidget(month_label)
        month_layout.addWidget(self.month_spinbox)
        month_layout.addStretch()
        
        main_layout.addLayout(month_layout)
        
        # 表单布局
        form_layout = QFormLayout()
        form_layout.setSpacing(10)
        
        # 创建表单字段
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("请输入姓名")
        
        self.base_salary_edit = QLineEdit()
        self.base_salary_edit.setPlaceholderText("例如: 5000")
        self.base_salary_edit.setValidator(QRegExpValidator(QRegExp(r"^\d{1,8}(\.\d{1,2})?$")))
        
        self.required_days_edit = QLineEdit()
        self.required_days_edit.setPlaceholderText("例如: 30")
        self.required_days_edit.setValidator(QRegExpValidator(QRegExp(r"^\d{1,3}$")))
        
        self.actual_days_edit = QLineEdit()
        self.actual_days_edit.setPlaceholderText("例如: 28")
        self.actual_days_edit.setValidator(QRegExpValidator(QRegExp(r"^\d{1,3}$")))
        
        self.night_shift_edit = QLineEdit()
        self.night_shift_edit.setPlaceholderText("例如: 200")
        self.night_shift_edit.setText("0")
        self.night_shift_edit.setValidator(QRegExpValidator(QRegExp(r"^-?\d{1,8}(\.\d{1,2})?$")))
        
        self.high_temp_edit = QLineEdit()
        self.high_temp_edit.setPlaceholderText("例如: 100")
        self.high_temp_edit.setText("0")
        self.high_temp_edit.setValidator(QRegExpValidator(QRegExp(r"^-?\d{1,8}(\.\d{1,2})?$")))
        
        self.late_fine_edit = QLineEdit()
        self.late_fine_edit.setPlaceholderText("例如: -50")
        self.late_fine_edit.setText("0")
        self.late_fine_edit.setValidator(QRegExpValidator(QRegExp(r"^-?\d{1,8}(\.\d{1,2})?$")))
        
        self.others_edit = QLineEdit()
        self.others_edit.setPlaceholderText("例如: 0")
        self.others_edit.setText("0")
        self.others_edit.setValidator(QRegExpValidator(QRegExp(r"^-?\d{1,8}(\.\d{1,2})?$")))
        
        # 添加表单字段到布局
        form_layout.addRow(QLabel("姓名 *"), self.name_edit)
        form_layout.addRow(QLabel("基本工资 *"), self.base_salary_edit)
        form_layout.addRow(QLabel("应出勤天数 *"), self.required_days_edit)
        form_layout.addRow(QLabel("实际出勤天数 *"), self.actual_days_edit)
        form_layout.addRow(QLabel("夜班补助"), self.night_shift_edit)
        form_layout.addRow(QLabel("高温补贴"), self.high_temp_edit)
        form_layout.addRow(QLabel("迟到罚款"), self.late_fine_edit)
        form_layout.addRow(QLabel("其他"), self.others_edit)
        
        main_layout.addLayout(form_layout)
        
        # 添加说明标签
        note_label = QLabel("* 为必填项")
        note_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(note_label)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 批量模式按钮
        self.batch_mode_button = QPushButton("切换到批量模式")
        self.batch_mode_button.setMinimumHeight(40)
        
        # 生成按钮
        self.generate_button = QPushButton("生成工资条")
        self.generate_button.setMinimumHeight(40)
        self.generate_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        
        # 清除按钮
        self.clear_button = QPushButton("清除所有")
        self.clear_button.setMinimumHeight(40)
        
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.batch_mode_button)
        button_layout.addWidget(self.generate_button)
        
        main_layout.addLayout(button_layout)
    
    def connect_signals(self):
        """连接信号和槽"""
        self.generate_button.clicked.connect(self.generate_payslip)
        self.clear_button.clicked.connect(self.clear_all)
        self.batch_mode_button.clicked.connect(self.switch_to_batch_mode)
        self.month_spinbox.valueChanged.connect(self.update_month)
    
    def update_month(self, month):
        """更新月份"""
        self.data_manager.set_current_month(month)
        print(f"已将月份设置为: {month}月")
    
    def generate_payslip(self):
        """生成工资条"""
        # 获取输入数据
        name = self.name_edit.text().strip()
        
        try:
            base_salary = float(self.base_salary_edit.text() or "0")
            required_days = int(self.required_days_edit.text() or "0")
            actual_days = int(self.actual_days_edit.text() or "0")
            night_shift = float(self.night_shift_edit.text() or "0")
            high_temp = float(self.high_temp_edit.text() or "0")
            late_fine = float(self.late_fine_edit.text() or "0")
            others = float(self.others_edit.text() or "0")
            month = self.month_spinbox.value()
        except (ValueError, TypeError) as e:
            QMessageBox.warning(self, "输入错误", "请输入有效的数字！")
            return
        
        # 验证输入
        if not name:
            QMessageBox.warning(self, "输入错误", "请输入姓名！")
            return
        
        if base_salary <= 0:
            QMessageBox.warning(self, "输入错误", "基本工资必须大于零！")
            return
        
        if required_days <= 0:
            QMessageBox.warning(self, "输入错误", "应出勤天数必须大于零！")
            return
        
        if actual_days < 0:
            QMessageBox.warning(self, "输入错误", "实际出勤天数不能为负数！")
            return
        
        # 计算缺勤扣款和实发工资
        absence_deduction = calculate_absence_deduction(base_salary, required_days, actual_days)
        net_salary = calculate_net_salary(base_salary, absence_deduction, night_shift, high_temp, late_fine, others)
        
        # 准备数据
        employee_data = {
            'name': name,
            'base_salary': base_salary,
            'required_days': required_days,
            'actual_days': actual_days,
            'night_shift': night_shift,
            'high_temp': high_temp,
            'late_fine': late_fine,
            'others': others,
            'absence_deduction': absence_deduction,
            'net_salary': net_salary,
            'month': month
        }
        
        # 保存数据到数据管理器
        self.data_manager.save_single_mode_data(employee_data)
        
        # 生成Excel文件
        try:
            file_path = generate_excel(employee_data)
            QMessageBox.information(
                self, 
                "成功", 
                f"工资条已成功生成！\n\n保存路径：{file_path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成工资条时出错：{str(e)}")
    
    def clear_all(self):
        """清除所有输入字段"""
        self.name_edit.clear()
        self.base_salary_edit.clear()
        self.required_days_edit.clear()
        self.actual_days_edit.clear()
        self.night_shift_edit.setText("0")
        self.high_temp_edit.setText("0")
        self.late_fine_edit.setText("0")
        self.others_edit.setText("0")
        
        # 清除数据管理器中的数据
        self.data_manager.clear_all_data()
    
    def switch_to_batch_mode(self):
        """切换到批量处理模式"""
        try:
            # 保存当前数据
            self.save_data()
            
            # 转换单人模式数据到批量模式
            self.data_manager.convert_single_to_batch()
            
            # 导入批量处理模块
            from ui.batch_payslip_ui import BatchPayslipWindow
            
            # 创建并显示批量模式窗口
            self.batch_window = BatchPayslipWindow()
            self.batch_window.show()
            self.close()
        except ImportError as e:
            QMessageBox.critical(self, "错误", f"无法加载批量处理模块：{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"切换模式时出错：{str(e)}")
    
    def save_data(self):
        """保存当前输入的数据"""
        # 只有当姓名字段有值时才保存数据
        if self.name_edit.text().strip():
            try:
                employee_data = {
                    'name': self.name_edit.text().strip(),
                    'base_salary': float(self.base_salary_edit.text() or "0"),
                    'required_days': int(self.required_days_edit.text() or "0"),
                    'actual_days': int(self.actual_days_edit.text() or "0"),
                    'night_shift': float(self.night_shift_edit.text() or "0"),
                    'high_temp': float(self.high_temp_edit.text() or "0"),
                    'late_fine': float(self.late_fine_edit.text() or "0"),
                    'others': float(self.others_edit.text() or "0"),
                    'month': self.month_spinbox.value()  # 保存月份
                }
                self.data_manager.save_single_mode_data(employee_data)
            except (ValueError, TypeError):
                # 如果数据无效，则不保存
                pass
    
    def load_data(self):
        """加载之前保存的数据"""
        data = self.data_manager.get_single_mode_data()
        
        if data.get('name'):
            self.name_edit.setText(data['name'])
            self.base_salary_edit.setText(str(data['base_salary']))
            self.required_days_edit.setText(str(data['required_days']))
            self.actual_days_edit.setText(str(data['actual_days']))
            self.night_shift_edit.setText(str(data['night_shift']))
            self.high_temp_edit.setText(str(data['high_temp']))
            self.late_fine_edit.setText(str(data['late_fine']))
            self.others_edit.setText(str(data['others']))
            
            # 设置月份
            if 'month' in data and 1 <= data['month'] <= 12:
                self.month_spinbox.setValue(data['month'])
    
    def closeEvent(self, event):
        """窗口关闭时保存数据"""
        self.save_data()
        super().closeEvent(event)