"""
批量工资条生成器GUI界面
"""

import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QFormLayout, QLabel, QLineEdit, 
                           QPushButton, QMessageBox, QDesktopWidget,
                           QTableWidget, QTableWidgetItem, QHeaderView,
                           QFileDialog, QSpinBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush

# 导入自定义模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.calculator import calculate_absence_deduction, calculate_net_salary, validate_input
from utils.data_manager import DataManager


class BatchPayslipWindow(QMainWindow):
    """批量工资条处理窗口"""
    
    def __init__(self):
        """初始化窗口"""
        super().__init__()
        self.setWindowTitle("工资条批量生成器")
        self.setMinimumSize(800, 600)
        
        # 获取数据管理器实例
        self.data_manager = DataManager.get_instance()
        
        # 员工数据列表
        self.employee_data = []
        
        # 设置UI
        self.setup_ui()
        
        # 连接信号
        self.connect_signals()
        
        # 加载保存的数据
        self.load_data()
        
        # 居中显示窗口
        self.center_window()
    
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
        title_label = QLabel("工资条批量生成器")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        main_layout.addWidget(title_label)
        
        # 工具栏
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(10)
        
        self.import_button = QPushButton("导入数据")
        self.export_template_button = QPushButton("导出模板")
        self.add_row_button = QPushButton("添加员工")
        self.delete_row_button = QPushButton("删除所选")
        
        toolbar_layout.addWidget(self.import_button)
        toolbar_layout.addWidget(self.export_template_button)
        toolbar_layout.addWidget(self.add_row_button)
        toolbar_layout.addWidget(self.delete_row_button)
        toolbar_layout.addStretch()
        
        main_layout.addLayout(toolbar_layout)
        
        # 在工具栏下方添加月份选择
        month_layout = QHBoxLayout()
        month_layout.setSpacing(10)
        
        month_label = QLabel("月份:")
        self.month_spinbox = QSpinBox()
        self.month_spinbox.setRange(1, 12)
        self.month_spinbox.setValue(self.data_manager.current_month)  # 默认当前月份
        self.month_spinbox.setFixedWidth(60)
        
        month_layout.addWidget(month_label)
        month_layout.addWidget(self.month_spinbox)
        month_layout.addStretch()
        
        main_layout.addLayout(month_layout)
        
        # 表格视图
        self.table_widget = QTableWidget()
        self.setup_table()
        main_layout.addWidget(self.table_widget)
        
        # 底部按钮
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.generate_button = QPushButton("生成汇总工资表")
        self.generate_button.setMinimumHeight(40)
        self.generate_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        
        self.generate_individual_button = QPushButton("生成个人工资条")
        self.generate_individual_button.setMinimumHeight(40)
        
        self.clear_button = QPushButton("清除所有数据")
        self.clear_button.setMinimumHeight(40)
        
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.generate_individual_button)
        button_layout.addWidget(self.generate_button)
        
        main_layout.addLayout(button_layout)
    
    def setup_table(self):
        """设置表格视图"""
        # 表头 - 增加月份列
        headers = ["姓名", "月份", "基本工资", "应出勤天数", "实际出勤天数", 
                  "夜班补助", "高温补贴", "迟到罚款", "其他", "缺勤扣款", "实发工资"]
        
        self.table_widget.setColumnCount(len(headers))
        self.table_widget.setHorizontalHeaderLabels(headers)
        
        # 设置表格属性
        self.table_widget.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_widget.setAlternatingRowColors(True)
        # 修改编辑触发方式为单击
        self.table_widget.setEditTriggers(QTableWidget.CurrentChanged | QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        
        # 设置列宽
        header = self.table_widget.horizontalHeader()
        for i in range(len(headers)):
            header.setSectionResizeMode(i, QHeaderView.Stretch)
    
    def connect_signals(self):
        """连接信号和槽"""
        self.import_button.clicked.connect(self.import_data)
        self.export_template_button.clicked.connect(self.export_template)
        self.add_row_button.clicked.connect(self.add_row)
        self.delete_row_button.clicked.connect(self.delete_rows)
        self.generate_button.clicked.connect(self.generate_summary)
        self.generate_individual_button.clicked.connect(self.generate_individual_payslips)
        self.clear_button.clicked.connect(self.clear_data)
        self.table_widget.cellChanged.connect(self.cell_changed)
        self.month_spinbox.valueChanged.connect(self.update_month)
    
    def update_month(self, month):
        """更新当前月份，仅影响新添加的行"""
        self.data_manager.set_current_month(month)
        print(f"已将默认月份设置为：{month}月")
    
    def import_data(self):
        """导入数据"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择数据文件", "", "Excel文件 (*.xlsx *.xls);;CSV文件 (*.csv);;所有文件 (*)"
        )
        
        if file_path:
            try:
                from utils.data_import import import_employee_data
                employees = import_employee_data(file_path)
                self.load_employees(employees)
                QMessageBox.information(self, "成功", f"成功导入{len(employees)}条员工数据")
                
                # 保存导入的数据
                self.save_data()
            except Exception as e:
                QMessageBox.critical(self, "导入错误", f"导入数据时出错：{str(e)}")
    
    def export_template(self):
        """导出模板"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存模板文件", "工资条导入模板.xlsx", "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                from utils.data_import import export_template
                export_template(file_path)
                QMessageBox.information(self, "成功", f"模板已成功保存到：{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "导出错误", f"导出模板时出错：{str(e)}")
    
    def add_row(self):
        """添加新行"""
        row_count = self.table_widget.rowCount()
        self.table_widget.insertRow(row_count)
        
        # 设置月份列的值为当前选择的月份
        month_item = QTableWidgetItem(str(self.month_spinbox.value()))
        self.table_widget.setItem(row_count, 1, month_item)
        
        # 设置缺勤扣款和实发工资单元格为只读
        for col in [9, 10]:  # 缺勤扣款和实发工资列
            item = QTableWidgetItem("0.00")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row_count, col, item)
        
        # 为数值列添加默认值0
        for col in range(2, 9):
            self.table_widget.setItem(row_count, col, QTableWidgetItem("0"))
    
    def delete_rows(self):
        """删除选中行"""
        selected_rows = sorted(set(index.row() for index in self.table_widget.selectedIndexes()))
        if not selected_rows:
            return
        
        for i, row in enumerate(selected_rows):
            self.table_widget.removeRow(row - i)  # 考虑删除后索引变化
        
        # 保存数据
        self.save_data()
    
    def cell_changed(self, row, column):
        """单元格内容变化时更新计算结果"""
        # 忽略缺勤扣款和实发工资列的变化，它们是自动计算的
        if column < 9 and column != 1:  # 排除月份列
            self.calculate_row(row)
            # 保存数据
            self.save_data()
        elif column == 1:
            # 月份发生变化，只保存数据但不重新计算
            self.save_data()
    
    def calculate_row(self, row):
        """计算指定行的缺勤扣款和实发工资"""
        try:
            # 获取输入值
            base_salary = self.get_cell_value(row, 2, 0.0)
            required_days = self.get_cell_value(row, 3, 0)
            actual_days = self.get_cell_value(row, 4, 0)
            night_shift = self.get_cell_value(row, 5, 0.0)
            high_temp = self.get_cell_value(row, 6, 0.0)
            late_fine = self.get_cell_value(row, 7, 0.0)
            others = self.get_cell_value(row, 8, 0.0)
            
            # 计算缺勤扣款和实发工资
            absence_deduction = calculate_absence_deduction(base_salary, required_days, actual_days)
            net_salary = calculate_net_salary(base_salary, absence_deduction, night_shift, high_temp, late_fine, others)
            
            # 更新表格
            self.update_cell_value(row, 9, f"{absence_deduction:.2f}")
            self.update_cell_value(row, 10, f"{net_salary:.2f}")
            
        except Exception as e:
            print(f"计算错误：{str(e)}")
    
    def get_cell_value(self, row, column, default=None):
        """获取单元格值"""
        item = self.table_widget.item(row, column)
        if item is None or not item.text().strip():
            return default
        
        try:
            if isinstance(default, int):
                return int(float(item.text()))
            return float(item.text())
        except (ValueError, TypeError):
            return default
    
    def update_cell_value(self, row, column, value):
        """更新单元格值"""
        item = QTableWidgetItem(str(value))
        if column in [9, 10]:  # 缺勤扣款和实发工资列设为只读
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        if column == 9:  # 缺勤扣款列，负值显示为红色
            try:
                if float(value) < 0:
                    item.setForeground(QBrush(QColor("red")))
            except:
                pass
        self.table_widget.setItem(row, column, item)
    
    def load_employees(self, employees):
        """加载员工数据到表格"""
        self.table_widget.setRowCount(0)  # 清除现有数据
        
        for employee in employees:
            row = self.table_widget.rowCount()
            self.table_widget.insertRow(row)
            
            # 填充数据
            self.table_widget.setItem(row, 0, QTableWidgetItem(employee.get('name', '')))
            self.table_widget.setItem(row, 1, QTableWidgetItem(str(employee.get('month', self.month_spinbox.value()))))
            self.table_widget.setItem(row, 2, QTableWidgetItem(str(employee.get('base_salary', '0'))))
            self.table_widget.setItem(row, 3, QTableWidgetItem(str(employee.get('required_days', '0'))))
            self.table_widget.setItem(row, 4, QTableWidgetItem(str(employee.get('actual_days', '0'))))
            self.table_widget.setItem(row, 5, QTableWidgetItem(str(employee.get('night_shift', '0'))))
            self.table_widget.setItem(row, 6, QTableWidgetItem(str(employee.get('high_temp', '0'))))
            self.table_widget.setItem(row, 7, QTableWidgetItem(str(employee.get('late_fine', '0'))))
            self.table_widget.setItem(row, 8, QTableWidgetItem(str(employee.get('others', '0'))))
            
            # 计算结果
            self.calculate_row(row)
            
            # 如果有月份信息，更新月份选择器（仅用于未来新行的默认值）
            if 'month' in employee and 1 <= employee['month'] <= 12:
                self.month_spinbox.setValue(employee['month'])
    
    def collect_employee_data(self):
        """收集表格中的所有员工数据"""
        employees = []
        
        for row in range(self.table_widget.rowCount()):
            # 获取姓名
            name_item = self.table_widget.item(row, 0)
            if name_item is None or not name_item.text().strip():
                continue
                
            # 获取月份
            try:
                month = int(self.get_cell_text(row, 1) or self.month_spinbox.value())
            except ValueError:
                month = self.month_spinbox.value()
                
            # 获取其他数据
            try:
                employee = {
                    'name': name_item.text().strip(),
                    'month': month,
                    'base_salary': self.get_cell_value(row, 2, 0.0),
                    'required_days': self.get_cell_value(row, 3, 0),
                    'actual_days': self.get_cell_value(row, 4, 0),
                    'night_shift': self.get_cell_value(row, 5, 0.0),
                    'high_temp': self.get_cell_value(row, 6, 0.0),
                    'late_fine': self.get_cell_value(row, 7, 0.0),
                    'others': self.get_cell_value(row, 8, 0.0),
                    'absence_deduction': self.get_cell_value(row, 9, 0.0),
                    'net_salary': self.get_cell_value(row, 10, 0.0)
                }
                
                # 验证关键数据
                if employee['base_salary'] <= 0 or employee['required_days'] <= 0:
                    print(f"行 {row+1} 的数据无效: 基本工资或应出勤天数必须大于零")
                    continue
                    
                employees.append(employee)
            except Exception as e:
                print(f"收集第 {row+1} 行数据时出错: {str(e)}")
        
        return employees
    
    def generate_summary(self):
        """生成汇总工资表"""
        # 先保存当前表格数据
        self.save_data()
        
        # 收集有效的员工数据
        employees = self.collect_employee_data()
        
        # 打印调试信息
        print(f"收集到 {len(employees)} 个有效员工数据")
        
        if not employees:
            QMessageBox.warning(self, "警告", "没有有效的员工数据！请确保至少有一行完整的员工信息，包括姓名、基本工资和出勤天数。")
            return
        
        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if not output_dir:
            return
        
        try:
            # 获取当前月份
            month = self.month_spinbox.value()
            
            from utils.excel import generate_summary_excel
            output_path = generate_summary_excel(
                employees, 
                month, 
                os.path.join(output_dir, f"{month}月工资表.xlsx")
            )
            
            QMessageBox.information(
                self, 
                "成功", 
                f"已成功生成{month}月工资汇总表！\n\n保存在：{output_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成工资表时出错：{str(e)}")
    
    def generate_individual_payslips(self):
        """生成个人工资条"""
        # 先保存当前表格数据
        self.save_data()
        
        # 收集有效的员工数据
        employees = self.collect_employee_data()
        
        if not employees:
            QMessageBox.warning(self, "警告", "没有有效的员工数据！请确保至少有一行完整的员工信息，包括姓名、基本工资和出勤天数。")
            return
        
        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
        if not output_dir:
            return
        
        try:
            from utils.excel import batch_generate_excel
            file_paths = batch_generate_excel(employees, output_dir)
            
            QMessageBox.information(
                self, 
                "成功", 
                f"已成功生成{len(file_paths)}份个人工资条！\n\n保存在：{output_dir}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成工资条时出错：{str(e)}")
    
    def clear_data(self):
        """清除所有数据"""
        if self.table_widget.rowCount() > 0:
            reply = QMessageBox.question(
                self, 
                "确认清除", 
                "确定要清除所有数据吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.table_widget.setRowCount(0)
                # 清除数据管理器中的批量模式数据
                self.data_manager.batch_mode_data = []
    
    def get_cell_text(self, row, column):
        """获取单元格文本内容"""
        item = self.table_widget.item(row, column)
        if item is None:
            return ""
        return item.text().strip()
    
    def save_data(self):
        """保存表格数据到数据管理器"""
        try:
            employees = []
            
            for row in range(self.table_widget.rowCount()):
                name = self.get_cell_text(row, 0)
                if not name:  # 跳过没有姓名的行
                    continue
                
                # 获取月份
                try:
                    month = int(self.get_cell_text(row, 1) or self.month_spinbox.value())
                except ValueError:
                    month = self.month_spinbox.value()
                    
                employee = {
                    'name': name,
                    'month': month,
                    'base_salary': self.get_cell_value(row, 2, 0.0),
                    'required_days': self.get_cell_value(row, 3, 0),
                    'actual_days': self.get_cell_value(row, 4, 0),
                    'night_shift': self.get_cell_value(row, 5, 0.0),
                    'high_temp': self.get_cell_value(row, 6, 0.0),
                    'late_fine': self.get_cell_value(row, 7, 0.0),
                    'others': self.get_cell_value(row, 8, 0.0),
                    'absence_deduction': self.get_cell_value(row, 9, 0.0),
                    'net_salary': self.get_cell_value(row, 10, 0.0)
                }
                employees.append(employee)
            
            self.data_manager.save_batch_mode_data(employees)
            print(f"已保存 {len(employees)} 条员工数据")
        except Exception as e:
            print(f"保存数据时出错: {str(e)}")
    
    def load_data(self):
        """从数据管理器加载数据到表格"""
        employees = self.data_manager.get_batch_mode_data()
        if employees:
            self.load_employees(employees)
            print(f"已加载 {len(employees)} 条员工数据")
    
    def closeEvent(self, event):
        """窗口关闭时保存数据"""
        self.save_data()
        super().closeEvent(event)