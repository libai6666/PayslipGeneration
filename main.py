"""
工资条生成器应用入口点
"""

import sys
import os
from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
import time

# 添加当前目录到系统路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 初始化数据管理器
from utils.data_manager import DataManager
data_manager = DataManager.get_instance()

# 导入资源路径辅助函数
from utils.resource_helper import resource_path

def show_splash_screen(app):
    """显示启动屏幕"""
    # 检查是否有启动图像
    splash_path = resource_path(os.path.join("assets", "splash.png"))
    if os.path.exists(splash_path):
        splash_pix = QPixmap(splash_path)
        splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
        splash.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        splash.show()
        app.processEvents()
        
        # 显示加载消息
        splash.showMessage("正在加载应用...", Qt.AlignBottom | Qt.AlignCenter, Qt.white)
        time.sleep(1)
        
        return splash
    
    return None

def main():
    """应用主入口点"""
    app = QApplication(sys.argv)
    app.setApplicationName("工资条生成器")
    
    # 显示启动画面
    splash = show_splash_screen(app)
    
    # 只使用批量模式
    from ui.batch_payslip_ui import BatchPayslipWindow
    window = BatchPayslipWindow()
    
    # 关闭启动画面并显示主窗口
    if splash:
        splash.finish(window)
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()