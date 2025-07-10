"""
资源文件路径处理模块
在PyInstaller打包后正确加载资源文件
"""

import os
import sys


def resource_path(relative_path):
    """
    获取资源的绝对路径
    
    参数:
        relative_path (str): 资源的相对路径
    
    返回:
        str: 资源的绝对路径
    """
    # PyInstaller创建临时文件夹并将资源提取到_MEIxxxx文件夹中
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    
    # 开发模式下，直接返回相对路径
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), relative_path) 