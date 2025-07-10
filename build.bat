@echo off
echo 开始打包工资条生成器...

rem 确保工作目录正确
cd /d %~dp0

rem 检查是否存在虚拟环境
if exist .venv (
    echo 使用现有的虚拟环境...
    call .venv\Scripts\activate
) else (
    echo 创建新的虚拟环境...
    python -m venv .venv
    call .venv\Scripts\activate
    
    echo 安装依赖包...
    pip install -r requirements.txt
)

rem 使用PyInstaller打包
echo 打包应用程序...
pyinstaller --onefile --windowed --name "工资条生成器" --icon=assets/icon.ico main.py

echo 打包完成！
echo 可执行文件位于dist目录下。
pause 