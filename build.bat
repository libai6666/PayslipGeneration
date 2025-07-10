@echo off
chcp 65001 >nul
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

rem 确保资源目录存在
if not exist assets mkdir assets

rem 检查是否有图标文件
set ICON_OPTION=
if exist assets\icon.ico (
    set ICON_OPTION=--icon=assets\icon.ico
) else (
    echo 注意：未找到图标文件 assets\icon.ico，将使用默认图标。
)

rem 准备数据文件选项
set DATA_OPTIONS=
if exist resources (
    set DATA_OPTIONS=--add-data "resources;resources"
)
if exist assets (
    set DATA_OPTIONS=%DATA_OPTIONS% --add-data "assets;assets"
)

echo 请选择打包模式:
echo 1. 单文件模式 (打包为单个exe文件，启动较慢)
echo 2. 目录模式 (打包为文件夹，启动较快)
choice /c 12 /m "请选择打包模式"

if errorlevel 2 goto DIR_MODE
if errorlevel 1 goto FILE_MODE

:FILE_MODE
echo 正在使用单文件模式打包...
pyinstaller --onefile --windowed --name "工资条生成器" %ICON_OPTION% %DATA_OPTIONS% --version-file=version.txt main.py
goto END

:DIR_MODE
echo 正在使用目录模式打包...
pyinstaller --windowed --name "工资条生成器" %ICON_OPTION% %DATA_OPTIONS% --version-file=version.txt main.py
echo 正在创建ZIP包...
powershell Compress-Archive -Path "dist\工资条生成器\*" -DestinationPath "dist\工资条生成器_目录模式.zip" -Force

:END
echo 打包完成！
echo 可执行文件位于dist目录下。
pause 