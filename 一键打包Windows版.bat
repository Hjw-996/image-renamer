@echo off
echo 正在检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.x 并勾选 "Add Python to PATH"！
    pause
    exit /b
)

echo 正在创建虚拟环境...
python -m venv venv
call venv\Scripts\activate

echo 正在安装依赖库 (PyInstaller, openpyxl)...
pip install pyinstaller openpyxl

echo 正在打包程序...
pyinstaller --name="图片重命名工具" --onefile --windowed --noconfirm image_renamer.py

echo.
echo ========================================================
echo 打包完成！
echo 请在当前目录下的 dist 文件夹中查找 "图片重命名工具.exe"
echo ========================================================
pause