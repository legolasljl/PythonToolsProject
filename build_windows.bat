@echo off
chcp 65001 >nul
echo ==========================================
echo   智能条款比对工具 - Windows 打包脚本
echo ==========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

:: 检查 PyInstaller
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装 PyInstaller...
    pip install pyinstaller
)

:: 检查必要文件
echo.
echo [检查] 检查必要文件...
if not exist "clause_diff_gui_ultimate_v16.py" (
    echo [错误] 未找到 clause_diff_gui_ultimate_v16.py
    pause
    exit /b 1
)
if not exist "clause_config_manager.py" (
    echo [错误] 未找到 clause_config_manager.py
    pause
    exit /b 1
)
if not exist "clause_config.json" (
    echo [错误] 未找到 clause_config.json
    pause
    exit /b 1
)

echo [OK] 所有必要文件已就绪

:: 检查图标
if not exist "icon.ico" (
    echo.
    echo [警告] 未找到 icon.ico，将使用默认图标
    echo [提示] 可使用在线工具将 icon.jpg 转换为 icon.ico
    echo.
)

:: 清理旧构建
echo.
echo [清理] 清理旧的构建文件...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

:: 开始打包
echo.
echo [打包] 开始打包，请稍候...
echo.

if exist "icon.ico" (
    pyinstaller --onefile --windowed --icon=icon.ico ^
        --add-data "clause_config_manager.py;." ^
        --add-data "clause_config.json;." ^
        --add-data "clause_config_client.json;." ^
        --name SmartClauseMatcher ^
        clause_diff_gui_ultimate_v16.py
) else (
    pyinstaller --onefile --windowed ^
        --add-data "clause_config_manager.py;." ^
        --add-data "clause_config.json;." ^
        --add-data "clause_config_client.json;." ^
        --name SmartClauseMatcher ^
        clause_diff_gui_ultimate_v16.py
)

:: 检查结果
echo.
if exist "dist\SmartClauseMatcher.exe" (
    echo ==========================================
    echo [成功] 打包完成！
    echo ==========================================
    echo.
    echo 程序位置: dist\SmartClauseMatcher.exe
    echo.
    echo 使用说明:
    echo   1. 双击 SmartClauseMatcher.exe 运行
    echo   2. 配置文件保存在 %%USERPROFILE%%\.clause_diff\
    echo   3. 可将 exe 复制到任意位置运行
    echo.

    :: 显示文件大小
    for %%A in ("dist\SmartClauseMatcher.exe") do echo 文件大小: %%~zA 字节

) else (
    echo.
    echo [失败] 打包失败，请检查错误信息
)

echo.
pause
