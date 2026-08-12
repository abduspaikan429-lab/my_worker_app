@echo off
chcp 65001 >nul
echo 正在为您生成桌面快捷方式...

cd /d "%~dp0"
python make_shortcut.py

echo.
echo ===================================================
echo   桌面快捷方式创建成功！
echo   快捷方式：[ 建筑劳务综合管理平台.lnk ]
echo   现在您可以直接在桌面双击图标启动系统。
echo ===================================================
echo.
pause
