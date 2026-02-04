@echo off
chcp 65001 >nul
title 元器件库存管理工具

echo 元器件库存管理工具
echo ========================
echo.

if not exist "dist\元器件库存管理工具.exe" (
    echo 错误：未找到可执行文件
    echo 请先运行打包命令
    pause
    exit /b 1
)

echo 找到可执行文件
echo 正在启动...
echo.
echo 请访问：http://localhost:5000
echo 按 Ctrl+C 停止程序
echo.

start "" "dist\元器件库存管理工具.exe"
ping -n 3 127.0.0.1 >nul
start http://localhost:5000

echo.
echo 程序已启动，可以开始使用
echo 此窗口可以关闭
pause