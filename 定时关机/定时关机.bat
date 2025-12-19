@echo off
echo 本脚本将会在今天 21:52 自动关机...
echo.

:: 获取当前时间（小时和分钟）
for /f "tokens=1-2 delims=: " %%a in ("%time%") do (
    set hh=%%a
    set mm=%%b
)

:: 目标关机时间（21:52）
set target_hh=21
set target_mm=55

:: 计算当前时间总分钟数
set /a now_total=%hh%*60 + %mm%

:: 计算目标时间总分钟数
set /a target_total=%target_hh%*60 + %target_mm%

:: 计算需要等待的总秒数
set /a diff_min=%target_total% - %now_total%
set /a diff_sec=%diff_min% * 60

:: 如果 diff_sec <= 0，说明今天 21:52 已经过了，不再关机
if %diff_sec% LEQ 0 (
    echo 现在时间已经超过 21:52，本脚本将不会关机。
    pause
    exit
)

echo 将在 %diff_sec% 秒后自动关机...
timeout /t %diff_sec%

shutdown -s -t 0
