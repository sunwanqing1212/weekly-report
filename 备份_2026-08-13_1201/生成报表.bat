@echo off
cd /d "%~dp0"

echo [1/3] 计算数据并导出 JSON ...
node dump_data.js
if errorlevel 1 goto :err

echo [2/3] 生成 xlsx 报表（原生图表） ...
tools\python\python.exe gen_xlsx.py
if errorlevel 1 goto :err

echo [3/3] 生成 HTML 报表 ...
node gen_report.js
if errorlevel 1 goto :err

echo.
echo ============================================
echo  生成完成（文件名含统计日期和周数）：
echo    分地市周报对比_用户数与流量_日期_第N周.xlsx
echo    分地市周报对比_用户数与流量_日期_第N周.html
echo ============================================
pause
exit /b 0

:err
echo.
echo 生成失败，请检查上方错误信息。
pause
exit /b 1
