@echo off
chcp 65001 >nul
cd /d "%~dp0"
set SRC=%CD%\data\工作计划-测试数据.xls
set ADV=%CD%\承载力报告\建议\advice_20260901.json
set OUT=%CD%\承载力报告
set OFFICE=国网青岛供电公司安全生产委员会办公室
echo 使用最新模板 + native 可编辑 docx，锚定 2026-09-01，出日/周/月...
python scripts\runReports.py --source "%SRC%" --all --report-date 2026-09-01 --advice "%ADV%" --office "%OFFICE%" --out-dir "%OUT%" --chart native
echo.
echo 成品在：%OUT%\承载力报告YYYYMMDD\
pause
