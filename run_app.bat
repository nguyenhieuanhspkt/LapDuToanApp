@echo off
title Khoi dong LapDuToanApp - Streamlit
echo ========================================================
echo   DANG KHOI DONG UNG DUNG QUAN LY VA LAP DU TOAN PCCC
echo ========================================================
echo.
cd /d "D:\TaskApp_kiet\LapDuToanApp"
call "%~dp0venv\Scripts\activate.bat"
streamlit run app.py
pause
