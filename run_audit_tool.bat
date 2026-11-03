@echo off
title THAM TRA DU LIEU THO PAKT - EVN
cd /d "D:\TaskApp_kiet\LapDuToanApp"

echo ==============================================================================
echo       HE THONG THAM TRA VA BOC TACH DU LIEU THO PAKT - EVN (STANDALONE)
echo ==============================================================================

netstat -ano | findstr :8502 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [THONG BAO] He thong da dang chay tren cong 8502.
    echo Dang mo trinh duyet Web tai: http://localhost:8502 ...
    start http://localhost:8502
    goto end
)

echo Dang khoi dong may chu Streamlit tren cong 8502...
start http://localhost:8502
streamlit run audit_app.py --server.port 8502

:end
