@echo off
chcp 65001 > nul
title CAD Quantity Pro v2.0

echo ==========================================
echo    CAD Quantity Pro v2.0 실행
echo ==========================================
echo.

python CAD_Quantity_Pro.py

if errorlevel 1 (
    echo.
    echo [오류] 프로그램 실행 중 문제가 발생했습니다.
    echo Python이 설치되어 있는지 확인하세요.
    pause
)