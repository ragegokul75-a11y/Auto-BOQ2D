@echo off
title AutoBOQ 2D Guided Wizard - Precision Takeoff, BOQ & Tender Steel
cd /d "%~dp0"
echo ====================================================================
echo   AutoBOQ 2D Guided Wizard (DXF / DWG / PDF)
echo   Step 1: Upload Drawing  -^>  Step 2: Click/Trace on Plan  -^>  Step 3: Export BOQ
echo   Opening browser at http://127.0.0.1:8095 ...
echo ====================================================================
start "" "http://127.0.0.1:8095"
python -m uvicorn app:app --host 127.0.0.1 --port 8095
pause
