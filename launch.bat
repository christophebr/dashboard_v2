@echo off

:: Activate conda environment
call conda activate streamlitenv

:: Change to the script directory
cd /d "%~dp0"

:: Run the app with streamlit
streamlit run app.py

pause 