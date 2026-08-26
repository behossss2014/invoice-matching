@echo off
cd /d "%~dp0"
echo Checking and installing requirements...
python -m pip install -r requirements.txt
echo.
echo Launching Streamlit App...
python -m streamlit run app.py
pause