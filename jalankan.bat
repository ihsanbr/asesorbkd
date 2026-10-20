@echo off
echo ========================================
echo  Dashboard Monitoring Asesor BKD
echo ========================================
echo.
echo Menginstall dependency jika belum ada...
pip install -r requirements.txt
echo.
echo Menjalankan aplikasi...
streamlit run dashboard_asesor_bkd.py
pause
