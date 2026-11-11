@echo off
cd /d "%~dp0"
echo Iniciando pagina web de Reseguimientos Amgen...
echo Abre http://localhost:8501 en el navegador.
echo Para que otras personas de la red la usen: http://NOMBRE-DE-ESTE-EQUIPO:8501
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
