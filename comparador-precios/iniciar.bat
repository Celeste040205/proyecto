@echo off
cd /d "%~dp0"
if not exist ".env" (
  echo Falta el archivo .env. Copia .env.example como .env y pega tu clave.
  pause
  exit /b
)
if not exist "venv\Scripts\python.exe" (
  echo Primera vez: instalando dependencias...
  python -m venv venv
  venv\Scripts\python -m pip install -r requirements.txt
)
venv\Scripts\python app.py
pause
