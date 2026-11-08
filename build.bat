@echo off
chcp 65001 >nul
echo 正在本地打包 Windows exe（离线，不联网运行）...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name MiniAssistant main.py
echo.
echo 打包完成，exe 在 dist\MiniAssistant.exe
pause
