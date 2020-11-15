rmdir build /S /Q 
rmdir dist /S /Q 
pyinstaller --noconsole -p "ui" -p "interface" -p "..\\Public"   docTrans.py
REM @echo D|xcopy "config" "dist\\Config" /s /e
pause