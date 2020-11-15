rmdir build /S /Q 
rmdir dist /S /Q 
::pyinstaller --noconsole --onefile  -p "ui" -p "..\\Public" DocxHandle.py
pyinstaller --noconsole --onefile  -p "ui" -p "..\\Public" docTrans.py