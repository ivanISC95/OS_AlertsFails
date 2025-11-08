Build EXE
pyinstaller --onefile --noconsole main.py

pyinstaller --onefile --noconsole --add-data "cp_data.pkl;." main.py
