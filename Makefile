

main: main.py gui.py merge.py tasks.py logs.py
	pyinstaller main.py --onefile --windowed --icon=favicon.ico -y