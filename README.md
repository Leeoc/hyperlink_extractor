# Hyperlink Extractor for docx

This script will extract all the links within a document and display them in a table for viewing. 
If you click on a  link within the table it will copy it to the clipboard for pasting later. 

## To run w/Python
- Clone the repo
- Cd into the project with `cd <.../.../.../>hyperlink_extractor`
- Create a venv with `python3 -m venv env`
- Activate venv with (windows) `env/Scripts/activate.bat` (linux/macos) `source env/bin/activate`
- Install requirements with `pip install requirements.txt`
- Finally run it using `python main.py`

## To compile into .exe 
- Clone the repo
- Cd into the project with `cd <.../.../.../>hyperlink_extractor`
- Create a venv with `python3 -m venv env`
- Activate venv with (windows) `env/Scripts/activate.bat` (linux/macos) `source env/bin/activate`
- Install requirements with `pip install requirements.txt`
- Use `pyinstaller --onefile .\main.py` to compile to `.exe`
- Open the `dist` folder and double click the `main.exe` file to run