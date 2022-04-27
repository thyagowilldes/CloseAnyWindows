
For Generate a new executable follws this steps:

    1- Create a virtual environment (This ensures only library on virtual environment will be in .exe)
        python -m venv venv

    2- Activate virtual environment All library installed on this project must
        .\venv\Scripts\Activate.ps1
    
    3- Instal necessary library (if not installed)

    4- Generate a new executable just run this command on terminal ( must Activate virtual environment):
        pyinstaller --onefile --noconsole --add-data="Tesseract-OCR\*;Tesseract-OCR"  CloseAnyWindow.py 
        pyinstaller --onefile --noconsole  CloseAnyWindow.py

    

