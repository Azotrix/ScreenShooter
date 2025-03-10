# ScreenShooter
A simple screen shooter application that captures screenshots for use in documents. Works up to 3 monitors.

Prequisites
-------------
The program uses the libraries specified in the `requirements.txt` file.
You can install them using the following terminal command:
```
pip install -r requirements.txt
```

Usage
-----
When launching the program, it creates a `screenshots/` folder in the directory of execution, where screenshots are stored for saving into a Word document.

The program works with the following hotkeys:

| Hotkey | Event | Comment |
|--------|-------|---------|
| Alt+0 | Capturing all monitors | If more monitors are present, their screen is combined into a single PNG image |
| Alt+1 | Capturing monitor 1 | |
| Alt+2 | Capturing monitor 2 | |
| Alt+3 | Capturing monitor 3 | |
| Alt+F9 | Saving screenshots into a Word document | This closes the application and opens the Word document, deleting the screenshots from `screenshots/` |
| Alt+F12 | Exiting application | The user gets promted whether to delete screenshots |


Create an .exe executable
-----
If you would like to create an .exe that you can run on Windows, you need `pyinstaller`. Install it with:
```
pip install pyinstaller
```

And then make the executable with (executed from the folder of the `screenShooter.py` script):
```
pyinstaller --onefile .\screenShooter.py
```

Remarks
-----------
The code might not work on all systems, and you may need to modify it to your needs. The program was tested on Windows 10 and Windows 11 systems.
