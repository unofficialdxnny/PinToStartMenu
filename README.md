# PinToStartMenu
Python Program to pin all apps installed to the start menu

----

This Python code snippet retrieves the names of all programs installed on a Windows system and creates shortcuts for them in the Start Menu folder. This makes it easy to access the programs from the Start Menu.

# Prerequisites

To run this code, you will need:

- Python 3.x installed on your Windows system
- The winshell and win32com libraries installed. You can install them using pip by running the following command:

```py
pip install winshell pywin32
```

----

# How to Run the Code

1. Open a text editor and copy the code snippet from above.
2. Paste the code into a new file and save it with a .py extension (e.g. `start_menu_shortcuts.py`).
3. Open a command prompt or terminal window and navigate to the directory where you saved the file.
4. Run the script by typing python start_menu_shortcuts.py and pressing Enter.

# How the code works
The code uses the winshell library to get the path to the Start Menu folder and a list of all programs installed on the system. It then creates a Windows Shell object using the win32com library and loops through each program, creating a shortcut in the Start Menu folder using the CreateShortcut method of the Windows Shell object.

Each shortcut is given a name based on the program's name, with the .lnk extension. The TargetPath and WorkingDirectory properties of the shortcut are set to the program's location to ensure that the shortcut points to the correct program.

# Conclusion

This code snippet provides a simple and convenient way to retrieve the names of all programs installed on a Windows system and pin them to the Start Menu for easy access.
