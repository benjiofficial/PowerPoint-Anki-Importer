import sys, os, pwd, subprocess

def get_username():
    return pwd.getpwuid(os.getuid())[0]

try:
    import pptx
except ImportError as e:
    print("PPTX is required")
    pass

try:
    import PyQt6
except ImportError as e:
    print("PyQt6 is required")
    pass

user_path = "/Users/" + str(get_username())
anki_path = user_path + "/Library/Application Support/Anki2/addons21/PPT_Anki/"
print("'" + anki_path + "'")
os.mkdir(anki_path)
subprocess.run(["ls", "-l"])
subprocess.run(["mv", "__init__.py", anki_path])
subprocess.run(["pyinstaller", "--onefile", "--windowed", "main.py", "--name", "PPTX to Anki", "--icon", "icon.ico"])
subprocess.run(["mv", "dist/PPTX to Anki.app", "./"])
subprocess.run(["rm", "-rf", "build"])
subprocess.run(["rm", "-rf", "dist"])
print("Installation Complete")
