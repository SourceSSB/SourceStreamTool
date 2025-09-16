from flask import Flask, render_template, request, redirect
import os
import sys
import winshell
from win32com.client import Dispatch
import traceback
import re
import webview

app = Flask(__name__)
app2_template = "http://www.example.com/index.html"

bestof_list = ["1","2","3","4","5","6","7","8","9","10"]

# Determine base directory for script or EXE
if getattr(sys, "frozen", False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Create and use a dedicated folder for text files
TEXT_FILES_FOLDER = os.path.join(base_dir, "Text Files")
os.makedirs(TEXT_FILES_FOLDER, exist_ok=True)

def get_file_path(filename):
    return os.path.join(TEXT_FILES_FOLDER, filename)

def read_or_default(filename, default=""):
    try:
        with open(get_file_path(filename), "r", encoding="utf-8") as f:
            content = f.read().strip()
            if filename == "bestof.txt":
                match = re.match(r"Best of (\d+)", content)
                if match:
                    return match.group(1)  # Return just the number part (e.g., "3")
                else:
                    return default  # Fallback if pattern doesn't match
                
            return content
    except FileNotFoundError:
        return default


def create_shortcut():
    try:
        shortcut_path = os.path.join(base_dir, "StreamController.lnk")

        # Handle icon path depending on if we're in a PyInstaller bundle or not
        if getattr(sys, "frozen", False):
            # In EXE mode: __MEIPASS is where PyInstaller extracts files
            icon_path = os.path.join(sys._MEIPASS, "static", "logo.ico")
        else:
            icon_path = os.path.join(base_dir, "static", "logo.ico")

        target_path = "C:\\Windows\\System32\\cmd.exe"
        arguments = '/c start http://localhost:5000/'

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = target_path
        shortcut.Arguments = arguments
        shortcut.WorkingDirectory = base_dir
        if os.path.exists(icon_path):
            shortcut.IconLocation = icon_path
        shortcut.Save()
    except Exception as e:
        print("Failed to create shortcut:")
        traceback.print_exc()

@app.route("/", methods=["GET", "POST"])

def index():
    if request.method == "POST":
        try:
            form_data = request.get_json()
            if not form_data:
                form_data = request.form

            for key, value in form_data.items():
                path = get_file_path(f"{key}.txt")
                with open(path, "w", encoding="utf-8") as f:
                    if key == "bestof":
                        f.write("Best of " + str(value))
                    else:
                        f.write(str(value))

            return "", 204
        except Exception as e:
            print("=== ERROR WHILE SAVING ===")
            traceback.print_exc()
            return f"Error saving data: {e}", 500

    # Handle GET — load data
    data = {
        "host": read_or_default("host.txt", "Host Name"),
        "round": read_or_default("round.txt", "Round"),
        "event": read_or_default("event.txt", "Event Name"),
        "player1": read_or_default("player1.txt", "Player 1"),
        "player2": read_or_default("player2.txt", "Player 2"),
        "player3": read_or_default("player3.txt", "Player 3"),
        "player4": read_or_default("player4.txt", "Player 4"),
        "score1": read_or_default("score1.txt", "0"),
        "score2": read_or_default("score2.txt", "0"),
        "score3": read_or_default("score3.txt", "0"),
        "score4": read_or_default("score4.txt", "0")
    }
    return render_template("index.html", **data)

webview.create_window("Test",app)

if __name__ == "__main__":
    create_shortcut()
    app.run(debug="True")
    webview.start()