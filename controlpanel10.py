import tkinter as tk
import os.path
from pathlib import Path

# pip install pywin32
from win32com.client import Dispatch

home = str(Path.home())
batch_folder = os.path.join(home, "AppData", "Roaming", "controlpanel10")
shortcuts = os.path.join(home, "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "controlpanel10")
commands_path = "config/defaults/commands.txt"

commands = {}

def read_commands():
    with open(commands_path) as file:
        for line in file.readlines():
            split = line.split(" - ", 1)
            commands[split[0]] = split[1]

def create_batch():
    # Get and makedirs for batch_folder
    Path(batch_folder).mkdir(parents=True, exist_ok=True)

    for command in commands.keys():
        batch_file = os.path.join(batch_folder, command + ".bat")
        with open(batch_file, "w") as file:
            file.write(f"start \"\" {commands[command]}")
        print(f"Batch program {batch_file} has been created")


def create_shortcuts():
    # Get and makedirs for shortcuts
    Path(shortcuts).mkdir(parents=True, exist_ok=True)
    
    # Create a shortcut for each command
    for command in commands.keys():
        if(commands[command]):
            path = os.path.join(shortcuts, command + ".lnk")

            # If the file doesn't exist, create it
            if not os.path.exists(path):
                f = open(path, "x")
            
            target = os.path.join(batch_folder, command + ".bat")

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.IconLocation = commands[command]
            shortcut.save()
            print(f"Shortcut {path} has been created")
    pass

def restore_control_panel():
    create_batch()
    create_shortcuts()

def start():
    read_commands()

    window = tk.Tk()
    window.title('controlpanel10')
    window.geometry('300x50')

    frame = tk.Frame(window)
    frame.pack()

    button_restore_control_panel = tk.Button(window,
                                            text='Restore Control Panel',
                                            command=restore_control_panel)
    button_restore_control_panel.place(relx=.5, rely=.5, anchor="center")

    window.mainloop()

if(__name__ == "__main__"):
    start()