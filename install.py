import os
import sys
import subprocess
import json
import platform
import shutil
from pathlib import Path

def create_shortcut(target_path, shortcut_name):
    """Create desktop shortcut based on operating system"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    
    if platform.system() == "Windows": # for windows
        import winshell
        from win32com.client import Dispatch
        
        shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = target_path
        shortcut.save()
        
    elif platform.system() == "Darwin":  # the install for macOS
        shortcut_path = os.path.join(desktop, f"{shortcut_name}.command")
        with open(shortcut_path, 'w') as f:
            f.write(f'#!/bin/bash\ncd "$(dirname "$0")"\npython3 "{target_path}"\n')
        os.chmod(shortcut_path, 0o755)
        
    elif platform.system() == "Linux": # for linux
        shortcut_path = os.path.join(desktop, f"{shortcut_name}.desktop")
        with open(shortcut_path, 'w') as f:
            f.write(f"[Desktop Entry]\nType=Application\nName={shortcut_name}\nExec=python3 {target_path}\nTerminal=false\n")
        os.chmod(shortcut_path, 0o755)
    
    print(f"Created shortcut at: {shortcut_path}")

def install_dependencies(): # all dependencies
    """Install required Python packages"""
    required_packages = [
        "pandas",
        "openpyxl",
        "pillow",
        "tkcalendar"
    ]
    
    if platform.system() == "Windows": # windows specific
        required_packages.append("pywin32")
        required_packages.append("winshell")
    
    print("Installing required packages...")
    for package in required_packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Successfully installed {package}")
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}. Please install it manually using: pip install {package}")

def create_folder_structure():
    """Create necessary folders for the application"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # create data directory
    data_dir = os.path.join(base_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    
    # create workplace directories
    workplaces = [
        "IT Service Center",
        "Esports Lounge (Shultz)",
        "Esports Lounge (Williams Center)"
    ]
    
    for workplace in workplaces:
        workplace_dir = os.path.join(data_dir, workplace)
        os.makedirs(workplace_dir, exist_ok=True)
        
        # Create subdirectories
        os.makedirs(os.path.join(workplace_dir, "workers"), exist_ok=True)
        os.makedirs(os.path.join(workplace_dir, "schedules"), exist_ok=True)
        
        # create initial config file
        config = {
            "name": workplace,
            "hours_of_operation": {}
        }
        
        if workplace == "IT Service Center":
            config["hours_of_operation"] = {
                "Monday": "12:00 PM - 7:30 PM",
                "Tuesday": "12:00 PM - 7:30 PM",
                "Wednesday": "12:00 PM - 7:30 PM",
                "Thursday": "12:00 PM - 7:30 PM",
                "Friday": "11:00 AM - 4:00 PM",
                "Saturday": "11:30 AM - 4:30 PM",
                "Sunday": "11:30 AM - 4:30 PM"
            }
            config["shift_times"] = {
                "Monday": ["12:00 PM - 3:00 PM", "3:00 PM - 6:00 PM", "6:00 PM - 8:00 PM", 
                           "12:00 PM - 2:00 PM", "2:00 PM - 5:00 PM", "5:00 PM - 8:00 PM"],
                "Tuesday": ["12:30 PM - 3:00 PM", "3:00 PM - 8:00 PM",
                            "12:00 PM - 3:00 PM", "3:00 PM - 6:00 PM", "6:00 PM - 8:00 PM"],
                "Wednesday": ["12:00 PM - 3:00 PM", "3:00 PM - 6:00 PM", "6:00 PM - 8:00 PM",
                              "12:00 PM - 5:00 PM", "5:00 PM - 8:00 PM"],
                "Thursday": ["12:30 PM - 3:00 PM", "3:00 PM - 6:00 PM", "6:00 PM - 8:00 PM",
                             "12:00 PM - 3:00 PM", "3:00 PM - 8:00 PM"],
                "Friday": ["11:00 AM - 2:00 PM", "2:00 PM - 4:00 PM"],
                "Saturday": ["12:00 PM - 5:00 PM", "12:00 PM - 5:00 PM"],
                "Sunday": ["12:00 PM - 5:00 PM", "12:00 PM - 5:00 PM"]
            }
        elif workplace == "Esports Lounge (Shultz)":
            config["hours_of_operation"] = {
                "Monday": "2:00 PM - 12:00 AM",
                "Tuesday": "2:00 PM - 12:00 AM",
                "Wednesday": "2:00 PM - 12:00 AM",
                "Thursday": "2:00 PM - 12:00 AM",
                "Friday": "2:00 PM - 12:00 AM",
                "Saturday": "12:00 PM - 12:00 AM",
                "Sunday": "12:00 PM - 12:00 AM"
            }
            config["shift_times"] = {
                "Monday": ["2:00 PM - 6:00 PM", "6:00 PM - 9:00 PM", "9:00 PM - 12:00 AM"],
                "Tuesday": ["2:00 PM - 5:00 PM", "5:00 PM - 8:00 PM", "8:00 PM - 12:00 AM"],
                "Wednesday": ["2:00 PM - 6:00 PM", "6:00 PM - 9:00 PM", "9:00 PM - 12:00 AM"],
                "Thursday": ["2:00 PM - 4:00 PM", "4:00 PM - 8:00 PM", "8:00 PM - 12:00 AM"],
                "Friday": ["2:00 PM - 7:00 PM", "7:00 PM - 9:00 PM", "9:00 PM - 12:00 AM"],
                "Saturday": ["12:00 PM - 4:00 PM", "4:00 PM - 7:00 PM", "7:00 PM - 10:00 PM", "10:00 PM - 12:00 AM"],
                "Sunday": ["12:00 PM - 4:00 PM", "4:00 PM - 8:00 PM", "8:00 PM - 12:00 AM"]
            }
        
        with open(os.path.join(workplace_dir, "config.json"), 'w') as f:
            json.dump(config, f, indent=4)
        
        # create empty workers.json file
        with open(os.path.join(workplace_dir, "workers", "workers.json"), 'w') as f:
            json.dump([], f, indent=4)
    
    # create user settings file
    user_settings = {
        "email": "admin@example.com"
    }
    
    with open(os.path.join(data_dir, "settings.json"), 'w') as f:
        json.dump(user_settings, f, indent=4)
    
    print(f"Created folder structure in {data_dir}")

def create_template_excel():
    """Create template Excel files for worker import"""
    try:
        import pandas as pd
        
        base_dir = os.path.dirname(os.path.abspath(__file__))
        templates_dir = os.path.join(base_dir, "templates")
        os.makedirs(templates_dir, exist_ok=True)
        
        # create worker template
        columns = [
            "First Name", "Last Name", 
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
            "Day(s) & Time not Available", "Day(s) & Time not Available", 
            "Email", "Work Study"
        ]
        
        df = pd.DataFrame(columns=columns)
        
        # add example data
        example_data = {
            "First Name": ["John", "Jane"],
            "Last Name": ["Doe", "Smith"],
            "Sunday": ["12pm - 5pm", "na"],
            "Monday": ["12pm - 8pm", "2pm - 6pm"],
            "Tuesday": ["12pm - 8pm", "na"],
            "Wednesday": ["12pm - 8pm", "2pm - 6pm"],
            "Thursday": ["12pm - 8pm", "na"],
            "Friday": ["11am - 4pm", "2pm - 6pm"],
            "Saturday": ["12pm - 5pm", "na"],
            "Day(s) & Time not Available": ["MWF 1pm - 2pm", "TR 12pm - 2pm"],
            "Day(s) & Time not Available.1": ["na", "na"],
            "Email": ["john.doe@example.com", "jane.smith@example.com"],
            "Work Study": ["Y", "N"]
        }
        
        example_df = pd.DataFrame(example_data)
        
        # save template
        template_path = os.path.join(templates_dir, "worker_template.xlsx")
        example_df.to_excel(template_path, index=False)
        
        print(f"Created template Excel file at {template_path}")
    except ImportError:
        print("Failed to create template Excel file. Please make sure pandas and openpyxl are installed.")

def copy_app_files():
    """Copy application files to the installation directory"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # create app files if they don't exist
    app_files = ["main.py", "scheduler.py", "utils.py"]
    
    for file in app_files:
        if not os.path.exists(os.path.join(base_dir, file)):
            print(f"Warning: {file} not found in the installation directory.")

def main():
    print("Starting installation of Workplace Scheduler App...")
    
    # install dependencies
    install_dependencies()
    
    # create folder structure
    create_folder_structure()
    
    # create template Excel files
    create_template_excel()
    
    # copy app files
    copy_app_files()
    
    # create desktop shortcut
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")
    create_shortcut(script_path, "Workplace Scheduler")
    
    print("\nInstallation completed successfully!")
    print("You can now run the application by clicking the desktop shortcut or by running main.py")

if __name__ == "__main__":
    main()
