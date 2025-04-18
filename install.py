import os
import sys
import subprocess
import json
import platform
import shutil
import ctypes
from pathlib import Path

def is_admin():
    """Check if the script is running with admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Re-run the script with admin privileges"""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )

def create_windows_shortcut(target_path, shortcut_name):
    """Create Windows desktop shortcut"""
    try:
        import winshell
        from win32com.client import Dispatch
        
        # Get desktop path
        desktop = winshell.desktop()
        
        # Create shortcut
        shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = sys.executable
        shortcut.save()
        
        print(f"Created desktop shortcut at: {shortcut_path}")
        return True
    except Exception as e:
        print(f"Error creating desktop shortcut: {e}")
        return False

def create_windows_batch_file(script_path, name):
    """Create a batch file to run the application"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    batch_path = os.path.join(base_dir, f"{name}.bat")
    
    try:
        with open(batch_path, 'w') as f:
            f.write(f'@echo off\n')
            f.write(f'echo Starting Workplace Scheduler...\n')
            f.write(f'python "{script_path}"\n')
            f.write(f'if %ERRORLEVEL% neq 0 pause\n')
        
        # Also create a desktop batch file
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        desktop_batch = os.path.join(desktop, f"{name}.bat")
        
        try:
            shutil.copy(batch_path, desktop_batch)
            print(f"Created desktop batch file at: {desktop_batch}")
        except Exception as e:
            print(f"Could not create desktop batch file: {e}")
        
        print(f"Created application batch file at: {batch_path}")
        return True
    except Exception as e:
        print(f"Warning: Could not create batch file: {e}")
        return False

def install_dependencies():
    """Install required Python packages"""
    required_packages = [
        "pandas",
        "openpyxl",
        "pillow",
        "tkcalendar",
        "pywin32",
        "winshell"
    ]
    
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
    
    # Create data directory
    data_dir = os.path.join(base_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    
    # Create workplace directories
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
        
        # Create initial config file
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
        
        # Create empty workers.json file
        with open(os.path.join(workplace_dir, "workers", "workers.json"), 'w') as f:
            json.dump([], f, indent=4)
    
    # Create user settings file
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
        
        # Create worker template
        columns = [
            "First Name", "Last Name", 
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
            "Day(s) & Time not Available", "Day(s) & Time not Available", 
            "Email", "Work Study"
        ]
        
        df = pd.DataFrame(columns=columns)
        
        # Add example data
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
        
        # Save template
        template_path = os.path.join(templates_dir, "worker_template.xlsx")
        example_df.to_excel(template_path, index=False)
        
        print(f"Created template Excel file at {template_path}")
    except ImportError:
        print("Failed to create template Excel file. Please make sure pandas and openpyxl are installed.")

def copy_app_files():
    """Copy application files to the installation directory"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Create app files if they don't exist
    app_files = ["main.py", "scheduler.py", "utils.py"]
    
    for file in app_files:
        if not os.path.exists(os.path.join(base_dir, file)):
            print(f"Warning: {file} not found in the installation directory.")

def main():
    # Check if running on Windows
    if platform.system() != "Windows":
        print("This installer is designed for Windows only.")
        sys.exit(1)
    
    print("Starting installation of Workplace Scheduler App...")
    
    # Install dependencies
    install_dependencies()
    
    # Create folder structure
    create_folder_structure()
    
    # Create template Excel files
    create_template_excel()
    
    # Copy app files
    copy_app_files()
    
    # Create desktop shortcut and batch file
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")
    
    # Try to create shortcut
    shortcut_created = create_windows_shortcut(script_path, "Workplace Scheduler")
    
    # Always create batch file as a backup
    create_windows_batch_file(script_path, "Workplace Scheduler")
    
    # If shortcut creation failed and not running as admin, offer to run as admin
    if not shortcut_created and not is_admin():
        print("\nCreating desktop shortcut failed. This might be due to permission restrictions.")
        response = input("Would you like to run the installer with administrator privileges? (y/n): ")
        
        if response.lower() in ['y', 'yes']:
            print("Restarting with admin privileges...")
            run_as_admin()
            sys.exit(0)
    
    print("\nInstallation completed successfully!")
    print("You can now run the application using one of these methods:")
    print("1. Desktop shortcut (if created successfully)")
    print("2. Desktop batch file 'Workplace Scheduler.bat'")
    print("3. Application folder batch file 'Workplace Scheduler.bat'")
    print("4. Run 'python main.py' from the command line in the application folder")

if __name__ == "__main__":
    main()
    
