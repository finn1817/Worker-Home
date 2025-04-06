import os
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
from tkcalendar import DateEntry
from PIL import Image, ImageTk

from utils import (
    load_json_data, save_json_data, get_workplace_path, get_workers_path,
    get_config_path, get_schedules_path, get_settings_path,
    import_workers_from_excel, export_workers_to_excel,
    parse_time_range, format_time_12hr, backup_data, restore_data
)
from scheduler import ScheduleGenerator

class WorkplaceSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Workplace Scheduler")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # set base directory
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # load settings
        self.settings_path = get_settings_path(self.base_dir)
        self.settings = load_json_data(self.settings_path, {"email": ""})
        
        # initialize current workplace
        self.current_workplace = None
        
        # create main frame
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # create styles
        self.create_styles()
        
        # show dashboard
        self.show_dashboard()
    
    def create_styles(self):
        """Create custom styles for the application"""
        style = ttk.Style()
        
        # configure button styles
        style.configure("TButton", font=("Arial", 12), padding=5)
        style.configure("Workplace.TButton", font=("Arial", 14, "bold"), padding=10)
        style.configure("Action.TButton", font=("Arial", 12), padding=8)
        
        # configure label styles
        style.configure("Title.TLabel", font=("Arial", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("Arial", 14, "bold"))
        style.configure("Info.TLabel", font=("Arial", 12))
        
        # configure frame styles
        style.configure("Card.TFrame", relief="raised", borderwidth=1)
    
    def clear_frame(self, frame):
        """Clear all widgets from a frame"""
        for widget in frame.winfo_children():
            widget.destroy()
    
    def show_dashboard(self):
        """Show the main dashboard"""
        # clear main frame
        self.clear_frame(self.main_frame)
        
        # create header frame
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # add title
        title_label = ttk.Label(header_frame, text="Workplace Scheduler", style="Title.TLabel")
        title_label.pack(side=tk.LEFT)
        
        # add email field
        email_frame = ttk.Frame(header_frame)
        email_frame.pack(side=tk.RIGHT)
        
        email_label = ttk.Label(email_frame, text="User Email:", style="Info.TLabel")
        email_label.pack(side=tk.LEFT, padx=(0, 5))
        
        email_var = tk.StringVar(value=self.settings.get("email", ""))
        email_entry = ttk.Entry(email_frame, textvariable=email_var, width=30)
        email_entry.pack(side=tk.LEFT)
        
        def save_email():
            self.settings["email"] = email_var.get()
            save_json_data(self.settings_path, self.settings)
            messagebox.showinfo("Success", "Email saved successfully!")
        
        save_button = ttk.Button(email_frame, text="Save", command=save_email)
        save_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # create workplaces frame
        workplaces_frame = ttk.Frame(self.main_frame)
        workplaces_frame.pack(fill=tk.BOTH, expand=True)
        
        # add workplaces title
        workplaces_label = ttk.Label(workplaces_frame, text="Workplaces", style="Subtitle.TLabel")
        workplaces_label.pack(anchor=tk.W, pady=(0, 10))
        
        # get list of workplaces
        workplaces_dir = os.path.join(self.base_dir, "data")
        workplaces = []
        
        if os.path.exists(workplaces_dir):
            workplaces = [d for d in os.listdir(workplaces_dir) 
                         if os.path.isdir(os.path.join(workplaces_dir, d)) and d != "backup"]
        
        # add workplace buttons
        for workplace in workplaces:
            workplace_button = ttk.Button(
                workplaces_frame, 
                text=workplace,
                style="Workplace.TButton",
                command=lambda w=workplace: self.open_workplace(w)
            )
            workplace_button.pack(fill=tk.X, pady=5)
        
        # add backup/restore frame
        backup_frame = ttk.Frame(self.main_frame)
        backup_frame.pack(fill=tk.X, pady=(20, 0))
        
        backup_button = ttk.Button(
            backup_frame,
            text="Backup Data",
            command=self.backup_data
        )
        backup_button.pack(side=tk.LEFT, padx=(0, 10))
        
        restore_button = ttk.Button(
            backup_frame,
            text="Restore from Backup",
            command=self.restore_data
        )
        restore_button.pack(side=tk.LEFT)
    
    def open_workplace(self, workplace_name):
        """Open a workplace"""
        self.current_workplace = workplace_name
        
        # check if workplace has workers
        workers_path = get_workers_path(self.base_dir, workplace_name)
        workers = load_json_data(workers_path, [])
        
        if not workers:
            # ask to import workers first
            result = messagebox.askyesno(
                "Import Workers",
                f"No workers found for {workplace_name}. Would you like to import workers from an Excel file?"
            )
            
            if result:
                self.import_workers_excel()
            else:
                # show workplace with empty workers
                self.show_workplace()
        else:
            # show workplace
            self.show_workplace()
    
    def show_workplace(self):
        """Show workplace page"""
        # clear main frame
        self.clear_frame(self.main_frame)
        
        # create header frame
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # add back button
        back_button = ttk.Button(
            header_frame,
            text="← Back to Dashboard",
            command=self.show_dashboard
        )
        back_button.pack(side=tk.LEFT)
        
        # add title
        title_label = ttk.Label(header_frame, text=self.current_workplace, style="Title.TLabel")
        title_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # create actions frame
        actions_frame = ttk.Frame(self.main_frame)
        actions_frame.pack(fill=tk.BOTH, expand=True)
        
        # add action buttons
        view_workers_button = ttk.Button(
            actions_frame,
            text="View Workers and Availability",
            style="Action.TButton",
            command=self.view_workers
        )
        view_workers_button.pack(fill=tk.X, pady=5)
        
        add_worker_button = ttk.Button(
            actions_frame,
            text="Add a New Worker Manually",
            style="Action.TButton",
            command=self.add_worker_manually
        )
        add_worker_button.pack(fill=tk.X, pady=5)
        
        import_workers_button = ttk.Button(
            actions_frame,
            text="Add Workers via Excel Doc",
            style="Action.TButton",
            command=self.import_workers_excel
        )
        import_workers_button.pack(fill=tk.X, pady=5)
        
        create_schedule_button = ttk.Button(
            actions_frame,
            text="Create a Schedule",
            style="Action.TButton",
            command=self.create_schedule
        )
        create_schedule_button.pack(fill=tk.X, pady=5)
        
        last_minute_button = ttk.Button(
            actions_frame,
            text="Find Last Minute Replacement",
            style="Action.TButton",
            command=self.find_replacement
        )
        last_minute_button.pack(fill=tk.X, pady=5)
        
        export_workers_button = ttk.Button(
            actions_frame,
            text="Export Workers to Excel",
            style="Action.TButton",
            command=self.export_workers_excel
        )
        export_workers_button.pack(fill=tk.X, pady=5)
    
    def view_workers(self):
        """View workers and their availability"""
        # clear main frame
        self.clear_frame(self.main_frame)
        
        # create header frame
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # add back button
        back_button = ttk.Button(
            header_frame,
            text="← Back to Workplace",
            command=self.show_workplace
        )
        back_button.pack(side=tk.LEFT)
        
        # add title
        title_label = ttk.Label(header_frame, text=f"Workers - {self.current_workplace}", style="Title.TLabel")
        title_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # load workers
        workers_path = get_workers_path(self.base_dir, self.current_workplace)
        workers = load_json_data(workers_path, [])
        
        # create workers frame
        workers_frame = ttk.Frame(self.main_frame)
        workers_frame.pack(fill=tk.BOTH, expand=True)
        
        # create scrollable frame
        canvas = tk.Canvas(workers_frame)
        scrollbar = ttk.Scrollbar(workers_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # add workers to scrollable frame
        if not workers:
            no_workers_label = ttk.Label(scrollable_frame, text="No workers found.", style="Info.TLabel")
            no_workers_label.pack(pady=20)
        else:
            # create header row
            header_frame = ttk.Frame(scrollable_frame)
            header_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(header_frame, text="Name", width=20, font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
            ttk.Label(header_frame, text="Email", width=30, font=("Arial", 12, "bold")).grid(row=0, column=1, padx=5)
            ttk.Label(header_frame, text="Work Study", width=10, font=("Arial", 12, "bold")).grid(row=0, column=2, padx=5)
            ttk.Label(header_frame, text="Actions", width=10, font=("Arial", 12, "bold")).grid(row=0, column=3, padx=5)
            
            # add separator
            ttk.Separator(scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 10))
            
            # add each worker
            for i, worker in enumerate(workers):
                worker_frame = ttk.Frame(scrollable_frame)
                worker_frame.pack(fill=tk.X, pady=5)
                
                name = f"{worker['first_name']} {worker['last_name']}"
                email = worker.get("email", "")
                work_study = "Yes" if worker.get("work_study", False) else "No"
                
                ttk.Label(worker_frame, text=name, width=20).grid(row=0, column=0, padx=5)
                ttk.Label(worker_frame, text=email, width=30).grid(row=0, column=1, padx=5)
                ttk.Label(worker_frame, text=work_study, width=10).grid(row=0, column=2, padx=5)
                
                # add view details button
                view_button = ttk.Button(
                    worker_frame,
                    text="Details",
                    command=lambda w=worker: self.view_worker_details(w)
                )
                view_button.grid(row=0, column=3, padx=5)
                
                # add remove button
                remove_button = ttk.Button(
                    worker_frame,
                    text="Remove",
                    command=lambda w=worker: self.remove_worker(w)
                )
                remove_button.grid(row=0, column=4, padx=5)
                
                # add separator after each worker except the last one
                if i < len(workers) - 1:
                    ttk.Separator(scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=(5, 0))
    
    def view_worker_details(self, worker):
        """View details for a specific worker"""
        # create a new top-level window
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Worker Details - {worker['first_name']} {worker['last_name']}")
        details_window.geometry("600x500")
        details_window.minsize(600, 500)
        
        # create main frame
        main_frame = ttk.Frame(details_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # add worker info
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(info_frame, text="Name:", style="Info.TLabel").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=f"{worker['first_name']} {worker['last_name']}").grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(info_frame, text="Email:", style="Info.TLabel").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text=worker.get("email", "")).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(info_frame, text="Work Study:", style="Info.TLabel").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text="Yes" if worker.get("work_study", False) else "No").grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        # add availability section
        ttk.Label(main_frame, text="Availability", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # create scrollable frame for availability
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # add availability for each day
        days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        
        for i, day in enumerate(days):
            day_frame = ttk.Frame(scrollable_frame)
            day_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(day_frame, text=f"{day}:", width=15, style="Info.TLabel").grid(row=0, column=0, sticky=tk.W)
            
            # get availability for this day
            availability = worker.get("availability", {}).get(day, [])
            
            if not availability:
                ttk.Label(day_frame, text="Not available").grid(row=0, column=1, sticky=tk.W)
            else:
                availability_text = ", ".join([
                    f"{format_time_12hr(time_range['start'])} - {format_time_12hr(time_range['end'])}"
                    for time_range in availability
                ])
                ttk.Label(day_frame, text=availability_text).grid(row=0, column=1, sticky=tk.W)
        
        # add unavailable times section
        ttk.Label(scrollable_frame, text="Unavailable Times", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(20, 10))
      
        # add unavailable times for each day
        unavailable = worker.get("unavailable", {})
        
        if not unavailable:
            ttk.Label(scrollable_frame, text="No unavailable times specified.").pack(anchor=tk.W)
        else:
            for day, times in unavailable.items():
                day_frame = ttk.Frame(scrollable_frame)
                day_frame.pack(fill=tk.X, pady=5)
                
                ttk.Label(day_frame, text=f"{day}:", width=15, style="Info.TLabel").grid(row=0, column=0, sticky=tk.W)
                
                times_text = ", ".join([
                    f"{format_time_12hr(start)} - {format_time_12hr(end)}"
                    for start, end in times
                ])
                ttk.Label(day_frame, text=times_text).grid(row=0, column=1, sticky=tk.W)
    
    def remove_worker(self, worker):
        """Remove a worker from the workplace"""
        # confirm removal
        result = messagebox.askyesno(
            "Confirm Removal",
            f"Are you sure you want to remove {worker['first_name']} {worker['last_name']}?"
        )
        
        if not result:
            return
        
        # load workers
        workers_path = get_workers_path(self.base_dir, self.current_workplace)
        workers = load_json_data(workers_path, [])
        
        # remove worker
        workers = [w for w in workers if w["id"] != worker["id"]]
        
        # save workers
        save_json_data(workers_path, workers)
        
        # refresh view
        self.view_workers()
    
    def add_worker_manually(self):
        """Add a new worker manually"""
        # create a new top-level window
        add_window = tk.Toplevel(self.root)
        add_window.title(f"Add Worker - {self.current_workplace}")
        add_window.geometry("800x600")
        add_window.minsize(800, 600)
        
        # create main frame
        main_frame = ttk.Frame(add_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # add title
        title_label = ttk.Label(main_frame, text="Add New Worker", style="Title.TLabel")
        title_label.pack(anchor=tk.W, pady=(0, 20))
        
        # create form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # create scrollable frame
        canvas = tk.Canvas(form_frame)
        scrollbar = ttk.Scrollbar(form_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # add form fields
        # basic info
        basic_info_frame = ttk.Frame(scrollable_frame)
        basic_info_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(basic_info_frame, text="Basic Information", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # first name
        first_name_frame = ttk.Frame(basic_info_frame)
        first_name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(first_name_frame, text="First Name:", width=15).pack(side=tk.LEFT)
        first_name_var = tk.StringVar()
        ttk.Entry(first_name_frame, textvariable=first_name_var, width=30).pack(side=tk.LEFT)
        
        # last name
        last_name_frame = ttk.Frame(basic_info_frame)
        last_name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(last_name_frame, text="Last Name:", width=15).pack(side=tk.LEFT)
        last_name_var = tk.StringVar()
        ttk.Entry(last_name_frame, textvariable=last_name_var, width=30).pack(side=tk.LEFT)
        
        # email
        email_frame = ttk.Frame(basic_info_frame)
        email_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(email_frame, text="Email:", width=15).pack(side=tk.LEFT)
        email_var = tk.StringVar()
        ttk.Entry(email_frame, textvariable=email_var, width=30).pack(side=tk.LEFT)
        
        # work study
        work_study_frame = ttk.Frame(basic_info_frame)
        work_study_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(work_study_frame, text="Work Study:", width=15).pack(side=tk.LEFT)
        work_study_var = tk.BooleanVar()
        ttk.Checkbutton(work_study_frame, variable=work_study_var).pack(side=tk.LEFT)
        
        # availability
        availability_frame = ttk.Frame(scrollable_frame)
        availability_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(availability_frame, text="Availability", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # add availability for each day
        days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        availability_vars = {}
        
        for day in days:
            day_frame = ttk.Frame(availability_frame)
            day_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(day_frame, text=f"{day}:", width=15).pack(side=tk.LEFT)
            
            day_var = tk.StringVar()
            ttk.Entry(day_frame, textvariable=day_var, width=30).pack(side=tk.LEFT)
            ttk.Label(day_frame, text="Format: 12pm - 8pm (or 'na' if not available)").pack(side=tk.LEFT, padx=(10, 0))
            
            availability_vars[day] = day_var
        
        # unavailable times
        unavailable_frame = ttk.Frame(scrollable_frame)
        unavailable_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(unavailable_frame, text="Unavailable Times", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # add unavailable time fields
        unavailable_vars = []
        
        for i in range(2):
            unavail_frame = ttk.Frame(unavailable_frame)
            unavail_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(unavail_frame, text=f"Unavailable {i+1}:", width=15).pack(side=tk.LEFT)
            
            unavail_var = tk.StringVar()
            ttk.Entry(unavail_frame, textvariable=unavail_var, width=30).pack(side=tk.LEFT)
            ttk.Label(unavail_frame, text="Format: MWF 1pm - 2pm (or 'na' if not applicable)").pack(side=tk.LEFT, padx=(10, 0))
            
            unavailable_vars.append(unavail_var)
        
        # add save button
        save_button = ttk.Button(
            scrollable_frame,
            text="Save Worker",
            style="Action.TButton",
            command=lambda: self.save_new_worker(
                first_name_var.get(),
                last_name_var.get(),
                email_var.get(),
                work_study_var.get(),
                {day: var.get() for day, var in availability_vars.items()},
                [var.get() for var in unavailable_vars],
                add_window
            )
        )
        save_button.pack(pady=20)
    
    def save_new_worker(self, first_name, last_name, email, work_study, availability, unavailable_times, window):
        """Save a new worker"""
        # validate input
        if not first_name or not last_name:
            messagebox.showerror("Error", "First name and last name are required.")
            return
        
        # load workers
        workers_path = get_workers_path(self.base_dir, self.current_workplace)
        workers = load_json_data(workers_path, [])
        
        # check for duplicate
        for worker in workers:
            if (worker["first_name"].lower() == first_name.lower() and 
                worker["last_name"].lower() == last_name.lower()):
                messagebox.showerror("Error", f"A worker with the name {first_name} {last_name} already exists.")
                return
        
        # process availability
        processed_availability = {}
        for day, time_range in availability.items():
            if not time_range or time_range.lower() == 'na':
                processed_availability[day] = []
            else:
                # split multiple time ranges if present
                ranges = [r.strip() for r in time_range.split(',')]
                day_ranges = []
                for r in ranges:
                    start_time, end_time = parse_time_range(r)
                    if start_time and end_time:
                        day_ranges.append({
                            "start": start_time,
                            "end": end_time
                        })
                processed_availability[day] = day_ranges
        
        # process unavailable times
        processed_unavailable = {}
        for unavail_str in unavailable_times:
            if not unavail_str or unavail_str.lower() == 'na':
                continue
                
            # extract day codes and time range
            import re
            match = re.match(r'([UMTWRFS]+)\s+(.*)', unavail_str)
            if not match:
                continue
                
            day_codes, time_range = match.groups()
            start_time, end_time = parse_time_range(time_range)
            
            if start_time and end_time:
                for day_code in day_codes:
                    day_name = {
                        'U': 'Sunday',
                        'M': 'Monday',
                        'T': 'Tuesday',
                        'W': 'Wednesday',
                        'R': 'Thursday',
                        'F': 'Friday',
                        'S': 'Saturday'
                    }.get(day_code.upper())
                    
                    if day_name:
                        if day_name not in processed_unavailable:
                            processed_unavailable[day_name] = []
                        processed_unavailable[day_name].append((start_time, end_time))
        
        # create worker object
        worker = {
            "id": f"{first_name.lower()}_{last_name.lower()}_{len(workers)}",
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "work_study": work_study,
            "availability": processed_availability,
            "unavailable": processed_unavailable,
            "weekly_hours": 0
        }
        
        # add worker to list
        workers.append(worker)
        
        # save workers
        save_json_data(workers_path, workers)
        
        # close window
        window.destroy()
        
        # show success message
        messagebox.showinfo("Success", f"Worker {first_name} {last_name} added successfully!")
        
        # refresh view if on workers page
        if hasattr(self, 'view_workers'):
            self.view_workers()
    
    def import_workers_excel(self):
        """Import workers from Excel file"""
        # open file dialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            # import workers
            new_workers = import_workers_from_excel(file_path)
            
            # load existing workers
            workers_path = get_workers_path(self.base_dir, self.current_workplace)
            existing_workers = load_json_data(workers_path, [])
            
            # check for duplicates
            existing_names = {(w["first_name"].lower(), w["last_name"].lower()) for w in existing_workers}
            
            added_count = 0
            skipped_count = 0
            
            for worker in new_workers:
                name_key = (worker["first_name"].lower(), worker["last_name"].lower())
                
                if name_key in existing_names:
                    skipped_count += 1
                    continue
                
                existing_workers.append(worker)
                existing_names.add(name_key)
                added_count += 1
            
            # save workers
            save_json_data(workers_path, existing_workers)
            
            # show success message
            messagebox.showinfo(
                "Import Complete",
                f"Import complete!\n\nAdded: {added_count} workers\nSkipped: {skipped_count} duplicates"
            )
            
            # show workplace
            self.show_workplace()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error importing workers: {str(e)}")
    
    def export_workers_excel(self):
        """Export workers to Excel file"""
        # load workers
        workers_path = get_workers_path(self.base_dir, self.current_workplace)
        workers = load_json_data(workers_path, [])
        
        if not workers:
            messagebox.showinfo("Export", "No workers to export.")
            return
        
        # open file dialog
        file_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            # export workers
            success = export_workers_to_excel(workers, file_path)
            
            if success:
                messagebox.showinfo("Export Complete", f"Workers exported successfully to {file_path}")
            else:
                messagebox.showerror("Export Error", "An error occurred while exporting workers.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting workers: {str(e)}")
    
    def create_schedule(self):
        """Create a new schedule"""
        # load workers
        workers_path = get_workers_path(self.base_dir, self.current_workplace)
        workers = load_json_data(workers_path, [])
        
        if not workers:
            messagebox.showinfo("Create Schedule", "No workers available to create a schedule.")
            return
        
        # create a new top-level window
        schedule_window = tk.Toplevel(self.root)
        schedule_window.title(f"Create Schedule - {self.current_workplace}")
        schedule_window.geometry("800x600")
        schedule_window.minsize(800, 600)
        
        # create main frame
        main_frame = ttk.Frame(schedule_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # add title
        title_label = ttk.Label(main_frame, text="Create Schedule", style="Title.TLabel")
        title_label.pack(anchor=tk.W, pady=(0, 20))
        
        # add instructions
        instructions_label = ttk.Label(
            main_frame, 
            text="Select workers to include in the schedule. All selected workers will be considered for shifts based on their availability.",
            wraplength=780
        )
        instructions_label.pack(anchor=tk.W, pady=(0, 20))
        
        # create workers frame
        workers_frame = ttk.Frame(main_frame)
        workers_frame.pack(fill=tk.BOTH, expand=True)
        
        # create scrollable frame
        canvas = tk.Canvas(workers_frame)
        scrollbar = ttk.Scrollbar(workers_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # add workers with checkboxes
        worker_vars = {}
        
        for worker in workers:
            worker_frame = ttk.Frame(scrollable_frame)
            worker_frame.pack(fill=tk.X, pady=5)
            
            var = tk.BooleanVar(value=True)
            worker_vars[worker["id"]] = var
            
            checkbox = ttk.Checkbutton(worker_frame, variable=var)
            checkbox.pack(side=tk.LEFT)
            
            name = f"{worker['first_name']} {worker['last_name']}"
            email = worker.get("email", "")
            work_study = "Work Study" if worker.get("work_study", False) else ""
            
            ttk.Label(worker_frame, text=name, width=20).pack(side=tk.LEFT)
            ttk.Label(worker_frame, text=email, width=30).pack(side=tk.LEFT)
            ttk.Label(worker_frame, text=work_study, width=15).pack(side=tk.LEFT)
        
        # add buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(20, 0))
        
        select_all_var = tk.BooleanVar(value=True)
        
        def toggle_all():
            for var in worker_vars.values():
                var.set(select_all_var.get())
        
        select_all_check = ttk.Checkbutton(
            buttons_frame, 
            text="Select All", 
            variable=select_all_var,
            command=toggle_all
        )
        select_all_check.pack(side=tk.LEFT)
        
        create_button = ttk.Button(
            buttons_frame,
            text="Create Schedule with Selected Workers",
            style="Action.TButton",
            command=lambda: self.generate_schedule(
                {id: var.get() for id, var in worker_vars.items()},
                schedule_window
            )
        )
        create_button.pack(side=tk.RIGHT)
    
    def generate_schedule(self, worker_selections, window):
        """Generate a schedule with selected workers"""
        # get selected worker IDs
        selected_worker_ids = [id for id, selected in worker_selections.items() if selected]
        
        if not selected_worker_ids:
            messagebox.showinfo("Create Schedule", "No workers selected. Please select at least one worker.")
            return
        
        # create schedule generator
        generator = ScheduleGenerator(self.base_dir, self.current_workplace)
        
        # generate schedule
        schedule = generator.generate_schedule(selected_worker_ids)
        
        # close worker selection window
        window.destroy()
        
        # show schedule editor
        self.show_schedule_editor(schedule)
    
    def show_schedule_editor(self, schedule):
        """Show schedule editor"""
        # create a new top-level window
        editor_window = tk.Toplevel(self.root)
        editor_window.title(f"Schedule Editor - {self.current_workplace}")
        editor_window.geometry("1000x700")
        editor_window.minsize(1000, 700)
        
        # create main frame
        main_frame = ttk.Frame(editor_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # add title
        title_label = ttk.Label(main_frame, text="Schedule Editor", style="Title.TLabel")
        title_label.pack(anchor=tk.W, pady=(0, 20))
        
        # create schedule frame
        schedule_frame = ttk.Frame(main_frame)
        schedule_frame.pack(fill=tk.BOTH, expand=True)
        
        # create notebook for days
        notebook = ttk.Notebook(schedule_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # add a tab for each day
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_frames = {}
        
        for day in days:
            day_frame = ttk.Frame(notebook)
            notebook.add(day_frame, text=day)
            day_frames[day] = day_frame
            
            # add shifts for this day
            if day in schedule["days"]:
                shifts = schedule["days"][day]
                
                # create shifts table
                shifts_frame = ttk.Frame(day_frame, padding=10)
                shifts_frame.pack(fill=tk.BOTH, expand=True)
                
                # add header row
                header_frame = ttk.Frame(shifts_frame)
                header_frame.pack(fill=tk.X, pady=(0, 10))
                
                ttk.Label(header_frame, text="Start Time", width=15, font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
                ttk.Label(header_frame, text="End Time", width=15, font=("Arial", 12, "bold")).grid(row=0, column=1, padx=5)
                ttk.Label(header_frame, text="Worker", width=30, font=("Arial", 12, "bold")).grid(row=0, column=2, padx=5)
                ttk.Label(header_frame, text="Duration", width=10, font=("Arial", 12, "bold")).grid(row=0, column=3, padx=5)
                
                # add separator
                ttk.Separator(shifts_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 10))
                
                # Add each shift
                for i, shift in enumerate(shifts):
                    shift_frame = ttk.Frame(shifts_frame)
                    shift_frame.pack(fill=tk.X, pady=5)
                    
                    start_time = format_time_12hr(shift["start_time"])
                    end_time = format_time_12hr(shift["end_time"])
                    worker_name = shift["worker_name"]
                    duration = f"{shift['duration_hours']:.2f} hrs"
                    
                    ttk.Label(shift_frame, text=start_time, width=15).grid(row=0, column=0, padx=5)
                    ttk.Label(shift_frame, text=end_time, width=15).grid(row=0, column=1, padx=5)
                    
                    # create worker selection dropdown
                    worker_var = tk.StringVar(value=worker_name)
                    
                    # get all workers
                    workers_path = get_workers_path(self.base_dir, self.current_workplace)
                    workers = load_json_data(workers_path, [])
                    
                    # create list of worker names
                    worker_names = ["UNASSIGNED"] + [f"{w['first_name']} {w['last_name']}" for w in workers]
                    
                    worker_dropdown = ttk.Combobox(shift_frame, textvariable=worker_var, values=worker_names, width=30)
                    worker_dropdown.grid(row=0, column=2, padx=5)
                    
                    # update shift when worker changes
                    def update_worker(event, day=day, shift_index=i, var=worker_var):
                        new_worker_name = var.get()
                        
                        if new_worker_name == "UNASSIGNED":
                            schedule["days"][day][shift_index]["worker_id"] = None
                            schedule["days"][day][shift_index]["worker_name"] = "UNASSIGNED"
                        else:
                            # find worker by name
                            for worker in workers:
                                if f"{worker['first_name']} {worker['last_name']}" == new_worker_name:
                                    schedule["days"][day][shift_index]["worker_id"] = worker["id"]
                                    schedule["days"][day][shift_index]["worker_name"] = new_worker_name
                                    break
                        
                        # update worker summary
                        self.update_worker_summary(schedule)
                    
                    worker_dropdown.bind("<<ComboboxSelected>>", update_worker)
                    
                    ttk.Label(shift_frame, text=duration, width=10).grid(row=0, column=3, padx=5)
                    
                    # add separator after each shift except the last one
                    if i < len(shifts) - 1:
                        ttk.Separator(shifts_frame, orient="horizontal").pack(fill=tk.X, pady=(5, 0))
            else:
                # no shifts for this day
                ttk.Label(day_frame, text="No shifts scheduled for this day.", padding=20).pack()
        
        # add worker summary
        summary_frame = ttk.Frame(main_frame)
        summary_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Label(summary_frame, text="Worker Summary", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))
        
        # create scrollable frame for worker summary
        summary_canvas = tk.Canvas(summary_frame, height=150)
        summary_scrollbar = ttk.Scrollbar(summary_frame, orient="vertical", command=summary_canvas.yview)
        summary_scrollable_frame = ttk.Frame(summary_canvas)
        
        summary_scrollable_frame.bind(
            "<Configure>",
            lambda e: summary_canvas.configure(scrollregion=summary_canvas.bbox("all"))
        )
        
        summary_canvas.create_window((0, 0), window=summary_scrollable_frame, anchor="nw")
        summary_canvas.configure(yscrollcommand=summary_scrollbar.set)
        
        summary_canvas.pack(side="left", fill="x", expand=True)
        summary_scrollbar.pack(side="right", fill="y")
        
        # add worker summary
        self.update_worker_summary(schedule, summary_scrollable_frame)
        
        # add buttons
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(20, 0))
        
        save_button = ttk.Button(
            buttons_frame,
            text="Finish Creating Schedule",
            style="Action.TButton",
            command=lambda: self.save_schedule(schedule, editor_window)
        )
        save_button.pack(side=tk.RIGHT)
    
    def update_worker_summary(self, schedule, frame=None):
        """Update worker summary in schedule editor"""
        # calculate hours for each worker
        worker_hours = {}
        
        for day, shifts in schedule["days"].items():
            for shift in shifts:
                worker_id = shift["worker_id"]
                if worker_id:
                    if worker_id not in worker_hours:
                        worker_hours[worker_id] = 0
                    worker_hours[worker_id] += shift["duration_hours"]
        
        # update worker summary in schedule
        for worker in schedule["workers"]:
            worker["weekly_hours"] = worker_hours.get(worker["id"], 0)
        
        # update UI if frame is provided
        if frame:
            # clear frame
            for widget in frame.winfo_children():
                widget.destroy()
            
            # add header row
            header_frame = ttk.Frame(frame)
            header_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(header_frame, text="Name", width=30, font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
            ttk.Label(header_frame, text="Work Study", width=15, font=("Arial", 12, "bold")).grid(row=0, column=1, padx=5)
            ttk.Label(header_frame, text="Weekly Hours", width=15, font=("Arial", 12, "bold")).grid(row=0, column=2, padx=5)
            ttk.Label(header_frame, text="Status", width=20, font=("Arial", 12, "bold")).grid(row=0, column=3, padx=5)
            
            # add separator
            ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=(0, 10))
            
            # add each worker
            for i, worker in enumerate(sorted(schedule["workers"], key=lambda w: w["name"])):
                worker_frame = ttk.Frame(frame)
                worker_frame.pack(fill=tk.X, pady=5)
                
                name = worker["name"]
                work_study = "Yes" if worker["work_study"] else "No"
                hours = f"{worker['weekly_hours']:.2f}"
                
                # determine status
                status = ""
                if worker["work_study"]:
                    if worker["weekly_hours"] < 5:
                        status = f"Needs {5 - worker['weekly_hours']:.2f} more hours"
                    elif worker["weekly_hours"] > 5:
                        status = f"{worker['weekly_hours'] - 5:.2f} hours over limit"
                    else:
                        status = "Perfect (5 hours)"
                
                ttk.Label(worker_frame, text=name, width=30).grid(row=0, column=0, padx=5)
                ttk.Label(worker_frame, text=work_study, width=15).grid(row=0, column=1, padx=5)
                ttk.Label(worker_frame, text=hours, width=15).grid(row=0, column=2, padx=5)
                ttk.Label(worker_frame, text=status, width=20).grid(row=0, column=3, padx=5)
                
                # add separator after each worker except the last one
                if i < len(schedule["workers"]) - 1:
                    ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=(5, 0))
    
    def save_schedule(self, schedule, window):
        """Save the schedule"""
        # create schedule generator
        generator = ScheduleGenerator(self.base_dir, self.current_workplace)
        
        # save schedule
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"schedule_{timestamp}.json"
        schedule_path = generator.save_schedule(schedule, filename)
        
        # export to Excel
        excel_path = os.path.join(generator.schedules_path, f"schedule_{timestamp}.xlsx")
        generator.export_schedule_to_excel(schedule, excel_path)
        
        # close window
        window.destroy()
        
        # show success message
        messagebox.showinfo(
            "Schedule Created",
            f"Schedule created successfully!\n\nJSON saved to: {os.path.basename(schedule_path)}\nExcel saved to: {os.path.basename(excel_path)}"
        )
        
        # ask if user wants to open Excel file
        result = messagebox.askyesno(
            "Open Excel File",
            "Would you like to open the Excel file?"
        )
        
        if result:
            # open Excel file
            import subprocess
            import platform
            
            if platform.system() == "Windows": # windows
                os.startfile(excel_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.call(["open", excel_path])
            else:  # linux
                subprocess.call(["xdg-open", excel_path])
    
    def find_replacement(self):
        """Find a replacement worker for a shift"""
        # create a new top-level window
        replacement_window = tk.Toplevel(self.root)
        replacement_window.title(f"Find Replacement - {self.current_workplace}")
        replacement_window.geometry("800x600")
        replacement_window.minsize(800, 600)
        
        # create main frame
        main_frame = ttk.Frame(replacement_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # add title
        title_label = ttk.Label(main_frame, text="Find Last Minute Replacement", style="Title.TLabel")
        title_label.pack(anchor=tk.W, pady=(0, 20))
        
        # add instructions
        instructions_label = ttk.Label(
            main_frame, 
            text="Enter the day and time of the shift to find available workers.",
            wraplength=780
        )
        instructions_label.pack(anchor=tk.W, pady=(0, 20))
        
        # create form frame
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X, pady=(0, 20))
        
        # day selection
        day_frame = ttk.Frame(form_frame)
        day_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(day_frame, text="Day:", width=15).pack(side=tk.LEFT)
        
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_var = tk.StringVar(value=days[0])
        day_dropdown = ttk.Combobox(day_frame, textvariable=day_var, values=days, width=20)
        day_dropdown.pack(side=tk.LEFT)
        
        # time range
        time_frame = ttk.Frame(form_frame)
        time_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(time_frame, text="Time Range:", width=15).pack(side=tk.LEFT)
        
        start_time_var = tk.StringVar()
        start_time_entry = ttk.Entry(time_frame, textvariable=start_time_var, width=10)
        start_time_entry.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text="to").pack(side=tk.LEFT, padx=5)
        
        end_time_var = tk.StringVar()
        end_time_entry = ttk.Entry(time_frame, textvariable=end_time_var, width=10)
        end_time_entry.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text="Format: 2pm or 2:00 PM").pack(side=tk.LEFT, padx=(10, 0))
        
        # add search button
        search_button = ttk.Button(
            form_frame,
            text="Find Available Workers",
            style="Action.TButton",
            command=lambda: self.show_available_workers(
                day_var.get(),
                start_time_var.get(),
                end_time_var.get(),
                results_frame
            )
        )
        search_button.pack(pady=10)
        
        # create results frame
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # add initial message
        ttk.Label(results_frame, text="Enter shift details and click 'Find Available Workers'").pack(pady=20)
    
    def show_available_workers(self, day, start_time, end_time, results_frame):
        """Show available workers for a shift"""
        # clear results frame
        for widget in results_frame.winfo_children():
            widget.destroy()
        
        # validate input
        if not day or not start_time or not end_time:
            ttk.Label(results_frame, text="Please enter day and time range.").pack(pady=20)
            return
        
        # parse time range
        from utils import convert_time_to_24hr
        start_time_24hr = convert_time_to_24hr(start_time)
        end_time_24hr = convert_time_to_24hr(end_time)
        
        if not start_time_24hr or not end_time_24hr:
            ttk.Label(results_frame, text="Invalid time format. Please use format like '2pm' or '2:00 PM'.").pack(pady=20)
            return
        
        # create schedule generator
        generator = ScheduleGenerator(self.base_dir, self.current_workplace)
        
        # find available workers
        available_workers = generator.find_replacement_workers(day, start_time_24hr, end_time_24hr)
        
        # show results
        ttk.Label(
            results_frame, 
            text=f"Available Workers for {day}, {format_time_12hr(start_time_24hr)} to {format_time_12hr(end_time_24hr)}:",
            style="Subtitle.TLabel"
        ).pack(anchor=tk.W, pady=(0, 10))
        
        if not available_workers:
            ttk.Label(results_frame, text="No workers available for this shift.").pack(pady=20)
            return
        
        # create scrollable frame
        canvas = tk.Canvas(results_frame)
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # add header row
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Name", width=30, font=("Arial", 12, "bold")).grid(row=0, column=0, padx=5)
        ttk.Label(header_frame, text="Email", width=30, font=("Arial", 12, "bold")).grid(row=0, column=1, padx=5)
        ttk.Label(header_frame, text="Work Study", width=15, font=("Arial", 12, "bold")).grid(row=0, column=2, padx=5)
        ttk.Label(header_frame, text="Weekly Hours", width=15, font=("Arial", 12, "bold")).grid(row=0, column=3, padx=5)
        
        # add separator
        ttk.Separator(scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 10))
        
        # add each worker
        for i, worker in enumerate(available_workers):
            worker_frame = ttk.Frame(scrollable_frame)
            worker_frame.pack(fill=tk.X, pady=5)
            
            name = f"{worker['first_name']} {worker['last_name']}"
            email = worker.get("email", "")
            work_study = "Yes" if worker.get("work_study", False) else "No"
            hours = f"{worker.get('weekly_hours', 0):.2f}"
            
            ttk.Label(worker_frame, text=name, width=30).grid(row=0, column=0, padx=5)
            ttk.Label(worker_frame, text=email, width=30).grid(row=0, column=1, padx=5)
            ttk.Label(worker_frame, text=work_study, width=15).grid(row=0, column=2, padx=5)
            ttk.Label(worker_frame, text=hours, width=15).grid(row=0, column=3, padx=5)
            
            # add separator after each worker except the last one
            if i < len(available_workers) - 1:
                ttk.Separator(scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=(5, 0))
    
    def backup_data(self):
        """Backup all data"""
        # open file dialog
        file_path = filedialog.asksaveasfilename(
            title="Save Backup",
            defaultextension=".zip",
            filetypes=[("Zip files", "*.zip")]
        )
        
        if not file_path:
            return
        
        # create backup
        success = backup_data(self.base_dir, file_path.replace(".zip", ""))
        
        if success:
            messagebox.showinfo("Backup Complete", f"Data backed up successfully to {file_path}")
        else:
            messagebox.showerror("Backup Error", "An error occurred while creating the backup.")
    
    def restore_data(self):
        """Restore data from backup"""
        # open file dialog
        file_path = filedialog.askopenfilename(
            title="Select Backup File",
            filetypes=[("Zip files", "*.zip")]
        )
        
        if not file_path:
            return
        
        # confirm restore
        result = messagebox.askyesno(
            "Confirm Restore",
            "Restoring from backup will replace all current data. Are you sure you want to continue?"
        )
        
        if not result:
            return
        
        # restore backup
        success = restore_data(self.base_dir, file_path)
        
        if success:
            messagebox.showinfo("Restore Complete", "Data restored successfully. The application will now restart.")
            
            # restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)
        else:
            messagebox.showerror("Restore Error", "An error occurred while restoring from backup.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WorkplaceSchedulerApp(root)
    root.mainloop()
  
