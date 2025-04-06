import os
import json
import datetime
import pandas as pd
from datetime import datetime, timedelta
import re

# time conversion functions
def convert_time_to_24hr(time_str):
    """Convert time string like '2:00 PM' to 24-hour format (14:00)"""
    if time_str.lower() == 'na' or not time_str:
        return None
    
    try:
        # handle various time formats
        if ' - ' in time_str:
            # if it's a range, just get the first part
            time_str = time_str.split(' - ')[0]
        
        # remove any non-time characters
        time_str = time_str.strip()
        
        # parse the time
        if 'am' in time_str.lower() or 'pm' in time_str.lower():
            # format like "2pm" or "2:00pm"
            time_str = time_str.lower().replace('am', ' AM').replace('pm', ' PM')
            time_str = time_str.replace('a.m.', ' AM').replace('p.m.', ' PM')
            
            # handle formats without a colon
            if ':' not in time_str and any(c.isdigit() for c in time_str):
                # extract the number part
                num_part = ''.join(c for c in time_str if c.isdigit())
                am_pm = 'AM' if 'am' in time_str.lower() else 'PM'
                time_str = f"{num_part}:00 {am_pm}"
            
            # parse with datetime
            time_obj = datetime.strptime(time_str, '%I:%M %p')
            return time_obj.strftime('%H:%M')
        else:
            # assume 24-hour format already
            return time_str
    except Exception as e:
        print(f"Error converting time '{time_str}': {e}")
        return None

def parse_time_range(time_range):
    """Parse a time range like '2:00 PM - 5:00 PM' into start and end times in 24hr format"""
    if not time_range or time_range.lower() == 'na':
        return None, None
    
    try:
        parts = time_range.split(' - ')
        if len(parts) != 2:
            return None, None
        
        start_time = convert_time_to_24hr(parts[0])
        end_time = convert_time_to_24hr(parts[1])
        
        return start_time, end_time
    except Exception as e:
        print(f"Error parsing time range '{time_range}': {e}")
        return None, None

def time_to_minutes(time_str):
    """Convert time string in 24hr format (HH:MM) to minutes since midnight"""
    if not time_str:
        return None
    
    try:
        hours, minutes = map(int, time_str.split(':'))
        return hours * 60 + minutes
    except Exception as e:
        print(f"Error converting time to minutes '{time_str}': {e}")
        return None

def minutes_to_time(minutes):
    """Convert minutes since midnight to time string in 24hr format (HH:MM)"""
    if minutes is None:
        return None
    
    hours = minutes // 60
    mins = minutes % 60
    return f"{hours:02d}:{mins:02d}"

def format_time_12hr(time_str):
    """Format time string from 24hr (HH:MM) to 12hr (H:MM AM/PM)"""
    if not time_str:
        return ""
    
    try:
        hours, minutes = map(int, time_str.split(':'))
        
        # convert to 12-hour format
        period = "AM" if hours < 12 else "PM"
        hours = hours % 12
        if hours == 0:
            hours = 12
            
        return f"{hours}:{minutes:02d} {period}"
    except Exception as e:
        print(f"Error formatting time '{time_str}': {e}")
        return time_str

def calculate_shift_duration(start_time, end_time):
    """Calculate duration between start and end times in minutes"""
    if not start_time or not end_time:
        return 0
    
    start_minutes = time_to_minutes(start_time)
    end_minutes = time_to_minutes(end_time)
    
    # handle overnight shifts
    if end_minutes < start_minutes:
        end_minutes += 24 * 60
        
    return end_minutes - start_minutes

def calculate_shift_duration_hours(start_time, end_time):
    """Calculate duration between start and end times in hours (decimal)"""
    minutes = calculate_shift_duration(start_time, end_time)
    return minutes / 60

def parse_day_code(day_code):
    """Convert day code (U,M,T,W,R,F,S) to full day name"""
    day_map = {
        'U': 'Sunday',
        'M': 'Monday',
        'T': 'Tuesday',
        'W': 'Wednesday',
        'R': 'Thursday',
        'F': 'Friday',
        'S': 'Saturday'
    }
    
    return day_map.get(day_code.upper(), None)

def parse_unavailable_time(unavailable_str):
    """Parse unavailable time string like 'MWF 1pm - 2pm'"""
    if not unavailable_str or unavailable_str.lower() == 'na':
        return {}
    
    result = {}
    try:
        # extract day codes and time range
        match = re.match(r'([UMTWRFS]+)\s+(.*)', unavailable_str)
        if not match:
            return result
        
        day_codes, time_range = match.groups()
        start_time, end_time = parse_time_range(time_range)
        
        if start_time and end_time:
            for day_code in day_codes:
                day_name = parse_day_code(day_code)
                if day_name:
                    if day_name not in result:
                        result[day_name] = []
                    result[day_name].append((start_time, end_time))
    except Exception as e:
        print(f"Error parsing unavailable time '{unavailable_str}': {e}")
    
    return result

# data handling functions
def load_json_data(file_path, default=None):
    """Load JSON data from file, return default if file doesn't exist"""
    if default is None:
        default = []
        
    if not os.path.exists(file_path):
        return default
    
    try:
        with open(file_path, 'r') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading JSON from {file_path}: {e}")
        return default

def save_json_data(file_path, data):
    """Save data to JSON file"""
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving JSON to {file_path}: {e}")
        return False

def get_workplace_path(base_dir, workplace_name):
    """Get the path to a workplace directory"""
    return os.path.join(base_dir, "data", workplace_name)

def get_workers_path(base_dir, workplace_name):
    """Get the path to a workplace's workers.json file"""
    workplace_path = get_workplace_path(base_dir, workplace_name)
    return os.path.join(workplace_path, "workers", "workers.json")

def get_config_path(base_dir, workplace_name):
    """Get the path to a workplace's config.json file"""
    workplace_path = get_workplace_path(base_dir, workplace_name)
    return os.path.join(workplace_path, "config.json")

def get_schedules_path(base_dir, workplace_name):
    """Get the path to a workplace's schedules directory"""
    workplace_path = get_workplace_path(base_dir, workplace_name)
    return os.path.join(workplace_path, "schedules")

def get_settings_path(base_dir):
    """Get the path to the user settings file"""
    return os.path.join(base_dir, "data", "settings.json")

def import_workers_from_excel(file_path):
    """Import workers from Excel file"""
    try:
        df = pd.read_excel(file_path)
        
        # validate required columns
        required_columns = [
            "First Name", "Last Name", 
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
            "Email", "Work Study"
        ]
        
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Required column '{col}' not found in Excel file")
        
        # process each row
        workers = []
        for _, row in df.iterrows():
            # skip rows with empty first name or last name
            if pd.isna(row["First Name"]) or pd.isna(row["Last Name"]):
                continue
                
            # process availability
            availability = {}
            for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
                time_range = row[day]
                if pd.isna(time_range) or time_range.lower() == 'na':
                    availability[day] = []
                else:
                    # split multiple time ranges if present
                    ranges = [r.strip() for r in str(time_range).split(',')]
                    day_ranges = []
                    for r in ranges:
                        start_time, end_time = parse_time_range(r)
                        if start_time and end_time:
                            day_ranges.append({
                                "start": start_time,
                                "end": end_time
                            })
                    availability[day] = day_ranges
            
            # process unavailable times
            unavailable = {}
            for col in df.columns:
                if "Day(s) & Time not Available" in col and not pd.isna(row[col]):
                    unavail_dict = parse_unavailable_time(str(row[col]))
                    for day, times in unavail_dict.items():
                        if day not in unavailable:
                            unavailable[day] = []
                        unavailable[day].extend(times)
            
            # create worker object
            worker = {
                "id": f"{row['First Name'].lower()}_{row['Last Name'].lower()}_{len(workers)}",
                "first_name": row["First Name"],
                "last_name": row["Last Name"],
                "email": row["Email"] if not pd.isna(row["Email"]) else "",
                "work_study": str(row["Work Study"]).lower() in ['y', 'yes', 'true', '1'],
                "availability": availability,
                "unavailable": unavailable,
                "weekly_hours": 0  # will be updated when scheduling
            }
            
            workers.append(worker)
        
        return workers
    except Exception as e:
        print(f"Error importing workers from Excel: {e}")
        raise

def export_workers_to_excel(workers, file_path):
    """Export workers to Excel file"""
    try:
        # create DataFrame
        data = []
        for worker in workers:
            row = {
                "First Name": worker["first_name"],
                "Last Name": worker["last_name"],
                "Email": worker["email"],
                "Work Study": "Y" if worker["work_study"] else "N"
            }
            
            # add availability
            for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
                if day in worker["availability"] and worker["availability"][day]:
                    time_ranges = []
                    for time_range in worker["availability"][day]:
                        start = format_time_12hr(time_range["start"])
                        end = format_time_12hr(time_range["end"])
                        time_ranges.append(f"{start} - {end}")
                    row[day] = ", ".join(time_ranges)
                else:
                    row[day] = "na"
            
            # add unavailable times
            unavail_strings = []
            for day, times in worker.get("unavailable", {}).items():
                day_code = day[0] if day != "Thursday" else "R"
                for start, end in times:
                    start_12hr = format_time_12hr(start)
                    end_12hr = format_time_12hr(end)
                    unavail_strings.append(f"{day_code} {start_12hr} - {end_12hr}")
            
            row["Day(s) & Time not Available"] = unavail_strings[0] if unavail_strings else "na"
            row["Day(s) & Time not Available.1"] = unavail_strings[1] if len(unavail_strings) > 1 else "na"
            
            data.append(row)
        
        # create DataFrame and save to Excel
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        
        return True
    except Exception as e:
        print(f"Error exporting workers to Excel: {e}")
        return False

def is_worker_available(worker, day, start_time, end_time):
    """Check if a worker is available for a specific shift"""
    # check if the day exists in availability
    if day not in worker["availability"]:
        return False
    
    # convert times to minutes for easier comparison
    shift_start_mins = time_to_minutes(start_time)
    shift_end_mins = time_to_minutes(end_time)
    
    # check if worker has any availability for this day
    if not worker["availability"][day]:
        return False
    
    # check if the shift falls within any of the worker's available time ranges
    for time_range in worker["availability"][day]:
        avail_start_mins = time_to_minutes(time_range["start"])
        avail_end_mins = time_to_minutes(time_range["end"])
        
        # if the shift is completely within the available range
        if avail_start_mins <= shift_start_mins and shift_end_mins <= avail_end_mins:
            # now check if the shift overlaps with any unavailable times
            if day in worker.get("unavailable", {}):
                for unavail_start, unavail_end in worker["unavailable"][day]:
                    unavail_start_mins = time_to_minutes(unavail_start)
                    unavail_end_mins = time_to_minutes(unavail_end)
                    
                    # check if there's any overlap with unavailable time
                    if not (shift_end_mins <= unavail_start_mins or shift_start_mins >= unavail_end_mins):
                        return False
            
            return True
    
    return False

def get_available_workers(workers, day, start_time, end_time, exclude_workers=None):
    """Get list of workers available for a specific shift"""
    if exclude_workers is None:
        exclude_workers = []
    
    available_workers = []
    for worker in workers:
        # skip excluded workers
        if worker["id"] in exclude_workers:
            continue
            
        if is_worker_available(worker, day, start_time, end_time):
            available_workers.append(worker)
    
    return available_workers

def backup_data(base_dir, backup_path):
    """Backup all data to a specified location"""
    try:
        # create backup directory if it doesn't exist
        os.makedirs(os.path.dirname(backup_path), exist_ok=True)
        
        # copy data directory to backup location
        data_dir = os.path.join(base_dir, "data")
        import shutil
        shutil.make_archive(backup_path, 'zip', data_dir)
        
        return True
    except Exception as e:
        print(f"Error creating backup: {e}")
        return False

def restore_data(base_dir, backup_path):
    """Restore data from a backup"""
    try:
        # extract backup to data directory
        data_dir = os.path.join(base_dir, "data")
        import shutil
        
        # remove existing data directory
        if os.path.exists(data_dir):
            shutil.rmtree(data_dir)
        
        # extract backup
        shutil.unpack_archive(backup_path, base_dir)
        
        return True
    except Exception as e:
        print(f"Error restoring from backup: {e}")
        return False
