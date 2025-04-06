import os
import json
import pandas as pd
import random
from datetime import datetime, timedelta
from utils import (
    load_json_data, save_json_data, 
    parse_time_range, time_to_minutes, minutes_to_time,
    calculate_shift_duration, calculate_shift_duration_hours,
    is_worker_available, get_available_workers
)

class ScheduleGenerator:
    def __init__(self, base_dir, workplace_name):
        self.base_dir = base_dir
        self.workplace_name = workplace_name
        self.workplace_path = os.path.join(base_dir, "data", workplace_name)
        self.config_path = os.path.join(self.workplace_path, "config.json")
        self.workers_path = os.path.join(self.workplace_path, "workers", "workers.json")
        self.schedules_path = os.path.join(self.workplace_path, "schedules")
        
        # load configuration
        self.config = load_json_data(self.config_path, {})
        self.workers = load_json_data(self.workers_path, [])
        
        # reset weekly hours for all workers
        for worker in self.workers:
            worker["weekly_hours"] = 0
    
    def generate_schedule(self, selected_worker_ids=None):
        """Generate a schedule based on worker availability and shift requirements"""
        if selected_worker_ids is None:
            selected_worker_ids = [w["id"] for w in self.workers]
        
        # filter workers by selected IDs
        selected_workers = [w for w in self.workers if w["id"] in selected_worker_ids]
        
        # initialize schedule
        schedule = {
            "workplace": self.workplace_name,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "days": {}
        }
        
        # get shift times from config
        shift_times = self.config.get("shift_times", {})
        
        # process each day
        for day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]:
            if day not in shift_times:
                continue
                
            day_shifts = []
            assigned_workers = set()  # track workers already assigned for this day
            
            # process each shift for this day
            for shift_time_str in shift_times[day]:
                start_time, end_time = parse_time_range(shift_time_str)
                if not start_time or not end_time:
                    continue
                
                # calculate shift duration
                duration_hours = calculate_shift_duration_hours(start_time, end_time)
                
                # find available workers for this shift
                available_workers = [
                    w for w in selected_workers 
                    if w["id"] not in assigned_workers and is_worker_available(w, day, start_time, end_time)
                ]
                
                # sorts workers by priority
                # Work study workers with less than 5 hours get priority
                # then non-work study workers
                work_study_workers = [w for w in available_workers if w["work_study"] and w["weekly_hours"] < 5]
                other_workers = [w for w in available_workers if not w["work_study"] or w["weekly_hours"] >= 5]
                
                # sort work study workers by current weekly hours (ascending)
                work_study_workers.sort(key=lambda w: w["weekly_hours"])
                
                # sort other workers by current weekly hours (ascending)
                other_workers.sort(key=lambda w: w["weekly_hours"])
                
                # combine sorted lists
                sorted_workers = work_study_workers + other_workers
                
                # assign worker to shift if available
                if sorted_workers:
                    assigned_worker = sorted_workers[0]
                    
                    # add shift to schedule
                    shift = {
                        "start_time": start_time,
                        "end_time": end_time,
                        "worker_id": assigned_worker["id"],
                        "worker_name": f"{assigned_worker['first_name']} {assigned_worker['last_name']}",
                        "duration_hours": duration_hours
                    }
                    
                    day_shifts.append(shift)
                    
                    # update worker's weekly hours
                    assigned_worker["weekly_hours"] += duration_hours
                    
                    # add worker to assigned set for this day
                    assigned_workers.add(assigned_worker["id"])
                else:
                    # no available worker for this shift
                    shift = {
                        "start_time": start_time,
                        "end_time": end_time,
                        "worker_id": None,
                        "worker_name": "UNASSIGNED",
                        "duration_hours": duration_hours
                    }
                    day_shifts.append(shift)
            
            # add day's shifts to schedule
            if day_shifts:
                schedule["days"][day] = day_shifts
        
        # add worker summary to schedule
        schedule["workers"] = []
        for worker in selected_workers:
            schedule["workers"].append({
                "id": worker["id"],
                "name": f"{worker['first_name']} {worker['last_name']}",
                "email": worker["email"],
                "work_study": worker["work_study"],
                "weekly_hours": worker["weekly_hours"]
            })
        
        return schedule
    
    def save_schedule(self, schedule, filename=None):
        """Save a generated schedule"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"schedule_{timestamp}.json"
        
        # ensure schedules directory exists
        os.makedirs(self.schedules_path, exist_ok=True)
        
        # save schedule as JSON
        schedule_path = os.path.join(self.schedules_path, filename)
        save_json_data(schedule_path, schedule)
        
        return schedule_path
    
    def export_schedule_to_excel(self, schedule, file_path):
        """Export schedule to Excel format"""
        try:
            # create a DataFrame for the schedule
            days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            hours = []
            
            # find all unique hours across all days
            for day in days:
                if day in schedule["days"]:
                    for shift in schedule["days"][day]:
                        start_time = shift["start_time"]
                        end_time = shift["end_time"]
                        if start_time and end_time and start_time not in hours:
                            hours.append(start_time)
                        if end_time and end_time not in hours:
                            hours.append(end_time)
            
            # sort hours
            hours = sorted(hours, key=lambda x: time_to_minutes(x))
            
            # create time slots (30-minute increments)
            time_slots = []
            for i in range(len(hours) - 1):
                start_mins = time_to_minutes(hours[i])
                end_mins = time_to_minutes(hours[i+1])
                
                # add 30-minute slots
                current_mins = start_mins
                while current_mins < end_mins:
                    time_slots.append(minutes_to_time(current_mins))
                    current_mins += 30
            
            # add the last hour
            time_slots.append(hours[-1])
            
            # remove duplicates and sort
            time_slots = sorted(set(time_slots), key=lambda x: time_to_minutes(x))
            
            # create DataFrame
            df = pd.DataFrame(index=time_slots, columns=days)
            
            # fill in schedule data
            for day in days:
                if day in schedule["days"]:
                    for shift in schedule["days"][day]:
                        start_time = shift["start_time"]
                        end_time = shift["end_time"]
                        worker_name = shift["worker_name"]
                        
                        # find all time slots that fall within this shift
                        for slot in time_slots:
                            slot_mins = time_to_minutes(slot)
                            start_mins = time_to_minutes(start_time)
                            end_mins = time_to_minutes(end_time)
                            
                            if start_mins <= slot_mins < end_mins:
                                df.at[slot, day] = worker_name
            
            # format the index to be more readable
            df.index = [self._format_time_for_display(t) for t in df.index]
            
            # add worker summary
            worker_summary = pd.DataFrame([
                {
                    "Name": w["name"],
                    "Email": w["email"],
                    "Work Study": "Yes" if w["work_study"] else "No",
                    "Weekly Hours": f"{w['weekly_hours']:.2f}"
                }
                for w in schedule["workers"]
            ])
            
            # save to Excel with multiple sheets
            with pd.ExcelWriter(file_path) as writer:
                df.to_excel(writer, sheet_name="Schedule")
                worker_summary.to_excel(writer, sheet_name="Worker Summary", index=False)
            
            return True
        except Exception as e:
            print(f"Error exporting schedule to Excel: {e}")
            return False
    
    def _format_time_for_display(self, time_str):
        """Format time string for display in Excel"""
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
        except:
            return time_str
    
    def find_replacement_workers(self, day, start_time, end_time):
        """Find workers available for a specific shift (for last-minute replacements)"""
        # convert day name to proper format if needed
        day_map = {
            "sunday": "Sunday",
            "monday": "Monday",
            "tuesday": "Tuesday",
            "wednesday": "Wednesday",
            "thursday": "Thursday",
            "friday": "Friday",
            "saturday": "Saturday"
        }
        day = day_map.get(day.lower(), day)
        
        # find available workers
        available_workers = get_available_workers(self.workers, day, start_time, end_time)
        
        # sort by work study status and weekly hours
        work_study_workers = [w for w in available_workers if w["work_study"] and w["weekly_hours"] < 5]
        other_workers = [w for w in available_workers if not w["work_study"] or w["weekly_hours"] >= 5]
        
        # sort work study workers by current weekly hours (ascending)
        work_study_workers.sort(key=lambda w: w["weekly_hours"])
        
        # sort other workers by current weekly hours (ascending)
        other_workers.sort(key=lambda w: w["weekly_hours"])
        
        # combine sorted lists
        sorted_workers = work_study_workers + other_workers
        
        return sorted_workers
      
