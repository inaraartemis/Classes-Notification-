import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import threading
import json
import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ------------------- Colors & Style -------------------
BG_COLOR = "#d9c7b8"        # Light coffee background
BUTTON_COLOR = "#8b5e3c"    # Dark coffee buttons
BUTTON_FG = "white"
LABEL_FG = "#3e2f2f"
FONT_TITLE = ("Helvetica", 18, "bold")
FONT_LABEL = ("Helvetica", 12)
FONT_BUTTON = ("Helvetica", 12, "bold")

# ------------------- Google Calendar Setup -------------------
SCOPES = ["https://www.googleapis.com/auth/calendar"]

def get_service():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    service = build("calendar", "v3", credentials=creds)
    return service

def delete_timetable_events(service):
    now = datetime.utcnow().isoformat() + "Z"
    events_result = service.events().list(
        calendarId="primary", timeMin=now, maxResults=2500, singleEvents=True
    ).execute()
    events = events_result.get("items", [])
    count = 0
    for event in events:
        if event.get("description") == "source: timetable-script":
            service.events().delete(calendarId="primary", eventId=event["id"]).execute()
            count += 1
    return count

# ------------------- Time Parsing -------------------
def parse_time_slot(current_date, slot_str):
    try:
        slot_str = slot_str.strip().upper()
        if "-" not in slot_str:
            raise ValueError("Invalid time slot format")
        start_str, end_str = slot_str.split("-")
        period = "AM" if "AM" in slot_str else "PM"
        start_hour = int("".join(c for c in start_str if c.isdigit()))
        end_hour = int("".join(c for c in end_str if c.isdigit()))

        if start_hour == 12:
            start_hour_24 = 0 if period == "AM" else 12
        else:
            start_hour_24 = start_hour if period == "AM" else start_hour + 12

        if end_hour == 12:
            end_hour_24 = 0 if period == "AM" else 12
        else:
            end_hour_24 = end_hour if period == "AM" else end_hour + 12

        start_dt = datetime.combine(current_date, datetime.min.time()) + timedelta(hours=start_hour_24)
        end_dt = datetime.combine(current_date, datetime.min.time()) + timedelta(hours=end_hour_24)
        if end_dt <= start_dt:
            end_dt = start_dt + timedelta(hours=1)
        return start_dt, end_dt
    except Exception as e:
        print(f"Error parsing time slot '{slot_str}': {e}")
        start_dt = datetime.combine(current_date, datetime.strptime("09:00 AM", "%I:%M %p").time())
        end_dt = start_dt + timedelta(hours=1)
        return start_dt, end_dt

def add_events_to_calendar(service, timetable, start_date, end_date, progress_callback):
    days_map = {"Monday":0, "Tuesday":1, "Wednesday":2, "Thursday":3, "Friday":4, "Saturday":5, "Sunday":6}
    current_date = datetime.combine(start_date, datetime.min.time())
    end_datetime = datetime.combine(end_date, datetime.min.time())
    events_to_add = []

    while current_date <= end_datetime:
        weekday_num = current_date.weekday()
        for cls in timetable:
            day_name = cls["day"].strip()
            cls_day_num = days_map.get(day_name)
            if cls_day_num is None or cls_day_num != weekday_num:
                continue
            start_dt, end_dt = parse_time_slot(current_date, cls["time_slot"])
            code = cls.get("code") or cls.get("subject","")  # course code first
            summary = f"{cls.get('time_slot','')} - {code}"  # time + course code
            events_to_add.append({
                "summary": summary,
                "location": cls.get("room", ""),
                "description": "source: timetable-script",
                "start": {"dateTime": start_dt.isoformat(), "timeZone":"Asia/Kolkata"},
                "end": {"dateTime": end_dt.isoformat(), "timeZone":"Asia/Kolkata"}
            })
        current_date += timedelta(days=1)

    total = len(events_to_add)
    for i, event in enumerate(events_to_add, 1):
        service.events().insert(calendarId="primary", body=event).execute()
        progress_callback(int(i/total*100), f"Adding {i}/{total} events")
    return total

# ------------------- GUI -------------------
class TimetableGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mi Clase Programadora")
        self.root.geometry("400x650")
        self.root.configure(bg=BG_COLOR)
        self.root.resizable(False, False)

        # Header
        tk.Label(root, text="☕ Mi Clase Programadora", bg=BG_COLOR, fg=LABEL_FG, font=FONT_TITLE).pack(pady=15)

        # Date selection
        date_frame = tk.Frame(root, bg=BG_COLOR)
        date_frame.pack(pady=10)
        tk.Label(date_frame, text="Start Date:", bg=BG_COLOR, fg=LABEL_FG, font=FONT_LABEL).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.start_date = DateEntry(date_frame, width=15, background='brown', foreground='white', borderwidth=2)
        self.start_date.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(date_frame, text="End Date:", bg=BG_COLOR, fg=LABEL_FG, font=FONT_LABEL).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.end_date = DateEntry(date_frame, width=15, background='brown', foreground='white', borderwidth=2)
        self.end_date.grid(row=1, column=1, padx=5, pady=5)

        # Buttons
        button_frame = tk.Frame(root, bg=BG_COLOR)
        button_frame.pack(pady=20)
        self.delete_btn = tk.Button(button_frame, text="🗑 Delete Old Events", bg=BUTTON_COLOR, fg=BUTTON_FG, font=FONT_BUTTON, width=25, command=self.delete_old)
        self.delete_btn.pack(pady=10)
        self.add_btn = tk.Button(button_frame, text="➕ Add Timetable Events", bg=BUTTON_COLOR, fg=BUTTON_FG, font=FONT_BUTTON, width=25, command=self.add_new)
        self.add_btn.pack(pady=10)

        # Progress
        tk.Label(root, text="Progress:", bg=BG_COLOR, fg=LABEL_FG, font=FONT_LABEL).pack(pady=(20,5))
        self.progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=5)
        self.progress_label = tk.Label(root, text="", bg=BG_COLOR, fg=LABEL_FG, font=FONT_LABEL)
        self.progress_label.pack(pady=(5,10))

        # Preview
        tk.Label(root, text="Upcoming Classes Preview:", bg=BG_COLOR, fg=LABEL_FG, font=FONT_LABEL).pack(pady=(10,5))
        self.preview_text = tk.Text(root, height=25, width=45, bg="white", fg="black")
        self.preview_text.pack(pady=5)
        self.preview_text.insert(tk.END, "Load timetable preview here...")
        self.preview_text.config(state=tk.DISABLED)

        # Load timetable JSON and flatten
        with open("timetable.json", "r") as f:
            data = json.load(f)

        self.timetable = []
        for day, slots in data["timetable"].items():
            for time_slot, cls_str in slots.items():
                parts = cls_str.split(" / ")
                cls_dict = {"day": day, "time_slot": time_slot}
                for part in parts:
                    if part.startswith("C:"):
                        cls_dict["code"] = part[2:]
                    elif part.startswith("R:"):
                        cls_dict["room"] = part[2:]
                    elif part.startswith("S:"):
                        cls_dict["subject"] = part[2:]
                self.timetable.append(cls_dict)

        self.load_preview()

    def update_progress(self, value, label_text=""):
        self.progress["value"] = value
        self.progress_label.config(text=label_text)
        self.root.update_idletasks()

    def load_preview(self):
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        days_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        for day in days_order:
            day_classes = [cls for cls in self.timetable if cls["day"] == day]
            if not day_classes:
                continue
            self.preview_text.insert(tk.END, f"{day}:\n")
            for cls in day_classes:
                code = cls.get("code") or cls.get("subject","")  # course code first
                line = f"  {cls.get('time_slot','')} - {code}\n"
                self.preview_text.insert(tk.END, line)
            self.preview_text.insert(tk.END, "\n")
        self.preview_text.config(state=tk.DISABLED)

    def delete_old(self):
        def run_delete():
            try:
                service = get_service()
                self.update_progress(0, "Deleting old events...")
                count = delete_timetable_events(service)
                self.update_progress(100, f"Deleted {count} old events")
                messagebox.showinfo("Done", f"Deleted {count} old timetable events.")
                self.update_progress(0)
            except Exception as e:
                messagebox.showerror("Error", str(e))
        threading.Thread(target=run_delete).start()

    def add_new(self):
        def run_add():
            try:
                service = get_service()
                self.update_progress(0, "Adding timetable events...")
                start_date = self.start_date.get_date()
                end_date = self.end_date.get_date()
                total = add_events_to_calendar(service, self.timetable, start_date, end_date, self.update_progress)
                self.update_progress(100, f"Added {total} events")
                messagebox.showinfo("Done", f"Added {total} timetable events.")
                self.update_progress(0)
            except Exception as e:
                messagebox.showerror("Error", str(e))
        threading.Thread(target=run_add).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = TimetableGUI(root)
    root.mainloop()
