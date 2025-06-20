import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkcalendar import Calendar
from roster_searcher import RosterSearcher
import uuid
import hashlib
from datetime import datetime, timedelta

class SchemaExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Roster Search")
        self.searcher = RosterSearcher()

        # File path entry
        self.file_label = tk.Label(root, text="Excel file path:")
        self.file_label.grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.file_entry = tk.Entry(root, width=60)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)
        self.open_browser_button = tk.Button(root, text="Open in Browser", command=self.open_in_browser, state='disabled')
        self.open_browser_button.grid(row=0, column=3, padx=5, pady=5)
        self.file_entry.bind('<KeyRelease>', self.check_sharepoint_link)

        # Name entry
        self.name_label = tk.Label(root, text="Name to search:")
        self.name_label.grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.name_entry = tk.Entry(root, width=50)
        self.name_entry.grid(row=1, column=1, padx=5, pady=5)
        self.name_entry.insert(0, "Naam")

        # Password entry
        self.pw_label = tk.Label(root, text="Excel password (optional):")
        self.pw_label.grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.pw_entry = tk.Entry(root, width=50, show='*')
        self.pw_entry.grid(row=2, column=1, padx=5, pady=5)

        # Search button
        self.search_button = tk.Button(root, text="Search", command=self.run_search)
        self.search_button.grid(row=3, column=1, pady=10)

        # Calendar area
        self.calendar_frame = tk.Frame(root)
        self.calendar_frame.grid(row=4, column=0, columnspan=4, padx=10, pady=5, sticky='nsew')
        self.calendar = Calendar(self.calendar_frame, selectmode='day', date_pattern='yyyy-mm-dd')
        self.calendar.pack(fill='both', expand=True)
        self.calendar.bind("<<CalendarSelected>>", self.on_calendar_date_selected)

        # Results area (text)
        self.results_text = scrolledtext.ScrolledText(root, width=80, height=8, state='disabled')
        self.results_text.grid(row=5, column=0, columnspan=3, padx=10, pady=5)

        # Save button
        self.save_button = tk.Button(root, text="Save Results", command=self.save_results, state='disabled')
        self.save_button.grid(row=6, column=1, pady=5)
        
        # Export to Calendar button
        self.export_cal_button = tk.Button(root, text="Export to Calendar", command=self.export_to_calendar, state='disabled')
        self.export_cal_button.grid(row=6, column=2, pady=5)

        self.results = []
        self.results_by_date = {}

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
        self.check_sharepoint_link()

    def check_sharepoint_link(self, event=None):
        url = self.file_entry.get().strip()
        if 'sharepoint.com' in url or 'onedrive.live.com' in url or '1drv.ms' in url:
            self.open_browser_button.config(state='normal')
        else:
            self.open_browser_button.config(state='disabled')

    def open_in_browser(self):
        import webbrowser
        url = self.file_entry.get().strip()
        if url:
            webbrowser.open(url)

    def run_search(self):
        file_path = self.file_entry.get().strip()
        name = self.name_entry.get().strip()
        password = self.pw_entry.get().strip() or None
        if not file_path or not name:
            messagebox.showerror("Error", "Please provide both file path and name to search.")
            return
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_text.insert(tk.END, f"Searching for '{name}' in {file_path}\n")
        self.results_text.config(state='disabled')
        self.calendar.calevent_remove('workday')
        self.results_by_date = {}
        try:
            results = self.searcher.search_person_schedule(file_path, name, password)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self.results = results
        self.display_results(results, name, file_path)

    def display_results(self, results, name, file_path):
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        self.results_by_date = {}
        self.calendar.calevent_remove('workday')
        if not results:
            self.results_text.insert(tk.END, "No matches found.\n")
        else:
            self.results_text.insert(tk.END, f"Found {len(results)} work assignments:\n" + "-"*80 + "\n")
            sheets = {}
            for result in results:
                sheet = result.get('sheet', 'Unknown')
                if sheet not in sheets:
                    sheets[sheet] = []
                sheets[sheet].append(result)
                date_str = result.get('date', None)
                if date_str:
                    try:
                        from datetime import datetime
                        dt = datetime.strptime(date_str[:10], "%Y-%m-%d") if '-' in date_str else datetime.strptime(date_str[:10], "%d-%m-%Y")
                        date_key = dt.strftime("%Y-%m-%d")
                        self.results_by_date.setdefault(date_key, []).append(result)
                        self.calendar.calevent_create(dt, 'Work', 'workday')
                    except Exception as e:
                        print(f"Could not parse date '{date_str}': {e}")
            for sheet_name, sheet_results in sheets.items():
                self.results_text.insert(tk.END, f"\nSheet: {sheet_name}\n" + "-"*40 + "\n")
                for result in sheet_results:
                    self.results_text.insert(tk.END, f"Date: {result.get('date', 'Unknown')}\n")
                    self.results_text.insert(tk.END, f"Context: {result.get('context', '')}\n")
                    self.results_text.insert(tk.END, f"Position: {result.get('position', '')}\n")
                    self.results_text.insert(tk.END, f"Table Type: {result.get('table_type', '')}\n\n")
            self.save_button.config(state='normal')
            self.export_cal_button.config(state='normal')
        self.results_text.config(state='disabled')

    def on_calendar_date_selected(self, event):
        selected_date = self.calendar.get_date()
        details = self.results_by_date.get(selected_date, [])
        if details:
            msg = f"Work assignments for {selected_date}:\n\n"
            for result in details:
                msg += f"Sheet: {result.get('sheet', 'Unknown')}\n"
                msg += f"Context: {result.get('context', '')}\n"
                msg += f"Position: {result.get('position', '')}\n"
                msg += f"Table Type: {result.get('table_type', '')}\n"
                msg += "-"*30 + "\n"
            messagebox.showinfo("Workday Details", msg)
        else:
            messagebox.showinfo("Workday Details", f"No assignments found for {selected_date}.")

    def save_results(self):
        if not self.results:
            return
        from tkinter import filedialog
        output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not output_file:
            return
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                for result in self.results:
                    f.write(f"Sheet: {result.get('sheet', 'Unknown')}\n")
                    f.write(f"Date: {result.get('date', 'Unknown')}\n")
                    f.write(f"Context: {result.get('context', '')}\n")
                    f.write(f"Position: {result.get('position', '')}\n")
                    f.write(f"Table Type: {result.get('table_type', '')}\n")
                    f.write("-"*30 + "\n")
            messagebox.showinfo("Saved", f"Results saved to: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {str(e)}")
            
    def export_to_calendar(self):
        """Export work schedule to iCalendar format for Google Calendar"""
        if not self.results:
            return
            
        output_file = filedialog.asksaveasfilename(
            defaultextension=".ics",
            initialfile="work_schedule.ics",
            filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")]
        )
        
        if not output_file:
            return
            
        try:
            calendar_content = []
            calendar_content.append("BEGIN:VCALENDAR")
            calendar_content.append("VERSION:2.0")
            calendar_content.append("PRODID:-//Excel Roster Search//Calendar Export//EN")
            calendar_content.append("CALSCALE:GREGORIAN")
            calendar_content.append("METHOD:PUBLISH")
            
            # Group results by date to avoid duplicates
            dates_processed = set()
            
            for result in self.results:
                date_str = result.get('date')
                if not date_str:
                    continue
                    
                # Convert to datetime object (handle both YYYY-MM-DD and DD-MM-YYYY formats)
                try:
                    if '-' in date_str:
                        if date_str[2] == '-' or date_str[1] == '-':  # DD-MM-YYYY format
                            dt = datetime.strptime(date_str[:10], "%d-%m-%Y")
                        else:  # YYYY-MM-DD format
                            dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
                    else:
                        continue  # Skip if no valid date format
                except Exception:
                    continue
                    
                # Check if we've already processed this date to avoid duplicates
                date_key = dt.strftime("%Y-%m-%d")
                name = result.get('name', '')
                unique_key = f"{date_key}_{name}"
                
                if unique_key in dates_processed:
                    continue  # Skip duplicate dates for the same person
                    
                dates_processed.add(unique_key)
                
                # Create event from 8:30 AM to 5:00 PM
                start_time = dt.replace(hour=8, minute=30, second=0)
                end_time = dt.replace(hour=17, minute=0, second=0)
                
                # Format dates for iCalendar
                start_str = start_time.strftime("%Y%m%dT%H%M%S")
                end_str = end_time.strftime("%Y%m%dT%H%M%S")
                
                # Create a unique identifier based on date and name
                uid_base = f"{date_key}_{name}_{result.get('context', '')}"
                uid = hashlib.md5(uid_base.encode()).hexdigest()
                
                # Add event to calendar
                calendar_content.append("BEGIN:VEVENT")
                calendar_content.append(f"UID:{uid}@excelrostersearch")
                calendar_content.append(f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%S')}")
                calendar_content.append(f"DTSTART:{start_str}")
                calendar_content.append(f"DTEND:{end_str}")
                calendar_content.append(f"SUMMARY:Work: {name}")
                calendar_content.append(f"DESCRIPTION:Sheet: {result.get('sheet', 'Unknown')}\nContext: {result.get('context', '')}\nPosition: {result.get('position', '')}")
                calendar_content.append("END:VEVENT")
                
            calendar_content.append("END:VCALENDAR")
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(calendar_content))
                
            messagebox.showinfo("Calendar Export", 
                f"Work schedule exported to {output_file}\n\n"
                f"To import into Google Calendar:\n"
                f"1. Open Google Calendar\n"
                f"2. Click the '+' next to 'Other calendars'\n"
                f"3. Select 'Import'\n"
                f"4. Upload the .ics file\n\n"
                f"Each shift is set from 8:30 AM to 5:00 PM\n"
                f"Duplicate imports will be prevented.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Could not export calendar: {str(e)}")

