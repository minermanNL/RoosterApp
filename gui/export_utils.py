"""
Export utilities for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import filedialog
import hashlib
from datetime import datetime
from .theme import show_info, show_error

class ExportUtils:
    """Utility class for exporting search results"""
    
    @staticmethod
    def save_results_to_file(results, parent_window=None):
        """Save search results to a text file
        
        Args:
            results (list): List of result dictionaries
            parent_window (tk.Widget, optional): Parent window for the file dialog
        
        Returns:
            bool: True if save was successful, False otherwise
        """
        if not results:
            show_error("Save Error", "No results to save.")
            return False
            
        output_file = filedialog.asksaveasfilename(
            parent=parent_window,
            title="Save Search Results",
            defaultextension=".txt",
            initialfile="search_results.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not output_file:
            return False
            
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                # Group results by date
                results_by_date = {}
                for result in results:
                    date = result.get('date', 'Unknown date')
                    if date not in results_by_date:
                        results_by_date[date] = []
                    results_by_date[date].append(result)
                
                # Write header
                f.write("EXCEL ROSTER SEARCH RESULTS\n")
                f.write("=" * 30 + "\n\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                # Write results grouped by date
                for date, date_results in sorted(results_by_date.items()):
                    f.write(f"DATE: {date}\n")
                    f.write("-" * 20 + "\n")
                    
                    for idx, result in enumerate(date_results):
                        sheet = result.get('sheet', 'Unknown sheet')
                        context = result.get('context', 'No context')
                        position = result.get('position', 'No position')
                        name = result.get('name', 'Unknown name')
                        
                        f.write(f"  Name: {name}\n")
                        f.write(f"  Sheet: {sheet}\n")
                        f.write(f"  Context: {context}\n")
                        if position:
                            f.write(f"  Position: {position}\n")
                        
                        if idx < len(date_results) - 1:
                            f.write("\n")
                    
                    f.write("\n\n")
                    
            show_info("Save Successful", f"Results saved to {output_file}")
            return True
            
        except Exception as e:
            show_error("Save Error", f"Could not save results: {str(e)}")
            return False
    
    @staticmethod
    def export_to_ical(results, parent_window=None):
        """Export search results to iCalendar format
        
        Args:
            results (list): List of result dictionaries
            parent_window (tk.Widget, optional): Parent window for the file dialog
        
        Returns:
            bool: True if export was successful, False otherwise
        """
        if not results:
            show_error("Export Error", "No results to export.")
            return False
            
        output_file = filedialog.asksaveasfilename(
            parent=parent_window,
            title="Export to iCalendar",
            defaultextension=".ics",
            initialfile="work_schedule.ics",
            filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")]
        )
        
        if not output_file:
            return False
            
        try:
            calendar_content = []
            calendar_content.append("BEGIN:VCALENDAR")
            calendar_content.append("VERSION:2.0")
            calendar_content.append("PRODID:-//Excel Roster Search//Calendar Export//EN")
            calendar_content.append("CALSCALE:GREGORIAN")
            calendar_content.append("METHOD:PUBLISH")
            
            # Group results by date to avoid duplicates
            dates_processed = set()
            
            for result in results:
                date_str = result.get('date')
                if not date_str:
                    continue
                    
                # Convert to datetime object (handle both YYYY-MM-DD and DD-MM-YYYY formats)
                try:
                    if '-' in date_str:
                        if len(date_str) >= 10 and (date_str[2] == '-' or date_str[1] == '-'):  # DD-MM-YYYY format
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
                calendar_content.append(f"DESCRIPTION:Sheet: {result.get('sheet', 'Unknown')}"
                                      f"\\nContext: {result.get('context', '')}"
                                      f"\\nPosition: {result.get('position', '')}")
                calendar_content.append("END:VEVENT")
                
            calendar_content.append("END:VCALENDAR")
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(calendar_content))
                
            show_info("Calendar Export", 
                f"Work schedule exported to {output_file}\n\n"
                f"To import into Google Calendar:\n"
                f"1. Open Google Calendar\n"
                f"2. Click the '+' next to 'Other calendars'\n"
                f"3. Select 'Import'\n"
                f"4. Upload the .ics file\n\n"
                f"Each shift is set from 8:30 AM to 5:00 PM\n"
                f"Duplicate imports will be prevented.")
            return True
                
        except Exception as e:
            show_error("Export Error", f"Could not export calendar: {str(e)}")
            return False
