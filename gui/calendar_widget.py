"""
Calendar component for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime
from .theme import COLORS, FONTS

class CalendarWidget(ttk.Frame):
    """Component for displaying and interacting with the calendar"""
    
    def __init__(self, parent, on_date_selected=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(style='TFrame', padding=(10, 5, 10, 10))
        self.on_date_selected = on_date_selected
        
        # Calendar header
        self.header_label = ttk.Label(
            self, 
            text="Calendar View", 
            style='Subheader.TLabel'
        )
        self.header_label.pack(anchor='w', pady=(0, 10))

        # Calendar container with border and elevation effect
        self.calendar_container = ttk.Frame(
            self,
            style='TFrame',
            padding=2
        )
        self.calendar_container.pack(fill='both', expand=True)
        
        # Add a subtle border using a canvas behind the calendar
        self.canvas = tk.Canvas(
            self.calendar_container,
            highlightthickness=1,
            highlightbackground=COLORS['secondary'],
            background=COLORS['background']
        )
        self.canvas.pack(fill='both', expand=True)
        
        # Calendar widget
        self.calendar = Calendar(
            self.canvas,
            selectmode='day',
            date_pattern='yyyy-mm-dd',
            background=COLORS['background'],
            foreground=COLORS['text'],
            headersbackground=COLORS['secondary'],
            headersforeground=COLORS['text_light'],
            normalbackground=COLORS['background'],
            normalforeground=COLORS['text'],
            weekendbackground="#e8f4f8",  # Light blue for weekends
            weekendforeground=COLORS['text'],
            selectbackground=COLORS['accent'],
            selectforeground=COLORS['text_light'],
            othermonthbackground="#f8f9fa",  # Very light gray for other month days
            othermonthwebackground="#f8f9fa",  # Same for other month weekends
            font=FONTS['body']
        )
        self.calendar.pack(fill='both', expand=True, padx=2, pady=2)
        
        # Bind selection event
        self.calendar.bind("<<CalendarSelected>>", self._on_date_selected)
        
        # Date info display
        self.date_info_frame = ttk.Frame(self, style='TFrame')
        self.date_info_frame.pack(fill='x', pady=(10, 0))
        
        self.date_label = ttk.Label(
            self.date_info_frame,
            text="No date selected",
            font=FONTS['body']
        )
        self.date_label.pack(side='left')
        
        # Initialize with current date
        today = datetime.now().strftime('%Y-%m-%d')
        self.calendar.selection_set(today)
        self._update_date_info(today)
    
    def _on_date_selected(self, event):
        """Handle date selection in the calendar"""
        selected_date = self.calendar.get_date()
        self._update_date_info(selected_date)
        
        if self.on_date_selected:
            self.on_date_selected(selected_date)
    
    def _update_date_info(self, date_str):
        """Update the date information display"""
        try:
            selected_date = datetime.strptime(date_str, '%Y-%m-%d')
            formatted_date = selected_date.strftime('%A, %d %B %Y')
            self.date_label.config(text=f"Selected: {formatted_date}")
        except ValueError:
            self.date_label.config(text="Invalid date format")
    
    def get_selected_date(self):
        """Return the currently selected date"""
        return self.calendar.get_date()
    
    def set_date(self, date_str):
        """Set the calendar to a specific date"""
        try:
            self.calendar.selection_set(date_str)
            self._update_date_info(date_str)
        except ValueError:
            pass  # Invalid date format, ignore
    
    def highlight_dates(self, dates, color=None):
        """Highlight specific dates in the calendar
        
        Args:
            dates (list): List of dates in 'yyyy-mm-dd' format
            color (str, optional): Color for highlighting. Defaults to accent color.
        """
        if color is None:
            color = COLORS['success']
            
        for date_str in dates:
            try:
                # tkcalendar format for marking dates needs to be year-month-day
                date = datetime.strptime(date_str, '%Y-%m-%d')
                self.calendar.calevent_create(
                    date, 
                    "Scheduled", 
                    "schedule"
                )
                self.calendar.tag_config(
                    "schedule", 
                    background=color, 
                    foreground=COLORS['text_light']
                )
            except ValueError:
                continue  # Skip invalid dates
