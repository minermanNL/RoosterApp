"""
Main application class for the Excel Roster Search GUI
"""
import tkinter as tk
from tkinter import ttk, messagebox

from .theme import setup_theme, COLORS
from .file_selector import FileSelector
from .search_panel import SearchPanel
from .calendar_widget import CalendarWidget
from .export_utils import ExportUtils

from roster_searcher import RosterSearcher

class RosterSearchApp:
    """Main application class for the Excel Roster Search GUI"""
    
    def __init__(self, root):
        """Initialize the application
        
        Args:
            root (tk.Tk): Root window for the application
        """
        self.root = root
        self.root.title("Excel Roster Search")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Set app icon if available
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass  # No icon available, continue without it
            
        # Initialize roster searcher
        self.searcher = RosterSearcher()
        
        # Apply theme
        setup_theme(self.root)
        
        # Create main container
        self.main_container = ttk.Frame(self.root, style='TFrame')
        self.main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Application header
        self.create_header()
        
        # Create components
        self.create_components()
        
        # Layout components
        self.layout_components()
        
        # Initialize data holders
        self.results = []
        self.results_by_date = {}
        
    def create_header(self):
        """Create application header"""
        header_frame = ttk.Frame(self.main_container, style='TFrame')
        header_frame.pack(fill='x', pady=(0, 10))
        
        app_title = ttk.Label(
            header_frame,
            text="Excel Roster Search",
            font=('Helvetica', 18, 'bold'),
            foreground=COLORS['primary']
        )
        app_title.pack(side='left')
        
        # Version info
        version_label = ttk.Label(
            header_frame,
            text="v2.0",
            font=('Helvetica', 10),
            foreground=COLORS['secondary']
        )
        version_label.pack(side='right')
        
    def create_components(self):
        """Create minimal application components"""
        # File selector
        self.file_selector = FileSelector(self.main_container)

        # Search panel
        self.search_panel = SearchPanel(
            self.main_container,
            on_search_callback=self.run_search
        )

        # Calendar widget (directly under main container)
        self.calendar_widget = CalendarWidget(
            self.main_container,
            on_date_selected=self.on_calendar_date_selected
        )
        
    def layout_components(self):
        """Layout components for minimal UI"""
        self.file_selector.pack(fill='x', pady=(0, 5))
        self.search_panel.pack(fill='x', pady=5)
        self.calendar_widget.pack(fill='both', expand=True, pady=5)
        
    def run_search(self, name=None):
        """Run the roster search
        
        Args:
            name (str, optional): Name to search for. If None, gets from search panel
        """
        file_path = self.file_selector.get_file_path()
        password = self.file_selector.get_password()
        
        if not name:
            name = self.search_panel.get_search_name()
            
        if not file_path or not name:
            messagebox.showerror(
                "Search Error", 
                "Please provide both file path and name to search."
            )
            self.search_panel.reset_status()
            return
            
        try:
            results = self.searcher.search_person_schedule(file_path, name, password)
            
            # Store results for later reference
            self.results = results
            
            # Group results by date for calendar highlighting
            self.results_by_date = {}
            highlight_dates = []
            for result in results:
                date_str = result.get('date')
                if date_str:
                    try:
                        # Standardize date format to YYYY-MM-DD for highlighting
                        if '-' in date_str:
                            if date_str[2] == '-' or date_str[1] == '-':  # DD-MM-YYYY format
                                from datetime import datetime
                                dt = datetime.strptime(date_str[:10], "%d-%m-%Y")
                                standard_date = dt.strftime("%Y-%m-%d")
                            else:  # Already YYYY-MM-DD format
                                standard_date = date_str[:10]
                                
                            if standard_date not in self.results_by_date:
                                self.results_by_date[standard_date] = []
                            self.results_by_date[standard_date].append(result)
                            highlight_dates.append(standard_date)
                    except Exception:
                        continue
            
            # Highlight dates with results in calendar
            self.calendar_widget.highlight_dates(highlight_dates)
            
            # Update status
            if results:
                self.search_panel.set_status(
                    f"Found {len(results)} results for '{name}'"
                )
            else:
                self.search_panel.set_status(f"No results found for '{name}'")
                
        except Exception as e:
            messagebox.showerror("Search Error", str(e))
            self.search_panel.reset_status()
    
    def on_calendar_date_selected(self, date_str):
        """Handle calendar date selection
        
        Args:
            date_str (str): Selected date in YYYY-MM-DD format
        """
        # If we have results for this date, filter and show only those
        if date_str in self.results_by_date:
            date_results = self.results_by_date[date_str]
            name = date_results[0].get('name', '')
            info_lines = [f"Results for {name} on {date_str}:"]
            for res in date_results:
                sheet = res.get('sheet','Unknown')
                context = res.get('context','')
                position = res.get('position','')
                info_lines.append(f"Sheet: {sheet}\nContext: {context}\nPosition: {position}\n---")
            messagebox.showinfo("Schedule", "\n".join(info_lines))
            self.search_panel.set_status(f"{len(date_results)} result(s) for {date_str}")
        elif self.results:
            self.search_panel.set_status(f"No results for {date_str}")
    
    def save_results(self, results):
        """Save search results to a file
        
        Args:
            results (list): List of result dictionaries
        """
        ExportUtils.save_results_to_file(results, self.root)
    
    def export_to_calendar(self, results):
        """Export search results to iCalendar format
        
        Args:
            results (list): List of result dictionaries
        """
        ExportUtils.export_to_ical(results, self.root)
