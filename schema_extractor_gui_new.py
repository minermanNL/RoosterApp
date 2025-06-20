"""
Main entry point for the Excel Roster Search GUI
This file preserves backward compatibility with existing imports
while delegating to the new modular implementation
"""
import tkinter as tk
from gui.app import RosterSearchApp

class SchemaExtractorGUI:
    """
    Legacy wrapper class that maintains backward compatibility
    while using the new modular GUI implementation
    """
    
    def __init__(self, root):
        """Initialize the GUI with the new implementation
        
        Args:
            root (tk.Tk): Root window for the application
        """
        # Create the new app
        self.app = RosterSearchApp(root)
        
        # Copy references to methods to maintain compatibility
        self.browse_file = self.app.file_selector._browse_file
        self.check_sharepoint_link = self.app.file_selector._check_sharepoint_link
        self.open_in_browser = self.app.file_selector._open_in_browser
        self.run_search = self.app.run_search
        self.display_results = self.app.results_display.display_results
        self.on_calendar_date_selected = self.app.on_calendar_date_selected
        self.save_results = self.app.save_results
        self.export_to_calendar = self.app.export_to_calendar

def launch():
    """Launch the GUI application"""
    root = tk.Tk()
    app = RosterSearchApp(root)
    root.mainloop()

if __name__ == "__main__":
    launch()
