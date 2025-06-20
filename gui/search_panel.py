"""
Search panel component for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import ttk
from .theme import create_button, create_label, create_entry, show_error

class SearchPanel(ttk.Frame):
    """Component for entering search criteria and executing searches"""
    
    def __init__(self, parent, on_search_callback=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(style='TFrame', padding=(10, 5, 10, 10))
        self.on_search_callback = on_search_callback
        
        # Create search input components
        self.search_frame = ttk.Frame(self, style='TFrame')
        self.search_frame.pack(fill='x', expand=True)
        
        # Name entry
        self.name_label = create_label(self.search_frame, "Name to search:")
        self.name_label.grid(row=0, column=0, sticky='e', padx=5, pady=5)
        
        self.name_entry = create_entry(self.search_frame, width=50)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        self.name_entry.insert(0, "Naam")  # Default placeholder
        
        # Search button with improved styling
        button_frame = ttk.Frame(self, style='TFrame')
        button_frame.pack(fill='x', pady=(10, 0))
        
        self.search_button = create_button(
            button_frame, 
            "Search Roster", 
            command=self._execute_search,
            width=15
        )
        self.search_button.pack(pady=5)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground="#3498db",  # Blue for status messages
            font=('Helvetica', 10, 'italic')
        )
        self.status_label.pack(fill='x', pady=(5, 0))
        
        # Configure grid layout
        self.search_frame.columnconfigure(1, weight=1)
        
        # Focus on name entry by default
        self.name_entry.focus_set()
        # Bind Enter key to search
        self.name_entry.bind("<Return>", lambda event: self._execute_search())

    def _execute_search(self):
        """Execute the search with the provided name"""
        search_name = self.name_entry.get().strip()
        
        if not search_name or search_name == "Naam":
            show_error("Search Error", "Please enter a name to search for.")
            return
        
        self.status_var.set("Searching...")
        
        if self.on_search_callback:
            self.on_search_callback(search_name)
    
    def get_search_name(self):
        """Return the currently entered search name"""
        return self.name_entry.get().strip()
    
    def reset_status(self):
        """Reset the status message"""
        self.status_var.set("")
    
    def set_status(self, message):
        """Set the status message"""
        self.status_var.set(message)
