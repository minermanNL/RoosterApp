"""
File selection component for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import filedialog, ttk
import webbrowser
from .theme import COLORS, FONTS, create_button, create_label, create_entry

class FileSelector(ttk.Frame):
    """Component for selecting Excel files or entering SharePoint/OneDrive links"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(style='TFrame', padding=(10, 10, 10, 5))
        
        # Create a header
        self.header_label = ttk.Label(
            self, 
            text="Excel File Selection", 
            style='Header.TLabel'
        )
        self.header_label.grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 10))
        
        # File path components
        self.file_label = create_label(self, "Excel file path:")
        self.file_label.grid(row=1, column=0, sticky='e', padx=5, pady=5)
        
        self.file_entry = create_entry(self, width=60)
        self.file_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.file_entry.bind('<KeyRelease>', self._check_sharepoint_link)
        
        self.browse_button = create_button(self, "Browse", command=self._browse_file, width=8)
        self.browse_button.grid(row=1, column=2, padx=5, pady=5)
        
        self.open_browser_button = create_button(
            self, 
            "Open in Browser", 
            command=self._open_in_browser, 
            width=15
        )
        self.open_browser_button.grid(row=1, column=3, padx=5, pady=5)
        self.open_browser_button.state(['disabled'])
        
        # Password entry components
        self.pw_label = create_label(self, "Excel password (optional):")
        self.pw_label.grid(row=2, column=0, sticky='e', padx=5, pady=5)
        
        self.pw_entry = ttk.Entry(self, width=40, show='*')
        self.pw_entry.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        # Configure grid layout
        self.columnconfigure(1, weight=1)
        
        # Tooltip frame for helpful information
        self.tip_frame = ttk.Frame(self, style='TFrame', padding=(5, 0, 5, 0))
        self.tip_frame.grid(row=3, column=0, columnspan=4, sticky='ew', pady=(5, 0))
        
        tip_text = "Accepts local Excel files or SharePoint/OneDrive links"
        self.tip_label = ttk.Label(
            self.tip_frame, 
            text=tip_text,
            font=FONTS['small'],
            foreground=COLORS['secondary']
        )
        self.tip_label.pack(anchor='w')

    def _browse_file(self):
        """Open file dialog to select an Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel Roster File",
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"), 
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
        self._check_sharepoint_link()

    def _check_sharepoint_link(self, event=None):
        """Check if the entered path is a SharePoint or OneDrive link"""
        url = self.file_entry.get().strip()
        is_cloud_link = any(domain in url.lower() for domain in [
            'sharepoint.com', 
            'onedrive.live.com', 
            '1drv.ms'
        ])
        
        if is_cloud_link:
            self.open_browser_button.state(['!disabled'])
        else:
            self.open_browser_button.state(['disabled'])

    def _open_in_browser(self):
        """Open the SharePoint or OneDrive link in the default browser"""
        url = self.file_entry.get().strip()
        if url:
            webbrowser.open(url)

    def get_file_path(self):
        """Return the currently selected file path"""
        return self.file_entry.get().strip()

    def get_password(self):
        """Return the entered password"""
        return self.pw_entry.get().strip() or None
