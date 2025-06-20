"""
Results display component for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import scrolledtext, ttk
from .theme import COLORS, FONTS, create_button, create_label

class ResultsDisplay(ttk.Frame):
    """Component for displaying search results"""
    
    def __init__(self, parent, on_save=None, on_export=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(style='TFrame', padding=(10, 5, 10, 10))
        self.on_save = on_save
        self.on_export = on_export
        
        # Results header with counter
        self.header_frame = ttk.Frame(self, style='TFrame')
        self.header_frame.pack(fill='x', expand=False)
        
        self.header_label = ttk.Label(
            self.header_frame, 
            text="Search Results", 
            style='Subheader.TLabel'
        )
        self.header_label.pack(side='left')
        
        self.result_count = tk.StringVar(value="(0 results)")
        self.count_label = ttk.Label(
            self.header_frame,
            textvariable=self.result_count,
            font=FONTS['body'],
            foreground=COLORS['secondary']
        )
        self.count_label.pack(side='left', padx=(10, 0))
        
        # Results text area with improved styling
        self.results_frame = ttk.Frame(self, style='TFrame')
        self.results_frame.pack(fill='both', expand=True, pady=(10, 0))
        
        # Custom style for text widget
        self.results_text = scrolledtext.ScrolledText(
            self.results_frame,
            width=80,
            height=10,
            font=FONTS['body'],
            bg='white',
            fg=COLORS['text'],
            padx=8,
            pady=8,
            wrap=tk.WORD,
            borderwidth=1,
            relief="solid"
        )
        self.results_text.pack(fill='both', expand=True)
        self.results_text.config(state='disabled')
        
        # Action buttons
        self.button_frame = ttk.Frame(self, style='TFrame')
        self.button_frame.pack(fill='x', expand=False, pady=(10, 0))
        
        self.save_button = create_button(
            self.button_frame, 
            "Save Results", 
            command=self._on_save_click,
            width=15
        )
        self.save_button.pack(side='left', padx=(0, 10))
        self.save_button.state(['disabled'])
        
        self.export_button = create_button(
            self.button_frame, 
            "Export to Calendar", 
            command=self._on_export_click,
            width=18
        )
        self.export_button.pack(side='left')
        self.export_button.state(['disabled'])
        
        # Store results data
        self.results_data = []
    
    def display_results(self, results_data, name=None, file_path=None):
        """Display search results in the text area
        
        Args:
            results_data (list): List of result dictionaries
            name (str, optional): Name that was searched
            file_path (str, optional): Path to the file that was searched
        """
        self.results_data = results_data
        self.results_text.config(state='normal')
        self.results_text.delete(1.0, tk.END)
        
        if not results_data:
            self.results_text.insert(tk.END, "No results found.\n")
            self.result_count.set("(0 results)")
            self._update_button_states(False)
            self.results_text.config(state='disabled')
            return
            
        # Format and display results
        self.results_text.insert(tk.END, f"Results for '{name}' in '{file_path}'\n")
        self.results_text.insert(tk.END, "-" * 60 + "\n\n")
        
        # Group results by date
        results_by_date = {}
        for result in results_data:
            date = result.get('date', 'Unknown date')
            if date not in results_by_date:
                results_by_date[date] = []
            results_by_date[date].append(result)
            
        # Display results grouped by date
        for date, date_results in sorted(results_by_date.items()):
            self.results_text.insert(tk.END, f"Date: {date}\n", "date_header")
            
            for idx, result in enumerate(date_results):
                sheet = result.get('sheet', 'Unknown sheet')
                context = result.get('context', 'No context')
                position = result.get('position', 'No position')
                
                self.results_text.insert(tk.END, f"  • Sheet: {sheet}\n", "result_item")
                self.results_text.insert(tk.END, f"    Context: {context}\n", "result_detail")
                if position:
                    self.results_text.insert(tk.END, f"    Position: {position}\n", "result_detail")
                
                if idx < len(date_results) - 1:
                    self.results_text.insert(tk.END, "\n")
            
            self.results_text.insert(tk.END, "\n")
        
        # Update result count
        result_count = len(results_data)
        self.result_count.set(f"({result_count} {'result' if result_count == 1 else 'results'})")
        
        # Enable buttons if we have results
        self._update_button_states(True)
            
        # Set tags for text formatting
        self.results_text.tag_configure("date_header", 
                                       font=FONTS['subheader'],
                                       foreground=COLORS['secondary'])
        self.results_text.tag_configure("result_item", 
                                       font=FONTS['body'],
                                       foreground=COLORS['text'])
        self.results_text.tag_configure("result_detail", 
                                       font=FONTS['small'],
                                       foreground=COLORS['text'])
            
        # Disable editing
        self.results_text.config(state='disabled')
    
    def _update_button_states(self, has_results):
        """Update button states based on whether we have results"""
        if has_results:
            self.save_button.state(['!disabled'])
            self.export_button.state(['!disabled'])
        else:
            self.save_button.state(['disabled'])
            self.export_button.state(['disabled'])
    
    def _on_save_click(self):
        """Handle save button click"""
        if self.on_save:
            self.on_save(self.results_data)
    
    def _on_export_click(self):
        """Handle export button click"""
        if self.on_export:
            self.on_export(self.results_data)
    
    def clear(self):
        """Clear the results display"""
        self.results_data = []
        self.results_text.config(state='normal')
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state='disabled')
        self.result_count.set("(0 results)")
        self._update_button_states(False)
