"""
Theme and styling configuration for the Excel Roster Search application
"""
import tkinter as tk
from tkinter import ttk
import sys
import os

# Color scheme
COLORS = {
    'primary': '#2c3e50',       # Dark blue-grey
    'secondary': '#3498db',     # Bright blue
    'accent': '#1abc9c',        # Turquoise
    'background': '#ecf0f1',    # Light grey
    'text': '#2c3e50',          # Dark blue-grey for text
    'text_light': '#ecf0f1',    # Light grey for text on dark backgrounds
    'success': '#2ecc71',       # Green
    'warning': '#f39c12',       # Orange
    'error': '#e74c3c',         # Red
}

# Font configurations
FONTS = {
    'header': ('Helvetica', 16, 'bold'),
    'subheader': ('Helvetica', 14, 'bold'),
    'body': ('Helvetica', 12),
    'small': ('Helvetica', 10),
    'button': ('Helvetica', 12),
}

def setup_theme(root):
    """Apply the application theme to the root window"""
    root.configure(bg=COLORS['background'])
    
    # Create styles for ttk widgets
    style = ttk.Style(root)
    style.theme_use('clam')  # Use clam as base theme
    
    # Configure button style
    style.configure(
        'TButton',
        font=FONTS['button'],
        background=COLORS['secondary'],
        foreground=COLORS['text_light'],
        borderwidth=1,
        focusthickness=3,
        focuscolor=COLORS['accent']
    )
    
    style.map(
        'TButton',
        background=[('active', COLORS['accent']), ('disabled', '#d1d8e0')],
        foreground=[('disabled', '#7f8c8d')],
        relief=[('pressed', 'sunken')]
    )
    
    # Configure entry style
    style.configure(
        'TEntry',
        font=FONTS['body'],
        fieldbackground=COLORS['background'],
        foreground=COLORS['text']
    )
    
    # Configure frame style
    style.configure(
        'TFrame',
        background=COLORS['background']
    )
    
    # Configure label style
    style.configure(
        'TLabel',
        font=FONTS['body'],
        background=COLORS['background'],
        foreground=COLORS['text']
    )
    
    style.configure(
        'Header.TLabel',
        font=FONTS['header'],
        background=COLORS['background'],
        foreground=COLORS['primary']
    )
    
    style.configure(
        'Subheader.TLabel',
        font=FONTS['subheader'],
        background=COLORS['background'],
        foreground=COLORS['primary']
    )
    
    # Configure scrollbar style
    style.configure(
        'TScrollbar',
        background=COLORS['background'],
        troughcolor=COLORS['background'],
        arrowcolor=COLORS['text']
    )

def create_button(parent, text, command=None, width=10, style='TButton'):
    """Creates a standard styled button"""
    button = ttk.Button(parent, text=text, command=command, style=style, width=width)
    return button

def create_label(parent, text, style='TLabel'):
    """Creates a standard styled label"""
    label = ttk.Label(parent, text=text, style=style)
    return label

def create_entry(parent, width=40):
    """Creates a standard styled entry widget"""
    entry = ttk.Entry(parent, width=width)
    return entry

def create_frame(parent, padding=(5, 5, 5, 5)):
    """Creates a standard styled frame"""
    frame = ttk.Frame(parent, padding=padding)
    frame.pack(fill='both', expand=True)
    return frame

def show_info(title, message):
    """Shows info message with consistent styling"""
    from tkinter import messagebox
    return messagebox.showinfo(title, message)

def show_error(title, message):
    """Shows error message with consistent styling"""
    from tkinter import messagebox
    return messagebox.showerror(title, message)

def show_warning(title, message):
    """Shows warning message with consistent styling"""
    from tkinter import messagebox
    return messagebox.showwarning(title, message)
