#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel-Based File Renamer

A GUI application for renaming files based on matching data in Excel spreadsheets.
"""

from src.ui.main_window import MainWindow
import tkinter as tk
from tkinter import ttk
import sys
import os

# Add the parent directory to sys.path to allow relative imports
sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '..')))


def setup_theme(root):
    """Set up a modern, feminine-oriented theme with pastel blue palette"""
    # Define our custom color scheme
    # Soft, pastel blue palette with professional look
    colors = {
        "primary": "#89CFF0",  # Pastel blue
        "primary_light": "#C5E3F6",  # Light pastel blue
        "secondary": "#A0D2EB",  # Secondary pastel blue
        "accent": "#BEE7E8",  # Mint/aqua accent
        "text": "#4A4A4A",  # Dark gray
        "text_light": "#7D7D7D",  # Medium gray
        "background": "#FFFFFF",  # White
        "background_alt": "#F0F8FF",  # Very light blue (Alice Blue)
        "success": "#87CEAB",  # Pastel green
        "warning": "#FFD580",  # Pastel amber
        "error": "#FFAFCC"  # Pastel pink for errors
    }

    # Configure ttk styles
    style = ttk.Style(root)

    # Try to use a more modern theme as base if available
    try:
        # First try 'clam' which is widely available and customizable
        style.theme_use('clam')
    except tk.TclError:
        # Fallback to default
        pass

    # Configure common elements
    style.configure('TFrame', background=colors["background"])
    style.configure(
        'TLabel', background=colors["background"], foreground=colors["text"])

    # Buttons with rounded corners effect
    style.configure('TButton',
                    background=colors["primary"],
                    foreground=colors["text"],
                    borderwidth=1,
                    relief='flat',
                    padding=[10, 6])
    style.map('TButton',
              background=[('active', colors["primary_light"]),
                          ('pressed', colors["primary_light"])],
              foreground=[('active', colors["text"]),
                          ('pressed', colors["text"])])

    # Accent buttons (for main actions)
    style.configure('Accent.TButton',
                    background=colors["accent"],
                    foreground=colors["text"],
                    padding=[10, 6])
    style.map('Accent.TButton',
              background=[('active', "#D1F0F0"),
                          ('pressed', "#D1F0F0")])

    # Entry fields
    style.configure('TEntry',
                    fieldbackground=colors["background_alt"],
                    foreground=colors["text"],
                    borderwidth=1,
                    relief='solid',
                    padding=[5, 3])

    # Combobox
    style.configure('TCombobox',
                    fieldbackground=colors["background_alt"],
                    foreground=colors["text"],
                    borderwidth=1,
                    relief='solid',
                    padding=[5, 3],
                    arrowsize=12)

    # Notebook (tabs)
    style.configure('TNotebook',
                    background=colors["background"],
                    tabmargins=[2, 5, 2, 0])
    style.configure('TNotebook.Tab',
                    background=colors["background_alt"],
                    foreground=colors["text"],
                    padding=[10, 2])
    style.map('TNotebook.Tab',
              background=[('selected', colors["primary_light"]),
                          ('active', colors["primary_light"])],
              foreground=[('selected', colors["text"]),
                          ('active', colors["text"])])

    # Treeview (for Excel data)
    style.configure("Treeview",
                    background=colors["background"],
                    foreground=colors["text"],
                    rowheight=25,
                    fieldbackground=colors["background"])
    style.configure("Treeview.Heading",
                    background=colors["primary_light"],
                    foreground=colors["text"],
                    padding=[5, 3])
    style.map("Treeview",
              background=[('selected', colors["primary"]),
                          ('active', colors["primary_light"])],
              foreground=[('selected', colors["text"]),
                          ('active', colors["text"])])

    # LabelFrame
    style.configure('TLabelframe',
                    background=colors["background"],
                    foreground=colors["text"],
                    borderwidth=1,
                    relief='groove')
    style.configure('TLabelframe.Label',
                    background=colors["background"],
                    foreground=colors["primary"],
                    font=('Arial', 10, 'bold'))

    # Scrollbars
    style.configure('TScrollbar',
                    background=colors["background_alt"],
                    troughcolor=colors["background"],
                    borderwidth=0,
                    arrowsize=12)

    # Set root window properties
    root.configure(bg=colors["background"])

    # Return colors for use in the app
    return colors


def main():
    """Main entry point for the application"""
    root = tk.Tk()

    # Set application title
    root.title("File Renamer & Organizer")

    # Make the window start maximized
    root.state('zoomed')  # Windows-specific maximized state

    # Setup custom theme
    theme_colors = setup_theme(root)

    # Create the application with theme
    app = MainWindow(root, theme_colors)

    # Start the application
    root.mainloop()


if __name__ == "__main__":
    main()
