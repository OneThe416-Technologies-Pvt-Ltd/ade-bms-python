import os
import subprocess  # Import subprocess module
import tkinter as tk
from tkinter import ttk
from gui.can_connect import CanConnection
from gui.rs_connect import RSConnection
from gui.can_battery_info import CanBatteryInfo
from gui.rs_battery_info import RSBatteryInfo
from gui.splash_screen import SplashScreen


# Get the base directory of the script
base_dir = os.path.dirname(__file__)

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("ADE BMS")
        self.center_window(500, 400)  # Center the window with the given width and height
        self.master.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
        self.master.resizable(False, False)
        
        # Create a notebook (tab container)
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill='both', expand=True)

        # Create CAN tab
        self.can_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.can_frame, text="CAN")

        # Initialize the CAN connection settings
        self.can_connection = CanConnection(self.can_frame, self)
        self.can_connection.pack(fill="both", expand=True)

        # Create RS232/RS422 tab
        self.rs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.rs_frame, text="RS232/RS422")

        # Initialize the RS232/RS422 connection settings
        self.rs_connection = RSConnection(self.rs_frame, self)
        self.rs_connection.pack(fill="both", expand=True)

        # Initialize the CanBatteryInfo and RSBatteryInfo frames separately
        self.can_battery_info = None
        self.rs_battery_info = None

    def center_window(self, width, height):
        # Calculate x and y coordinates to center the window
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.master.geometry(f'{width}x{height}+{x}+{y}')
    
    def show_can_battery_info(self, event=None):
        if self.rs_battery_info:
            self.rs_battery_info.main_frame.pack_forget()
        
        self.notebook.pack_forget()
        # Maximize the window
        self.master.overrideredirect(False)  # Ensure window decorations (close, min, max) are shown
        self.master.state('zoomed')  # Maximize the window to full screen

        if not self.can_battery_info:
            self.can_battery_info = CanBatteryInfo(self.master, self)  # Pass the reference to the MainWindow
        else:
            self.can_battery_info.main_frame.pack(fill="both", expand=True)
        
        # Show the dashboard by default
        self.can_battery_info.show_dashboard()
    
    def show_rs_battery_info(self, event=None):
        if self.can_battery_info:
            self.can_battery_info.main_frame.pack_forget()
        
        self.notebook.pack_forget()
        # Maximize the window
        self.master.overrideredirect(False)  # Ensure window decorations (close, min, max) are shown
        self.master.state('zoomed')  # Maximize the window to full screen
        if not self.rs_battery_info:
            self.rs_battery_info = RSBatteryInfo(self.master, self)  # Pass the reference to the MainWindow
        else:
            self.rs_battery_info.main_frame.pack(fill="both", expand=True)

    def show_main_window(self):
        if self.can_battery_info:
            self.can_battery_info.main_frame.pack_forget()
        if self.rs_battery_info:
            self.rs_battery_info.main_frame.pack_forget()
        
        # Set the desired size for the main window when showing the notebook again
        self.master.state('normal')  # Reset from maximized to normal window size
        self.center_window(500, 400)  # Set the window to a specific size (500x400)
        self.master.resizable(False, False)  # Disable window resizing
        self.notebook.pack(fill="both", expand=True)

            
if __name__ == "__main__":
    # Create the root window
    root = tk.Tk()
    # Hide the main window while the splash screen is displayed
    root.withdraw()

    # Create and show the splash screen
    splash = SplashScreen(root)
    splash.update()

    # After the splash screen is closed, show the main window
    root.after(5000, lambda: (splash.destroy(), root.deiconify()))

    # Initialize and run the main application
    app = MainWindow(root)
    root.mainloop()
