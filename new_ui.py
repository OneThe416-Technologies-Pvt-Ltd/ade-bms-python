import os
import subprocess  # Import subprocess module
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from helpers.utils import load_and_resize_image, create_image_button
from pcan_api.custom_pcan_methods import pcan_write_read
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
        self.master.title("Connection Settings")
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

         # Initialize the CanBatteryInfo frame
        self.battery_info = None

    def center_window(self, width, height):
        # Calculate x and y coordinates to center the window
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.master.geometry(f'{width}x{height}+{x}+{y}')
    
    def show_can_battery_info(self, event=None):
        self.notebook.pack_forget()
        # Set the desired size for the battery info window
        self.center_window(1200, 600)
        if not self.battery_info:
            self.battery_info = CanBatteryInfo(self.master, self)  # Pass the reference to the MainWindow
        else:
            self.battery_info.main_frame.pack(fill="both", expand=True)
        
        # Show the dashboard by default
        self.battery_info.show_dashboard()
    
    def show_rs_battery_info(self, event=None):
        self.notebook.pack_forget()
    
        # Set the desired size for the battery info window
        self.center_window(1200, 600)  # Adjust the width and height as necessary
        if not self.battery_info:
            self.battery_info = RSBatteryInfo(self.master, self)  # Pass the reference to the MainWindow
        else:
            self.battery_info.main_frame.pack(fill="both", expand=True)
        
        # Show the dashboard by default
        # self.battery_info.show_dashboard()

        # Path to the .exe file
        #exe_path = os.path.join(base_dir, 'C:\Program Files\PostgreSQL\16\pgAdmin 4\runtime\pgAdmin4.exe')  # Replace with the actual path to the .exe file
        # exe_path = r'C:\Program Files\PostgreSQL\16\pgAdmin 4\runtime\pgAdmin4.exe'
        # # Open the .exe file
        # try:
        #     subprocess.Popen(exe_path)  # This will open the .exe file
        # except Exception as e:
        #     messagebox.showerror("Error", f"Failed to open the .exe file: {e}")

    def show_main_window(self):
        if self.battery_info:
            self.battery_info.main_frame.pack_forget()
        # Set the desired size for the main window when showing the notebook again
        self.center_window(500, 400)
        self.notebook.pack(fill="both", expand=True)

            
if __name__ == "__main__":
    # Create the root window
    root = tk.Tk()
    # Hide the main window while the splash screen is displayed
    # root.withdraw()

    # # Create and show the splash screen
    # splash = SplashScreen(root)
    # splash.update()

    # # After the splash screen is closed, show the main window
    # root.after(5000, lambda: (splash.destroy(), root.deiconify()))

    # Initialize and run the main application
    app = MainWindow(root)
    root.mainloop()
