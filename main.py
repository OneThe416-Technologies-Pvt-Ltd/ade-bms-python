import os
import subprocess  # Import subprocess module for managing additional system processes
import tkinter as tk
import ttkbootstrap as ttk  # Use ttkbootstrap for themed widgets
from gui.can_connect import CanConnection  # Import CAN connection handling
from gui.rs_connect import RSConnection  # Import RS232/RS422 connection handling
from gui.settings import Settings  # Import settings configuration
from gui.can_battery_info import CanBatteryInfo  # Import CAN battery info display
from gui.rs_battery_info import RSBatteryInfo  # Import RS battery info display
from gui.splash_screen import SplashScreen  # Import splash screen functionality
from helpers.config import *  # Import configuration helper functions
from helpers.logger import logger  # Import custom logging configuration

# Get the base directory of the script for resource paths
base_dir = os.path.dirname(__file__)

# Main application window class``
class MainWindow:
    def __init__(self, master):
        self.master = master  # Reference to the root window
        self.master.title("ADE BMS")  # Set the main window title
        self.center_window(500, 400)  # Center the window with specified width and height
        # Set application icon
        self.master.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
        self.master.resizable(False, False)  # Disable resizing of the main window

        load_config()  # Load configuration settings at startup

        # Initialize notebook for tabs
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill='both', expand=True)

        # Create CAN tab and initialize connection settings
        self.can_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.can_frame, text="CAN")
        self.can_connection = CanConnection(self.can_frame, self)  # CAN connection instance
        self.can_connection.pack(fill="both", expand=True)
        logger.info("CAN tab initialized.")

        # Create RS232/RS422 tab and initialize connection settings
        self.rs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.rs_frame, text="RS232/RS422")
        self.rs_connection = RSConnection(self.rs_frame, self)  # RS232/RS422 connection instance
        self.rs_connection.pack(fill="both", expand=True)
        logger.info("RS232/RS422 tab initialized.")

        # Create Settings tab and initialize settings
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Settings")
        self.settings = Settings(self.settings_frame, self)  # Settings instance
        self.settings.pack(fill="both", expand=True)
        logger.info("Settings tab initialized.")

        # Prepare instances for CAN and RS battery information displays
        self.can_battery_info = None
        self.rs_battery_info = None
    
        # Bind tab change event to handle switching
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)
        logger.info("Tab change event bound.")

    def on_tab_change(self, event):
        # Detect and handle the selected tab for updating information as needed
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if selected_tab == "CAN":
            logger.info("Switched to CAN tab.")
            self.can_connection.update_displayed_values()  # Refresh displayed values

    def create_menu(self):
        # Set up main menu bar with Settings menu
        menu_bar = tk.Menu(self.master)
        self.master.config(menu=menu_bar)
        
        # Add "Settings" option to the menu bar
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)
        
        # Add command to open the settings window
        settings_menu.add_command(label="Open Settings", command=self.open_settings_window)
        logger.info("Settings menu created.")

    def open_settings_window(self):
        # Open a new window specifically for Settings
        settings_window = tk.Toplevel(self.master)
        settings_window.title("Settings")
        settings_window.geometry("400x300")  # Define window size
        
        # Display the Settings content in this new window
        settings_frame = ttk.Frame(settings_window)
        settings_frame.pack(fill="both", expand=True, padx=10, pady=10)
        settings = Settings(settings_frame, self)  # Pass main window reference if needed
        settings.pack(fill="both", expand=True)
        logger.info("Settings window opened.")

    def center_window(self, width, height):
        # Calculate and set coordinates to center the window on the screen
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.master.geometry(f'{width}x{height}+{x}+{y}')
        logger.info(f"Window centered with size {width}x{height}.")

    def show_can_battery_info(self, event=None):
        # Display CAN battery info in maximized mode, hiding RS battery info if open
        if self.rs_battery_info:
            self.rs_battery_info.main_frame.pack_forget()
            logger.info("RS Battery Info frame hidden.")
        
        self.notebook.pack_forget()  # Hide the main notebook tabs
        self.master.overrideredirect(False)  # Show window decorations (close, min, max buttons)
        self.master.state('zoomed')  # Maximize window
        logger.info("Window maximized for CAN Battery Info.")

        # Initialize or show CAN battery information dashboard
        if not self.can_battery_info:
            self.can_battery_info = CanBatteryInfo(self.master, self)
        else:
            self.can_battery_info.main_frame.pack(fill="both", expand=True)
        self.can_battery_info.show_dashboard()  # Show default dashboard view
        logger.info("CAN Battery Info dashboard displayed.")

        # self.can_battery_info.auto_refresh()  # Start auto-refresh for CAN battery info
        logger.info("Auto refresh started for CAN Battery Info.")
    
    def show_rs_battery_info(self, event=None):
        # Display RS battery info in maximized mode, hiding CAN battery info if open
        if self.can_battery_info:
            self.can_battery_info.main_frame.pack_forget()
            logger.info("CAN Battery Info frame hidden.")
        
        self.notebook.pack_forget()  # Hide main notebook tabs
        self.master.overrideredirect(False)  # Show window decorations
        self.master.state('zoomed')  # Maximize window
        logger.info("Window maximized for RS Battery Info.")

        # Initialize or show RS battery information
        if not self.rs_battery_info:
            self.rs_battery_info = RSBatteryInfo(self.master, self)
        else:
            self.rs_battery_info.main_frame.pack(fill="both", expand=True)
        logger.info("RS Battery Info displayed.")

    def show_main_window(self):
        # Restore main notebook view and reset window size
        if self.can_battery_info:
            self.can_battery_info.main_frame.pack_forget()
            logger.info("CAN Battery Info frame hidden.")
        if self.rs_battery_info:
            self.rs_battery_info.main_frame.pack_forget()
            logger.info("RS Battery Info frame hidden.")
        
        self.master.state('normal')  # Restore window to normal size
        self.center_window(500, 400)  # Center the window with default dimensions
        self.master.resizable(False, False)  # Disable resizing
        self.notebook.pack(fill="both", expand=True)  # Show notebook tabs
        logger.info("Main window displayed.")

            
if __name__ == "__main__":
    # Create the main root window
    root = tk.Tk()
    root.withdraw()  # Hide main window initially for splash screen

    # Show splash screen
    splash = SplashScreen(root)
    splash.update()  # Update splash screen display

    # After 5 seconds, close splash screen and display main window
    root.after(5000, lambda: (splash.destroy(), root.deiconify()))
    logger.info("Splash screen shown.")

    # Initialize and start the main application loop
    app = MainWindow(root)
    logger.info("Main application started.")
    root.mainloop()
