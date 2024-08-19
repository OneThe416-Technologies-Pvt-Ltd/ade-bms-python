# main.py
import os
import sys

if hasattr(sys, '_MEIPASS'):
    os.environ['TCL_LIBRARY'] = os.path.join(sys._MEIPASS, 'tcl')
    os.environ['TK_LIBRARY'] = os.path.join(sys._MEIPASS, 'tk')

import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from helpers.utils import load_and_resize_image, create_image_button
from gui.can_connect import CanConnection
from gui.rs_connect import RSConnection
from gui.can_battery_info import CanBatteryInfo
from gui.splash_screen import SplashScreen

# Get the base directory of the script
base_dir = os.path.dirname(__file__)

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("ADE BMS")
        self.master.geometry("900x600")
        self.master.resizable(False, False)

        icon_path = os.path.join(base_dir, 'assets/logo/drdo_icon.ico')
        self.master.iconbitmap(icon_path)

        self.can_connected = False
        self.rs_connected = False

        self.background_image_path = os.path.join(base_dir, 'assets/images/bg_main1.png')
        self.bg_image = Image.open(self.background_image_path)
        self.bg_image = self.bg_image.resize((900, 600), Image.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        self.canvas = tk.Canvas(self.master, width=900, height=600)
        self.canvas.pack(fill='both', expand=True)
        self.canvas.create_image(0, 0, image=self.bg_photo, anchor='nw')

        self.button_coords = {
            'can': (570, 95, 800, 180),
            'rs': (580, 220, 800, 340),
            'can_battery': (690, 100, 950, 160),
            'rs_battery': (700, 220, 960, 310)
        }
        self.colors = {
            'can': '#33B4B4',
            'rs': '#33B4B4',
            'can_battery': '#33B4B4',
            'rs_battery': '#33B4B4'
        }

        self.image_refs = {}
        self.canvas_items = {}

        self.create_buttons()

        self.battery_info = None  # Initialize as None

    def create_buttons(self):
        images = {
            'can': os.path.join(base_dir, 'assets/images/pcan_btn_img.png'),
            'rs': os.path.join(base_dir, 'assets/images/moxa_rs232_btn_img.png'),
            'can_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png'),
            'rs_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png')
        }

        sizes = {
            'can': (65, 65),
            'rs': (80, 70),
            'can_battery': (100, 90),
            'rs_battery': (100, 90)
        }

        handlers = {
            'can': self.handle_can_btn_click,
            'rs': self.handle_rs_btn_click,
            'can_battery': self.handle_can_battery_btn_click,
            'rs_battery': self.handle_rs_battery_btn_click
        }

        for key, (x1, y1, x2, y2) in self.button_coords.items():
            image_tk = load_and_resize_image(images[key], sizes[key], border_width=15, border_color=self.colors.get(key.split('_')[0]))
            if image_tk:
                self.image_refs[key] = image_tk 
                position = ((x2 + x1) // 2, (y2 + y1) // 2)
                item_id = create_image_button(self.canvas, image_tk, position, handlers[key])
                self.canvas_items[key] = item_id

    def show_can_battery_info(self, event=None):
        self.canvas.pack_forget()
        if not self.battery_info:
            self.battery_info = CanBatteryInfo(self.master, self)  # Pass the reference to the MainWindow
        else:
            self.battery_info.main_frame.pack(fill="both", expand=True)
        
        # Show the dashboard by default
        self.battery_info.show_dashboard()

    def show_main_window(self):
        if self.battery_info:
            self.battery_info.main_frame.pack_forget()
        self.canvas.pack(fill="both", expand=True)

    def handle_can_btn_click(self, event):
        if self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 already connected")
        else:
            main_window_x = self.master.winfo_rootx()
            main_window_y = self.master.winfo_rooty()
            main_window_width = self.master.winfo_width()
            main_window_height = self.master.winfo_height()

            can_window_width = 400
            can_window_height = 400

            can_window_x = main_window_x + (main_window_width - can_window_width) // 2
            can_window_y = main_window_y + (main_window_height - can_window_height) // 2

            can_window = tk.Toplevel(self.master)
            can_window.title("CAN Connection Settings")
            can_window.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)

            can_connection = CanConnection(can_window)
            can_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            can_window.protocol("WM_DELETE_WINDOW", lambda: self.on_can_window_close(can_window, can_connection))

            can_window.grab_set()


    def update_colors(self):
        if self.can_connected:
            self.colors['can'] = 'green'
            self.colors['can_battery'] = 'green'
        else:
            self.colors['can'] = '#33B4B4'
            self.colors['can_battery'] = '#33B4B4'

        if self.rs_connected:
            self.colors['rs'] = 'green'
            self.colors['rs_battery'] = 'green'
        else:
            self.colors['rs'] = '#33B4B4'
            self.colors['rs_battery'] = '#33B4B4'

        images = {
            'can': os.path.join(base_dir, 'assets/images/pcan_btn_img.png'),
            'rs': os.path.join(base_dir, 'assets/images/moxa_rs232_btn_img.png'),
            'can_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png'),
            'rs_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png')
        }
        sizes = {
            'can': (70, 70),
            'rs': (80, 70),
            'can_battery': (100, 90),
            'rs_battery': (100, 90)
        }

        for key in ['can', 'rs', 'can_battery', 'rs_battery']:
            if key in self.canvas_items:
                image_path = images[key]
                size = sizes[key]
                border_color = self.colors[key]

                image_tk = load_and_resize_image(image_path, size, border_width=15, border_color=border_color)
                if image_tk:
                    self.image_refs[key] = image_tk
                    self.canvas.itemconfig(self.canvas_items[key], image=image_tk)
                else:
                    print(f"Failed to load image for {key} with color {border_color}")
            else:
                print(f"Key {key} not found in canvas_items")


    def on_can_window_close(self, window, can_connection):
        self.can_connected = can_connection.get_connection_status()
        self.update_colors()
        window.grab_release()
        window.destroy()
        self.master.attributes("-disabled", False)
        self.master.lift()
        self.master.focus_force()


    def handle_rs_btn_click(self, event):
        if self.can_connected:
            messagebox.showerror("ERROR!", "CAN already connected")
        else:
            main_window_x = self.master.winfo_rootx()
            main_window_y = self.master.winfo_rooty()
            main_window_width = self.master.winfo_width()
            main_window_height = self.master.winfo_height()

            rs_window_width = 400
            rs_window_height = 400

            rs_window_x = main_window_x + (main_window_width - rs_window_width) // 2
            rs_window_y = main_window_y + (main_window_height - rs_window_height) // 2

            rs_window = tk.Toplevel(self.master)
            rs_window.title("RS Connection Settings")
            rs_window.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
            rs_window.geometry(f"{rs_window_width}x{rs_window_height}+{rs_window_x}+{rs_window_y}")
            rs_window.resizable(False, False)

            rs_connection = RSConnection(rs_window)
            rs_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            rs_window.protocol("WM_DELETE_WINDOW", lambda: self.on_rs_window_close(rs_window, rs_connection))

            rs_window.grab_set()


    def on_rs_window_close(self, window, rs_connection):
        self.rs_connected = rs_connection.get_connection_status()
        self.update_colors()
        window.grab_release()
        window.destroy()
        self.master.attributes("-disabled", False)
        self.master.lift()


    def handle_rs_battery_btn_click(self, event):
        if self.can_connected:
            messagebox.showerror("ERROR!", "CAN already connected")
        elif not self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 not connected")
        else:
            print("RS232 Battery connect")


    def handle_can_battery_btn_click(self, event):
        # self.show_can_battery_info()
        if self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 already connected")
        elif not self.can_connected:
            messagebox.showerror("ERROR!", "CAN not connected")
        else:
            self.show_can_battery_info()


if __name__ == "__main__":
    # Create the root window
    root = tk.Tk()
    icon_path = os.path.join(base_dir, 'assets/logo/drdo_icon.ico')
    icon_image = Image.open(icon_path)
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(False, icon_photo)

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
