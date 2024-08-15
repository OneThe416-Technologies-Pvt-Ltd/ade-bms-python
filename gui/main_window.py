# gui/main_window.py
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import messagebox
from gui.can_connect import CanConnection
from pcan_api.custom_pcan_methods import *
from helpers.utils import load_and_resize_image, create_image_button
from gui.can_battery_info import BatteryInfo

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("ADE BMS")
        self.master.geometry("1000x600")
        self.master.resizable(False, False)  # Disable maximize

        # Set custom icon
        self.master.iconbitmap('asserts/logo/drdo_icon.ico')

        self.can_connected = False
        self.rs_connected = False

        # Load background image
        self.background_image_path = 'asserts/images/bg_main1.png'
        self.bg_image = Image.open(self.background_image_path)
        self.bg_image = self.bg_image.resize((1000, 600), Image.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)
        self.canvas = tk.Canvas(self.master, width=1000, height=600)
        self.canvas.pack(fill='both', expand=True)
        self.canvas.create_image(0, 0, image=self.bg_photo, anchor='nw')

        # Define coordinates for buttons
        self.button_coords = {
            'can': (720, 100, 800, 180),
            'rs': (730, 220, 800, 340),
            'can_battery': (870, 100, 950, 160),
            'rs_battery': (880, 220, 960, 310)
        }
        self.colors = {
            'can': '#33B4B4',
            'rs': '#33B4B4',
            'can_battery': '#33B4B4',
            'rs_battery': '#33B4B4'
        }

        self.image_refs = {}  # Store references to PhotoImage objects
        self.canvas_items = {}  # Store references to canvas item IDs

        # Load and create buttons
        self.images = {}
        self.create_buttons()

    def create_buttons(self):
        images = {
            'can': 'asserts/images/pcan_btn_img.png',
            'rs': 'asserts/images/moxa_rs232_btn_img.png',
            'can_battery': 'asserts/images/brenergy_battery_img.png',
            'rs_battery': 'asserts/images/brenergy_battery_img.png'
        }

        sizes = {
            'can': (80, 70),
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
                 self.image_refs[key] = image_tk  # Hold reference to the PhotoImage object
                 position = ((x2 + x1) // 2, (y2 + y1) // 2)
                 item_id = create_image_button(self.canvas, image_tk, position, handlers[key])
                 self.canvas_items[key] = item_id

    def handle_can_btn_click(self, event):
        if self.rs_connected:
            messagebox.showerror("ERROR!","RS232/RS422 already connected")
        else:
            main_window_x = self.master.winfo_rootx ()
            main_window_y = self.master.winfo_rooty()
            main_window_width = self.master.winfo_width()
            main_window_height = self.master.winfo_height()

            can_window_width = 400
            can_window_height = 400

            # Calculate the center position
            can_window_x = main_window_x + (main_window_width - can_window_width) // 2
            can_window_y = main_window_y + (main_window_height - can_window_height) // 2

            # Create a Toplevel window for CAN connection settings
            can_window = tk.Toplevel(self.master)
            can_window.title("CAN Connection Settings")
            can_window.iconbitmap('asserts/logo/drdo_icon.ico')
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)  # Disable maximize

            # Instantiate CanConnection class from can_connect.py
            can_connection = CanConnection(can_window)
            can_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            # Show the main window again when can_window is closed
            can_window.protocol("WM_DELETE_WINDOW", lambda: self.on_can_window_close(can_window,can_connection))

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
            'can': 'asserts/images/pcan_btn_img.png',
            'rs': 'asserts/images/moxa_rs232_btn_img.png',
            'can_battery': 'asserts/images/brenergy_battery_img.png',
            'rs_battery': 'asserts/images/brenergy_battery_img.png'
        }
        sizes = {
            'can': (80, 70),
            'rs': (80, 70),
            'can_battery': (100, 90),
            'rs_battery': (100, 90)
        }

        # Update button images on the canvas
        for key in ['can', 'rs', 'can_battery', 'rs_battery']:
            if key in self.canvas_items:
                image_path = images[key]
                size = sizes[key]
                border_color = self.colors[key]

                # Load and resize image with updated color
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
        window.grab_release()  # Release the grab set by can_window
        window.destroy()  # Destroy the Toplevel window
        self.master.attributes("-disabled", False)  # Re-enable the main window
        self.master.lift()  # Bring the main window to the front
        self.master.focus_force()  # Ensure focus returns to the main window

    def handle_rs_btn_click(self, event):
        if self.can_connected:
            messagebox.showerror("ERROR!","CAN already connected")
        else:
            print("RS232 connect")
            main_window_x = self.master.winfo_rootx ()
            main_window_y = self.master.winfo_rooty()
            main_window_width = self.master.winfo_width()
            main_window_height = self.master.winfo_height()

            can_window_width = 400
            can_window_height = 400

            # Calculate the center position
            can_window_x = main_window_x + (main_window_width - can_window_width) // 2
            can_window_y = main_window_y + (main_window_height - can_window_height) // 2

            # Create a Toplevel window for CAN connection settings
            can_window = tk.Toplevel(self.master)
            can_window.title("RS Connection Settings")
            can_window.iconbitmap('asserts/logo/drdo_icon.ico')
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)  # Disable maximize

            # Instantiate CanConnection class from can_connect.py
            rs_connection = CanConnection(can_window)
            rs_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            # Show the main window again when can_window is closed
            can_window.protocol("WM_DELETE_WINDOW", lambda: self.on_rs_window_close(can_window,rs_connection))

            can_window.grab_set()

    def on_rs_window_close(self, window, rs_connection):
        self.rs_connected = rs_connection.get_connection_status()
        self.update_colors()
        window.grab_release()  # Release the grab set by can_window
        window.destroy()  # Destroy the Toplevel window
        self.master.attributes("-disabled", False)  # Re-enable the main window
        self.master.lift()  # Bring the main window to the front
        #self.master.focus_force()

    def handle_can_battery_btn_click(self, event):
        # if self.rs_connected:
        #     messagebox.showerror("ERROR!","RS232/RS422 already connected")
        # elif not self.can_connected:
        #     messagebox.showerror("ERROR!","CAN not connected")
        # else:
        print("CAN Battery connect")
        self.master.withdraw()
        can_battery_info = BatteryInfo()
        can_battery_info.protocol("WM_DELETE_WINDOW", lambda: self.on_can_battery_window_close(can_battery_info))
        can_battery_info.mainloop()

    def on_can_battery_window_close(self, window):
        window.destroy()  # Destroy the Toplevel window
        self.master.deiconify()  # Show the main window again
        self.master.attributes("-disabled", False)  # Re-enable the main window
        self.master.lift()
        print("CAN Battery window closed")

    def handle_rs_battery_btn_click(self, event):
        if self.can_connected:
            messagebox.showerror("ERROR!","CAN already connected")
        elif not self.rs_connected:
            messagebox.showerror("ERROR!","RS232/RS422 not connected")
        else:
            print("RS232 Battery connect")
