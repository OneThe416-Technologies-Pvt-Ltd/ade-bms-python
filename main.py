import os
import tkinter as tk
from tkinter import filedialog, PhotoImage
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
from tkinter import messagebox
import ttkbootstrap as ttk
from customtkinter import CTk, CTkFrame, CTkButton, CTkLabel, CTkEntry, set_appearance_mode
from helpers.methods import load_and_resize_image, create_image_button
from PCAN_API.custom_pcan_methods import pcan_write
from gui.can_connect import CanConnection

# Get the base directory of the script
base_dir = os.path.dirname(__file__)

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("ADE BMS")
        self.master.geometry("1000x600")
        self.master.resizable(False, False)

        # Set custom icon
        icon_path = os.path.join(base_dir, 'assets/logo/drdo_icon.ico')
        self.master.iconbitmap(icon_path)

        self.can_connected = False
        self.rs_connected = False

        # Load background image
        self.background_image_path = os.path.join(base_dir, 'assets/images/bg_main1.png')
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

        # Battery details content
        self.battery_info_frame = None
        self.create_battery_info_frame()

    def create_buttons(self):
        images = {
            'can': os.path.join(base_dir, 'assets/images/pcan_btn_img.png'),
            'rs': os.path.join(base_dir, 'assets/images/moxa_rs232_btn_img.png'),
            'can_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png'),
            'rs_battery': os.path.join(base_dir, 'assets/images/brenergy_battery_img.png')
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

    def create_battery_info_frame(self):
        self.battery_info_frame = CTkFrame(self.master, fg_color="#ffffff")
        self.battery_info_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.side_menu = CTkFrame(self.battery_info_frame, corner_radius=10, fg_color="#009688")
        self.side_menu.pack(side="left", fill="y", padx=15, pady=10)

        self.content_area = CTkFrame(self.battery_info_frame, corner_radius=10, fg_color="#33b4b4")
        self.content_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.create_side_menu_options()

        self.active_button = None
        self.show_content("Dashboard")  # Show Dashboard by default

    def create_side_menu_options(self):
        self.buttons = []

        # Home button
        home_button = self.create_custom_button("Home", self.go_home)
        self.buttons.append(("Home", home_button))
        home_button.pack(pady=5, padx=5)

        # Dashboard button
        dashboard_button = self.create_custom_button("Dashboard", lambda: self.show_content("Dashboard"))
        self.buttons.append(("Dashboard", dashboard_button))
        dashboard_button.pack(pady=5, padx=5)

        options = [
            # ("Temperature", lambda: self.show_content("Temperature")),
            ("Current", lambda: self.show_content("Current")),
            ("Capacity", lambda: self.show_content("Capacity")),
            ("Voltage", lambda: self.show_content("Voltage")),
            ("Info", lambda: self.show_content("Info")),
        ]
        for text, command in options:
            button = self.create_custom_button(text, command)
            self.buttons.append((text, button))
            button.pack(pady=5, padx=5)

        device_name = "device name"
        serial_number = 1234
        manufacture_date = 122022
        manufacturer_name = "manufacturer_name"
        battery_status = "Good"
        cycle_count =20
        design_capacity = 200
        design_voltage = 200

        fields = [
            f"Device Name: {device_name}",
            f"Serial Number: {serial_number}",
            f"Manufacture Date: {manufacture_date}",
            f"Manufacturer Name: {manufacturer_name}",
            f"Battery Status: {battery_status}",
            f"Cycle Count: {cycle_count}",
            f"Design Capacity: {design_capacity}",
            f"Design Voltage: {design_voltage}"
        ]
        
        for field in reversed(fields):
            label = CTkLabel(self.side_menu, text=field, font=("Palatino Linotype", 12))
            label.pack(side="bottom", pady=4, padx=4)


    def create_initial_content(self):
        self.content_frame = CTkFrame(self.content_area, corner_radius=10, fg_color="#ffffff")
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=10)

        title_label = CTkLabel(
            self.content_frame,
            text="Other - Info",
            font=("Palatino Linotype", 20, "bold"),
            text_color="#333333"
        )
        title_label.pack(pady=20)

        info_frame = CTkFrame(self.content_frame, corner_radius=10, fg_color="#f8f8f8", border_color="#d0d0d0", border_width=2)
        info_frame.pack(padx=10, pady=10, anchor="n")

        at_rate = 100
        at_rate_time_to_full = 120
        at_rate_time_to_empty = 90
        at_rate_ok = 1
        at_rate_ok_text = "Yes" if at_rate_ok == 1 else "No"
        rel_state_of_charge = 80
        abs_state_of_charge = 75
        run_time_to_empty = 60
        avg_time_to_empty = 70
        avg_time_to_full = 50
        max_error = 20

        # Update labels with variables
        labels = [
            f"At Rate: {at_rate} mA",
            f"At Rate Time To Full: {at_rate_time_to_full} mins",
            f"At Rate Time To Empty: {at_rate_time_to_empty} mins",
            f"At Rate OK: {at_rate_ok_text}",
            f"Rel State of Charge: {rel_state_of_charge} %",
            f"Absolute State of Charge: {abs_state_of_charge} %",
            f"Run Time To Empty: {run_time_to_empty} mins",
            f"Avg Time To Empty: {avg_time_to_empty} mins",
            f"Avg Time To Full: {avg_time_to_full} mins",
            f"Max Error: {max_error} %"
        ]

        for label_text in labels:
            label = CTkLabel(info_frame, text=label_text, font=("Palatino Linotype", 14), text_color="#333333")
            label.pack(padx=10, pady=10, anchor="w")

    def get_bootstyle(self, value):
        if value < 20:
            return "success"
        elif value < 50:
            return "info"
        elif value < 80:
            return "warning"
        else:
            return "danger"

    def create_dashboard(self, parent_frame):
        temp = 50.5
        current = 50
        capacity = 70
        voltage = 240
        meter_info = {
            "Temperature": {"amountused": temp, "subtext": "Temperature", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Current": {"amountused": current, "subtext": "Current", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Capacity": {"amountused": capacity, "subtext": "Capacity", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
            "Voltage": {"amountused": voltage, "subtext": "Voltage", "textright": "V", "bootstyle": "dark","stripethickness":5,"amounttotal":300}
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#ffffff", bg_color="#33b4b4")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # def download_data():
        #     # Implement your download functionality here
        #     file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        #     if file_path:
        #         with open(file_path, 'w') as file:
        #             file.write("Download the data you want here")

        image_path = os.path.join(os.path.dirname(__file__), "assets/images/download_icon.png")
        if not os.path.isfile(image_path):
            raise FileNotFoundError(f"Image file not found: {image_path}")

        # Load the download icon image
        download_icon = PhotoImage(file=image_path)

        # Create and place the download button in the top right corner
        download_button = CTkButton(dashboard_frame, image=download_icon, fg_color="#ff6600", text="", width=30, height=30)
        download_button.place(relx=1.0, rely=0.0, anchor="ne")
        row, col = 0, 0
        for idx, (name, config) in enumerate(meter_info.items()):
            if idx % 2 == 0:
                row += 1
                col = 0
            else:
                col += 1

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#f8f8f8")
            meter_frame.grid(row=row, column=col, padx=60, pady=10)

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"), text_color="#333333")
            title_label.pack(pady=(10, 10))

            bootstyle = self.get_bootstyle(config["amountused"])

            meter = ttk.Meter(
                master=meter_frame,
                metersize=200,
                amountused=config["amountused"],
                meterthickness=10,
                metertype="semi",
                subtext=config["subtext"],
                textright=config["textright"],
                amounttotal=config["amounttotal"],
                bootstyle=bootstyle,
                stripethickness=config["stripethickness"],
                subtextfont='-size 10'
            )
            meter.pack()

    # def create_temperature_content(self, parent_frame):
    #     temp = 70.6
    #     content_frame = CTkFrame(self.content_area, corner_radius=10, fg_color="#ffffff")
    #     content_frame.pack(fill="both", expand=True, padx=20, pady=10)

    #     title_label = CTkLabel(content_frame, text="Temperature", font=("Helvetica", 16, "bold"), text_color="#333333")
    #     title_label.pack(pady=(20, 10))

    #     meter_frame = CTkFrame(content_frame, corner_radius=10, fg_color="#ffffff")  # Set fg_color to match the background
    #     meter_frame.pack(padx=10, pady=10, anchor="n")

    #     meter = ttk.Meter(
    #         master=meter_frame,
    #         metersize=200,
    #         amountused=temp,
    #         amounttotal=100,
    #         meterthickness=10,
    #         metertype="semi",
    #         subtext="Temperature",
    #         textright="°C",
    #     )
    #     meter.pack(pady=20)

    def create_current_content(self, parent_frame):
        avg_current = 50.5
        current = 50
        charging_current = 70
        meter_info = {
            "Avg Current": {"amountused": avg_current, "subtext": "Avg Current", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Current": {"amountused": current, "subtext": "Current", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Charging Current": {"amountused": charging_current, "subtext": "Charging Current", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#ffffff", bg_color="#33b4b4")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        row, col = 0, 0

        for idx, (name, config) in enumerate(meter_info.items()):
            if idx == 0:
                row, col = 0, 0
                padding_x = 30
            elif idx == 1:
                row, col = 0, 1 
                padding_x = 30
            elif idx == 2:
                row, col = 1, 0
                padding_x = 60

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#d0d0d0")
            meter_frame.grid(row=row, column=col, padx=padding_x, pady=10)

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"), text_color="#333333")
            title_label.pack(pady=(10, 10))

            meter = ttk.Meter(
                master=meter_frame,
                metersize=200,
                amountused=config["amountused"],
                meterthickness=10,
                metertype="semi",
                subtext=config["subtext"],
                textright=config["textright"],
                amounttotal=config["amounttotal"],
                bootstyle=config["bootstyle"],
                stripethickness=config["stripethickness"],
                subtextfont='-size 10'
            )
            meter.pack()

    def create_capacity_content(self, parent_frame):
        remaing_capacity = 50.5
        full_charge_capacity = 50
        design_capacity = 70

        meter_info = {
            "Remaing Capacity": {"amountused": remaing_capacity, "subtext": "Remaing Capacity", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Full Charge Capacity": {"amountused": full_charge_capacity, "subtext": "Full Charge Capacity", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Design Capacity": {"amountused": design_capacity, "subtext": "Design Capacity", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#ffffff", bg_color="#33b4b4")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        row, col = 0, 0

        for idx, (name, config) in enumerate(meter_info.items()):
            if idx == 0:
                row, col = 0, 0
                padding_x = 30
            elif idx == 1:
                row, col = 0, 1 
                padding_x = 30
            elif idx == 2:
                row, col = 1, 0
                padding_x = 60

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#d0d0d0")
            meter_frame.grid(row=row, column=col, padx=padding_x, pady=10)

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"),text_color="#333333")
            title_label.pack(pady=(10, 10))

            meter = ttk.Meter(
                master=meter_frame,
                metersize=200,
                amountused=config["amountused"],
                meterthickness=10,
                metertype="semi",
                subtext=config["subtext"],
                textright=config["textright"],
                amounttotal=config["amounttotal"],
                bootstyle=config["bootstyle"],
                stripethickness=config["stripethickness"],
                subtextfont='-size 10'
            )
            meter.pack()

    def create_voltage_content(self, parent_frame):
        voltage = 50.5
        charging_voltage = 50
        design_voltage = 70

        meter_info = {
            "Voltage": {"amountused": voltage, "subtext": "Voltage", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Charging Voltage": {"amountused": charging_voltage, "subtext": "Charging Voltage", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Design Voltage": {"amountused": design_voltage, "subtext": "Design Voltage", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#ffffff", bg_color="#33b4b4")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        row, col = 0, 0

        for idx, (name, config) in enumerate(meter_info.items()):
            if idx == 0:
                row, col = 0, 0
                padding_x = 30
            elif idx == 1:
                row, col = 0, 1 
                padding_x = 30
            elif idx == 2:
                row, col = 1, 0
                padding_x = 60

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#d0d0d0")
            meter_frame.grid(row=row, column=col, padx=padding_x, pady=10)

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"),text_color="#333333")
            title_label.pack(pady=(10, 10))

            meter = ttk.Meter(
                master=meter_frame,
                metersize=200,
                amountused=config["amountused"],
                meterthickness=10,
                metertype="semi",
                subtext=config["subtext"],
                textright=config["textright"],
                amounttotal=config["amounttotal"],
                bootstyle=config["bootstyle"],
                stripethickness=config["stripethickness"],
                subtextfont='-size 10'
            )
            meter.pack()

    def show_content(self, content_name):
        for widget in self.content_area.winfo_children():
            widget.destroy()

        for text, button in self.buttons:
            if text == content_name:
                button.configure(fg_color="#003f6b", hover_color="#001f3b")
                self.active_button = button
            else:
                button.configure(fg_color="#007acc", hover_color="#005fa3")

        if content_name == "Info":
            self.create_initial_content()
        elif content_name == "Dashboard":
            self.create_dashboard(self.content_area)
        elif content_name == "Temperature":
            self.create_temperature_content(self.content_area)
        elif content_name == "Current":
           self.create_current_content(self.content_area)
        elif content_name == "Capacity":
           self.create_capacity_content(self.content_area)
        elif content_name == "Voltage":
            self.create_voltage_content(self.content_area)


    def create_custom_button(self, text, command):
        button = CTkButton(self.side_menu, text=text, command=lambda: self.on_button_click(text, command), corner_radius=10, fg_color="#007acc", hover_color="#005fa3")
        return button

    def on_button_click(self, text, command):
        if self.active_button:
            self.active_button.configure(fg_color="#007acc", hover_color="#005fa3")

        for t, button in self.buttons:
            if t == text:
                button.configure(fg_color="#003f6b", hover_color="#001f3b")
                self.active_button = button
        command()

    def show_battery_info(self, event=None):
        self.canvas.pack_forget()
        self.battery_info_frame.pack(fill="both", expand=True)
        self.show_content("Dashboard")

    def show_main_window(self):
        self.battery_info_frame.pack_forget()
        self.canvas.pack(fill="both", expand=True)

    def go_home(self):
        self.show_main_window()

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

            # Calculate the center position
            can_window_x = main_window_x + (main_window_width - can_window_width) // 2
            can_window_y = main_window_y + (main_window_height - can_window_height) // 2

            # Create a Toplevel window for CAN connection settings
            can_window = tk.Toplevel(self.master)
            can_window.title("CAN Connection Settings")
            can_window.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)  # Disable maximize

            # Instantiate CanConnection class from can_connect.py
            can_connection = CanConnection(can_window)
            can_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            # Show the main window again when can_window is closed
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
            messagebox.showerror("ERROR!", "CAN already connected")
        else:
            print("RS232 connect")
            main_window_x = self.master.winfo_rootx()
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
            can_window.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)  # Disable maximize

            # Instantiate CanConnection class from can_connect.py
            rs_connection = CanConnection(can_window)
            rs_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            # Show the main window again when can_window is closed
            can_window.protocol("WM_DELETE_WINDOW", lambda: self.on_rs_window_close(can_window, rs_connection))

            can_window.grab_set()

    def on_rs_window_close(self, window, rs_connection):
        self.rs_connected = rs_connection.get_connection_status()
        self.update_colors()
        window.grab_release()  # Release the grab set by can_window
        window.destroy()  # Destroy the Toplevel window
        self.master.attributes("-disabled", False)  # Re-enable the main window
        self.master.lift()  # Bring the main window to the front
        # self.master.focus_force()

    def handle_rs_battery_btn_click(self, event):
        if self.can_connected:
            messagebox.showerror("ERROR!", "CAN already connected")
        elif not self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 not connected")
        else:
            print("RS232 Battery connect")

    def handle_can_battery_btn_click(self, event):
        if self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 already connected")
        elif not self.can_connected:
            messagebox.showerror("ERROR!", "CAN not connected")
        else:
            # result = pcan_write('serial_no') 
            # print("can serial value",result)
            self.show_battery_info()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
