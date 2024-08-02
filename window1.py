import tkinter as tk
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
from tkinter import messagebox
import ttkbootstrap as ttk
from customtkinter import CTk, CTkFrame, CTkButton, CTkLabel, CTkEntry
from helpers.methods import load_and_resize_image, create_image_button
from gui.can_connect import CanConnection

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("ADE BMS")
        self.master.geometry("1000x600")
        self.master.resizable(False, False)

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

        # Battery details content
        self.battery_info_frame = None
        self.create_battery_info_frame()

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

    def create_battery_info_frame(self):
        self.battery_info_frame = CTkFrame(self.master, fg_color="#33b4b4", corner_radius=20)
        self.battery_info_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.side_menu = CTkFrame(self.battery_info_frame, corner_radius=0, fg_color="#009688")
        self.side_menu.pack(side="left", fill="y", padx=10, pady=10)

        self.separator = CTkFrame(self.battery_info_frame, width=2, fg_color="#33b4b4")
        self.separator.pack(side="left", fill="y", padx=5)

        self.content_area = CTkFrame(self.battery_info_frame, corner_radius=0, fg_color="#33b4b4")
        self.content_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.create_side_menu_options()

        self.active_button = None
        self.show_content("Dashboard")  # Show Dashboard by default

    def create_side_menu_options(self):
        self.buttons = []

        # Home button
        home_button = self.create_custom_button("Home", self.go_home)
        self.buttons.append(("Home", home_button))
        home_button.pack(pady=10, padx=5)

        # Dashboard button
        dashboard_button = self.create_custom_button("Dashboard", lambda: self.show_content("Dashboard"))
        self.buttons.append(("Dashboard", dashboard_button))
        dashboard_button.pack(pady=10, padx=5)

        # Device Info button
        # device_info_button = self.create_custom_button("Battery Info", lambda: self.show_content("Device Info"))
        # self.buttons.append(("Device Info", device_info_button))
        # device_info_button.pack(pady=5, padx=5)

        options = [
            ("Temperature", lambda: self.show_content("Temperature")),
            ("Current", lambda: self.show_content("Current")),
            ("Capacity", lambda: self.show_content("Capacity")),
            ("Voltage", lambda: self.show_content("Voltage")),
            ("Info", lambda: self.show_content("Info")),
        ]
        for text, command in options:
            button = self.create_custom_button(text, command)
            self.buttons.append((text, button))
            button.pack(pady=6, padx=6)

        fields = [
            "Device Name: Example Device",
            "Serial Number: 123456789",
            "Manufacture Date: 2023-01-01",
            "Manufacturer Name: Example Manufacturer",
            "Battery Status: Good",
            "Cycle Count: 150",
            "Design Capacity: 1.0",
            "Design Voltage: 2.0"
        ]

        # fields = [
        #     f"Device Name: {device_name}",
        #     f"Serial Number: {serial_number}",
        #     f"Manufacture Date: {manufacture_date}",
        #     f"Manufacturer Name: {manufacturer_name}",
        #     f"Battery Status: {battery_status}",
        #     f"Cycle Count: {cycle_count}",
        #     f"Design Capacity: {design_capacity}",
        #     f"Design Voltage: {design_voltage}"
        # ]
        
        for field in reversed(fields):
            label = CTkLabel(self.side_menu, text=field, font=("Palatino Linotype", 12), text_color="#000000", fg_color="transparent")
            label.pack(side="bottom", pady=4, padx=4)


    def create_initial_content(self):
        self.content_frame = CTkFrame(self.content_area, corner_radius=10, fg_color="#ffffff")
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        title_label = CTkLabel(
            self.content_frame,
            text="Other - Info",
            font=("Palatino Linotype", 20, "bold"),
            text_color="#333333"
        )
        title_label.pack(pady=20)

        info_frame = CTkFrame(self.content_frame, corner_radius=10, fg_color="#f8f8f8", border_color="#d0d0d0", border_width=2)
        info_frame.pack(padx=40, pady=40, anchor="n")

        labels = [
            "At Rate: 100mA",
            "At Rate Time To Full: 120 mins",
            "At Rate Time To Empty: 90 mins",
            "At Rate OK: 1",
            "Rel State of Charge: 80 %",
            "Absolute State of Charge: 75 %",
            "Run Time To Empty: 60 mins",
            "Avg Time To Empty: 70 mins",
            "Avg Time To Full: 50 mins"
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

        dashboard_frame = CTkFrame(parent_frame, corner_radius=0, fg_color="#ffffff", bg_color="#33b4b4")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        row, col = 0, 0
        for idx, (name, config) in enumerate(meter_info.items()):
            if idx % 2 == 0:
                row += 1
                col = 0
            else:
                col += 1

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#f8f8f8")
            meter_frame.grid(row=row, column=col, padx=60, pady=10)

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"))
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

    def create_temperature_content(self, parent_frame):
        temp = 70.6
        content_frame = CTkFrame(self.content_area, corner_radius=10, fg_color="#ffffff")
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        title_label = CTkLabel(content_frame, text="Temperature", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(20, 10))

        meter_frame = CTkFrame(content_frame, corner_radius=10, fg_color="#ffffff")  # Set fg_color to match the background
        meter_frame.pack(padx=20, pady=20, anchor="n")

        meter = ttk.Meter(
            master=meter_frame,
            metersize=200,
            amountused=temp,
            amounttotal=100,
            meterthickness=10,
            metertype="semi",
        )
        meter.pack(pady=20)

        # def validate_numeric_input(P):
        #     if str.isdigit(P) or P == "":
        #         return True
        #     else:
        #         return False

        # vcmd = root.register(validate_numeric_input)

        # input_frame = CTkFrame(content_frame, fg_color="#ffffff")
        # input_frame.pack(pady=10)

        # max_label = CTkLabel(input_frame, text="Max Value", font=("Helvetica", 12))
        # max_label.grid(row=0, column=1, padx=(10, 0))
        # self.max_entry = CTkEntry(input_frame, validate="key", validatecommand=(vcmd, '%P'))
        # self.max_entry.grid(row=1, column=1, padx=(10, 0))

        # set_button = CTkButton(input_frame, text="Set Min and Max", command=self.set_min_max)
        # set_button.grid(row=1, column=2, padx=(10, 0))  # Place button to the right of max_entry

    def create_current_content(self, parent_frame):
        avg_current = 50.5
        current = 50
        charging_current = 70
        meter_info = {
            "Avg Current": {"amountused": avg_current, "subtext": "Avg Current", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Current": {"amountused": current, "subtext": "Current", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Charging Current": {"amountused": charging_current, "subtext": "Charging Current", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=0, fg_color="#ffffff")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=5)

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

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"))
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

        dashboard_frame = CTkFrame(parent_frame, corner_radius=0, fg_color="#ffffff")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=5)

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

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"))
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

        dashboard_frame = CTkFrame(parent_frame, corner_radius=0, fg_color="#ffffff")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=5)

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

            title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"))
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
    
    # def create_temperature_content(self, parent_frame):
    #     temp = 50.5
    #     current = 50
    #     capacity = 70
    #     voltage = 240
    #     meter_info = {
    #         "Temperature": {"amountused": temp, "subtext": "Temperature", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
    #         "Current": {"amountused": current, "subtext": "Current", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
    #         "Capacity": {"amountused": capacity, "subtext": "Capacity", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
    #         "Voltage": {"amountused": voltage, "subtext": "Voltage", "textright": "V", "bootstyle": "dark","stripethickness":5,"amounttotal":300}
    #     }

    #     dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#f8f8f8")
    #     dashboard_frame.pack(fill="both", expand=True, padx=120, pady=5)

    #     row, col = 0, 0
    #     for idx, (name, config) in enumerate(meter_info.items()):
    #         if idx % 2 == 0:
    #             row += 1
    #             col = 0
    #         else:
    #             col += 1

    #         meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#ffffff", border_width=0, border_color="#d0d0d0")
    #         meter_frame.grid(row=row, column=col, padx=30, pady=10)

    #         title_label = CTkLabel(meter_frame, text=f"{name}", font=("Helvetica", 16, "bold"))
    #         title_label.pack(pady=(20, 20))

    #         bootstyle = self.get_bootstyle(config["amountused"])

    #         meter = ttk.Meter(
    #             master=meter_frame,
    #             metersize=200,
    #             amountused=config["amountused"],
    #             meterthickness=10,
    #             metertype="semi",
    #             subtext=config["subtext"],
    #             textright=config["textright"],
    #             amounttotal=config["amounttotal"],
    #             bootstyle=bootstyle,
    #             stripethickness=config["stripethickness"],
    #             subtextfont='-size 10'
    #         )
    #         meter.pack()


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
            can_window.iconbitmap('asserts/logo/drdo_icon.ico')
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
            can_window.iconbitmap('asserts/logo/drdo_icon.ico')
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
            self.show_battery_info()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
