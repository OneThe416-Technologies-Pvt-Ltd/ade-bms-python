import os
import tkinter as tk
import csv
import time
import threading
from datetime import datetime
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
from tkinter import messagebox
import ttkbootstrap as ttk
import customtkinter as CTk
from customtkinter import CTkFrame, CTkButton, CTkLabel, CTkComboBox
from helpers.methods import load_and_resize_image, create_image_button
from PCAN_API.custom_pcan_methods import device_data, update_device_data, label_mapping,pcan_write_control
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
                self.image_refs[key] = image_tk 
                position = ((x2 + x1) // 2, (y2 + y1) // 2)
                item_id = create_image_button(self.canvas, image_tk, position, handlers[key])
                self.canvas_items[key] = item_id
    

    def create_battery_info_frame(self):
        self.battery_info_frame = CTkFrame(self.master, fg_color="#e2eff1")
        self.battery_info_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.side_menu = CTkFrame(self.battery_info_frame, corner_radius=10, fg_color="#3d5af1")
        self.side_menu.pack(side="left", fill="y", padx=15, pady=10)

        self.content_area = CTkFrame(self.battery_info_frame, corner_radius=10, fg_color="#3d5af1")
        self.content_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.create_side_menu_options()

        self.active_button = None
        self.show_content("Dashboard")


    def create_side_menu_options(self):
        self.buttons = []

        # Home button
        home_button = self.create_custom_button("Home", self.go_home)
        self.buttons.append(("Home", home_button))
        home_button.pack(pady=10, padx=20)

        # Dashboard button
        dashboard_button = self.create_custom_button("Dashboard", lambda: self.show_content("Dashboard"))
        self.buttons.append(("Dashboard", dashboard_button))
        dashboard_button.pack(pady=10, padx=20)

        options = [
            ("Info", lambda: self.show_content("Info")),
            ("Controls", lambda: self.show_content("Controls")),
        ]
        for text, command in options:
            button = self.create_custom_button(text, command)
            self.buttons.append((text, button))
            button.pack(pady=10, padx=20)


    def table_content(self):
        content_frame = CTk.CTkFrame(self.content_area, corner_radius=10, fg_color="#e2eff1")
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Add a heading label
        heading_label = CTk.CTkLabel(content_frame,
                                     text="Battery Information",
                                     font=("Helvetica", 20, "bold"),
                                     text_color="#333333")
        heading_label.pack(pady=10)

        def on_checkbox_change():
            if self.checkbox_var.get():  # If checkbox is selected (True)
                update_device_data(continuous=True)
            else:
                update_device_data(continuous=False)

        row_frame = CTk.CTkFrame(content_frame, corner_radius=10, fg_color="#e2eff1")
        row_frame.pack(fill="x", pady=5)

        # Add the checkbox to the row frame
        self.checkbox_var = CTk.BooleanVar()  # Variable to store the checkbox state
        checkbox = CTk.CTkCheckBox(row_frame,
                                   text="Auto Refresh",
                                   variable=self.checkbox_var,
                                   onvalue=True,
                                   offvalue=False,
                                   command=on_checkbox_change,
                                   text_color="#333333")
        checkbox.pack(side="left", padx=5)

        # Add the button to the row frame
        update_button = CTk.CTkButton(row_frame,
                                      text="Refresh Now",
                                      command=lambda: update_device_data(continuous=False))
        update_button.pack(side="left", padx=5)
        # Add the button to the row frame
        start_logging_button = CTk.CTkButton(row_frame,
                                      text="Start Logging",
                                      command=self.start_logging
                                      )
        start_logging_button.pack(side="left", padx=5) 
        # Add the button to the row frame
        stop_logging_button = CTk.CTkButton(row_frame,
                                      text="Stop Logging",
                                    #   command=lambda: update_device_data(continuous=False)
                                      )
        stop_logging_button.pack(side="left", padx=5)
        # Create a frame to hold the Treeview
        table_frame = CTk.CTkFrame(content_frame, corner_radius=10, fg_color="#e2eff1")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Style the Treeview widget
        style = ttk.Style()
        style.configure("Treeview",
                        font=("Helvetica", 12),  # Set font for rows
                        rowheight=25,  # Set row height
                        background="#f5f5f5",  # Set background color
                        fieldbackground="#f5f5f5",  # Set field background color
                        borderwidth=1,  # Simulate border width
                        foreground="#333333")  # Set text color
        style.configure("Treeview.Heading",
                        font=("Helvetica", 14, "bold"),  # Set font for headings
                        background="#007acc",  # Set background color for heading
                        borderwidth=1,  # Simulate border width
                        foreground="white")  # Set text color for heading
        style.map("Treeview",
                  background=[('selected', '#007acc')],  # Set selected row color
                  foreground=[('selected', 'white')])  # Set selected row text color

        # Create the Treeview widget with centered text
        tree = ttk.Treeview(table_frame, columns=('Label', 'Value'), show='headings', style="Treeview")
        tree.heading('Label', text='Field', anchor='center')
        tree.heading('Value', text='Value', anchor='center')
        tree.column('Label', width=200, anchor='center')
        tree.column('Value', width=300, anchor='center')

        # Add data to the Treeview
        for index, (key, value) in enumerate(device_data.items()):
            label = label_mapping.get(key, key)
            bg_color = "#f0f0f0" if index % 2 == 0 else "#e2eff1"
            tree.insert('', 'end', values=(label, value), tags=('evenrow' if index % 2 == 0 else 'oddrow',))

        # Apply tag configurations for alternating row colors
        tree.tag_configure('evenrow', background="#f0f0f0")
        tree.tag_configure('oddrow', background="#e2eff1")


        # Add a scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        tree.pack(fill="both", expand=True)


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
        meter_info = {
            "Temperature": {"amountused": device_data['temperature'], "subtext": "Temperature", "textright": "°C", "bootstyle": "danger","stripethickness":0,"amounttotal":100},
            "Current": {"amountused": device_data['current'], "subtext": "Current", "textright": "mA", "bootstyle": "info","stripethickness":10,"amounttotal":100},
            "Capacity": {"amountused": device_data['remaining_capacity'], "subtext": "Capacity", "textright": "%", "bootstyle": "warning","stripethickness":8,"amounttotal":100},
            "Voltage": {"amountused": device_data['voltage'], "subtext": "Voltage", "textright": "V", "bootstyle": "dark","stripethickness":5,"amounttotal":300}
        }

        dashboard_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#e2eff1", bg_color="#3d5af1")
        dashboard_frame.pack(fill="both", expand=True, padx=10, pady=10)

        row, col = 0, 0
        for idx, (name, config) in enumerate(meter_info.items()):
            if idx % 2 == 0:
                row += 1
                col = 0
            else:
                col += 1

            meter_frame = CTkFrame(dashboard_frame, corner_radius=10, fg_color="#e2eff1")
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


    def control_option(self, parent_frame):
        control_option_frame = CTkFrame(parent_frame, corner_radius=10, fg_color="#e2eff1", bg_color="#33b4b4")
        control_option_frame.pack(fill="both", expand=True, padx=10, pady=10)

        heading_label = CTk.CTkLabel(control_option_frame,
                                     text="Controls Information",
                                     font=("Helvetica", 20, "bold"),
                                     text_color="#333333")
        heading_label.pack(pady=10)

        self.dropdown_values = {
                "Charge FET ON and Discharge FET OFF": "charge_on",
                "Discharge FET ON and Charge FET OFF": "discharge_on",
                "Both FET ON": "both_on",
                "Both FET OFF": "both_off"
            }

        # Dropdown menu with four values
        self.dropdown = CTkComboBox(control_option_frame, values=list(self.dropdown_values.keys()), width=250)
        self.dropdown.pack(pady=10)

        # Button that calls a method with the selected dropdown value
        action_button = CTkButton(control_option_frame, text="Execute", command=self.execute_action)
        action_button.pack(pady=10)
  

    def execute_action(self):
        selected_display_value = self.dropdown.get()
        # Get the actual value from the dictionary
        actual_value = self.dropdown_values[selected_display_value]
        pcan_write_control(actual_value)

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
            self.table_content()
        elif content_name == "Dashboard":
            self.create_dashboard(self.content_area)
        elif content_name == "Controls":
            self.control_option(self.content_area)


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
        update_device_data(continuous=False)
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
            'can': (80, 70),
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

            can_window_width = 400
            can_window_height = 400

            can_window_x = main_window_x + (main_window_width - can_window_width) // 2
            can_window_y = main_window_y + (main_window_height - can_window_height) // 2

            can_window = tk.Toplevel(self.master)
            can_window.title("RS Connection Settings")
            can_window.iconbitmap(os.path.join(base_dir, 'assets/logo/drdo_icon.ico'))
            can_window.geometry(f"{can_window_width}x{can_window_height}+{can_window_x}+{can_window_y}")
            can_window.resizable(False, False)

            rs_connection = CanConnection(can_window)
            rs_connection.pack(fill="both", expand=True)

            self.master.attributes("-disabled", True)
            can_window.protocol("WM_DELETE_WINDOW", lambda: self.on_rs_window_close(can_window, rs_connection))

            can_window.grab_set()


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
        if self.rs_connected:
            messagebox.showerror("ERROR!", "RS232/RS422 already connected")
        elif not self.can_connected:
            messagebox.showerror("ERROR!", "CAN not connected")
        else:
            self.show_battery_info()


    def start_logging():
        # Create a CSV file with a unique name based on the current timestamp
        filename = f"device_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

        # Open the CSV file in write mode
        with open(filename, 'w', newline='') as csvfile:
            fieldnames = list(device_data.keys())
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            # Write the header
            writer.writeheader()

            def log_data():
                while True:
                    # Fetch and update device data
                    update_device_data(continuous=False)

                    # Write the current data to the CSV file
                    writer.writerow(device_data)

                    # Wait for 5 seconds before logging the next entry
                    time.sleep(5)

            # Start logging in a new thread
            logging_thread = threading.Thread(target=log_data)
            logging_thread.daemon = True  # This will ensure the thread stops when the main program exits
            logging_thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
