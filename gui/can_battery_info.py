#can_battery_info.py
import customtkinter as ctk
import tkinter as tk
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
import datetime
import threading
import time
import asyncio
from tkinter import messagebox
from openpyxl import Workbook
from pcan_api.custom_pcan_methods import *
from load_api.chroma_api import *


class CanBatteryInfo:
    def __init__(self, master, main_window=None):
        self.master = master
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button
        self.master.resizable(True, False)
        # self.center_window(1200, 600)  # Center the window with specified dimensions

        # Calculate one-fourth of the main window width for the side menu
        self.side_menu_width_ratio = 0.20  # 20% for side menu
        self.update_frame_sizes()  # Set initial size

        # Variables to store control states for each battery
        self.charger_control_var_battery_1 = tk.BooleanVar(value=False)  # Default to False
        self.charger_control_var_battery_2 = tk.BooleanVar(value=False)
        self.discharger_control_var_battery_1 = tk.BooleanVar(value=False)
        self.discharger_control_var_battery_2 = tk.BooleanVar(value=False)

        self.is_connected = True  # Initialize is_connected to True by default

        # Track the selected battery
        self.selected_battery = "Battery 1"

         # Default to Battery 1
        self.charger_control_var = self.charger_control_var_battery_1
        self.discharger_control_var = self.discharger_control_var_battery_1

        self.logging_active = False
        # Initialize the limited attribute to avoid AttributeError
        self.limited = False

        self.first_time_dashboard = False
        self.mode_var = tk.StringVar(value="Testing Mode")
         # Directory paths
        base_path = os.path.dirname(os.path.abspath(__file__))
        self.assets_path = os.path.join(base_path, "../assets/images/")

        style = ttk.Style() 
        # Configure the Labelframe title font to be bold
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold")) 

        self.style = ttk.Style()
        self.configure_styles()

        self.device_data = None
        self.battery_status_flags = None  

        self.device_data = device_data_battery_1
        self.battery_status_flags = battery_1_status_flags

        self.check_battery_log_for_cycle_count()

        # Create the main container frame
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True)

        # Create the side menu frame with 1/4 width of the main window
        self.side_menu_frame = ttk.Frame(self.main_frame, bootstyle="info")
        self.side_menu_frame.place(x=0, y=0, width=self.side_menu_width, relheight=1.0)

        # Create the content frame on the right side with the remaining 3/4 width
        self.content_frame = ttk.Frame(self.main_frame, bootstyle="light")
        self.content_frame.place(x=self.side_menu_width, y=0, width=self.content_frame_width, relheight=1.0)

        # Load and resize icons using Pillow
        self.disconnect_icon = self.load_icon_menu(os.path.join(self.assets_path, "disconnect_button.png"))
        self.control_icon = self.load_icon_menu(os.path.join(self.assets_path, "load_icon.png"))
        self.controls_icon = self.load_icon_menu(os.path.join(self.assets_path, "settings.png"))
        self.info_icon = self.load_icon_menu(os.path.join(self.assets_path, "info.png"))
        self.help_icon = self.load_icon_menu(os.path.join(self.assets_path, "help.png"))
        self.dashboard_icon = self.load_icon_menu(os.path.join(self.assets_path, "menu.png"))
        self.download_icon = self.load_icon_menu(os.path.join(self.assets_path, "download.png"))

        self.side_menu_heading = ttk.Label(
            self.side_menu_frame,
            text="ADE BMS",
            font=("Courier New", 18, "bold"),
            bootstyle="inverse-info",  # Optional: change color style
            anchor="center"
        )
        self.side_menu_heading.pack(fill="x", pady=(10, 20))

        self.dashboard_button =  ttk.Button(
            self.side_menu_frame,
            text=" Dashboard",
            command=lambda: self.select_button(self.dashboard_button, self.show_dashboard),
            image=self.dashboard_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.dashboard_button.pack(fill="x", pady=5)

        self.control_button = ttk.Button(
            self.side_menu_frame,
            text=" Control       ",
            command=lambda: self.select_button(self.control_button,self.show_control),
            image=self.control_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.control_button.pack(fill="x", pady=2)

        self.info_button = ttk.Button(
            self.side_menu_frame,
            text=" Info             ",
            command=lambda: self.select_button(self.info_button, self.show_info),
            image=self.info_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.info_button.pack(fill="x", pady=5)

        self.report_button = ttk.Button(
            self.side_menu_frame,
            text=" Report         ",
            command=lambda: self.select_button(self.report_button,self.show_report),
            image=self.download_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.report_button.pack(fill="x", pady=2)

        self.help_button = ttk.Button(
            self.side_menu_frame,
            text=" Help            ",
            command=lambda: self.select_button(self.help_button),
            image=self.help_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.help_button.pack(fill="x", pady=(0,10))

        # Mode Label
        self.mode_label = ttk.Label(
            self.side_menu_frame,
            bootstyle="inverse-info",
            text="Mode:",
        )
        self.mode_label.pack(pady=(110, 5),padx=(0,5))

        # Mode Dropdown
        self.mode_dropdown = ttk.Combobox(
            self.side_menu_frame,
            textvariable=self.mode_var,
            values=["Testing Mode", "Maintenance Mode"],
            state="readonly"
        )
        self.mode_dropdown.pack(padx=(5, 20))
        self.mode_dropdown.bind("<<ComboboxSelected>>", self.update_info)

        # Add Battery Selection Label (conditionally displayed)
        self.battery_selection_label = ttk.Label(
            self.side_menu_frame,
            text="Select Battery:",
            bootstyle="inverse-info"
        )
        
        # Add Battery Selection Dropdown (conditionally displayed)
        self.battery_var = tk.StringVar(value="Battery 1")  # Default to Battery 1
        self.battery_dropdown = ttk.Combobox(
            self.side_menu_frame,
            textvariable=self.battery_var,
            values=["Battery 1", "Battery 2"],
            state="readonly"
        )
        self.battery_dropdown.bind("<<ComboboxSelected>>", self.update_battery_selection)
    
        if device_data_battery_2['serial_number'] == 0:
            self.battery_selection_label.pack_forget()
            self.battery_dropdown.pack_forget()
        else:
            self.battery_selection_label.pack(pady=(10, 5))
            self.battery_dropdown.pack(padx=(5, 20), pady=(0, 10))

        # Disconnect button
        self.disconnect_button = ttk.Button(
            self.side_menu_frame,
            text=" Disconnect",
            command=lambda: self.select_button(self.disconnect_button, self.on_disconnect),
            image=self.disconnect_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="danger"
        )
        self.disconnect_button.pack(side="bottom",pady=20)

        if(self.battery_status_flags.get('charge_fet_test') == 0):
            self.charge_fet_status = True
        else:
            self.charge_fet_status = False
        if(device_data['current'] > 0):
            self.discharge_fet_status = True
        else:
            self.discharge_fet_status = False
        
        # self.auto_refresh()
        self.show_dashboard()

        # Bind the window resize event to adjust frame sizes dynamically
        self.master.bind('<Configure>', self.on_window_resize)
    
    def update_battery_selection(self, event=None):
        """Update the controls based on the selected battery."""
        self.selected_battery = self.battery_var.get()
        if self.selected_battery == "Battery 1":
            self.device_data = device_data_battery_1
            self.battery_status_flags = battery_1_status_flags
            self.charger_control_var = self.charger_control_var_battery_1
            self.discharger_control_var = self.discharger_control_var_battery_1
        elif self.selected_battery == "Battery 2":
            self.device_data = device_data_battery_2
            self.battery_status_flags = battery_2_status_flags
            self.charger_control_var = self.charger_control_var_battery_2
            self.discharger_control_var = self.discharger_control_var_battery_2
            self.check_battery_log_for_cycle_count()

        if self.selected_button == self.dashboard_button:
            self.show_dashboard()
        elif self.selected_button == self.info_button:
            self.show_info()
        elif self.selected_button == self.control_button:
            self.show_control()
        elif self.selected_button == self.report_button:
            self.show_report()


    def show_dashboard(self):
        self.clear_content_frame()

        # Create a Frame for the battery info details at the bottom
        info_frame = ttk.Labelframe(self.content_frame, text="Battery Information", bootstyle="dark", borderwidth=10, relief="solid")
        info_frame.pack(fill="x", expand=True, padx=5, pady=5)

        # Device Name
        ttk.Label(info_frame, text="Device Name:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.device_name_label = ttk.Label(info_frame, text=self.device_data.get('device_name', 'N/A'), font=("Helvetica", 10))
        self.device_name_label.pack(side="left", padx=5)

        # Serial Number
        ttk.Label(info_frame, text="Serial Number:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.serial_number_label = ttk.Label(info_frame, text=self.device_data.get('serial_number', 'N/A'), font=("Helvetica", 10))
        self.serial_number_label.pack(side="left", padx=5)

        # Manufacturer Name
        ttk.Label(info_frame, text="Manufacturer:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.manufacturer_label = ttk.Label(info_frame, text=self.device_data.get('manufacturer_name', 'N/A'), font=("Helvetica", 10))
        self.manufacturer_label.pack(side="left", padx=5)

        # Manual Cycle Count
        ttk.Label(info_frame, text="Cycle Count:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.cycle_count_label = ttk.Label(info_frame, text=self.device_data.get('cycle_count', 'N/A'), font=("Helvetica", 10))
        self.cycle_count_label.pack(side="left", padx=5)

        # Charging Status
        ttk.Label(info_frame, text="Charging Status:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)

        self.status_var = tk.StringVar()
        charging_bootstyle = 'dark'
        discharging_bootstyle = "dark"

        # Determine the status and set the corresponding style
        charging_status = self.device_data.get('charging_battery_status', 'Off').lower()
        if charging_status == "charging":
            bootstyle = "success"
            charging_bootstyle = 'success'
            self.status_var.set("charging")
        elif charging_status == "discharging":
            bootstyle = "danger"
            discharging_bootstyle = "danger"
            self.status_var.set("discharging")
        else:  # Assuming 'Idle' or any other status means the device is off
            bootstyle = "dark"
            charging_bootstyle = 'dark'
            discharging_bootstyle = "dark"
            self.status_var.set("off")

        # Create a Radiobutton to visually represent the status
        self.status_indicator = ttk.Radiobutton(
            info_frame,
            text=charging_status.capitalize(),  # Display the charging status text
            variable=self.status_var,  # Use the StringVar
            value=charging_status,  # Set the value that matches the current status
            bootstyle=bootstyle, # Apply the bootstyle based on the charging status
        )
        self.status_indicator.pack(side="left", padx=5)

        # Create a main frame to hold Charging, Discharge, and Battery Health sections
        main_frame = ttk.Frame(self.content_frame)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Charging Section
        charging_frame = ttk.Labelframe(main_frame, text="Charging", bootstyle=charging_bootstyle, borderwidth=10, relief="solid")
        charging_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=2, sticky="nsew")

        # Charging Current (Row 0, Column 0)
        charging_current_frame = ttk.Frame(charging_frame)
        charging_current_frame.grid(row=0, column=0, padx=70, pady=2, sticky="nsew", columnspan=1)
        charging_current_label = ttk.Label(charging_current_frame, text="Charging Current", font=("Helvetica", 10, "bold"))
        charging_current_label.pack(pady=(5, 5))
        charging_current_meter = ttk.Meter(
            master=charging_current_frame,
            metersize=180,
            amountused=self.device_data['charging_current'],
            meterthickness=10,
            metertype="semi",
            subtext="Charging Current",
            textright="A",
            amounttotal=120,
            bootstyle=self.get_gauge_style(self.device_data['charging_current'], "charging_current"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        charging_current_meter.pack()

        # Charging Voltage (Row 0, Column 1)
        charging_voltage_frame = ttk.Frame(charging_frame)
        charging_voltage_frame.grid(row=0, column=1, padx=70, pady=2, sticky="nsew", columnspan=1)
        charging_voltage_label = ttk.Label(charging_voltage_frame, text="Charging Voltage", font=("Helvetica", 10, "bold"))
        charging_voltage_label.pack(pady=(5, 5))
        charging_voltage_meter = ttk.Meter(
            master=charging_voltage_frame,
            metersize=180,
            amountused=self.device_data['charging_voltage'],
            meterthickness=10,
            metertype="semi",
            subtext="Charging Voltage",
            textright="V",
            amounttotal=35,
            bootstyle=self.get_gauge_style(self.device_data['charging_voltage'], "charging_voltage"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        charging_voltage_meter.pack()

        # Battery Health Section
        battery_health_frame = ttk.Labelframe(main_frame, text="Battery Health", bootstyle='dark', borderwidth=10, relief="solid")
        battery_health_frame.grid(row=0, column=2, rowspan=2, padx=5, pady=5, sticky="nsew")

        # Temperature (Top of Battery Health)
        temp_frame = ttk.Frame(battery_health_frame)
        temp_frame.pack(padx=10, pady=5, fill="both", expand=True)
        temp_label = ttk.Label(temp_frame, text="Temperature", font=("Helvetica", 10, "bold"))
        temp_label.pack(pady=(10, 10))
        temp_meter = ttk.Meter(
            master=temp_frame,
            metersize=180,
            amountused=self.device_data['temperature'],
            meterthickness=10,
            metertype="semi",
            subtext="Temperature",
            textright="°C",
            amounttotal=100,
            bootstyle=self.get_gauge_style(self.device_data['temperature'], "temperature"),
            stripethickness=8,
            subtextfont='-size 10'
        )
        temp_meter.pack()

        # Capacity (Bottom of Battery Health)
        capacity_frame = ttk.Frame(battery_health_frame)
        capacity_frame.pack(padx=10, pady=5, fill="both", expand=True)
        capacity_label = ttk.Label(capacity_frame, text="Capacity", font=("Helvetica", 10, "bold"))
        capacity_label.pack(pady=(10, 10))
        full_charge_capacity = self.device_data.get('full_charge_capacity', 0)
        remaining_capacity = self.device_data.get('remaining_capacity', 0)
        amountused = (remaining_capacity / full_charge_capacity) * 100 if full_charge_capacity else 0
        capacity_meter = ttk.Meter(
            master=capacity_frame,
            metersize=180,
            amountused=round(amountused, 1),
            meterthickness=10,
            metertype="semi",
            subtext="Capacity",
            textright="%",
            amounttotal=100,
            bootstyle=self.get_gauge_style(amountused, "capacity"),
            stripethickness=8,
            subtextfont='-size 10'
        )
        capacity_meter.pack()


        # Discharging Section
        discharging_frame = ttk.Labelframe(main_frame, text="Discharging", bootstyle=discharging_bootstyle, borderwidth=10, relief="solid")
        discharging_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky="nsew")

        def update_max_value():
            if self.check_discharge_max.get():
                self.discharge_max_value.set(450)
            else:
                self.discharge_max_value.set(100)
            self.discharging_current_meter.configure(amounttotal=self.discharge_max_value.get())

        # Discharging Current (Row 1, Column 0)
        discharging_current_frame = ttk.Frame(discharging_frame)
        discharging_current_frame.grid(row=1, column=0, padx=70, pady=2, sticky="nsew", columnspan=1)
        discharging_current_label = ttk.Label(discharging_current_frame, text="Discharging Current", font=("Helvetica", 10, "bold"))
        discharging_current_label.pack(pady=(5, 5))
        # Variable to control the max value of the discharging current gauge
        self.discharge_max_value = tk.IntVar(value=100)  # Default max value

        # Checkbox to toggle the max value of the discharging current gauge
        self.check_discharge_max = tk.BooleanVar()
        discharge_max_checkbox = ttk.Checkbutton(
            discharging_current_frame, 
            text="Set Max to 450A", 
            variable=self.check_discharge_max, 
            command=update_max_value
        )
        discharge_max_checkbox.pack(pady=(5, 10))

        # Discharging Current Gauge
        self.discharging_current_meter = ttk.Meter(
            master=discharging_current_frame,
            metersize=180,
            amountused=self.device_data['current'],  # Assuming this is the discharging current
            meterthickness=10,
            metertype="semi",
            subtext="Discharging Current",
            textright="A",
            amounttotal=self.discharge_max_value.get(),
            bootstyle=self.get_gauge_style(self.device_data['current'], "current"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        self.discharging_current_meter.pack()

        # Discharging Voltage (Row 1, Column 1)
        discharging_voltage_frame = ttk.Frame(discharging_frame)
        discharging_voltage_frame.grid(row=1, column=1, padx=70, pady=2, sticky="nsew", columnspan=1)
        discharging_voltage_label = ttk.Label(discharging_voltage_frame, text="Discharging Voltage", font=("Helvetica", 10, "bold"))
        discharging_voltage_label.pack(pady=(5, 35))
        discharging_voltage_meter = ttk.Meter(
            master=discharging_voltage_frame,
            metersize=180,
            amountused=self.device_data['voltage'],  # Assuming this is the discharging voltage
            meterthickness=10,
            metertype="semi",
            subtext="Discharging Voltage",
            textright="V",
            amounttotal=100,
            bootstyle=self.get_gauge_style(self.device_data['voltage'], "voltage"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        discharging_voltage_meter.pack()

        # Configure row and column weights to ensure even distribution
        for i in range(2):
            main_frame.grid_rowconfigure(i, weight=1)
        for j in range(3):
            main_frame.grid_columnconfigure(j, weight=1)

        # Pack the content_frame itself
        self.content_frame.pack(fill="both", expand=True)
        self.select_button(self.dashboard_button)

    def prompt_manual_cycle_count(self):
        # Create a new top-level window for input
            input_window = tk.Toplevel(self.master)
            input_window.title("Cycle Count")
        
            # Calculate the position for the window to be centered
            screen_width = self.master.winfo_screenwidth()
            screen_height = self.master.winfo_screenheight()
            window_width = 300
            window_height = 150
            x = (screen_width // 2) - (window_width // 2)
            y = (screen_height // 2) - (window_height // 2)
        
            # Set the geometry of the window directly in the center
            input_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            input_window.transient(self.master)  # Set to always be on top of the parent window
            input_window.grab_set()  # Make the input window modal
        
            # Create and pack the widgets
            label = ttk.Label(input_window, text="Enter Cycle Count:", font=("Helvetica", 12))
            label.pack(pady=10)
        
            cycle_count_var = tk.IntVar()  # Use IntVar to store integer value
            entry = ttk.Entry(input_window, textvariable=cycle_count_var)
            entry.pack(pady=5)
        
            # Function to handle OK button click
            def handle_ok():
                # Get the selected battery from the dropdown
                if self.selected_battery == "Battery 1":
                    # Update cycle count for Battery 1
                    device_data_battery_1['cycle_count'] = cycle_count_var.get()
                    update_cycle_count_in_excel(device_data_battery_1['serial_number'], device_data_battery_1['cycle_count'])
                elif self.selected_battery == "Battery 2":
                    # Update cycle count for Battery 2
                    device_data_battery_2['cycle_count'] = cycle_count_var.get()
                    update_cycle_count_in_excel(device_data_battery_2['serial_number'], device_data_battery_2['cycle_count'])

                input_window.destroy()  # Close the input window
                self.show_dashboard()
        
            # OK Button
            ok_button = ttk.Button(input_window, text="OK", command=handle_ok)
            ok_button.pack(pady=10)
        
            # Focus the input window and entry
            input_window.focus()
            entry.focus_set()
        
            # Bind the Return key to trigger the OK button
            input_window.bind('<Return>', lambda event: handle_ok())

    def show_help(self):
        self.clear_content_frame()
        label = ttk.Label(self.content_frame, text="Help", font=("Helvetica", 16))
        label.pack(pady=20)
        self.select_button(self.help_button)
    
    def show_report(self):
        self.clear_content_frame()
        latest_data = get_latest_battery_log(self.device_data.get('serial_number'))

        # Create the main container LabelFrame
        main_frame = ttk.LabelFrame(self.content_frame, text="Report", borderwidth=5)
        main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Configure the grid for centering the main_frame within content_frame
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Configure the main_frame with 4 rows and 4 columns
        main_frame.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)
        main_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        # Button to open the folder with PDF files (aligned to the top-right)
        open_folder_button = ttk.Button(main_frame, text="View Battery Reports", command=open_pdf_folder)
        open_folder_button.grid(row=0, column=11, padx=10, pady=10, sticky="ne")

        # Battery Info Frame (row 1-2, col 0-2)
        battery_info_frame = ttk.LabelFrame(main_frame, text="Battery Info", borderwidth=10)
        battery_info_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

        # Add a Label for the Project Dropdown
        ttk.Label(battery_info_frame, text="Project").grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="e")

        # Add the Combobox for the Project Dropdown
        project_options = ["TAPAS", "ARCHER", "SETB"]  # List of project options
        project_combobox = ttk.Combobox(battery_info_frame, values=project_options, state="readonly")  # Dropdown with readonly state
        project_combobox.grid(row=0, column=3, columnspan=2, padx=10, pady=10, sticky="w")
        project_combobox.set("Select a Project")

        # Add individual fields for Battery Info
        ttk.Label(battery_info_frame, text="Device Name").grid(row=1,  column=1,columnspan=2, padx=10, pady=10, sticky="e")
        device_name_entry = ttk.Entry(battery_info_frame)
        device_name_entry.insert(0, str(latest_data.get("Device Name", "")))
        device_name_entry.grid(row=1, column=3,columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(battery_info_frame, text="Serial Number").grid(row=2,  column=1,columnspan=2, padx=10, pady=10, sticky="e")
        serial_number_entry = ttk.Entry(battery_info_frame)
        serial_number_entry.insert(0, str(latest_data.get("Serial Number", "")))
        serial_number_entry.grid(row=2, column=3,columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(battery_info_frame, text="Manufacturer Name").grid(row=3,  column=1,columnspan=2, padx=10, pady=10, sticky="e")
        manufacturer_entry = ttk.Entry(battery_info_frame)
        manufacturer_entry.insert(0, str(latest_data.get("Manufacturer Name", "")))
        manufacturer_entry.grid(row=3, column=3,columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(battery_info_frame, text="Cycle Count").grid(row=4,  column=1,columnspan=2, padx=10, pady=10, sticky="e")
        cycle_count_entry = ttk.Entry(battery_info_frame)
        cycle_count_entry.insert(0, str(latest_data.get("Cycle Count", "")))
        cycle_count_entry.grid(row=4, column=3,columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(battery_info_frame, text="Battery Capacity").grid(row=5,  column=1,columnspan=2, padx=10, pady=10, sticky="e")
        capacity_entry = ttk.Entry(battery_info_frame)
        capacity_entry.insert(0, str(103))
        capacity_entry.grid(row=5, column=3,columnspan=2, padx=10, pady=10, sticky="w")
        capacity_entry.config(state="readonly")  # Disable editing

        duration, date = fetch_charging_info(self.device_data.get('serial_number'))

        # Charging Info Frame (row 1-2, col 2-4)
        charging_frame = ttk.LabelFrame(main_frame, text="Charging Info", borderwidth=10)
        charging_frame.grid(row=1, column=4, columnspan=4, padx=10, pady=10, sticky="nsew")

        ttk.Label(charging_frame, text="Duration").grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        duration_entry = ttk.Entry(charging_frame)
        duration_entry.insert(0, str(latest_data.get("Charging Duration", 50)))  # Insert the fetched duration or a default message
        duration_entry.grid(row=0, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(charging_frame, text="Date").grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        date_entry = ttk.Entry(charging_frame)
        date_entry.insert(0, str(latest_data.get("Charging Date", 50)))  # Insert the fetched date or a default message
        date_entry.grid(row=2, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        # Add individual fields for Discharging Info
        ttk.Label(charging_frame, text="OCV").grid(row=3, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        ocv_entry = ttk.Entry(charging_frame)
        ocv_entry.insert(0, str(latest_data.get("OCV", 0)))
        ocv_entry.grid(row=3, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        # # Discharging Info Frame (row 2-3, col 0-4)
        discharging_frame = ttk.LabelFrame(main_frame, text="Load Test Info", borderwidth=10)
        discharging_frame.grid(row=1, column=8, columnspan=4, padx=10, pady=10, sticky="nsew")

        # Add individual fields for Discharging Info
        ttk.Label(discharging_frame, text="Current Loaded").grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        discharging_current_loaded_entry = ttk.Entry(discharging_frame)
        discharging_current_loaded_entry.insert(0, str(latest_data.get("Current Loaded", 50)))
        discharging_current_loaded_entry.grid(row=0, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(discharging_frame, text="Duration").grid(row=1, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        discharging_duration_entry = ttk.Entry(discharging_frame)
        discharging_duration_entry.insert(0, str(latest_data.get("Discharging Duration",0)))
        discharging_duration_entry.grid(row=1, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(discharging_frame, text="0th Second Voltage").grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        zero_second_voltage_entry = ttk.Entry(discharging_frame)
        zero_second_voltage_entry.insert(0, str(latest_data.get("0th Second Voltage", 0)))
        zero_second_voltage_entry.grid(row=2, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(discharging_frame, text="5th Second Voltage").grid(row=3, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        fifth_second_voltage_entry = ttk.Entry(discharging_frame)
        fifth_second_voltage_entry.insert(0, str(latest_data.get("5th Second Voltage",0)))
        fifth_second_voltage_entry.grid(row=3, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(discharging_frame, text="30th Second Voltage").grid(row=4, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        thirtieth_second_voltage_entry = ttk.Entry(discharging_frame)
        thirtieth_second_voltage_entry.insert(0, str(latest_data.get("30th Second Voltage", 0)))
        thirtieth_second_voltage_entry.grid(row=4, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(discharging_frame, text="Date").grid(row=5, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        discharging_date_entry = ttk.Entry(discharging_frame)
        discharging_date_entry.insert(0, str(latest_data.get("Discharging Date", 0)))
        discharging_date_entry.grid(row=5, column=3, columnspan=2, padx=10, pady=10, sticky="w")

        def collect_data_and_generate_report():
            # Collecting all the data from the fields
            report_data = {
                "Project": project_combobox.get(),  # Collecting the project name
                "Device Name": device_name_entry.get(),
                "Serial Number": serial_number_entry.get(),
                "Manufacturer Name": manufacturer_entry.get(),
                "Cycle Count": cycle_count_entry.get(),
                "Battery Capacity": capacity_entry.get(),
                "Charging Duration": duration_entry.get(),
                "Charging Date": date_entry.get(),
                "Current Loaded": discharging_current_loaded_entry.get(),
                "Discharging Duration": discharging_duration_entry.get(),
                "OCV": ocv_entry.get(),
                "0th Second Voltage": zero_second_voltage_entry.get(),
                "5th Second Voltage": fifth_second_voltage_entry.get(),
                "30th Second Voltage": thirtieth_second_voltage_entry.get(),
                "Discharging Date": discharging_date_entry.get()
            }
            # Call the function to generate the PDF and update Excel
            generate_pdf_and_update_excel(report_data,self.device_data.get('serial_number'))

        # Generate Button (centered at the bottom, row 4, col 2)
        generate_button = ttk.Button(main_frame, text="Generate", command=collect_data_and_generate_report)
        generate_button.grid(row=2, column=4, pady=10, sticky="ew")

        # Select the report button (if necessary for your UI)
        self.select_button(self.report_button)

    def show_control(self):
        """Show the load control screen."""
        self.clear_content_frame()

        load_frame = ttk.Frame(self.content_frame, borderwidth=10, relief="solid")
        load_frame.grid(row=0, column=0, columnspan=8, rowspan=8, padx=10, pady=10, sticky="nsew")

        # Section: Right Control Frame (Change to Grid Layout)
        right_control_frame = ttk.Labelframe(load_frame, text="Controls", bootstyle="dark", borderwidth=5, relief="solid")
        right_control_frame.grid(row=0, column=1, rowspan=1, padx=5, pady=5, sticky="nsew")  # Positioned to the right

        # Load icons
        self.reset_icon = self.load_icon(os.path.join(self.assets_path, "reset.png"))
        self.activate_icon = self.load_icon(os.path.join(self.assets_path, "activate.png"))

        if self.selected_battery == 'Battery 1':
            # Charger Control
            charger_control_label = ttk.Label(right_control_frame, text="Charging On/Off", font=("Helvetica", 12))
            charger_control_label.grid(row=0, column=0, padx=10, pady=5)
            charger_control_button = ttk.Checkbutton(
                right_control_frame,
                variable=self.charger_control_var,
                bootstyle="success-round-toggle" if self.charger_control_var.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.discharger_control_var_battery_1, charger_control_button, 'charge', self.selected_battery)
            )
            charger_control_button.grid(row=0, column=1, padx=10, pady=5)

            # Discharger Control
            discharger_control_label = ttk.Label(right_control_frame, text="Discharging On/Off", font=("Helvetica", 12))
            discharger_control_label.grid(row=1, column=0, padx=10, pady=5)
            discharger_control_button = ttk.Checkbutton(
                right_control_frame,
                variable=self.discharger_control_var,
                bootstyle="success-round-toggle" if self.discharger_control_var.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.charger_control_var_battery_1, discharger_control_button, 'discharge', self.selected_battery)
            )
            discharger_control_button.grid(row=1, column=1, padx=10, pady=5)
        elif self.selected_battery == 'Battery 2':
            # Charger Control
            charger_control_label = ttk.Label(right_control_frame, text="Charging On/Off", font=("Helvetica", 12))
            charger_control_label.grid(row=0, column=0, padx=10, pady=5)
            charger_control_button = ttk.Checkbutton(
                right_control_frame,
                variable=self.charger_control_var,
                bootstyle="success-round-toggle" if self.charger_control_var.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.discharger_control_var_battery_2, charger_control_button, 'charge', self.selected_battery)
            )
            charger_control_button.grid(row=0, column=1, padx=10, pady=5)
    
            # Discharger Control
            discharger_control_label = ttk.Label(right_control_frame, text="Discharging On/Off", font=("Helvetica", 12))
            discharger_control_label.grid(row=1, column=0, padx=10, pady=5)
            discharger_control_button = ttk.Checkbutton(
                right_control_frame,
                variable=self.discharger_control_var,
                bootstyle="success-round-toggle" if self.discharger_control_var.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.charger_control_var_battery_2, discharger_control_button, 'discharge', self.selected_battery)
            )
            discharger_control_button.grid(row=1, column=1, padx=10, pady=5)

        # Row 2: BMS Reset Button
        self.bms_reset_button = ctk.CTkButton(
            right_control_frame, text="Reset", image=self.reset_icon, compound="left",
            command=self.bmsreset, fg_color="#ff6361", hover_color="#d74a49"
        )
        self.bms_reset_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # Row 3: Activate Heater Button
        self.activate_heater_button = ctk.CTkButton(
            right_control_frame, text="Activate Heater", image=self.activate_icon, compound="left",
            command=self.activate_heater, fg_color="#ff6361", hover_color="#d74a49"
        )
        self.activate_heater_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        # Section:  Device Connection
        device_connect_frame = ttk.Labelframe(load_frame, bootstyle='dark', text="Connection", padding=10, borderwidth=10, relief="solid")
        device_connect_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Connect Button
        self.connect_button = ttk.Button(device_connect_frame, text="Connect to Chroma", command=self.connect_device)
        self.connect_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        # Status Label
        self.status_label = ttk.Label(device_connect_frame, text="Status: Disconnected")
        self.status_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Check the selected mode
        selected_mode = self.mode_var.get()

        if selected_mode == "Testing Mode":
            # Section: Testing Mode
            testing_frame = ttk.Labelframe(load_frame, text="Testing Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
            testing_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

            test_button = ttk.Button(testing_frame, text="Set 50A and Turn ON Load", command=self.testing_mode_load)
            test_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

            # Turn Off Load Button
            turn_off_button = ttk.Button(testing_frame, text="Turn OFF Load", command=self.load_off)
            turn_off_button.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

        elif selected_mode == "Maintenance Mode":
            # Section: Maintenance Mode
            maintenance_frame = ttk.Labelframe(load_frame, text="Maintenance Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
            maintenance_frame.grid(row=1, column=0,columnspan=4, padx=10, pady=10, sticky="nsew")

            maintenance_button = ttk.Button(maintenance_frame, text="Set 100A and Turn ON Load", command=self.maintains_mode_load)
            maintenance_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

            # Turn Off Load Button
            turn_off_button = ttk.Button(maintenance_frame, text="Turn OFF Load", command=self.load_off)
            turn_off_button.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="ew")
    

        # Section: Custom Mode
        custom_frame = ttk.Labelframe(load_frame, text="Custom Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
        custom_frame.grid(row=3, column=0,columnspan=4, padx=5, pady=5, sticky="nsew")

        custom_label = ttk.Label(custom_frame, text="Set Custom Current (A):", font=("Helvetica", 12), width=20, anchor="e")
        custom_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        def save_custom_value():
            """Saves the custom current value and sets it to the Chroma device."""
            custom_value = self.custom_current_entry.get()
            print(f"{custom_value} load value")
            if custom_value.isdigit():  # Basic validation to ensure it's a number
                set_custom_l1_value(int(custom_value))
            else:
                print("Invalid input: Please enter a valid number.")

        self.custom_current_entry = ttk.Entry(custom_frame)
        self.custom_current_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        custom_button = ttk.Button(custom_frame, text="Save", command=save_custom_value)
        custom_button.grid(row=0, column=2, columnspan=1, padx=5, pady=5, sticky="ew")

        # Create a toggle button for turning load ON and OFF
        self.load_status = tk.BooleanVar()  # Boolean variable to track load status (ON/OFF)
        self.load_status.set(False)  # Initial state is OFF

        def toggle_load():
            if self.load_status.get():
                turn_load_off
                toggle_button.config(text="Turn ON Load", bootstyle="success")  # Change to green when OFF
                self.load_status.set(False)  # Update state to OFF
            else:
                turn_load_on
                toggle_button.config(text="Turn OFF Load", bootstyle="danger")  # Change to red when ON
                self.load_status.set(True)  # Update state to ON

        # Toggle button for Load ON/OFF
        toggle_button = ttk.Button(custom_frame, text="Turn ON Load", command=toggle_load, bootstyle='success')
        toggle_button.grid(row=0, column=3, columnspan=2, padx=5, pady=5, sticky="ew")

        # Make the rows and columns expand as needed for content_frame and load_frame
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)  # Ensures the load_frame expands fully
        self.content_frame.grid_columnconfigure(0, weight=1)

        load_frame.grid_rowconfigure(0, weight=1)  # Device Connection frame
        load_frame.grid_rowconfigure(1, weight=1)  # Testing frame
        load_frame.grid_rowconfigure(2, weight=1)  # Maintenance frame
        load_frame.grid_rowconfigure(3, weight=1)  # Custom frame
        load_frame.grid_columnconfigure(0, weight=1)  # Ensure columns expand to fill space

        # Highlight the control button in the side menu
        self.select_button(self.control_button)


    def connect_device(self):
        """Connect to the Chroma device."""
        result = find_and_connect()  # This can be the same method or a separate one for specific device connection
        self.status_label.config(text=result)

    def update_info(self, event=None):
        """Update the Info page content based on the selected mode."""
        if self.selected_button == self.info_button:
            self.show_control()
            self.show_info()
        if self.selected_button == self.control_button:
            self.show_info()
            self.show_control()
    
    def configure_styles(self):
            # Configure custom styles
            self.style.configure("Custom.Treeview", font=("Helvetica", 10), rowheight=25, background="#f0f0f0", fieldbackground="#e8f4ec")
            self.style.configure("Custom.Treeview.Heading", font=("Helvetica", 12, "bold"), background="#333", foreground="#e8f4ec", padding=5)
            self.style.map("Custom.Treeview", background=[("selected", "#0078d7")], foreground=[("selected", "white")])

    def center_window(self, width, height):
            """Centers the window on the screen."""
            screen_width = self.master.winfo_screenwidth()
            screen_height = self.master.winfo_screenheight()
            center_x = int(screen_width / 2 - width / 2)
            center_y = int(screen_height / 2 - height / 2)
            self.master.geometry(f'{width}x{height}+{center_x}+{center_y}')
            self.master.update_idletasks()

    def load_icon(self, path, size=(15, 15)):
        image = Image.open(path)
        image = image.resize(size, Image.Resampling.LANCZOS)
        ctk_image = ctk.CTkImage(light_image=image, dark_image=image, size=size)
        return ctk_image
    
    def load_icon_menu(self, path, size=(24, 24)):
        icon = Image.open(path)
        icon = icon.resize(size, Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(icon)

    def update_frame_sizes(self):
        width = self.master.winfo_width()
        self.side_menu_width = int(width * self.side_menu_width_ratio)
        self.content_frame_width = width - self.side_menu_width

    def on_window_resize(self, event):
        self.update_frame_sizes()
        self.side_menu_frame.place(x=0, y=0, width=self.side_menu_width, relheight=1.0)
        self.content_frame.place(x=self.side_menu_width, y=0, width=self.content_frame_width, relheight=1.0)

    def select_button(self, button, command=None):
        # Check if the previously selected button still exists
        if self.selected_button and self.selected_button.winfo_exists():
            # Only reset style if the previously selected button is not the Disconnect button
            if self.selected_button != self.disconnect_button:
                self.selected_button.configure(bootstyle="info")  # Reset the previously selected button's style

        # Highlight the new selected button
        if button != self.disconnect_button:
            button.configure(bootstyle="primary")  # Apply the selected style, unless it's the Disconnect button
        self.selected_button = button

        # Execute the button's command if provided
        if command:
            command()

    def clear_content_frame(self):
        # Clear the content frame before displaying new content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def on_disconnect(self):
        self.mode_var.set("Testing Mode")
        if device_data_battery_2['serial_number'] != 0:
            pcan_write_control('both_off',1)
            pcan_write_control('both_off',2)
        else:
            pcan_write_control('both_off',1)
        time.sleep(1)
        # Cancel the scheduled auto-refresh task
        if hasattr(self, 'auto_refresh_task'):
            self.master.after_cancel(self.auto_refresh_task)
            print("Auto-refresh stopped.")
        pcan_uninitialize()
        self.first_time_dashboard = True
        self.main_frame.pack_forget()
        self.main_window.show_main_window()
        # Perform disconnection logic here (example: print disconnect message)
        print("Disconnecting...")

    def show_info(self, event=None):
        self.clear_content_frame()

        # Determine if we are in Testing Mode or Maintenance Mode
        selected_mode = self.mode_var.get()
        self.limited = selected_mode == "Testing Mode"   
        # Add a new right-side frame for controls     
        self.refresh_icon = self.load_icon(os.path.join(self.assets_path, "refresh.png"))
        self.start_icon = self.load_icon(os.path.join(self.assets_path, "start.png"))
        self.stop_icon = self.load_icon(os.path.join(self.assets_path, "stop.png"))
        self.file_icon = self.load_icon(os.path.join(self.assets_path, "folder.png"))

        # Create a Frame to hold the checkbox and buttons in one row with a white background
        control_frame = ttk.Labelframe(self.content_frame, text="Battery Data Controls", bootstyle="dark", borderwidth=5, relief="solid")
        control_frame.pack(fill="x", expand=True, padx=5, pady=1)


        self.refresh_button = ctk.CTkButton(control_frame, text="Refresh", image=self.refresh_icon, compound="left", command=self.refresh_info, width=6, fg_color="#5188d4",
            hover_color="#4263cc")
        self.refresh_button.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        timer_value_label = ttk.Label(control_frame, text="Timer")
        timer_value_label.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        self.timer_value = tk.StringVar(value=5)
        self.timer = ttk.Spinbox(
            control_frame,
            from_=1,
            to=30,
            width=5,
            values=(1, 5, 10, 15, 20, 30),
            textvariable=self.timer_value,
        )
        self.timer.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        self.start_logging_button = ctk.CTkButton(control_frame, text="Start Log", image=self.start_icon, compound="left", command=self.start_logging, fg_color="#72b043",
            hover_color="#007f4e")
        self.start_logging_button.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        self.stop_logging_button = ctk.CTkButton(control_frame, text="Stop Log", image=self.stop_icon, compound="left", command=self.stop_logging, fg_color="#ff6361",
            hover_color="#d74a49", state=tk.DISABLED)
        self.stop_logging_button.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        file_button = ctk.CTkButton(control_frame, image=self.file_icon, text='', compound="left", command=self.folder_open, fg_color="#72b043",
            hover_color="#007f4e")
        file_button.pack(side="left", padx=5, pady=1, fill="x", expand=True)

        if self.logging_active:
            self.start_logging_button.configure(state=tk.DISABLED)
            self.start_logging_button.configure(state=tk.NORMAL)
        else:
            self.start_logging_button.configure(state=tk.NORMAL)
            self.stop_logging_button.configure(state=tk.DISABLED)   
        # Frame to hold the Treeview and Scrollbar
        table_frame = ttk.Frame(self.content_frame, borderwidth=5, relief="solid")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)   
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")   
        columns = ('Name', 'Value', 'Units')
        self.info_table = ttk.Treeview(table_frame, columns=columns, show='headings', style="Custom.Treeview", yscrollcommand=scrollbar.set)   
        scrollbar.config(command=self.info_table.yview)   
        self.info_table.column('Name', anchor='center', width=200)
        self.info_table.column('Value', anchor='center', width=200)
        self.info_table.column('Units', anchor='center', width=100)   
        self.info_table.heading('Name', text='Name', anchor='center')
        self.info_table.heading('Value', text='Value', anchor='center')
        self.info_table.heading('Units', text='Units', anchor='center')   
        self.info_table.pack(fill="both", expand=True)   
        if self.limited:
            # Insert limited data for Testing Mode
            limited_data_keys = [
                 'device_name', 
                 'serial_number', 
                 'manufacturer_name', 
                 'cycle_count',
                 'remaining_capacity', 
                 'temperature', 
                 'current', 
                 'voltage', 
                 'charging_current', 
                 'charging_voltage', 
                 'charging_battery_status', 
                 'rel_state_of_charge'
             ]
            for key in limited_data_keys:
                name = ' '.join(word.title() for word in key.split('_'))
                value = self.device_data.get(key, 'N/A')
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit))
        else:
            # Insert full data for Maintenance Mode
            for key, value in self.device_data.items():
                name = ' '.join(word.title() for word in key.split('_'))
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit))               
        self.status_frame = ttk.Labelframe(self.content_frame, text="Battery Status", bootstyle="dark", borderwidth=5, relief="solid")
        self.status_frame.pack(fill="both", expand=True, padx=10, pady=10)   
        if self.limited:
            limited_status_flags = {
                 "Over Temperature Alarm": self.battery_status_flags.get('over_temperature_alarm'),
                 "Fully Charged": self.battery_status_flags.get('fully_charged'),
                 "Fully Discharged": self.battery_status_flags.get('fully_discharged')
            }
            self.display_status_labels(self.status_frame, limited_status_flags)
        else:
            self.display_status_labels(self.status_frame, self.battery_status_flags)   
        self.select_button(self.info_button)

    def display_status_labels(self, frame, battery_status_flags):
        max_columns = 3  # Number of columns before wrapping to the next row
        current_column = 0
        current_row = 0

        for label_text, flag_value in battery_status_flags.items():
            transformed_label_text = ' '.join(word.title() for word in label_text.split('_'))
            if transformed_label_text == "Error Codes":
                self.create_error_code_label(frame, flag_value, row=current_row, column=current_column)
            else:
                created = self.create_status_label(frame, transformed_label_text, flag_value, row=current_row, column=current_column)
                if created:
                    current_column += 1
                    if current_column >= max_columns:
                        current_column = 0
                        current_row += 1

        # Ensure that the rows and columns expand evenly
        for i in range(current_row + 1):
            frame.grid_rowconfigure(i, weight=1)
        for j in range(max_columns):
            frame.grid_columnconfigure(j, weight=1)

    def create_status_label(self, frame, label_text, flag_value, row, column):
        # Only show "Fully Charged" and "Fully Discharged" if the value is 1
        if label_text == "Fully Charged" and flag_value == 1:
            status_text = "Fully Charged"
            bootstyle = "inverse-success"
        elif label_text == "Fully Discharged" and flag_value == 1:
            status_text = "Fully Discharged"
            bootstyle = "inverse-danger"
        else:
            # Skip creating the label if the flag is not active (except for Fully Charged/Discharged)
            if flag_value != 1:
                return False
            status_text = ""  # Just show the label without Active/Inactive text
            bootstyle = "inverse-success"

        # Create and place the label
        self.battery_status_label = ttk.Label(frame, text=f"{label_text}", bootstyle=bootstyle, width=35)
        self.battery_status_label.grid(row=row, column=column, padx=5, pady=5, sticky="ew")

        # Make the column expand to fill space
        frame.grid_columnconfigure(column, weight=1)

        return True  # Return True if a label was created
    
    def create_error_code_label(self, frame, error_code, row, column):
        # Determine the bootstyle and text based on the error code
        error_texts = {
            0: ("OK", "success"),
            1: ("Busy", "warning"),
            2: ("Reserved Command", "secondary"),
            3: ("Unsupported Command", "danger"),
            4: ("Access Denied", "info"),
            5: ("Overflow/Underflow", "dark"),
            6: ("Bad Size", "primary"),
            7: ("Unknown Error", "danger"),
        }
        error_text, bootstyle = error_texts.get(error_code, ("Unknown", "danger"))
        self.battery_status_label = ttk.Label(frame, text=f"Error Code: {error_text}", bootstyle=f"inverse-{bootstyle}", width=40)
        self.battery_status_label.grid(row=row, column=column, padx=5, pady=2, sticky="w")

    def auto_refresh(self):
        if self.is_connected:  # Only refresh if the device is connected
            asyncio.run(update_device_data())

            if self.selected_battery == "Battery 1":
                self.device_data = device_data_battery_1
                self.battery_status_flags = battery_1_status_flags
            elif self.selected_battery == "Battery 2":
                self.device_data = device_data_battery_2
                self.battery_status_flags = battery_2_status_flags

            # Schedule the next refresh and store the task ID in self.auto_refresh_task
            self.auto_refresh_task = self.master.after(1000, self.auto_refresh)

    def refresh_info(self):
        asyncio.run(update_device_data())
        # Clear the existing table rows
        for item in self.info_table.get_children():
            self.info_table.delete(item)

        # Check if Testing Mode is active
        if self.limited:
            # Insert limited data for Testing Mode
            limited_data_keys = [
                'device_name', 
                'serial_number', 
                'manufacturer_name', 
                'cycle_count',
                'remaining_capacity', 
                'temperature', 
                'current', 
                'voltage', 
                'charging_current', 
                'charging_voltage', 
                'charging_battery_status', 
                'rel_state_of_charge'
            ]
            for index, key in enumerate(limited_data_keys):
                name = ' '.join(word.title() for word in key.split('_'))
                value = self.device_data.get(key, 'N/A')
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
        else:
            # Insert full data for Maintenance Mode
            for index, (key, value) in enumerate(self.device_data.items()):
                name = ' '.join(word.title() for word in key.split('_'))
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
        # Update the battery status flags
        if self.limited:
            limited_status_flags = {
                "Over Temperature Alarm": battery_status_flags.get('over_temperature_alarm'),
                "Fully Charged": battery_status_flags.get('fully_charged'),
                "Fully Discharged": battery_status_flags.get('fully_discharged')
            }
            self.display_status_labels(self.status_frame, limited_status_flags)
        else:
            self.display_status_labels(self.status_frame, battery_status_flags)
    
    def start_logging(self):
        self.logging_active = True
        # Disable the Start Logging button and enable the Stop Logging button
        self.start_logging_button.configure(state=tk.DISABLED)
        self.stop_logging_button.configure(state=tk.NORMAL)

        # Create the folder path
        folder_path = os.path.join(os.path.expanduser("~"), "Documents", "ADE Log Folder")
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Create a timestamped file name
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"CAN_Connector_Log_{self.device_data['charging_battery_status']}_{timestamp}.xlsx"
        self.file_path = os.path.join(folder_path, file_name)

        # Create a new Excel workbook and worksheet
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "BMS Data"

        # Write the headers (first two columns for date and time, then the device names)
        headers = ["Date", "Time"] + [name_mapping.get(key, key) for key in self.device_data.keys()]
        self.sheet.append(headers)

        # Start the logging loop
        self.logging_active = True
        self.log_data()

    def log_data(self):
        if self.logging_active:
            # Update device data
            asyncio.run(update_device_data())

            for item in self.info_table.get_children():
                self.info_table.delete(item)

            # Repopulate the table with updated data
            for index, (key, value) in enumerate(self.device_data.items()):
                name = ' '.join(word.title() for word in key.split('_'))
                unit = unit_mapping.get(key, key)
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))

            # Get the current date and time
            current_datetime = datetime.datetime.now()
            date = current_datetime.strftime("%Y-%m-%d")
            time = current_datetime.strftime("%H:%M:%S")

            # Collect the row data: date, time, and device values
            row_data = [date, time] + [self.device_data[key] for key in self.device_data.keys()]

            # Write the row data to the Excel sheet
            self.sheet.append(row_data)

            # Save the workbook after each update
            self.workbook.save(self.file_path)

            self.log_interval = int(self.timer_value.get()) * 1000

            # Schedule the next log
            self.master.after(self.log_interval, self.log_data)
    
    def stop_logging(self):
        # Stop logging process and update button states
        self.logging_active = False
        # Disable the Stop Logging button and enable the Start Logging button
        self.stop_logging_button.configure(state=tk.DISABLED)
        self.start_logging_button.configure(state=tk.NORMAL)

        self.logging_active = False  # Stop the logging loop
        if hasattr(self, 'workbook'):
            self.workbook.save(self.file_path)  # Ensure the final save
            self.workbook.close()  # Close the workbook
        folder_path = os.path.dirname(self.file_path)
        os.startfile(folder_path)

    def folder_open(self):
        folder_path = os.path.join(os.path.expanduser("~"), "Documents", "ADE Log Folder")
        if os.path.exists(folder_path):
            os.startfile(folder_path)
        else:
            messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")

    def get_gauge_style(self, value, gauge_type):
        if gauge_type == "voltage":
            if value < 20:
                return "success"
            elif 20 <= value < 30:
                return "warning"
            elif 30 <= value < 35:
                return "danger"
            else:
                return "success"
        elif gauge_type == "current":
            if value < 20:
                return "success"
            elif 20 <= value < 50:
                return "info"
            elif 50 <= value < 120:
                return "warning"
            else:
                return "danger"
        elif gauge_type == "temperature":
            if value < 30:
                return "success"
            elif 30 <= value < 69:
                return "warning"
            elif value >= 70:
                return "danger"
            else:
                return "info"
        elif gauge_type == "capacity":
            if value < 20:
                return "danger"
            elif 20 <= value < 50:
                return "warning"
            elif 50 <= value < 80:
                return "info"
            else:
                return "success"
        elif gauge_type == "charging_current":
            if value < 20:
                return "success"
            elif 20 <= value < 50:
                return "info"
            elif 50 <= value < 80:
                return "warning"
            else:
                return "danger"
        else:
            return "info"  # Default style if gauge_type is unknown
 
    def toggle_button_style(self, var, button, control_type, active_battery):
        """
        Toggle button style and control charging/discharging for the given active battery.

        Parameters:
        - var: The BooleanVar tracking the state of the toggle (True/False).
        - button: The button object whose style will be updated.
        - control_type: 'charge' or 'discharge' to specify the type of control.
        - active_battery: 'Battery 1' or 'Battery 2' to specify which battery is being controlled.
        """
        # Determine which battery's control to operate on
        if active_battery == "Battery 1":
            charge_on = self.charger_control_var_battery_1.get()
            discharge_on = self.discharger_control_var_battery_1.get()
        elif active_battery == "Battery 2":
            charge_on = self.charger_control_var_battery_2.get()
            discharge_on = self.discharger_control_var_battery_2.get()

        # Ensure only one operation (charge/discharge) is active at a time
        if control_type == 'charge':
            if charge_on and discharge_on:
                messagebox.showwarning("Warning", "Please turn off discharging before enabling charging.")
                var.set(False)
                return
            if active_battery == "Battery 1":
                self.charge_fet_status_battery_1 = charge_on
                if charge_on:
                    pcan_write_control('charge_on', battery_no=1)
                    start_fetching_current(battery_no=1)
                    self.discharger_control_var_battery_1.set(False)  # Turn off discharging
                else:
                    pcan_write_control('both_off', battery_no=1)
                    stop_fetching_current()
            elif active_battery == "Battery 2":
                self.charge_fet_status_battery_2 = charge_on
                if charge_on:
                    pcan_write_control('charge_on', battery_no=2)
                    start_fetching_current(battery_no=2)
                    self.discharger_control_var_battery_2.set(False)  # Turn off discharging
                else:
                    pcan_write_control('both_off', battery_no=2)
        elif control_type == 'discharge':
            if charge_on and discharge_on:
                messagebox.showwarning("Warning", "Please turn off charging before enabling discharging.")
                var.set(False)
                return
            if active_battery == "Battery 1":
                self.discharge_fet_status_battery_1 = discharge_on
                if discharge_on:
                    pcan_write_control('discharge_on', battery_no=1)
                    self.charger_control_var_battery_1.set(False)  # Turn off charging
                else:
                    pcan_write_control('both_off', battery_no=1)
            elif active_battery == "Battery 2":
                self.discharge_fet_status_battery_2 = discharge_on
                if discharge_on:
                    pcan_write_control('discharge_on', battery_no=2)
                    self.charger_control_var_battery_2.set(False)  # Turn off charging
                else:
                    pcan_write_control('both_off', battery_no=2)

        # Update button styles
        button.config(bootstyle="success round-toggle" if var.get() else "danger round-toggle")

        # Show alert message
        action = "Charging" if control_type == 'charge' else "Discharging"
        messagebox.showinfo("Action", f"{action} function for {active_battery} activated. Please wait for 10 seconds.")

        # Disable the button for 10 seconds
        button.config(state=tk.DISABLED)

        def reenable_button():
            # Wait for 10 seconds in a separate thread
            def wait_and_enable():
                time.sleep(10)
                try:
                    # Ensure the widget is still valid
                    if button.winfo_exists():
                        button.config(state=tk.NORMAL)
                except tk.TclError:
                    # The button no longer exists, so just pass
                    pass

            threading.Thread(target=wait_and_enable).start()

        # Start the cooldown timer
        reenable_button()

    def bmsreset(self):
        # # Disable the main window
        # self.master.attributes("-disabled", True)

        # Create a new top-level window for the progress bar without a title bar
        popup = ctk.CTkToplevel(self.master)
        popup.title("BMS Reset Progress")
        popup.geometry("300x100")
        popup.resizable(False, False)
        popup.overrideredirect(True)  # Remove the title bar (no minimize or close buttons)

        # Center the popup window on the screen
        popup.update_idletasks()
        x = (self.master.winfo_screenwidth() - popup.winfo_width()) // 2
        y = (self.master.winfo_screenheight() - popup.winfo_height()) // 2
        popup.geometry(f'+{x}+{y}')

        # Add a label and a progress bar to the popup
        label = ctk.CTkLabel(popup, text="Resetting BMS, please wait...", font=("Helvetica", 12))
        label.pack(pady=10)
        progress_bar = ctk.CTkProgressBar(popup, mode="determinate", progress_color="#45a049")
        progress_bar.pack(fill="x", padx=20, pady=10)

        # Function to update the progress bar gradually over 15 seconds
        def update_progress_bar():
            for i in range(101):  # 0% to 100%
                progress_bar.set(i / 100)  # Update progress bar value
                time.sleep(0.15)  # Sleep to simulate progress over time

            popup.destroy()  # Close the popup window
            # self.master.attributes("-disabled", False)  # Re-enable the main window

        # Run the update_progress_bar function in a separate thread to avoid blocking the UI
        threading.Thread(target=update_progress_bar).start()

        # Start the BMS reset process (this can be moved to the update_progress_bar function if needed)
        pcan_write_control('bms_reset')
        
    def activate_heater(self):
        # Change the button to indicate the heater is activated
        self.activate_heater_button.configure(text="Heater Activated", fg_color="#4CAF50", hover_color="#45a049")

        # Call pcan_write_control to activate the heater
        pcan_write_control('heater_on')

        # Start a timer to revert the button after 10 seconds
        self.master.after(10000, self.reset_heater_button)

    def reset_heater_button(self):
        # Revert the button to its original state
        self.activate_heater_button.configure(text="Activate Heater", fg_color="#ff6361", hover_color="#d74a49")

    def check_battery_log_for_cycle_count(self):
        """
        Check the battery log file to see if the current battery's serial number exists
        and if the cycle count is 0. If cycle count is 0, show a popup to update it.
        """
        # Define the folder and file path for the battery log
        folder_path = os.path.join(os.path.expanduser("~"), "Documents", "Battery_Logs")
        file_path = os.path.join(folder_path, "Can_Log.xlsx")

        # Check if the file exists
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "Battery log file not found.")
            return False  # Return False if the file doesn't exist

        # Open the existing Excel file
        workbook = load_workbook(file_path)
        sheet = workbook.active

        # Loop through the rows to find the correct row by Serial Number
        serial_number_column = 8  # Assuming Serial Number is in column H (8th column)
        cycle_count_column = 10  # Assuming Cycle Count is in column J (10th column)

        battery_serial_number = self.device_data.get('serial_number')

        serial_found = False
        for row in range(2, sheet.max_row + 1):  # Start at 2 to skip the header
            if sheet.cell(row=row, column=serial_number_column).value == battery_serial_number:
                serial_found = True  # Serial number found
                cycle_count = sheet.cell(row=row, column=cycle_count_column).value

                # If the cycle count is 0, update the dashboard flag and show popup
                if cycle_count == 0:
                    print('test')
                    self.first_time_dashboard = True
                    self.prompt_manual_cycle_count()  # Show the input dialog for cycle count update
                else:
                    print('test1')
                    self.device_data['cycle_count'] = cycle_count
                    self.first_time_dashboard = False
                break

        if not serial_found:
            messagebox.showwarning("Warning", f"Serial number {battery_serial_number} not found.")
            return False  # Return False if the serial number is not found

        return True  # Return True if the serial number is found and checked

