#can_battery_info.py
import customtkinter as ctk
import tkinter as tk
import os
import customtkinter as ctk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
import datetime
import threading
import asyncio
from tkinter import messagebox
from openpyxl import Workbook
from pcan_api.custom_pcan_methods import *

class CanBatteryInfo:
    def __init__(self, master, main_window=None):
        self.master = master
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button

        self.center_window(1200, 600)  # Center the window with specified dimensions

        # Calculate one-fourth of the main window width for the side menu
        self.side_menu_width_ratio = 0.20  # 20% for side menu
        self.update_frame_sizes()  # Set initial size
        self.auto_refresh_var = tk.BooleanVar()
        self.logging_active = False
        self.charge_fet_status = False
        self.discharge_fet_status = False
        self.first_time_dashboard = True
        self.mode_var = tk.StringVar(value="Testing Mode")  # Initialize mode_var here

        self.style = ttk.Style()
        self.configure_styles()

        # Directory paths
        base_path = os.path.dirname(os.path.abspath(__file__))
        self.assets_path = os.path.join(base_path, "../assets/images/")

        style = ttk.Style() 
        # Configure the Labelframe title font to be bold
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold"))   


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
        self.loads_icon = self.load_icon_menu(os.path.join(self.assets_path, "load_icon.png"))
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
        self.info_button = ttk.Button(
            self.side_menu_frame,
            text=" Information",
            command=lambda: self.select_button(self.info_button, self.show_info),
            image=self.info_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.info_button.pack(fill="x", pady=5)
        self.load_button = ttk.Button(
            self.side_menu_frame,
            text=" E-Load       ",
            command=lambda: self.select_button(self.load_button),
            image=self.loads_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.load_button.pack(fill="x", pady=2)
        self.help_button = ttk.Button(
            self.side_menu_frame,
            text=" Help            ",
            command=lambda: self.select_button(self.help_button),
            image=self.help_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.help_button.pack(fill="x", pady=(0,10))

                # Add the Download PDF button
        self.download_pdf_button = ttk.Button(
            self.side_menu_frame, 
            text="Download Summary PDF",
            image=self.download_icon,)
        self.download_pdf_button.pack(fill="x", pady=20)

        # Mode Label
        self.mode_label = ttk.Label(
            self.side_menu_frame,
            bootstyle="inverse-info",
            text="Mode:",
        )
        self.mode_label.pack(pady=(130, 5),padx=(0,5))

        # Mode Dropdown
        self.mode_dropdown = ttk.Combobox(
            self.side_menu_frame,
            textvariable=self.mode_var,
            values=["Testing Mode", "Maintenance Mode"],
            state="readonly"
        )
        self.mode_dropdown.pack(padx=(5, 20))
        self.mode_dropdown.bind("<<ComboboxSelected>>", self.update_info)

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
         
        self.show_dashboard()

        # Bind the window resize event to adjust frame sizes dynamically
        self.master.bind('<Configure>', self.on_window_resize)

    def show_dashboard(self):
        self.clear_content_frame()

        # Check if this is the first time the dashboard is being opened
        if self.first_time_dashboard:
            self.first_time_dashboard = False  # Set the flag to False after the first time
            self.prompt_manual_cycle_count()  # Show the input dialog

        # Create a Frame for the battery info details at the bottom
        info_frame = ttk.Labelframe(self.content_frame, text="Battery Information", bootstyle="dark")
        info_frame.pack(fill="x", expand=True, padx=10, pady=10)

        # Device Name
        ttk.Label(info_frame, text="Device Name:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.device_name_label = ttk.Label(info_frame, text=device_data.get('device_name', 'N/A'), font=("Helvetica", 10))
        self.device_name_label.pack(side="left", padx=5)

        # Serial Number
        ttk.Label(info_frame, text="Serial Number:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.serial_number_label = ttk.Label(info_frame, text=device_data.get('serial_number', 'N/A'), font=("Helvetica", 10))
        self.serial_number_label.pack(side="left", padx=5)

        # Manufacturer Name
        ttk.Label(info_frame, text="Manufacturer:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.manufacturer_label = ttk.Label(info_frame, text=device_data.get('manufacturer_name', 'N/A'), font=("Helvetica", 10))
        self.manufacturer_label.pack(side="left", padx=5)

        # Manual Cycle Count
        ttk.Label(info_frame, text="Manual Cycle Count:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        self.cycle_count_label = ttk.Label(info_frame, text=device_data.get('manual_cycle_count', 'N/A'), font=("Helvetica", 10))
        self.cycle_count_label.pack(side="left", padx=5)

        # Charging Status
        ttk.Label(info_frame, text="Charging Status:", font=("Helvetica", 10, "bold")).pack(side="left", padx=5)

        self.status_var = tk.StringVar()
        charging_bootstyle = 'dark'
        discharging_bootstyle = "dark"

        # Determine the status and set the corresponding style
        charging_status = device_data.get('charging_battery_status', 'Idle').lower()
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
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Charging Section
        charging_frame = ttk.Labelframe(main_frame, text="Charging", bootstyle=charging_bootstyle)
        charging_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        # Charging Current (Row 0, Column 0)
        charging_current_frame = ttk.Frame(charging_frame)
        charging_current_frame.grid(row=0, column=0, padx=70, pady=5, sticky="nsew", columnspan=1)
        charging_current_label = ttk.Label(charging_current_frame, text="Charging Current", font=("Helvetica", 12, "bold"))
        charging_current_label.pack(pady=(5, 5))
        charging_current_meter = ttk.Meter(
            master=charging_current_frame,
            metersize=180,
            amountused=device_data['charging_current'],
            meterthickness=10,
            metertype="semi",
            subtext="Charging Current",
            textright="A",
            amounttotal=120,
            bootstyle=self.get_gauge_style(device_data['charging_current'], "charging_current"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        charging_current_meter.pack()

        # Charging Voltage (Row 0, Column 1)
        charging_voltage_frame = ttk.Frame(charging_frame)
        charging_voltage_frame.grid(row=0, column=1, padx=70, pady=5, sticky="nsew", columnspan=1)
        charging_voltage_label = ttk.Label(charging_voltage_frame, text="Charging Voltage", font=("Helvetica", 12, "bold"))
        charging_voltage_label.pack(pady=(5, 5))
        charging_voltage_meter = ttk.Meter(
            master=charging_voltage_frame,
            metersize=180,
            amountused=device_data['charging_voltage'],
            meterthickness=10,
            metertype="semi",
            subtext="Charging Voltage",
            textright="V",
            amounttotal=35,
            bootstyle=self.get_gauge_style(device_data['charging_voltage'], "charging_voltage"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        charging_voltage_meter.pack()

        # Battery Health Section
        battery_health_frame = ttk.Labelframe(main_frame, text="Battery Health", bootstyle='dark')
        battery_health_frame.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky="nsew")

        # Temperature (Top of Battery Health)
        temp_frame = ttk.Frame(battery_health_frame)
        temp_frame.pack(padx=10, pady=5, fill="both", expand=True)
        temp_label = ttk.Label(temp_frame, text="Temperature", font=("Helvetica", 12, "bold"))
        temp_label.pack(pady=(10, 10))
        temp_meter = ttk.Meter(
            master=temp_frame,
            metersize=180,
            amountused=device_data['temperature'],
            meterthickness=10,
            metertype="semi",
            subtext="Temperature",
            textright="°C",
            amounttotal=100,
            bootstyle=self.get_gauge_style(device_data['temperature'], "temperature"),
            stripethickness=8,
            subtextfont='-size 10'
        )
        temp_meter.pack()

        # Capacity (Bottom of Battery Health)
        capacity_frame = ttk.Frame(battery_health_frame)
        capacity_frame.pack(padx=10, pady=5, fill="both", expand=True)
        capacity_label = ttk.Label(capacity_frame, text="Capacity", font=("Helvetica", 12, "bold"))
        capacity_label.pack(pady=(10, 10))
        full_charge_capacity = device_data.get('full_charge_capacity', 0)
        remaining_capacity = device_data.get('remaining_capacity', 0)
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
        discharging_frame = ttk.Labelframe(main_frame, text="Discharging", bootstyle=discharging_bootstyle)
        discharging_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        def update_max_value():
            if self.check_discharge_max.get():
                self.discharge_max_value.set(450)
            else:
                self.discharge_max_value.set(100)
            self.discharging_current_meter.configure(amounttotal=self.discharge_max_value.get())

        # Discharging Current (Row 1, Column 0)
        discharging_current_frame = ttk.Frame(discharging_frame)
        discharging_current_frame.grid(row=1, column=0, padx=70, pady=5, sticky="nsew", columnspan=1)
        discharging_current_label = ttk.Label(discharging_current_frame, text="Discharging Current", font=("Helvetica", 12, "bold"))
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
            amountused=device_data['current'],  # Assuming this is the discharging current
            meterthickness=10,
            metertype="semi",
            subtext="Discharging Current",
            textright="A",
            amounttotal=self.discharge_max_value.get(),
            bootstyle=self.get_gauge_style(device_data['current'], "current"),
            stripethickness=10,
            subtextfont='-size 10'
        )
        self.discharging_current_meter.pack()

        # Discharging Voltage (Row 1, Column 1)
        discharging_voltage_frame = ttk.Frame(discharging_frame)
        discharging_voltage_frame.grid(row=1, column=1, padx=70, pady=5, sticky="nsew", columnspan=1)
        discharging_voltage_label = ttk.Label(discharging_voltage_frame, text="Discharging Voltage", font=("Helvetica", 12, "bold"))
        discharging_voltage_label.pack(pady=(5, 35))
        discharging_voltage_meter = ttk.Meter(
            master=discharging_voltage_frame,
            metersize=180,
            amountused=device_data['voltage'],  # Assuming this is the discharging voltage
            meterthickness=10,
            metertype="semi",
            subtext="Discharging Voltage",
            textright="V",
            amounttotal=100,
            bootstyle=self.get_gauge_style(device_data['voltage'], "voltage"),
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
            input_window.title("Manual Cycle Count")
        
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
            label = ttk.Label(input_window, text="Enter Manual Cycle Count:", font=("Helvetica", 12))
            label.pack(pady=10)
        
            manual_cycle_count_var = tk.IntVar()  # Use IntVar to store integer value
            entry = ttk.Entry(input_window, textvariable=manual_cycle_count_var)
            entry.pack(pady=5)
        
            # Function to handle OK button click
            def handle_ok():
                device_data['manual_cycle_count'] = manual_cycle_count_var.get()
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

    def show_controls(self):
        self.clear_content_frame()

        # label = ttk.Label(self.content_frame, text="Controls", font=("Helvetica", 16))
        # label.pack(pady=20, anchor="center")
        self.reset_icon = self.load_icon(os.path.join(self.assets_path, "reset.png"))
        
        bms_control_label = ttk.Label(self.content_frame, text="BMS Controls", font=("Helvetica", 16, "bold"))
        bms_control_label.pack(pady=(20, 5))

        # Create a frame for the first Checkbutton
        check_frame1 = ttk.Frame(self.content_frame)
        check_frame1.pack(pady=10, anchor="center")

        check_label1 = ttk.Label(check_frame1, text="Charge FET ON/OFF:", font=("Helvetica", 12), width=20, anchor="e")
        check_label1.pack(side="left", padx=10)

        self.check_var1 = tk.BooleanVar()
        self.check_var1.set(self.charge_fet_status)
        check_button1 = ttk.Checkbutton(
            check_frame1,
            variable=self.check_var1,
            width=15,
            bootstyle="success-round-toggle" if self.check_var1.get() else "danger-round-toggle",
            command=lambda: self.toggle_button_style(self.check_var1, check_button1,'charge')
        )
        check_button1.pack(side="right", padx=10)

        # Create a frame for the second Checkbutton
        check_frame2 = ttk.Frame(self.content_frame)
        check_frame2.pack(pady=10, anchor="center")

        check_label2 = ttk.Label(check_frame2, text="Discharge FET ON/OFF:", font=("Helvetica", 12), width=20, anchor="e")
        check_label2.pack(side="left", padx=10)

        self.check_var2 = tk.BooleanVar()
        self.check_var2.set(self.discharge_fet_status)
        check_button2 = ttk.Checkbutton(
            check_frame2,
            variable=self.check_var2,
            width=15,
            bootstyle="success-round-toggle" if self.check_var2.get() else "danger-round-toggle",
            command=lambda: self.toggle_button_style(self.check_var2, check_button2, 'discharge')
        )
        check_button2.pack(side="right",pady=20, padx=10)

        bms_reset_label = ttk.Label(self.content_frame, text="BMS Reset", font=("Helvetica", 16, "bold"))
        bms_reset_label.pack(pady=(5, 5))

        self.bms_reset_button = ctk.CTkButton(self.content_frame, text="Reset", image=self.reset_icon, compound="left", command=self.bmsreset, width=12, fg_color="#ff6361",
            hover_color="#d74a49")
        self.bms_reset_button.pack(pady=(5, 5))

        activate_heater_label = ttk.Label(self.content_frame, text="Activate Heater", font=("Helvetica", 16, "bold"))
        activate_heater_label.pack(pady=(5, 5))

        self.activate_heater_button = ctk.CTkButton(self.content_frame, text="OFF", compound="left", command=self.activate_heater, width=12, fg_color="#ff6361",
            hover_color="#d74a49")
        self.activate_heater_button.pack(pady=(5, 5))

        # Pack the content_frame itself (if necessary, but typically this frame is already packed)
        self.content_frame.pack(fill="both", expand=True)

        # Highlight the control button in the side menu
        self.select_button(self.control_button)
    
    def show_load(self):
        self.clear_content_frame()

        label = ttk.Label(self.content_frame, text="Electronic Load", font=("Helvetica", 16))
        label.pack(pady=20, anchor="center")
        # self.reset_icon = self.load_icon("assets/images/reset.png")
        
        # bms_control_label = ttk.Label(self.content_frame, text="BMS Controls", font=("Helvetica", 16, "bold"))
        # bms_control_label.pack(pady=(20, 5))

        # # Create a frame for the first Checkbutton
        # check_frame1 = ttk.Frame(self.content_frame)
        # check_frame1.pack(pady=10, anchor="center")

        # check_label1 = ttk.Label(check_frame1, text="Charge FET ON/OFF:", font=("Helvetica", 12), width=20, anchor="e")
        # check_label1.pack(side="left", padx=10)

        # self.check_var1 = tk.BooleanVar()
        # self.check_var1.set(self.charge_fet_status)
        # check_button1 = ttk.Checkbutton(
        #     check_frame1,
        #     variable=self.check_var1,
        #     width=15,
        #     bootstyle="success-round-toggle" if self.check_var1.get() else "danger-round-toggle",
        #     command=lambda: self.toggle_button_style(self.check_var1, check_button1,'charge')
        # )
        # check_button1.pack(side="right", padx=10)

        # # Create a frame for the second Checkbutton
        # check_frame2 = ttk.Frame(self.content_frame)
        # check_frame2.pack(pady=10, anchor="center")

        # check_label2 = ttk.Label(check_frame2, text="Discharge FET ON/OFF:", font=("Helvetica", 12), width=20, anchor="e")
        # check_label2.pack(side="left", padx=10)

        # self.check_var2 = tk.BooleanVar()
        # self.check_var2.set(self.discharge_fet_status)
        # check_button2 = ttk.Checkbutton(
        #     check_frame2,
        #     variable=self.check_var2,
        #     width=15,
        #     bootstyle="success-round-toggle" if self.check_var2.get() else "danger-round-toggle",
        #     command=lambda: self.toggle_button_style(self.check_var2, check_button2, 'discharge')
        # )
        # check_button2.pack(side="right",pady=20, padx=10)

        # bms_reset_label = ttk.Label(self.content_frame, text="BMS Reset", font=("Helvetica", 16, "bold"))
        # bms_reset_label.pack(pady=(5, 5))

        # self.bms_reset_button = ttk.Button(self.content_frame, text="Reset", image=self.reset_icon, compound="left", command=self.bmsreset, width=12, bootstyle="danger")
        # self.bms_reset_button.pack(pady=(5, 5))

        # activate_heater_label = ttk.Label(self.content_frame, text="Activate Heater", font=("Helvetica", 16, "bold"))
        # activate_heater_label.pack(pady=(5, 5))

        # self.activate_heater_button = ttk.Button(self.content_frame, text="OFF", compound="left", command=self.activate_heater, width=12, bootstyle="danger")
        # self.activate_heater_button.pack(pady=(5, 5))

        # # Pack the content_frame itself (if necessary, but typically this frame is already packed)
        # self.content_frame.pack(fill="both", expand=True)

        # Highlight the control button in the side menu
        self.select_button(self.load_button)

    def update_info(self, event=None):
        """Update the Info page content based on the selected mode."""
        if self.selected_button == self.info_button:
            self.show_info()
    
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
        pcan_write_control('both_off')
        time.sleep(1)
        pcan_uninitialize()
        self.first_time_dashboard = True
        self.main_frame.pack_forget()
        self.main_window.show_main_window()
        # Perform disconnection logic here (example: print disconnect message)
        print("Disconnecting...")
        # self.update_widgets()

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
        control_frame = ttk.Labelframe(self.content_frame, text="Battery Data Controls", bootstyle="dark")
        control_frame.pack(fill="x", expand=True, padx=10, pady=10)

        auto_refresh_checkbox = ttk.Checkbutton(
            control_frame, 
            text="Auto Refresh", 
            variable=self.auto_refresh_var, 
            width=12, 
            command=self.auto_refresh  # Link the checkbox command to start/stop auto-refresh
        )
        auto_refresh_checkbox.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        self.refresh_button = ctk.CTkButton(control_frame, text="Refresh", image=self.refresh_icon, compound="left", command=self.refresh_info, width=6, fg_color="#5188d4",
            hover_color="#4263cc")
        self.refresh_button.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        timer_value_label = ttk.Label(control_frame, text="Timer")
        timer_value_label.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        self.timer_value = tk.StringVar(value=5)
        self.timer = ttk.Spinbox(
            control_frame,
            from_=1,
            to=30,
            width=5,
            values=(1, 5, 10, 15, 20, 30),
            textvariable=self.timer_value,
        )
        self.timer.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        self.start_logging_button = ctk.CTkButton(control_frame, text="Start Log", image=self.start_icon, compound="left", command=self.start_logging, fg_color="#72b043",
            hover_color="#007f4e")
        self.start_logging_button.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        self.stop_logging_button = ctk.CTkButton(control_frame, text="Stop Log", image=self.stop_icon, compound="left", command=self.stop_logging, fg_color="#ff6361",
            hover_color="#d74a49", state=tk.DISABLED)
        self.stop_logging_button.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        file_button = ctk.CTkButton(control_frame, image=self.file_icon, text='', compound="left", command=self.folder_open, fg_color="#72b043",
            hover_color="#007f4e")
        file_button.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        if self.logging_active:
            self.start_logging_button.configure(state=tk.DISABLED)
            self.start_logging_button.configure(state=tk.NORMAL)
        else:
            self.start_logging_button.configure(state=tk.NORMAL)
            self.stop_logging_button.configure(state=tk.DISABLED)   
        # Frame to hold the Treeview and Scrollbar
        table_frame = ttk.Frame(self.content_frame)
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
                 'manual_cycle_count',
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
                name = name_mapping.get(key, key)
                value = device_data.get(key, 'N/A')
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit))
        else:
            # Insert full data for Maintenance Mode
            for key, value in device_data.items():
                name = name_mapping.get(key, key)
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit))   
        if not self.limited:
            right_control_frame = ttk.Labelframe(self.content_frame, text="Controls", bootstyle="dark")
            right_control_frame.pack(fill="both", expand=True, padx=5, pady=5, side="right")   
            self.reset_icon = self.load_icon(os.path.join(self.assets_path, "reset.png"))
            self.activate_icon = self.load_icon(os.path.join(self.assets_path, "activate.png"))  
            check_label1 = ttk.Label(right_control_frame, text="Charging On/Off:", font=("Helvetica", 8), width=10, anchor="center")
            check_label1.pack(pady=5)   
            self.check_var1 = tk.BooleanVar()
            self.check_var1.set(self.charge_fet_status)
            check_button1 = ttk.Checkbutton(
                right_control_frame,
                variable=self.check_var1,
                bootstyle="success-round-toggle" if self.check_var1.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.check_var1, check_button1, 'charge')
            )
            check_button1.pack(pady=5)   
            check_label2 = ttk.Label(right_control_frame, text="Discharging On/Off:", font=("Helvetica", 8), width=10, anchor="center")
            check_label2.pack(pady=5)   
            self.check_var2 = tk.BooleanVar()
            self.check_var2.set(self.discharge_fet_status)
            check_button2 = ttk.Checkbutton(
                right_control_frame,
                variable=self.check_var2,
                bootstyle="success-round-toggle" if self.check_var2.get() else "danger-round-toggle",
                command=lambda: self.toggle_button_style(self.check_var2, check_button2, 'discharge')
            )
            check_button2.pack(pady=5)   
            self.bms_reset_button = ctk.CTkButton(right_control_frame, text="Reset", image=self.reset_icon, compound="left", command=self.bmsreset, width=10, fg_color="#ff6361",
                hover_color="#d74a49")
            self.bms_reset_button.pack(pady=5)   
            self.activate_heater_button = ctk.CTkButton(right_control_frame, text="Activate Heater", image=self.activate_icon, compound="left", command=self.activate_heater, width=10, fg_color="#ff6361",
                hover_color="#d74a49")
            self.activate_heater_button.pack(pady=5)    
        self.status_frame = ttk.Labelframe(self.content_frame, text="Battery Status", bootstyle="dark")
        self.status_frame.pack(fill="both", expand=True, padx=10, pady=10)   
        if self.limited:
            limited_status_flags = {
                 "Over Temperature Alarm": battery_status_flags.get('over_temperature_alarm'),
                 "Fully Charged": battery_status_flags.get('fully_charged'),
                 "Fully Discharged": battery_status_flags.get('fully_discharged')
            }
            self.display_status_labels(self.status_frame, limited_status_flags)
        else:
            self.display_status_labels(self.status_frame, battery_status_flags)   
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
        if label_text == "Charge Fet Test":
            if flag_value == 1:
                status_text = "Discharging or At Rest"
                bootstyle = "inverse-success"
            else:
                status_text = "Charging"
                bootstyle = "inverse-info"
        elif label_text == "Fully Charged":
            if flag_value == 1:
                status_text = "Fully Charged"
                bootstyle = "inverse-success"
            else:
                status_text = "Not Fully Charged"
                bootstyle = "inverse-danger"
        elif label_text == "Fully Discharged":
            if flag_value == 1:
                status_text = "Fully Discharged"
                bootstyle = "inverse-danger"
            else:
                status_text = "OK"
                bootstyle = "inverse-success"
        else:
            if flag_value != 1:
                return False  # Skip creating the label if the status is not "Active"
            status_text = "Active"
            bootstyle = "inverse-success"

        # Create and place the label
        self.battery_status_label = ttk.Label(frame, text=f"{label_text}: {status_text}", bootstyle=bootstyle, width=35)
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
        print(f"Auto-refresh mode: {'Testing Mode' if self.limited else 'Maintenance Mode'}")
        if self.auto_refresh_var.get():
            asyncio.run(update_device_data())

        # Check if the info_table widget still exists before trying to clear it
        if self.info_table.winfo_exists():
            # Clear the existing table rows
            for item in self.info_table.get_children():
                self.info_table.delete(item)
        else:
            print("info_table does not exist anymore.")
            return  # Exit the function if the widget does not exist

        # Check if Testing Mode is active
        if self.limited:
            # Insert limited data for Testing Mode
            limited_data_keys = [
                'device_name', 
                'serial_number', 
                'manufacturer_name', 
                'manual_cycle_count',
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
                name = name_mapping.get(key, key)
                value = device_data.get(key, 'N/A')
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
        else:
            # Insert full data for Maintenance Mode
            for index, (key, value) in enumerate(device_data.items()):
                name = name_mapping.get(key, key)
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

        # Schedule the next refresh
        self.master.after(5000, self.auto_refresh)

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
                'manual_cycle_count',
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
                name = name_mapping.get(key, key)
                value = device_data.get(key, 'N/A')
                unit = unit_mapping.get(key, '')
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
        else:
            # Insert full data for Maintenance Mode
            for index, (key, value) in enumerate(device_data.items()):
                name = name_mapping.get(key, key)
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
        file_name = f"CAN_Connector_Log_{device_data['charging_battery_status']}_{timestamp}.xlsx"
        self.file_path = os.path.join(folder_path, file_name)

        # Create a new Excel workbook and worksheet
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "BMS Data"

        # Write the headers (first two columns for date and time, then the device names)
        headers = ["Date", "Time"] + [name_mapping.get(key, key) for key in device_data.keys()]
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
            for index, (key, value) in enumerate(device_data.items()):
                name = name_mapping.get(key, key)
                unit = unit_mapping.get(key, key)
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))

            # Get the current date and time
            current_datetime = datetime.datetime.now()
            date = current_datetime.strftime("%Y-%m-%d")
            time = current_datetime.strftime("%H:%M:%S")

            # Collect the row data: date, time, and device values
            row_data = [date, time] + [device_data[key] for key in device_data.keys()]

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
 
    def toggle_button_style(self, var, button, control_type):
        charge_on = self.check_var1.get()
        discharge_on = self.check_var2.get()

        # Ensure only one operation is active at a time
        if control_type == 'charge':
            if charge_on and discharge_on:
                messagebox.showwarning("Warning", "Please turn off discharging before enabling charging.")
                var.set(False)
                return
            self.charge_fet_status = charge_on
            if charge_on:
                pcan_write_control('charge_on')
                self.check_var2.set(False)  # Turn off discharging
            else:
                pcan_write_control('both_off' if not discharge_on else 'discharge_on')
        elif control_type == 'discharge':
            if discharge_on and charge_on:
                messagebox.showwarning("Warning", "Please turn off charging before enabling discharging.")
                var.set(False)
                return
            self.discharge_fet_status = discharge_on
            if discharge_on:
                pcan_write_control('discharge_on')
                self.check_var1.set(False)  # Turn off charging
            else:
                pcan_write_control('both_off' if not charge_on else 'charge_on')

        # Update button styles
        button.config(bootstyle="success round-toggle" if var.get() else "danger round-toggle")

        # Show alert message
        if control_type == 'charge':
            action = "Charging" if charge_on else "Charging turned off"
        else:
            action = "Discharging" if discharge_on else "Discharging turned off"

        messagebox.showinfo("Action", f"{action} function activated. Please wait for 10 seconds.")

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
        # Disable the main window
        self.master.attributes("-disabled", True)

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
            self.master.attributes("-disabled", False)  # Re-enable the main window

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


