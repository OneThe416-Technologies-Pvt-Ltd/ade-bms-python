#can_battery_info.py
import customtkinter as ctk
import tkinter as tk
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
import datetime
import time
import asyncio
from tkinter import messagebox
from openpyxl import Workbook
from pcan_api.custom_pcan_methods import *
from pcomm_api.pcomm import *
from load_api.chroma_api import *

class RSBatteryInfo:
    def __init__(self, master, main_window=None):
        self.master = master
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button

        self.center_window(1200, 600)  # Center the window with specified dimensions

        # Calculate one-fourth of the main window width for the side menu
        self.side_menu_width_ratio = 0.20  # 20% for side menu
        self.update_frame_sizes()  # Set initial size
        self.auto_refresh_var = tk.BooleanVar()
        self.device_data = {}
        self.logging_active = False
        self.charge_fet_status = False
        self.discharge_fet_status = False
        self.first_time_dashboard = True
        self.rs232_flag = False
        self.rs422_flag = False
        self.refresh_active = True
        self.mode_var = tk.StringVar(value="Testing Mode")  # Initialize mode_var here

        self.auto_refresh_interval = 5000

        self.style = ttk.Style()
        self.configure_styles()

        # Directory paths
        base_path = os.path.dirname(os.path.abspath(__file__))
        self.assets_path = os.path.join(base_path, "../assets/images/")

        # Create the main container frame
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True)

        active_protocol = get_active_protocol()
        print(f"{active_protocol} protocol")
        # Step 3: Update the flags based on the active protocol
        if active_protocol == "RS-232":
            self.rs232_flag = True
            self.rs422_flag = False
        elif active_protocol == "RS-422":
            self.rs232_flag = False
            self.rs422_flag = True

        print(f"{self.rs232_flag} self.rs232_flag")
        print(f"{self.rs422_flag} self.rs422_flag")

        if self.rs232_flag:
            self.device_data = rs232_device_data
            print("RS232 data updated")
        elif self.rs422_flag:
            self.device_data = rs422_device_data
            print("RS422 data updated")

        
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
            command=lambda: self.select_button(self.dashboard_button,self.show_dashboard),
            image=self.dashboard_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.dashboard_button.pack(fill="x", pady=5)
        self.info_button = ttk.Button(
            self.side_menu_frame,
            text=" Info        ",
            command=lambda: self.select_button(self.info_button, self.show_info),
            image=self.info_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.info_button.pack(fill="x", pady=5)
        if self.rs232_flag:
            self.control_button = ttk.Button(
                self.side_menu_frame,
                text=" Controls       ",
                command=lambda: self.select_button(self.control_button,self.show_control),
                image=self.loads_icon,
                compound="left",  # Place the icon to the left of the text
                bootstyle="info"
            )
            self.control_button.pack(fill="x", pady=2)

            self.report_button = ttk.Button(
                self.side_menu_frame,
                text=" Reports    ",
                command=lambda: self.select_button(self.report_button,self.show_report),
                image=self.download_icon,
                compound="left",  # Place the icon to the left of the text
                bootstyle="info"
            )
            self.report_button.pack(fill="x", pady=2)

        self.help_button = ttk.Button(
            self.side_menu_frame,
            text=" Help    ",
            command=lambda: self.select_button(self.help_button),
            image=self.help_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.help_button.pack(fill="x", pady=2)

        # Disconnect button
        self.disconnect_button = ttk.Button(
            self.side_menu_frame,
            text=" Disconnect",
            command=lambda: self.select_button(self.disconnect_button, self.on_disconnect),
            image=self.disconnect_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="danger"
        )
        self.disconnect_button.pack(side="bottom",pady=15)
         
        self.periodic_refresh()
        self.show_dashboard()
        

        # Bind the window resize event to adjust frame sizes dynamically
        self.master.bind('<Configure>', self.on_window_resize)
        # # Set up the auto-refresh to call `update_device_data` periodically
        # self.master.after(self.auto_refresh_interval, self.periodic_refresh)
        
        # Other initialization code...
    
    def periodic_refresh(self, initial_call=True):
        """Periodically refresh device data based on the selected flag (RS-232 or RS-422)."""
        # Step 1: Check if refresh is still active
        if not self.refresh_active:
            return  # Exit if the refresh has been stopped

        # Step 4: Update device data based on the selected protocol
        if self.rs232_flag:
            self.device_data = rs232_device_data
            print("RS232 data updated")
        elif self.rs422_flag:
            self.device_data = rs422_device_data
            print("RS422 data updated")

        # Step 5: Immediately update UI or display data  # If it's the first call, update immediately

        # Call this method again after the set interval, only if refresh is active
        self.master.after(3000, self.periodic_refresh, False)

    def update_info(self, event=None):
        """Update the Info page content based on the selected mode."""
        if self.selected_button == self.info_button:
            self.show_info()
        elif self.selected_button == self.dashboard_button:
            self.show_dashboard()
        elif self.selected_button == self.control_button:
            self.show_control()

    def start_logging(self):
        # Your existing start logging method...
        # Ensure that this also gets updated when logging
        self.periodic_refresh()

    def show_dashboard(self):
        self.clear_content_frame()

        if self.rs232_flag:
            # Create a Frame for the battery info details at the bottom
            info_frame = ttk.Labelframe(self.content_frame, text="Battery Information", bootstyle="dark", borderwidth=5, relief="solid")
            info_frame.pack(fill="x", expand=True, padx=10, pady=10)

            # Create and place the Battery Info labels within the info_frame
            battery_id_label = ttk.Label(info_frame, text="Battery ID:", font=("Arial", 10))
            battery_id_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")   
            battery_id_value = ttk.Label(info_frame, text=self.device_data.get('battery_id', 'N/A'), font=("Arial", 10, "bold"))
            battery_id_value.grid(row=0, column=1, padx=10, pady=5, sticky="w")   
            battery_sn_label = ttk.Label(info_frame, text="Battery S.No.:", font=("Arial", 10))
            battery_sn_label.grid(row=0, column=2, padx=10, pady=5, sticky="e")   
            battery_sn_value = ttk.Label(info_frame, text=self.device_data.get('serial_number', 'N/A'), font=("Arial", 10, "bold"))
            battery_sn_value.grid(row=0, column=3, padx=10, pady=5, sticky="w")   
            hw_ver_label = ttk.Label(info_frame, text="HW Ver:", font=("Arial", 10))
            hw_ver_label.grid(row=0, column=4, padx=10, pady=5, sticky="e")   
            hw_ver_value = ttk.Label(info_frame, text=self.device_data.get('hw_version', 'N/A'), font=("Arial", 10, "bold"))
            hw_ver_value.grid(row=0, column=5, padx=10, pady=5, sticky="w")   
            sw_ver_label = ttk.Label(info_frame, text="SW Ver:", font=("Arial", 10))
            sw_ver_label.grid(row=0, column=6, padx=10, pady=5, sticky="e")   
            sw_ver_value = ttk.Label(info_frame, text=self.device_data.get('sw_version', 'N/A'), font=("Arial", 10, "bold"))
            sw_ver_value.grid(row=0, column=7, padx=10, pady=5, sticky="w")

            for i in range(8):
                info_frame.grid_columnconfigure(i, weight=1)

            # Create a main frame to hold Charging, Discharge, and Battery Health sections
            main_frame = ttk.Frame(self.content_frame)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Charging Section
            bus1_frame = ttk.Labelframe(main_frame, text="Bus 1", bootstyle='dark', borderwidth=10, relief="solid")
            bus1_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

            # Charging Current (Row 0, Column 0)
            bus1_voltage_frame = ttk.Frame(bus1_frame)
            bus1_voltage_frame.grid(row=0, column=0, padx=70, pady=5, sticky="nsew", columnspan=1)
            bus1_voltage_label = ttk.Label(bus1_voltage_frame, text="After Diode Voltage", font=("Helvetica", 10, "bold"))
            bus1_voltage_label.pack(pady=(5, 5))
            bus1_voltage_meter = ttk.Meter(
                master=bus1_voltage_frame,
                metersize=180,
                amountused=self.device_data['bus_1_voltage_after_diode'],
                meterthickness=10,
                metertype="semi",
                subtext="After Diode Voltage",
                textright="V",
                amounttotal=120,
                bootstyle=self.get_gauge_style(self.device_data['bus_1_voltage_after_diode'], "bus_1_voltage_after_diode"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            bus1_voltage_meter.pack()

            # Charging Voltage (Row 0, Column 1)
            charging_voltage_frame = ttk.Frame(bus1_frame)
            charging_voltage_frame.grid(row=0, column=1, padx=70, pady=5, sticky="nsew", columnspan=1)
            charging_voltage_label = ttk.Label(charging_voltage_frame, text="Current Sensor 1", font=("Helvetica", 10, "bold"))
            charging_voltage_label.pack(pady=(5, 5)) 
            charging_voltage_meter = ttk.Meter(
                master=charging_voltage_frame,
                metersize=180,
                amountused=self.device_data['bus_1_current_sensor1'],
                meterthickness=10,
                metertype="semi",
                subtext="Current Sensor 1",
                textright="A",
                amounttotal=35,
                bootstyle=self.get_gauge_style(self.device_data['bus_1_current_sensor1'], "bus_1_current_sensor1"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            charging_voltage_meter.pack()

            # Battery Health Section
            charger_frame = ttk.Labelframe(main_frame, text="Charger", bootstyle='dark', borderwidth=10, relief="solid")
            charger_frame.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky="nsew")

            # Temperature (Top of Battery Health)
            charger_current_frame = ttk.Frame(charger_frame)
            charger_current_frame.pack(padx=10, pady=5, fill="both", expand=True)
            charger_current_label = ttk.Label(charger_current_frame, text="Charging Current", font=("Helvetica", 10, "bold"))
            charger_current_label.pack(pady=(10, 10))
            charger_current_meter = ttk.Meter(
                master=charger_current_frame,
                metersize=180,
                amountused=self.device_data['charger_input_current'],
                meterthickness=10,
                metertype="semi",
                subtext="Charging Current",
                textright="A",
                amounttotal=100,
                bootstyle=self.get_gauge_style(self.device_data['charger_input_current'], "charger_input_current"),
                stripethickness=8,
                subtextfont='-size 10'
            )
            charger_current_meter.pack()

            # Capacity (Bottom of Battery Health)
            charger_voltage_frame = ttk.Frame(charger_frame)
            charger_voltage_frame.pack(padx=10, pady=5, fill="both", expand=True)
            charger_voltage_label = ttk.Label(charger_voltage_frame, text="Charging Voltage", font=("Helvetica", 10, "bold"))
            charger_voltage_label.pack(pady=(10, 10))
            charger_voltage_meter = ttk.Meter(
                master=charger_voltage_frame,
                metersize=180,
                amountused=self.device_data['charger_output_current'],
                meterthickness=10,
                metertype="semi",
                subtext="Charging Voltage",
                textright="%",
                amounttotal=100,
                bootstyle=self.get_gauge_style(self.device_data['charger_output_current'], "charger_output_current"),
                stripethickness=8,
                subtextfont='-size 10'
            )
            charger_voltage_meter.pack()

            # Discharging Section
            bus2_frame = ttk.Labelframe(main_frame, text="Bus 2", bootstyle='dark', borderwidth=10, relief="solid")
            bus2_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

            # Discharging Current (Row 1, Column 0)
            bus2_current_frame = ttk.Frame(bus2_frame)
            bus2_current_frame.grid(row=1, column=0, padx=70, pady=5, sticky="nsew", columnspan=1)
            discharging_current_label = ttk.Label(bus2_current_frame, text="After Diode Voltage", font=("Helvetica", 10, "bold"))
            discharging_current_label.pack(pady=(5, 5))
            # Variable to control the max value of the discharging current gauge

            # Discharging Current Gauge
            self.discharging_current_meter = ttk.Meter(
                master=bus2_current_frame,
                metersize=180,
                amountused=self.device_data['bus_2_voltage_after_diode'],  # Assuming this is the discharging current
                meterthickness=10,
                metertype="semi",
                subtext="After Diode Voltage",
                textright="V",
                amounttotal=100,
                bootstyle=self.get_gauge_style(self.device_data['bus_2_voltage_after_diode'], "bus_2_voltage_after_diode"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            self.discharging_current_meter.pack()

            # Discharging Voltage (Row 1, Column 1)
            discharging_voltage_frame = ttk.Frame(bus2_frame)
            discharging_voltage_frame.grid(row=1, column=1, padx=70, pady=5, sticky="nsew", columnspan=1)
            discharging_voltage_label = ttk.Label(discharging_voltage_frame, text="Current Sensor 1", font=("Helvetica", 10, "bold"))
            discharging_voltage_label.pack(pady=(5, 5))
            discharging_voltage_meter = ttk.Meter(
                master=discharging_voltage_frame,
                metersize=180,
                amountused=self.device_data['bus_2_current_sensor1'],  # Assuming this is the discharging voltage
                meterthickness=10,
                metertype="semi",
                subtext="Current Sensor 1",
                textright="A",
                amounttotal=100,
                bootstyle=self.get_gauge_style(self.device_data['bus_2_current_sensor1'], "bus_2_current_sensor1"),
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
        
        elif self.rs422_flag:
            # Create a main frame to hold Charging, Discharge, and Battery Health sections
            main_frame = ttk.Frame(self.content_frame)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Charging Section
            bus1_frame = ttk.Labelframe(main_frame, text="Bus", bootstyle='dark', borderwidth=10, relief="solid")
            bus1_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

            # Charging Current (Row 0, Column 0)
            bus1_voltage_frame = ttk.Frame(bus1_frame)
            bus1_voltage_frame.grid(row=0, column=0, padx=130, pady=5, sticky="nsew", columnspan=1)
            bus1_voltage_label = ttk.Label(bus1_voltage_frame, text="EB1 Current", font=("Helvetica", 10, "bold"))
            bus1_voltage_label.pack(pady=(5, 5))
            bus1_voltage_meter = ttk.Meter(
                master=bus1_voltage_frame,
                metersize=200,
                amountused=self.device_data['eb_1_current'],
                meterthickness=10,
                metertype="semi",
                subtext="EB1 Current",
                textright="A",
                amounttotal=120,
                bootstyle=self.get_gauge_style_422(self.device_data['eb_1_current'], "eb_1_current"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            bus1_voltage_meter.pack()

            # Charging Voltage (Row 0, Column 1)
            charging_voltage_frame = ttk.Frame(bus1_frame)
            charging_voltage_frame.grid(row=0, column=1, padx=130, pady=5, sticky="nsew", columnspan=1)
            charging_voltage_label = ttk.Label(charging_voltage_frame, text="EB2 Current", font=("Helvetica", 10, "bold"))
            charging_voltage_label.pack(pady=(5, 5)) 
            charging_voltage_meter = ttk.Meter(
                master=charging_voltage_frame,
                metersize=200,
                amountused=self.device_data['eb_2_current'],
                meterthickness=10,
                metertype="semi",
                subtext="EB2 Current",
                textright="A",
                amounttotal=35,
                bootstyle=self.get_gauge_style_422(self.device_data['eb_2_current'], "eb_2_current"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            charging_voltage_meter.pack()

            # Discharging Section
            health_frame = ttk.Labelframe(main_frame, text="Health Status", bootstyle='dark', borderwidth=10, relief="solid")
            health_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

            # Discharging Current (Row 1, Column 0)
            bus2_current_frame = ttk.Frame(health_frame)
            bus2_current_frame.grid(row=1, column=0, padx=130, pady=5, sticky="nsew", columnspan=1)
            discharging_current_label = ttk.Label(bus2_current_frame, text="Charging Current", font=("Helvetica", 10, "bold"))
            discharging_current_label.pack(pady=(5, 5))
            self.discharging_current_meter = ttk.Meter(
                master=bus2_current_frame,
                metersize=200,
                amountused=self.device_data['charge_current'],  # Assuming this is the discharging current
                meterthickness=10,
                metertype="semi",
                subtext="Charging Current",
                textright="A",
                amounttotal=100,
                bootstyle=self.get_gauge_style_422(self.device_data['charge_current'], "charge_current"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            self.discharging_current_meter.pack()

            # Discharging Voltage (Row 1, Column 1)
            discharging_voltage_frame = ttk.Frame(health_frame)
            discharging_voltage_frame.grid(row=1, column=1, padx=130, pady=5, sticky="nsew", columnspan=1)
            discharging_voltage_label = ttk.Label(discharging_voltage_frame, text="Temperature", font=("Helvetica", 10, "bold"))
            discharging_voltage_label.pack(pady=(5,5))
            discharging_voltage_meter = ttk.Meter(
                master=discharging_voltage_frame,
                metersize=200,
                amountused=self.device_data['temperature'],  # Assuming this is the discharging voltage
                meterthickness=10,
                metertype="semi",
                subtext="Temperature",
                textright="C",
                amounttotal=100,
                bootstyle=self.get_gauge_style_422(self.device_data['temperature'], "temperature"),
                stripethickness=10,
                subtextfont='-size 10'
            )
            discharging_voltage_meter.pack()

            # Configure row and column weights to ensure even distribution
            for i in range(2):
                main_frame.grid_rowconfigure(i, weight=1)
            for j in range(2):
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
        label.pack(pady=15)
        # Create a text box (entry widget)
        self.write_entry = tk.Entry(self.content_frame, width=50)
        self.write_entry.pack(pady=15)

        # Create a button to send the entered text
        send_button = tk.Button(self.content_frame, text="Send", command=self.send_serial_data)
        send_button.pack(pady=10)

        self.select_button(self.help_button)

    def send_serial_data(self):
        """Retrieve data from the entry field and send it via the serial port."""
        data = self.write_entry.get()  # Get the data entered in the entry widget
        if data:
            print(f"Sending data: {data}")
            # start_periodic_sending()
            # read_data()
        else:
            messagebox.showwarning("Input Error", "Please enter data to send.")
    
    def show_report(self):
        self.clear_content_frame()
        latest_data = {}
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
        ttk.Label(battery_info_frame, text="Project").grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="e")
        # Add the Combobox for the Project Dropdown
        project_options = ["TAPAS", "ARCHER", "SETB"]  # List of project options
        project_combobox = ttk.Combobox(battery_info_frame, values=project_options, state="readonly", width=23)  # Dropdown with readonly state
        project_combobox.grid(row=0, column=2, columnspan=2, padx=10, pady=10, sticky="w")
        project_combobox.set("Select a Project")
        # Add individual fields for Battery Info
        ttk.Label(battery_info_frame, text="Nomenclature").grid(row=1,  column=0,columnspan=2, padx=10, pady=10, sticky="e")
        battery_nomenclature_entry = ttk.Entry(battery_info_frame, width=25)
        battery_nomenclature_entry.insert(0, str("28V Lithium Ion Battery Unit"))
        battery_nomenclature_entry.grid(row=1, column=2,columnspan=2, padx=10, pady=10, sticky="w")
        battery_nomenclature_entry.config(state="readonly")
        ttk.Label(battery_info_frame, text="Serial Number").grid(row=2,  column=0,columnspan=2, padx=10, pady=10, sticky="e")
        serial_number_entry = ttk.Entry(battery_info_frame, width=25)
        serial_number_entry.insert(0, str(latest_data.get("Serial Number", "")))
        serial_number_entry.grid(row=2, column=2,columnspan=2, padx=10, pady=10, sticky="w")
        ttk.Label(battery_info_frame, text="Battery Chemistry").grid(row=3,  column=0,columnspan=2, padx=10, pady=10, sticky="e")
        battery_chemistry_entry = ttk.Entry(battery_info_frame, width=25)
        battery_chemistry_entry.insert(0, str("Lithium Ion"))
        battery_chemistry_entry.grid(row=3, column=2,columnspan=2, padx=10, pady=10, sticky="w")
        battery_chemistry_entry.config(state="readonly")
        ttk.Label(battery_info_frame, text="Battery Capacity(AH)").grid(row=4,  column=0,columnspan=2, padx=10, pady=10, sticky="e")
        battery_capacity_entry = ttk.Entry(battery_info_frame, width=25)
        battery_capacity_entry.insert(0, str("55 Ah"))
        battery_capacity_entry.grid(row=4, column=2,columnspan=2, padx=10, pady=10, sticky="w")
        battery_capacity_entry.config(state="readonly")
        ttk.Label(battery_info_frame, text="Cycle Count").grid(row=5,  column=0,columnspan=2, padx=10, pady=10, sticky="e")
        cycle_count_entry = ttk.Entry(battery_info_frame, width=25)
        cycle_count_entry.insert(0, str(0))
        cycle_count_entry.grid(row=5, column=2,columnspan=2, padx=10, pady=10, sticky="w")

        # Charging Info Frame (row 1-2, col 2-4)
        charging_frame = ttk.LabelFrame(main_frame, text="Charging Info", borderwidth=10)
        charging_frame.grid(row=1, column=4, columnspan=4, padx=10, pady=10, sticky="nsew")
        ttk.Label(charging_frame, text="Duration").grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        duration_entry = ttk.Entry(charging_frame)
        duration_entry.insert(0, str(latest_data.get("Charging Duration", 0)))  # Insert the fetched duration or a default message
        duration_entry.grid(row=0, column=3, columnspan=2, padx=10, pady=10, sticky="w")
        ttk.Label(charging_frame, text="Date").grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="e")
        date_entry = ttk.Entry(charging_frame)
        date_entry.insert(0, str(latest_data.get("Charging Date", 0)))  # Insert the fetched date or a default message
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
                "Battery Nomenclature": battery_nomenclature_entry.get(),
                "Serial Number": serial_number_entry.get(),
                "Battery Chemistry": battery_chemistry_entry.get(),
                "Cycle Count": cycle_count_entry.get(),
                "Battery Capacity": battery_capacity_entry.get(),
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
    
        # Section: Device Connection
        device_connect_frame = ttk.Labelframe(load_frame, bootstyle='dark', text="Connection", padding=10, borderwidth=10, relief="solid")
        device_connect_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew", columnspan=2)
    
        # Connect Button
        self.connect_button = ttk.Button(device_connect_frame, text="Connect to Chroma", command=self.connect_device)
        self.connect_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
    
        # Status Label
        self.status_label = ttk.Label(device_connect_frame, text="Status: Disconnected")
        self.status_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Function to toggle Bus 1
        def toggle_bus1():
            if rs232_write['cmd_byte_1'] & 0x01 == 0:  # If Bus 1 is OFF (Bit 0 = 0)
                rs232_write['cmd_byte_1'] |= 0x01  # Set Bit 0 to 1 to turn ON Bus 1
                time.sleep(3)
                bus1_status = self.device_data.get('bus1_status', 0)
                if bus1_status == 1:
                    bus1_control_status.config(text="ON", bootstyle="success")
                else:
                    bus1_control_status.config(text="OFF", bootstyle="danger")
            else:
                rs232_write['cmd_byte_1'] &= 0xFE  # Reset Bit 0 to 0 to turn OFF Bus 1
                time.sleep(3)
                bus1_status = self.device_data.get('bus1_status', 0)
                if bus1_status == 1:
                    bus1_control_status.config(text="ON", bootstyle="success")
                else:
                    bus1_control_status.config(text="OFF", bootstyle="danger")
            print(f"Bus 1 Status: {rs232_write['cmd_byte_1']:08b}")

        # Function to toggle Bus 2
        def toggle_bus2():
            if rs232_write['cmd_byte_1'] & 0x02 == 0:  # If Bus 2 is OFF (Bit 1 = 0)
                rs232_write['cmd_byte_1'] |= 0x02  # Set Bit 1 to 1 to turn ON Bus 2
                time.sleep(3)
                bus2_status = self.device_data.get('bus2_status', 0)
                if bus2_status == 1:
                    bus2_control_status.config(text="ON", bootstyle="success")
                else:
                    bus2_control_status.config(text="OFF", bootstyle="danger")
            else:
                rs232_write['cmd_byte_1'] &= 0xFD  # Reset Bit 1 to 0 to turn OFF Bus 2
                time.sleep(3)
                bus2_status = self.device_data.get('bus2_status', 0)
                if bus2_status == 0:
                    bus2_control_status.config(text="OFF", bootstyle="danger")
                else:
                    bus2_control_status.config(text="ON", bootstyle="success")
            print(f"Bus 2 Status: {rs232_write['cmd_byte_1']:08b}")

        # Function to toggle Charger
        def toggle_charger():
            if rs232_write['cmd_byte_1'] & 0x04 == 0:  # If Charger is OFF (Bit 2 = 0)
                rs232_write['cmd_byte_1'] |= 0x04  # Set Bit 2 to 1 to turn ON Charger
                time.sleep(3)
                charger_status = self.device_data.get('charging_on_off_status', 0)
                if charger_status == 1:
                    charger_control_status.config(text="ON", bootstyle="success")
                else:
                    charger_control_status.config(text="OFF", bootstyle="danger")
            else:
                rs232_write['cmd_byte_1'] &= 0xFB  # Reset Bit 2 to 0 to turn OFF Charger
                time.sleep(3)
                charger_status = self.device_data.get('charging_on_off_status', 0)
                if charger_status == 1:
                    charger_control_status.config(text="OFF", bootstyle="danger")
                else:
                    charger_control_status.config(text="ON", bootstyle="success")
            print(f"Charger Status: {rs232_write['cmd_byte_1']:08b}")

        # Function to toggle Charger Output Relay
        def toggle_output_relay():
            if rs232_write['cmd_byte_1'] & 0x08 == 0:  # If Output Relay is OFF (Bit 3 = 0)
                rs232_write['cmd_byte_1'] |= 0x08  # Set Bit 3 to 1 to turn ON Output Relay
                time.sleep(3)
                charger_relay_status = self.device_data.get('charger_relay_status', 0)
                if charger_relay_status == 1:
                    output_relay_control_status.config(text="ON", bootstyle="success")
                else:
                    output_relay_control_status.config(text="OFF", bootstyle="danger")
            else:
                rs232_write['cmd_byte_1'] &= 0xF7  # Reset Bit 3 to 0 to turn OFF Output Relay
                time.sleep(3)
                charger_relay_status = self.device_data.get('charger_relay_status', 0)
                if charger_relay_status == 0:
                    output_relay_control_status.config(text="OFF", bootstyle="danger")
                else:
                    output_relay_control_status.config(text="ON", bootstyle="success")
            print(f"Output Relay Status: {rs232_write['cmd_byte_1']:08b}")

        def toggle_reset():
            # Write the reset command value (0x23) to cmd_byte_2
            rs232_write['cmd_byte_2'] = 0x23  # Set Byte 2 to the reset command value
        
            # Update the reset button UI (could be any relevant UI change)
            messagebox.showinfo("Reset Triggered", "The reset command has been successfully triggered.")
            
            # Simulate sending the reset command and waiting for some time (if necessary)
            time.sleep(3)  # Short delay to simulate the reset process
            
            # You can reset the UI back to normal after the reset process completes
            reset_button.config(text="RESET", bootstyle="danger")
            
            print(f"Reset Command Triggered: {rs232_write['cmd_byte_2']:08b}")


        # Section: Charger Control Frame
        charger_control_frame = ttk.Labelframe(load_frame, text="Charger Controls", bootstyle="dark", borderwidth=5, relief="solid")
        charger_control_frame.grid(row=0, column=2, padx=5, pady=5, sticky="nsew", columnspan=2)
    
        ttk.Label(charger_control_frame, text="Charger On/Off").grid(row=0, column=0, padx=5, pady=5)
        charger_control_status = ttk.Button(charger_control_frame, text="OFF", bootstyle="danger",command=toggle_charger)
        charger_control_status.grid(row=0, column=1, padx=5, pady=5)
    
        ttk.Label(charger_control_frame, text="Output Relay On/Off").grid(row=1, column=0, padx=5, pady=5)
        output_relay_control_status = ttk.Button(charger_control_frame, text="OFF", bootstyle="danger", command=toggle_output_relay)
        output_relay_control_status.grid(row=1, column=1, padx=5, pady=5)
    
        # Section: Right Control Frame (BUS Controls)
        right_control_frame = ttk.Labelframe(load_frame, text="Controls", bootstyle="dark", borderwidth=5, relief="solid")
        right_control_frame.grid(row=0, column=4, padx=5, pady=5, sticky="nsew", columnspan=2)  # Spans across 2 columns
    
        # BUS1 Control
        ttk.Label(right_control_frame, text="BUS1 On/Off").grid(row=0, column=0, padx=5, pady=5)
        bus1_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger",command=toggle_bus1)
        bus1_control_status.grid(row=0, column=1, padx=5, pady=5)
    
        # BUS2 Control
        ttk.Label(right_control_frame, text="BUS2 On/Off").grid(row=1, column=0, padx=5, pady=5)
        bus2_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger",command=toggle_bus2)
        bus2_control_status.grid(row=1, column=1, padx=5, pady=5)
    
        # Reset Button
        ttk.Label(right_control_frame, text="Reset").grid(row=2, column=0, padx=5, pady=5)
        reset_button = ttk.Button(right_control_frame, text="RESET", bootstyle="danger",command=toggle_reset)
        reset_button.grid(row=2, column=1, padx=5, pady=5)
    
        # Check the selected mode
        selected_mode = self.mode_var.get()
    
        # Section: Testing Mode
        testing_frame = ttk.Labelframe(load_frame, text="Testing Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
        testing_frame.grid(row=1, column=0, columnspan=6, padx=10, pady=10, sticky="nsew")  # Spanning across 6 columns
        test_button = ttk.Button(testing_frame, text="Set 25A and Turn ON Load", command=set_l1_50a_and_turn_on)
        test_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
    
        maintenance_button = ttk.Button(testing_frame, text="Set 50A and Turn ON Load", command=set_l1_100a_and_turn_on)
        maintenance_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
    
        # Turn Off Load Button
        turn_off_button = ttk.Button(testing_frame, text="Turn OFF Load", command=custom_turn_off)
        turn_off_button.grid(row=0, column=2, padx=10, pady=10, sticky="ew")
    
        # Section: Custom Mode
        custom_frame = ttk.Labelframe(load_frame, text="Custom Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
        custom_frame.grid(row=2, column=0, columnspan=6, padx=5, pady=5, sticky="nsew")  # Spanning across 6 columns
    
        custom_label = ttk.Label(custom_frame, text="Set Custom Current (A):", font=("Helvetica", 12), width=20, anchor="e")
        custom_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    
        def save_custom_value():
            """Saves the custom current value and sets it to the Chroma device."""
            custom_value = self.custom_current_entry.get()
            if custom_value.isdigit():  # Basic validation to ensure it's a number
                set_custom_l1_value(int(custom_value))
            else:
                print("Invalid input: Please enter a valid number.")
    
        self.custom_current_entry = ttk.Entry(custom_frame)
        self.custom_current_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
        custom_button = ttk.Button(custom_frame, text="Save", command=save_custom_value)
        custom_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    
        # Create a toggle button for turning load ON and OFF
        self.load_status = tk.BooleanVar()  # Boolean variable to track load status (ON/OFF)
        self.load_status.set(False)  # Initial state is OFF
    
        def toggle_load():
            if self.load_status.get():
                stop_fetching_voltage()
                custom_turn_off()  # Call the Turn OFF function
                toggle_button.config(text="Turn ON Load", bootstyle="success")  # Change to green when OFF
                self.load_status.set(False)  # Update state to OFF
            else:
                custom_turn_on()  # Call the Turn ON function
                toggle_button.config(text="Turn OFF Load", bootstyle="danger")  # Change to red when ON
                self.load_status.set(True)  # Update state to ON
    
        # Toggle button for Load ON/OFF
        toggle_button = ttk.Button(custom_frame, text="Turn ON Load", command=toggle_load, bootstyle='success')
        toggle_button.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
    
        # Make the rows and columns expand as needed for content_frame and load_frame
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)  # Ensures the load_frame expands fully
        self.content_frame.grid_columnconfigure(0, weight=1)
    
        load_frame.grid_rowconfigure(0, weight=1)  # Device Connection frame
        load_frame.grid_rowconfigure(1, weight=1)  # Testing frame
        load_frame.grid_rowconfigure(2, weight=1)  # Custom frame
        load_frame.grid_columnconfigure(0, weight=1)  # Ensure columns expand to fill space
    
        # Highlight the control button in the side menu
        self.select_button(self.control_button)

    def maintains_mode_load(self):
        set_l1_50a_and_turn_on()

    def testing_mode_load(self):
        set_l1_25a_and_turn_on()

    def load_off(self):
        stop_fetching_voltage()
        custom_turn_off()

    def find_device(self):
        """Find available devices."""
        result = self.chroma.find_and_connect()  # Attempt to find the Chroma device
        self.status_label.config(text=result)
    
    def create_entries(self, parent_frame, fields, latest_data):
        """Helper function to create label and entry pairs inside a given frame."""
        entries = {}
        for i, field in enumerate(fields):
            label = ttk.Label(parent_frame, text=field)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="e")

            value = latest_data.get(field, "")
            entry = ttk.Entry(parent_frame)
            entry.insert(0, value)  # Insert the current value into the entry field
            entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")

            entries[field] = entry  # Store the entry widget for later retrieval
        return entries

    def connect_device(self):
        """Connect to the Chroma device."""
        result = self.chroma.find_and_connect()  # This can be the same method or a separate one for specific device connection
        self.status_label.config(text=result)

    def set_load_and_start_timer(self):
        current = self.current_entry.get()
        time_duration = int(self.time_entry.get())
        result = self.chroma.set_load_with_timer(current, time_duration)
        self.status_label.config(text=result)

    def manual_turn_on(self):
        result = self.chroma.turn_load_on()
        self.status_label.config(text=result)

    def manual_turn_off(self):
        result = self.chroma.turn_load_off()
        self.status_label.config(text=result)

    
    def configure_styles(self):
            # Configure custom styles
            self.style.configure("Custom.Treeview", font=("Helvetica", 10), rowheight=25, background="#f0f0f0", fieldbackground="#e8f4ec")
            self.style.configure("Custom.Treeview.Heading", font=("Helvetica", 10, "bold"), background="#333", foreground="#e8f4ec", padding=5)
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
        messagebox.showinfo("Info!", "Connection Disconnect!")
        self.rs232_flag = False
        self.rs422_flag = False
        stop_communication()  # Stop RS-232/422 communication

        self.clear_content_frame()  # Clear the current UI
        self.refresh_active = False  # Stop periodic refresh

        if self.main_window.rs_battery_info is not None:
            del self.main_window.rs_battery_info
            self.main_window.rs_battery_info = None
    
        self.main_frame.pack_forget()  # Hide the current window
        self.main_window.show_main_window()  # Show the main window

    def show_info(self, event=None):
        self.clear_content_frame()

        style = ttk.Style()

        # Configure the Labelframe title font to be bold
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold"))

        # Determine if we are in Testing Mode or Maintenance Mode
        selected_mode = self.mode_var.get()
        self.limited = selected_mode == "Testing Mode"

        if self.rs232_flag:
            # Create a Frame to hold the battery info at the top of the content frame
            info_frame = ttk.Labelframe(self.content_frame, text="Battery Info", bootstyle="dark",borderwidth=5, relief="solid")
            info_frame.pack(fill="x", padx=10, pady=10)

            # Create and place the Battery Info labels within the info_frame
            battery_id_label = ttk.Label(info_frame, text="Battery ID:", font=("Arial", 10))
            battery_id_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")   
            battery_id_value = ttk.Label(info_frame, text=self.device_data.get('battery_id', 'N/A'), font=("Arial", 10, "bold"))
            battery_id_value.grid(row=0, column=1, padx=10, pady=5, sticky="w")   
            battery_sn_label = ttk.Label(info_frame, text="Battery S.No.:", font=("Arial", 10))
            battery_sn_label.grid(row=0, column=2, padx=10, pady=5, sticky="e")   
            battery_sn_value = ttk.Label(info_frame, text=self.device_data.get('serial_number', 'N/A'), font=("Arial", 10, "bold"))
            battery_sn_value.grid(row=0, column=3, padx=10, pady=5, sticky="w")   
            hw_ver_label = ttk.Label(info_frame, text="HW Ver:", font=("Arial", 10))
            hw_ver_label.grid(row=0, column=4, padx=10, pady=5, sticky="e")   
            hw_ver_value = ttk.Label(info_frame, text=self.device_data.get('hw_version', 'N/A'), font=("Arial", 10, "bold"))
            hw_ver_value.grid(row=0, column=5, padx=10, pady=5, sticky="w")   
            sw_ver_label = ttk.Label(info_frame, text="SW Ver:", font=("Arial", 10))
            sw_ver_label.grid(row=0, column=6, padx=10, pady=5, sticky="e")   
            sw_ver_value = ttk.Label(info_frame, text=self.device_data.get('sw_version', 'N/A'), font=("Arial", 10, "bold"))
            sw_ver_value.grid(row=0, column=7, padx=10, pady=5, sticky="w")

            # Configure equal weight for the columns to distribute space equally
            for i in range(8):
                info_frame.grid_columnconfigure(i, weight=1)

            # Create a main frame to hold Charging, Discharge, and Battery Health sections
            main_frame = ttk.Frame(self.content_frame)
            main_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Battery section
            battery_frame = ttk.Labelframe(main_frame, text="Battery", bootstyle='dark',borderwidth=5, relief="solid")
            battery_frame.grid(row=0, column=0, columnspan=6, rowspan=5, padx=10, pady=5, sticky="nsew")

            # Parameters Frame inside Battery section
            parameters_frame = ttk.Labelframe(battery_frame, text="Parameters", bootstyle='dark',borderwidth=5, relief="solid")
            parameters_frame.grid(row=0, column=0, columnspan=6, rowspan=4, padx=15, pady=25, sticky="nsew")

            # Create labels for Voltage, Temperature, and Status headers
            ttk.Label(parameters_frame, text="").grid(row=0, column=0, padx=5, pady=5)  # Empty top-left cell
            ttk.Label(parameters_frame, text="Voltage").grid(row=1, column=0, padx=5, pady=5)
            ttk.Label(parameters_frame, text="Temperature").grid(row=2, column=0, padx=5, pady=5)
            ttk.Label(parameters_frame, text="Status").grid(row=3, column=0, padx=5, pady=5)  # Adding a Status label


            # Create labels and entries for each BAT (Voltage, Temperature, Status)
            for i in range(7):
                ttk.Label(parameters_frame, text=f"BAT{i+1}").grid(row=0, column=i+1, padx=5, pady=5)

                # Set the values from the rs232_device_data dictionary
                voltage_value = self.device_data.get(f'cell_{i+1}_voltage', 'N/A')
                temp_value = self.device_data.get(f'cell_{i+1}_temp', 'N/A')

                # Voltage label styled like an Entry (with border)
                voltage_label = ttk.Label(parameters_frame, text=voltage_value, relief="solid", borderwidth=2, width=10, anchor="center")
                voltage_label.grid(row=1, column=i+1, padx=5, pady=5)

                # Temperature label styled like an Entry (with border)
                charger_current_label = ttk.Label(parameters_frame, text=temp_value, relief="solid", borderwidth=2, width=10, anchor="center")
                charger_current_label.grid(row=2, column=i+1, padx=5, pady=5)

                cell_status_circle_color = "green" if self.device_data.get('cell_{i+1}_status') == 1 else "red"
                # Status label with a round dot representing status
                status_label = ttk.Label(parameters_frame, text="●", font=("Arial", 24), foreground=cell_status_circle_color, anchor="center")
                status_label.grid(row=3, column=i+1, padx=5, pady=5)
  

            # Configure equal weight for the columns to distribute space equally
            for i in range(8):
                parameters_frame.grid_columnconfigure(i, weight=1)     

            # IC Temperature Frame
            ic_charger_current_frame = ttk.Labelframe(battery_frame, text="IC Temperature", bootstyle='dark',borderwidth=5, relief="solid")
            ic_charger_current_frame.grid(row=4, column=0, columnspan=8, rowspan=1, padx=15, pady=10, sticky="nsew")
            ic_charger_current_label = ttk.Label(ic_charger_current_frame, text=self.device_data.get('ic_temp'), relief="solid", borderwidth=2, width=10, anchor="center")
            ic_charger_current_label.grid(row=0, column=1, padx=5, pady=5)

            # Configure equal weight for the columns to distribute space equally
            ic_charger_current_frame.grid_columnconfigure(1, weight=1)
            
            battery_frame.grid_columnconfigure(1, weight=1)

            # Charger section
            charger_frame = ttk.Labelframe(main_frame, text="Charger", bootstyle='dark',borderwidth=5, relief="solid")
            charger_frame.grid(row=4, column=6, columnspan=6, rowspan=8, padx=10, pady=5, sticky="nsew")

            ttk.Label(charger_frame, text="Input").grid(row=1, column=1, padx=5, pady=15)

            ttk.Label(charger_frame, text="Output").grid(row=1, column=3, padx=5, pady=15)

            ttk.Label(charger_frame, text="Voltage").grid(row=2, column=0, padx=5, pady=15)

            charger_voltage_output = ttk.Label(charger_frame, text=self.device_data.get('charger_output_voltage'), relief="solid", borderwidth=2, width=10, anchor="center")
            charger_voltage_output.grid(row=2, column=3, padx=5, pady=15)

            ttk.Label(charger_frame, text="Current").grid(row=3, column=0, padx=5, pady=15)

            charger_current_input = ttk.Label(charger_frame, text=self.device_data.get('charger_input_current'), relief="solid", borderwidth=2, width=10, anchor="center")
            charger_current_input.grid(row=3, column=1, padx=5, pady=15)

            charger_current_output = ttk.Label(charger_frame, text=self.device_data.get('charger_output_current'), relief="solid", borderwidth=2, width=10, anchor="center")
            charger_current_output.grid(row=3, column=3, padx=5, pady=15)

            # Configure equal weight for the columns to distribute space equally
            for i in range(2):
                charger_frame.grid_columnconfigure(i, weight=1)

            # Bus section
            bus_frame = ttk.Labelframe(main_frame, text="BUS", bootstyle='dark',borderwidth=5, relief="solid")
            bus_frame.grid(row=0, column=6, columnspan=6, rowspan=4, padx=10, pady=5, sticky="nsew")

            ttk.Label(bus_frame, text="BUS1").grid(row=1, column=1, padx=5, pady=15)

            ttk.Label(bus_frame, text="BUS2").grid(row=1, column=3, padx=5, pady=15)

            ttk.Label(bus_frame, text="Voltage Before Diode").grid(row=2, column=0, padx=5, pady=15)
            voltage_before_diode_bus1 = ttk.Label(bus_frame, text=self.device_data.get('bus_1_voltage_before_diode'), relief="solid", borderwidth=2, width=10, anchor="center")
            voltage_before_diode_bus1.grid(row=2, column=1, padx=5, pady=15,sticky="nsew")
            voltage_before_diode_bus2 = ttk.Label(bus_frame, text=self.device_data.get('bus_2_voltage_before_diode'), relief="solid", borderwidth=2, width=10, anchor="center")
            voltage_before_diode_bus2.grid(row=2, column=3, padx=5, pady=15,sticky="nsew")

            ttk.Label(bus_frame, text="Voltage After Diode").grid(row=3, column=0, padx=5, pady=15)
            voltage_after_diode_bus1 = ttk.Label(bus_frame, text=self.device_data.get('bus_1_voltage_after_diode'), relief="solid", borderwidth=2, width=10, anchor="center")
            voltage_after_diode_bus1.grid(row=3, column=1, padx=5, pady=15,sticky="nsew")
            voltage_after_diode_bus2 = ttk.Label(bus_frame, text=self.device_data.get('bus_2_voltage_after_diode'), relief="solid", borderwidth=2, width=10, anchor="center")
            voltage_after_diode_bus2.grid(row=3, column=3, padx=5, pady=15,sticky="nsew")

            ttk.Label(bus_frame, text="Current").grid(row=4, column=0, padx=5, pady=15)
            current_sensor1_bus1 = ttk.Label(bus_frame, text=self.device_data.get('bus_1_current_sensor1'), relief="solid", borderwidth=2, width=10, anchor="center")
            current_sensor1_bus1.grid(row=4, column=1, padx=5, pady=15, sticky="nsew")

            current_sensor2_bus2 = ttk.Label(bus_frame, text=self.device_data.get('bus_2_current_sensor1'), relief="solid", borderwidth=2, width=10, anchor="center")
            current_sensor2_bus2.grid(row=4, column=3, padx=5, pady=15, sticky="nsew")

            # Configure equal weight for the columns to distribute space equally
            for i in range(3):
                bus_frame.grid_columnconfigure(i, weight=1)

            # Heater Pad section
            heater_frame = ttk.Labelframe(main_frame, text="Heater Pad", bootstyle='dark',borderwidth=5, relief="solid")
            heater_frame.grid(row=9, column=0, padx=5, pady=5, sticky="nsew")

            ttk.Label(heater_frame, text="On/Off Status").grid(row=0, column=0, padx=5, pady=5)
            heater_circle_color = "green" if self.device_data.get('heater_pad') == 1 else "red"
            ttk.Label(heater_frame, text="●", foreground=heater_circle_color, font=("Arial", 24)).grid(row=0, column=1, padx=5, pady=5)

            # Configure equal weight for the columns to distribute space equally
            heater_frame.grid_columnconfigure(1, weight=1)

            # Heater Pad section
            status_frame = ttk.Labelframe(main_frame, text="Status", bootstyle='dark',borderwidth=5, relief="solid")
            status_frame.grid(row=6, column=1,rowspan=6,columnspan=5, padx=2, pady=5, sticky="nsew")

            ttk.Label(status_frame, text="Constant Voltage Mode").grid(row=0, column=0, padx=5, pady=0)
            con_volt_circle_color = "green" if self.device_data.get('constant_voltage_mode') == 1 else "red"
            ttk.Label(status_frame, text="●", foreground=con_volt_circle_color, font=("Arial", 24)).grid(row=0, column=1, padx=5, pady=0)

            ttk.Label(status_frame, text="Constant Current Mode").grid(row=1, column=0, padx=5, pady=5)
            con_current_circle_color = "green" if self.device_data.get('constant_current_mode') == 1 else "red"
            ttk.Label(status_frame, text="●", foreground=con_current_circle_color, font=("Arial", 24)).grid(row=1, column=1, padx=5, pady=5)

            ttk.Label(status_frame, text="Input Under Voltage").grid(row=0, column=3, padx=5, pady=5)
            under_volt_circle_color = "green" if self.device_data.get('input_under_voltage') == 1 else "red"
            ttk.Label(status_frame, text="●", foreground=under_volt_circle_color, font=("Arial", 24)).grid(row=0, column=4, padx=5, pady=5)

            ttk.Label(status_frame, text="Output Over Current").grid(row=1, column=3, padx=5, pady=5)
            over_current__circle_color = "green" if self.device_data.get('output_over_current') == 1 else "red"
            ttk.Label(status_frame, text="●", foreground=over_current__circle_color, font=("Arial", 24)).grid(row=1, column=4, padx=5, pady=5)

            # Configure equal weight for the columns to distribute space equally
            for i in range(4):
                status_frame.grid_columnconfigure(i, weight=1)


            # Configure row and column weights for the main frame
            for i in range(12):
                main_frame.grid_rowconfigure(i, weight=1)
            for j in range(12):
                main_frame.grid_columnconfigure(j, weight=1)
        
        if self.rs422_flag:
            # Create a Frame to hold the battery info at the top of the content frame
            rs422_info_frame = ttk.Labelframe(self.content_frame, text="RS-422", bootstyle="dark", borderwidth=5, relief="solid")
            rs422_info_frame.pack(fill="x", padx=5, pady=5)

            # Heater Pad/Charger Relay Status Frame
            heater_pad_frame = ttk.Labelframe(rs422_info_frame, text="Heater Pad/Charger Relay Status", bootstyle='dark', borderwidth=5, relief="solid")
            heater_pad_frame.grid(row=0, column=0, columnspan=6,rowspan=2, padx=15, pady=10, sticky="nsew")
            ttk.Label(heater_pad_frame, text="Heater Pad/Charger Relay Status").grid(row=0, column=0, padx=5, pady=5)
            heater_status_circle_color = "green" if self.device_data.get('heater_pad_charger_relay_status') == 1 else "red"
            ttk.Label(heater_pad_frame, text="●", foreground=heater_status_circle_color, font=("Arial", 24)).grid(row=0, column=1, padx=5, pady=5)
            for i in range(4):
                heater_pad_frame.grid_columnconfigure(i, weight=1)

            # Charger Status Frame
            charger_frame = ttk.Labelframe(rs422_info_frame, text="Charger Status", bootstyle='dark', borderwidth=5, relief="solid")
            charger_frame.grid(row=0, column=6, columnspan=6, rowspan=2, padx=15, pady=10, sticky="nsew")

            ttk.Label(charger_frame, text="Charger On/Off").grid(row=0, column=0, padx=5, pady=5)
            charger_status_circle_color = "green" if self.device_data.get('charger_status') == 1 else "red"
            ttk.Label(charger_frame, text="●", foreground=charger_status_circle_color, font=("Arial", 24)).grid(row=0, column=1, padx=5, pady=5)
            for i in range(4):
                charger_frame.grid_columnconfigure(i, weight=1)

            # Battery RS422 Frame
            battery_rs422_frame = ttk.Labelframe(rs422_info_frame, text="Battery", bootstyle='dark', borderwidth=5, relief="solid")
            battery_rs422_frame.grid(row=2, column=0, columnspan=6, rowspan=9, padx=15, pady=35, sticky="nsew")

            # Add Entries for Voltage, Current EB1, Current EB2, Charge Current, Temperature, State of Charge

            ttk.Label(battery_rs422_frame, text="Current EB1:").grid(row=0, column=0, padx=5, pady=35, sticky="e")
            current_eb1_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('eb_1_current'), relief="solid", borderwidth=2, width=10, anchor="center")
            current_eb1_entry.grid(row=0, column=1, padx=5, pady=35, sticky="ew")

            # Second row: Current EB2 and Charge Current
            ttk.Label(battery_rs422_frame, text="Current EB2:").grid(row=0, column=2, padx=5, pady=35, sticky="e")
            current_eb2_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('eb_2_current'), relief="solid", borderwidth=2, width=10, anchor="center")
            current_eb2_entry.grid(row=0, column=3, padx=5, pady=35, sticky="ew")

            # First row: Voltage and Current EB1
            ttk.Label(battery_rs422_frame, text="Voltage:").grid(row=4, column=0, padx=5, pady=35, sticky="e")
            voltage_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('voltage'), relief="solid", borderwidth=2, width=10, anchor="center")
            voltage_entry.grid(row=4, column=1, padx=5, pady=35, sticky="ew")

            ttk.Label(battery_rs422_frame, text="Charge Current:").grid(row=4, column=2, padx=5, pady=35, sticky="e")
            charge_current_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('charge_current'), relief="solid", borderwidth=2, width=10, anchor="center")
            charge_current_entry.grid(row=4, column=3, padx=5, pady=35, sticky="ew")

            # Third row: Temperature and State of Charge
            ttk.Label(battery_rs422_frame, text="Temperature:").grid(row=8, column=0, padx=5, pady=35, sticky="e")
            temperature_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('temperature'), relief="solid", borderwidth=2, width=10, anchor="center")
            temperature_entry.grid(row=8, column=1, padx=5, pady=35, sticky="ew")

            ttk.Label(battery_rs422_frame, text="Channel Selected:").grid(row=8, column=2, padx=5, pady=35, sticky="e")
            soc_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('channel_selected'), relief="solid", borderwidth=2, width=10, anchor="center")
            soc_entry.grid(row=8, column=3, padx=5, pady=35, sticky="ew")

            def toggle_eb1_relay():
                """Toggle EB1 Relay On/Off."""
                if rs422_write['cmd_byte_1'] & 0x01 == 0:  # If EB1 is OFF (Bit 0 = 0)
                    rs422_write['cmd_byte_1'] |= 0x01  # Set Bit 0 to 1 to turn ON EB1
                    self.master.after(3000, lambda: update_relay_status('eb_1', eb1_relay_control_status))
                else:
                    rs422_write['cmd_byte_1'] &= 0xFE  # Reset Bit 0 to 0 to turn OFF EB1
                    self.master.after(3000, lambda: update_relay_status('eb_1', eb1_relay_control_status))
                print(f"EB1 Relay Status: {rs422_write['cmd_byte_1']}")
            
            def toggle_eb2_relay():
                """Toggle EB2 Relay On/Off."""
                if rs422_write['cmd_byte_1'] & 0x04 == 0:  # If EB2 is OFF (Bit 2 = 0)
                    rs422_write['cmd_byte_1'] |= 0x04  # Set Bit 2 to 1 to turn ON EB2
                    self.master.after(3000, lambda: update_relay_status('eb_2', eb2_relay_control_status))
                else:
                    rs422_write['cmd_byte_1'] &= 0xFB  # Reset Bit 2 to 0 to turn OFF EB2
                    self.master.after(3000, lambda: update_relay_status('eb_2', eb2_relay_control_status))
                print(f"EB2 Relay Status: {rs422_write['cmd_byte_1']}")
            
            def toggle_shutdown():
                """Toggle Shutdown On/Off."""
                if rs422_write['cmd_byte_1'] & 0x10 == 0:  # If Shutdown is OFF (Bit 4 = 0)
                    rs422_write['cmd_byte_1'] |= 0x10  # Set Bit 4 to 1 to turn ON Shutdown
                    shutdown_control_status.config(text="Shutdown", bootstyle="danger")
                else:
                    rs422_write['cmd_byte_1'] &= 0xEF  # Reset Bit 4 to 0 to turn OFF Shutdown
                    shutdown_control_status.config(text="OFF", bootstyle="success")
                print(f"Shutdown Status: {rs422_write['cmd_byte_1']}")
            
            def toggle_master_slave():
                """Toggle Master/Slave Mode."""
                if rs422_write['cmd_byte_1'] & 0x08 == 0:  # If in Slave mode (Bit 3 = 0)
                    rs422_write['cmd_byte_1'] |= 0x08  # Set Bit 3 to 1 to switch to Master mode
                    self.master.after(3000, lambda: update_master_slave_status('master'))
                else:
                    rs422_write['cmd_byte_1'] &= 0xF7  # Reset Bit 3 to 0 to switch to Slave mode
                    self.master.after(3000, lambda: update_master_slave_status('slave'))
                print(f"Master/Slave Status: {rs422_write['cmd_byte_1']}")
            
            def update_relay_status(relay_name, control_button):
                """Update the status of the relay after toggling."""
                status = self.device_data.get(f'{relay_name}_relay_status')
                if status == 1:
                    control_button.config(text="ON", bootstyle="success")
                else:
                    control_button.config(text="OFF", bootstyle="danger")
            
            def update_master_slave_status(mode):
                """Update Master/Slave mode button after toggle."""
                if mode == 'master':
                    master_slave_control_status.config(text="Master", bootstyle="success")
                else:
                    master_slave_control_status.config(text="Slave", bootstyle="danger")

            # def set_eb1_auto_mode():
            #     # Set EB1 relay to Auto (11 -> 0x03)
            #     rs422_write['cmd_byte_1'] |= 0x03  # Set bits 0 and 1 to 1 (Auto Mode)
            #     eb1_relay_control_status.config(text="Auto", bootstyle="info", state="disabled")  # Disable the button and show "Auto" status

            #     eb1_relay_control_status.config(text="OFF", bootstyle="success")
            #     print(f"EB1 Relay Auto Mode: {rs422_write['cmd_byte_1']}")

            # def set_eb2_auto_mode():
            #     # Set EB2 relay to Auto (11 -> 0x0C)
            #     rs422_write['cmd_byte_1'] |= 0x0C  # Set bits 2 and 3 to 1 (Auto Mode)
            #     eb2_relay_control_status.config(text="Auto", bootstyle="info", state="disabled")  # Disable the button and show "Auto" status
            #     eb2_relay_control_status.config(text="OFF", bootstyle="success")
            #     print(f"EB2 Relay Auto Mode: {rs422_write['cmd_byte_1']}")

            # Callback function for when EB1 checkbox is toggled
            # def on_eb1_checkbox_toggled():
            #     if eb1_checkbox_var.get():  # If checked, set auto mode and disable the button
            #         set_eb1_auto_mode()
            #     else:  # If unchecked, enable the button for manual ON/OFF control
            #         eb1_relay_control_status.config(state="normal")  # Enable the button back for manual control
            #         toggle_eb1_relay()

            # # Callback function for when EB2 checkbox is toggled
            # def on_eb2_checkbox_toggled():
            #     if eb2_checkbox_var.get():  # If checked, set auto mode and disable the button
            #         set_eb2_auto_mode()
            #     else:  # If unchecked, enable the button for manual ON/OFF control
            #         eb2_relay_control_status.config(state="normal")  # Enable the button back for manual control
            #         toggle_eb2_relay()


            right_control_frame = ttk.Labelframe(rs422_info_frame, text="Controls", bootstyle="dark", borderwidth=5, relief="solid")
            right_control_frame.grid(row=2, column=6, columnspan=6, rowspan=9, padx=5, pady=35, sticky="nsew")  # 

            ttk.Label(right_control_frame, text="EB1 Relay On/Off").grid(row=0, column=0, padx=5, pady=35)
            eb1_checkbox_var = tk.BooleanVar()
            # ttk.Checkbutton(right_control_frame, text="Auto", variable=eb1_checkbox_var, command=on_eb1_checkbox_toggled).grid(row=0, column=1, padx=5, pady=35)
            eb1_relay_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger",command=toggle_eb1_relay)
            eb1_relay_control_status.grid(row=0, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="EB1 Relay On/Off").grid(row=1, column=0, padx=5, pady=35)
            eb2_checkbox_var = tk.BooleanVar()
            # ttk.Checkbutton(right_control_frame, text="Auto", variable=eb2_checkbox_var, command=on_eb2_checkbox_toggled).grid(row=1, column=1, padx=5, pady=35)
            eb2_relay_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger",command=toggle_eb2_relay)
            eb2_relay_control_status.grid(row=1, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="Shutdown").grid(row=2, column=0, padx=5, pady=5)
            shutdown_control_status = ttk.Button(right_control_frame, text="Shutdown", bootstyle="danger",command=toggle_shutdown)
            shutdown_control_status.grid(row=2, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="Master/Slave").grid(row=3, column=0, padx=5, pady=5)
            master_slave_control_status = ttk.Button(right_control_frame, text="Slave", bootstyle="danger",command=toggle_master_slave)
            master_slave_control_status.grid(row=3, column=2, padx=5, pady=35)

            # Configure column weights to distribute space equally in the battery frame
            for i in range(5):
                battery_rs422_frame.grid_columnconfigure(i, weight=1)

            # Configure equal weight for the rs422_info_frame columns to distribute space equally
            for i in range(12):
                rs422_info_frame.grid_columnconfigure(i, weight=1)

            for i in range(4):
                right_control_frame.grid_columnconfigure(i, weight=1)

            

        # Pack the content_frame itself
        self.content_frame.pack(fill="both", expand=True)
            
        self.select_button(self.info_button) 

    def get_gauge_style(self, value, gauge_type):
        if gauge_type == "bus_1_voltage_after_diode":
            if value < 20:
                return "success"
            elif 20 <= value < 30:
                return "warning"
            elif 30 <= value < 35:
                return "danger"
            else:
                return "success"
        elif gauge_type == "bus_2_voltage_after_diode":
            if value < 20:
                return "success"
            elif 20 <= value < 50:
                return "info"
            elif 50 <= value < 120:
                return "warning"
            else:
                return "danger"
        elif gauge_type == "bus_1_current_sensor1":
            if value < 30:
                return "success"
            elif 30 <= value < 69:
                return "warning"
            elif value >= 70:
                return "danger"
            else:
                return "info"
        elif gauge_type == "bus_2_current_sensor1":
            if value < 20:
                return "danger"
            elif 20 <= value < 50:
                return "warning"
            elif 50 <= value < 80:
                return "info"
            else:
                return "success"
        elif gauge_type == "charger_input_current":
            if value < 20:
                return "success"
            elif 20 <= value < 50:
                return "info"
            elif 50 <= value < 80:
                return "warning"
            else:
                return "danger"
        elif gauge_type == "charger_output_voltage":
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
 
    def get_gauge_style_422(self, value, gauge_type):
        if gauge_type == "eb_1_current":
            if value < 20:
                return "success"
            elif 20 <= value < 30:
                return "warning"
            elif 30 <= value < 35:
                return "danger"
            else:
                return "success"
        elif gauge_type == "eb_1_current":
            if value < 20:
                return "success"
            elif 20 <= value < 50:
                return "info"
            elif 50 <= value < 120:
                return "warning"
            else:
                return "danger"
        elif gauge_type == "charge_current":
            if value < 30:
                return "success"
            elif 30 <= value < 69:
                return "warning"
            elif value >= 70:
                return "danger"
            else:
                return "info"
        elif gauge_type == "temperature":
            if value < 20:
                return "danger"
            elif 20 <= value < 50:
                return "warning"
            elif 50 <= value < 80:
                return "info"
            else:
                return "success"
        else:
            return "info"  # Default style if gauge_type is unknown
 