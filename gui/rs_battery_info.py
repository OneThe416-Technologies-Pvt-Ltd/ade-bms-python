#can_battery_info.py
import customtkinter as ctk
import tkinter as tk
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
import datetime
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

        if self.rs232_flag:
            self.device_data = rs232_device_data
            print("RS232 data updated")
        elif self.rs422_flag:
            self.device_data = rs422_device_data
            print("RS422 data updated")
            
        active_protocol = get_active_protocol()

        # Step 3: Update the flags based on the active protocol
        if active_protocol == "RS-232":
            self.rs232_flag = True
            self.rs422_flag = False
        elif active_protocol == "RS-422":
            self.rs232_flag = False
            self.rs422_flag = True
        
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
            command=lambda: self.select_button(self.help_button,self.show_help),
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

        # Step 2: Get the active protocol from pcomm.py
        active_protocol = get_active_protocol()

        # Step 3: Update the flags based on the active protocol
        if active_protocol == "RS-232":
            self.rs232_flag = True
            self.rs422_flag = False
        elif active_protocol == "RS-422":
            self.rs232_flag = False
            self.rs422_flag = True

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

        # Section: Right Control Frame (Change to Grid Layout)
        right_control_frame = ttk.Labelframe(load_frame, text="Controls", bootstyle="dark", borderwidth=5, relief="solid")
        right_control_frame.grid(row=0, column=1, rowspan=1,columnspan=3, padx=5, pady=5, sticky="nsew")  # Positioned to the right

        if self.rs232_flag:
            ttk.Label(right_control_frame, text="BUS1 On/Off").grid(row=0, column=0, padx=5, pady=5)
            bus1_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger")
            bus1_control_status.grid(row=0, column=1, padx=5, pady=5)

            ttk.Label(right_control_frame, text="BUS2 On/Off").grid(row=1, column=0, padx=5, pady=5)
            bus2_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger")
            bus2_control_status.grid(row=1, column=1, padx=5, pady=5)

            ttk.Label(right_control_frame, text="Charger On/Off").grid(row=2, column=0, padx=5, pady=5)
            charger_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger")
            charger_control_status.grid(row=2, column=1, padx=5, pady=5)
            def toggle_output_relay():
                if rs232_write['cmd_byte_1'] == 0x00:
                    rs232_write['cmd_byte_1'] = 0xF0  # Set to 0x08 when turning ON
                    output_relay_control_status.config(text="ON", bootstyle="success")  # Update button text and style
                    print("Output Relay ON")
                else:
                    rs232_write['cmd_byte_1'] = 0x00  # Reset to 0x00 when turning OFF
                    output_relay_control_status.config(text="OFF", bootstyle="danger")  # Update button text and style
                    print("Output Relay OFF")

            ttk.Label(right_control_frame, text="Output Relay On/Off").grid(row=3, column=0, padx=5, pady=5)
            output_relay_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger", command=toggle_output_relay)
            output_relay_control_status.grid(row=3, column=1, padx=5, pady=5)

            ttk.Label(right_control_frame, text="Reset").grid(row=4, column=0, padx=5, pady=5)
            reset_button = ttk.Button(right_control_frame, text="RESET", bootstyle="danger")
            reset_button.grid(row=4, column=1, padx=5, pady=5)

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

        # Section: Testing Mode
        testing_frame = ttk.Labelframe(load_frame, text="Testing Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
        testing_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")
        test_button = ttk.Button(testing_frame, text="Set 25A and Turn ON Load", command=set_l1_50a_and_turn_on)
        test_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        maintenance_button = ttk.Button(testing_frame, text="Set 50A and Turn ON Load", command=set_l1_100a_and_turn_on)
        maintenance_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        # Turn Off Load Button
        turn_off_button = ttk.Button(testing_frame, text="Turn OFF Load", command=custom_turn_off)
        turn_off_button.grid(row=0, column=2, columnspan=2, padx=10, pady=10, sticky="ew")

        # Section: Custom Mode
        custom_frame = ttk.Labelframe(load_frame, text="Custom Mode", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
        custom_frame.grid(row=3, column=0,columnspan=4, padx=5, pady=5, sticky="nsew")

        custom_label = ttk.Label(custom_frame, text="Set Custom Current (A):", font=("Helvetica", 12), width=20, anchor="e")
        custom_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        def save_custom_value():
            """Saves the custom current value and sets it to the Chroma device."""
            custom_value = self.custom_current_entry.get()
            if custom_value.isdigit():  # Basic validation to ensure it's a number
                if self.selected_battery == "Battery 1":
                    start_fetching_voltage(battery_no=1,load_value=int(custom_value))
                elif self.selected_battery == "Battery 2":
                     start_fetching_voltage(battery_no=2,load_value=int(custom_value))
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
                stop_fetching_voltage()
                custom_turn_off  # Call the Turn OFF function
                toggle_button.config(text="Turn ON Load", bootstyle="success")  # Change to green when OFF
                self.load_status.set(False)  # Update state to OFF
            else:
                custom_turn_on  # Call the Turn ON function
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

    def maintains_mode_load(self):
        if self.selected_battery == "Battery 1":
            start_fetching_voltage(battery_no=1,load_value=100)
        elif self.selected_battery == "Battery 2":
             start_fetching_voltage(battery_no=2,load_value=100)
        set_l1_50a_and_turn_on()

    def testing_mode_load(self):
        if self.selected_battery == "Battery 1":
            start_fetching_voltage(battery_no=1,load_value=50)
        elif self.selected_battery == "Battery 2":
             start_fetching_voltage(battery_no=2,load_value=50)
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
        stop_communication()
        # Stop periodic refresh
        self.refresh_active = False  # Stop the periodic refresh
        self.main_frame.pack_forget()
        self.main_window.show_main_window()
        # Perform disconnection logic here (example: print disconnect message)
        print("Disconnecting...")
        # self.update_widgets()

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

            def toggle_eb1_relay():
                if rs422_write['cmd_byte_1'] == 0x00:
                    rs422_write['cmd_byte_1'] = 0x0F
                    eb1_relay_control_status.config(text="ON", bootstyle="success")  # Update button text and style
                    print(f"EB1 Relay ON: {rs422_write['cmd_byte_1']}")
                else:
                    rs422_write['cmd_byte_1'] = 0x00
                    eb1_relay_control_status.config(text="OFF", bootstyle="danger")  # Update button text and style
                    print(f"EB1 Relay OFF: {rs422_write['cmd_byte_1']}")

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
            soc_entry = ttk.Label(battery_rs422_frame, text=self.device_data.get('state_oc_charge'), relief="solid", borderwidth=2, width=10, anchor="center")
            soc_entry.grid(row=8, column=3, padx=5, pady=35, sticky="ew")

            right_control_frame = ttk.Labelframe(rs422_info_frame, text="Controls", bootstyle="dark", borderwidth=5, relief="solid")
            right_control_frame.grid(row=2, column=6, columnspan=6, rowspan=9, padx=5, pady=35, sticky="nsew")  # 

            ttk.Label(right_control_frame, text="EB1 Relay On/Off").grid(row=0, column=0, padx=5, pady=35)
            eb1_checkbox_var = tk.BooleanVar()
            ttk.Checkbutton(right_control_frame, text="Auto", variable=eb1_checkbox_var).grid(row=0, column=1, padx=5, pady=35)
            eb1_relay_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger",command=toggle_eb1_relay)
            eb1_relay_control_status.grid(row=0, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="EB1 Relay On/Off").grid(row=1, column=0, padx=5, pady=35)
            eb2_checkbox_var = tk.BooleanVar()
            ttk.Checkbutton(right_control_frame, text="Auto", variable=eb2_checkbox_var).grid(row=1, column=1, padx=5, pady=35)
            eb2_relay_control_status = ttk.Button(right_control_frame, text="OFF", bootstyle="danger")
            eb2_relay_control_status.grid(row=1, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="Shutdown").grid(row=2, column=0, padx=5, pady=5)
            shutdown_control_status = ttk.Button(right_control_frame, text="Shutdown", bootstyle="danger")
            shutdown_control_status.grid(row=2, column=2, padx=5, pady=35)

            ttk.Label(right_control_frame, text="Master/Slave").grid(row=3, column=0, padx=5, pady=5)
            master_slave_control_status = ttk.Button(right_control_frame, text="Slave", bootstyle="danger")
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
 