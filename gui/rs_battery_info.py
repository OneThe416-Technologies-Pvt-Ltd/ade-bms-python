#can_battery_info.py
import customtkinter as ctk
import tkinter as tk
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
Image.CUBIC = Image.BICUBIC
import time
from tkinter import messagebox
from pcomm_api.pcomm import *
from load_api.chroma_api import *
import helpers.config as config

class RSBatteryInfo:
    def __init__(self, master, main_window=None):
        self.master = master
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button

        # self.center_window(1200, 600)  # Center the window with specified dimensions
        self.master.resizable(True, False)
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
        self.mode_var = tk.StringVar(value="Testing")  # Initialize mode_var here
        self.show_dynamic_fields_var = tk.BooleanVar()  # Track if dynamic fields are shown

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
        self.config_icon = self.load_icon_menu(os.path.join(self.assets_path, "settings.png"))
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

        # Mode selection dropdown
        self.mode_label = ttk.Label(self.side_menu_frame, bootstyle="inverse-info", text="Mode:")
        self.mode_label.pack(pady=(140, 5), padx=(0, 5))

        # Mode Dropdown
        self.mode_dropdown = ttk.Combobox(
            self.side_menu_frame,
            textvariable=self.mode_var,
            values=["Testing", "Maintenance"],
            state="readonly"
        )
        self.mode_dropdown.pack(padx=(5, 20))
        self.mode_dropdown.bind("<<ComboboxSelected>>", self.update_mode)

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

    def update_mode(self, event=None):
        """Update the UI components based on the selected mode (Testing or Maintenance)."""
        selected_mode = self.mode_var.get()

        if selected_mode == "Testing" and self.selected_button == self.dashboard_button:
            # Hide the checkbox in Testing
            if hasattr(self, 'discharge_max_checkbox'):
                self.discharge_max_checkbox.pack_forget()
        elif selected_mode == "Maintenance" and self.selected_button == self.dashboard_button:
            if hasattr(self, 'discharge_max_checkbox'):
                self.discharge_max_checkbox.pack(pady=(5, 10), before=self.discharging_current_meter)

        # Check if we are in Testing or Maintenance
        if selected_mode == "Maintenance":
            self.limited = False  # Show full data in Maintenance Mode
            # Show the Configuration button
            if not hasattr(self, 'config_button'):  # Ensure the button exists
                self.config_button = ttk.Button(
                    self.side_menu_frame,
                    text=" Configuration",
                    command=lambda: self.select_button(self.config_button,self.show_config),
                    image=self.config_icon,
                    compound="left",  # Place the icon to the left of the text
                    bootstyle="info"
                )
            self.config_button.pack(fill="x", pady=2, before=self.help_button)
        else:
            self.limited = True  # Show limited data in Testing
            # Hide the Configuration button if in Testing
            if hasattr(self, 'config_button'):
                self.config_button.pack_forget()
        """Update the Info page content based on the selected mode."""
        if self.selected_button == self.info_button:
            self.show_control()
            self.show_info()
        elif self.selected_button == self.control_button:
            self.show_info()
            self.show_control()
        elif self.selected_button == self.config_button:
            self.show_dashboard()
        elif self.selected_button == self.report_button:
            self.show_report()

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
                self.device_data['manual_cycle_count'] = manual_cycle_count_var.get()
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

    def show_config(self):
        """Show the configuration settings page where users can view and update configurations."""
        self.clear_content_frame()  # Clear the content area before loading the config page

        # Create a frame for the config page content
        config_frame = ttk.Frame(self.content_frame, borderwidth=10, relief="solid")
        config_frame.pack(fill="x")

        # Create a new Labelframe for configuration settings
        config_labelframe = ttk.Labelframe(self.content_frame, text="Configuration Settings", bootstyle="dark", borderwidth=10, relief="solid")
        config_labelframe.pack(fill="x", padx=10, pady=10)

        # Validation functions for float and int
        def validate_numeric_input(value):
            if value == "" or value.isdigit() or value == "-" or value.replace(".", "", 1).isdigit():
                return True
            return False

        vcmd = (self.content_frame.register(validate_numeric_input), "%P")

        # Define configuration fields, using values from the config file
        config_values = {
            "Logging Time (ms)": config.config_values['rs_config'].get('logging_time'),
            "Discharge Cutoff Voltage (V)": config.config_values['rs_config'].get('discharge_cutoff_volt'),
            "Charge Cutoff Current (A)": config.config_values['rs_config'].get('charging_cutoff_curr'),
            "Charge Cutoff Voltage (V)": config.config_values['rs_config'].get('charging_cutoff_volt'),
            "Charge Cutoff Capacity (%)": config.config_values['rs_config'].get('charging_cutoff_capacity'),
            "Temperature Caution (°C)": config.config_values['rs_config'].get('temperature_caution'),
            "Temperature Alarm (°C)": config.config_values['rs_config'].get('temperature_alarm'),
            "Temperature Critical (°C)": config.config_values['rs_config'].get('temperature_critical')
        }

        # Create labels and entry fields for each config
        row = 0
        self.config_entries = {}  # Store entry fields to update values later
        for config_name, config_value in config_values.items():
            # Label for each configuration field
            ttk.Label(config_labelframe, text=f"{config_name}:", font=("Helvetica", 10, "bold")).grid(row=row, column=0, padx=10, pady=5, sticky="e")

            # Entry field for each configuration value
            config_entry = ttk.Entry(config_labelframe, validate="key", validatecommand=vcmd)
            config_entry.insert(0, config_value)  # Insert the existing value into the entry
            config_entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")

            # Save the entry fields in a dictionary for later access
            self.config_entries[config_name] = config_entry

            row += 1

        # Add an "Update" button to save the modified values inside the labelframe
        update_button = ttk.Button(config_labelframe, text="Update Configurations", command=self.update_config_values, bootstyle='success')
        update_button.grid(row=row, column=0, columnspan=2, padx=10, pady=20, sticky="ew")

        # Make the columns stretch to fill available space
        config_labelframe.grid_columnconfigure(0, weight=1)
        config_labelframe.grid_columnconfigure(1, weight=2)

        # Create a new Labelframe for project management
        project_frame = ttk.Labelframe(self.content_frame, text="Project Management", bootstyle="dark", borderwidth=10, relief="solid")
        project_frame.pack(fill="x", padx=10, pady=10)

        row += 1

        # Add the Combobox for the Project Dropdown
        ttk.Label(project_frame, text="Select Project:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.project_var = tk.StringVar()
        self.project_combobox = ttk.Combobox(project_frame, textvariable=self.project_var, values=config.config_values['rs_config']['projects'], state="readonly")
        self.project_combobox.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.project_combobox.set("Select a Project")

        # Add buttons for Add, Edit, and Delete project
        add_project_button = ttk.Button(project_frame, text="Add Project", command=self.add_project, bootstyle='success')
        add_project_button.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        edit_project_button = ttk.Button(project_frame, text="Edit Project", command=self.edit_project)
        edit_project_button.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        delete_project_button = ttk.Button(project_frame, text="Delete Project", command=self.delete_project, bootstyle='danger')
        delete_project_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        # Make columns stretch to fill available space
        project_frame.grid_columnconfigure(0, weight=1)
        project_frame.grid_columnconfigure(1, weight=2)

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
        selected_mode = self.mode_var.get()  # Get the selected mode

        # Container for the content
        main_frame = ttk.Frame(self.content_frame)
        main_frame.grid(row=0, column=0, columnspan=12, rowspan=6, padx=10, pady=10, sticky="nsew")

        # Configure the content frame for expansion and add equal space on both sides
        self.content_frame.grid_columnconfigure(0, weight=1)  # Allow the content frame to expand
        self.content_frame.grid_rowconfigure(0, weight=1)     # Allow the main_frame to expand vertically

        # Configure the main_frame for equal spacing
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_columnconfigure(2, weight=1)

        # Row 0: View Reports Button
        generate_button = ttk.Button(main_frame, text="View Reports", command=self.show_report)
        generate_button.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        # Battery Info Frame (row 1-2, col 0-2 inside the main_frame)
        battery_info_frame = ttk.LabelFrame(main_frame, text="Battery Info", borderwidth=5)
        battery_info_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        # Add a Label for the Project Dropdown
        ttk.Label(battery_info_frame, text="Project").grid(row=0, column=0, padx=15, pady=5, sticky="e")

        # Add the Combobox for the Project Dropdown
        project_options = config.config_values['rs_config']['projects']
        project_combobox = ttk.Combobox(battery_info_frame, values=project_options, state="readonly")  # Dropdown with readonly state
        project_combobox.grid(row=0, column=1, padx=15, pady=5, sticky="w")
        project_combobox.set("Select a Project")

        # Add individual fields for Battery Info
        ttk.Label(battery_info_frame, text="Battery Nomenclature").grid(row=0, column=2, padx=15, pady=5, sticky="e")
        device_name_entry = ttk.Entry(battery_info_frame)
        device_name_entry.insert(0, "Battery Nomenclature")
        device_name_entry.grid(row=0, column=3, padx=15, pady=5, sticky="w")
        device_name_entry.config(state="readonly")  # Disable editing

        ttk.Label(battery_info_frame, text="Serial Number").grid(row=0, column=4, padx=15, pady=5, sticky="e")
        serial_number_entry = ttk.Entry(battery_info_frame)
        serial_number_entry.insert(0, str(latest_data.get("Serial Number", "")))
        serial_number_entry.grid(row=0, column=5, padx=15, pady=5, sticky="w")
        serial_number_entry.config(state="readonly")  # Disable editing

        ttk.Label(battery_info_frame, text="Cycle Count").grid(row=1, column=0, padx=15, pady=5, sticky="e")
        cycle_count_entry = ttk.Entry(battery_info_frame)
        cycle_count_entry.insert(0, str(latest_data.get("Cycle Count", "")))
        cycle_count_entry.grid(row=1, column=1, padx=15, pady=5, sticky="w")

        ttk.Label(battery_info_frame, text="Battery Capacity").grid(row=1, column=2, padx=15, pady=5, sticky="e")
        capacity_entry = ttk.Entry(battery_info_frame)
        capacity_entry.insert(0, str(55))
        capacity_entry.grid(row=1, column=3, padx=15, pady=5, sticky="w")
        capacity_entry.config(state="readonly")  # Disable editing

        # Row 3-4: Charging Info Frame
        charging_frame = ttk.LabelFrame(main_frame, text="Charging Info", borderwidth=5)
        charging_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

        ttk.Label(charging_frame, text="Date :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        charging_date_entry = ttk.Entry(charging_frame)
        charging_date_entry.insert(0, "09/09/2024")
        charging_date_entry.grid(row=0, column=0, padx=50, pady=5, sticky="w")

        ttk.Label(charging_frame, text="OCV :").grid(row=0, column=0, padx=200, pady=5, sticky="w")
        charging_ocv_entry = ttk.Entry(charging_frame)
        charging_ocv_entry.insert(0, 0)
        charging_ocv_entry.grid(row=0, column=0, padx=250, pady=5, sticky="w")

        
        def load_rs_charging_data():
            # Define the path to the CAN charging data Excel file inside the AppData directory
            rs_charging_data_file = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database", "rs_charging_data.xlsx")

            # Read the Excel file using pandas
            try:
                df_rs_charging = pd.read_excel(rs_charging_data_file)
            except FileNotFoundError:
                print(f"Error: {rs_charging_data_file} not found.")
                return

            # Retrieve the serial number from self.device_data (assuming it contains a key 'serial_number')
            device_serial_number = self.device_data.get('serial_number')

            # Filter the DataFrame where the serial number matches
            filtered_data = df_rs_charging[df_rs_charging['Serial Number'] == device_serial_number]

            # If no matching serial number is found, you can add a fallback or a message
            if filtered_data.empty:
                print(f"No data found for Serial Number: {device_serial_number}")
                return

            # Clear any existing data in the Treeview
            for item in charging_table.get_children():
                charging_table.delete(item)

            # Insert the filtered data into the Treeview
            for _, row in filtered_data.iterrows():
                charging_table.insert('', 'end', values=(row['Hour'], row['Voltage'], row['Current'], row['Temperature'], row['State of Charge']))



        # Create Treeview for Charging Info Table
        columns = ('Hour', 'Voltage (V)', 'Current (A)', 'Temperature (°C)', 'State of Charge (%)')
        charging_table = ttk.Treeview(charging_frame, columns=columns, show='headings', height=5)

        charging_table.heading('Hour', text='Hour')
        charging_table.heading('Voltage (V)', text='Voltage (V)')
        charging_table.heading('Current (A)', text='Current (A)')
        charging_table.heading('Temperature (°C)', text='Temperature (°C)')
        charging_table.heading('State of Charge (%)', text='State of Charge (%)')

        charging_table.column('Hour', width=100, anchor='center')
        charging_table.column('Voltage (V)', width=100, anchor='center')
        charging_table.column('Current (A)', width=100, anchor='center')
        charging_table.column('Temperature (°C)', width=150, anchor='center')
        charging_table.column('State of Charge (%)', width=150, anchor='center')

        # Adding a vertical scrollbar to the charging table
        scrollbar_charging = ttk.Scrollbar(charging_frame, orient="vertical", command=charging_table.yview)
        charging_table.configure(yscrollcommand=scrollbar_charging.set)

        charging_table.grid(row=1, column=0, sticky='nsew')
        scrollbar_charging.grid(row=1, column=1, sticky='ns')

        charging_frame.grid_rowconfigure(1, weight=1)
        charging_frame.grid_columnconfigure(0, weight=1)

        # Call the function to load CAN charging data from the Excel file
        load_rs_charging_data()

        # Add the Discharge Info Frame based on the mode (Testing vs Maintenance)
        discharge_info_frame = ttk.LabelFrame(main_frame, text="Discharge Info", borderwidth=5)
        discharge_info_frame.grid(row=3, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

        ttk.Label(discharge_info_frame, text="Date :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        discharging_date_entry = ttk.Entry(discharge_info_frame)
        discharging_date_entry.insert(0, "09/09/2024")
        discharging_date_entry.grid(row=0, column=0, padx=50, pady=5, sticky="w")

        ttk.Label(discharge_info_frame, text="OCV :").grid(row=0, column=0, padx=200, pady=5, sticky="w")
        discharging_ocv_entry = ttk.Entry(discharge_info_frame)
        discharging_ocv_entry.insert(0, 0)
        discharging_ocv_entry.grid(row=0, column=0, padx=250, pady=5, sticky="w")

        def save_values():
            # Collect values from the entry boxes into arrays
            project = project_combobox.get()  # Get the selected project
            battery_nomenclature = device_name_entry.get()
            serial_number = serial_number_entry.get()
            cycle_count = cycle_count_entry.get()
            full_charge_capacity = capacity_entry.get()
            charging_date = charging_date_entry.get()  # New field
            ocv_before_charging = charging_ocv_entry.get()  # New field
            discharging_date = discharging_date_entry.get()  # New field
            ocv_before_discharging = discharging_ocv_entry.get()  # New field

            # Create an array to hold the values
            data_to_save = [
                project,  # Add the project value at the beginning
                battery_nomenclature,
                serial_number,
                cycle_count,
                full_charge_capacity,
                charging_date,
                ocv_before_charging,
                discharging_date,
                ocv_before_discharging
            ]

            # Call the method to update the Excel file
            # update_excel_and_download_pdf(data_to_save)


        def load_rs_discharging_data():
            """
            Load the CAN discharging data from an Excel file and display it in the Treeview.
            """
            # Define the path to the CAN discharging data Excel file inside the AppData directory
            rs_discharging_data_file = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database", "rs_discharging_data.xlsx")

            # Read the Excel file using pandas
            try:
                df_rs_discharging = pd.read_excel(rs_discharging_data_file)
            except FileNotFoundError:
                messagebox.showerror("Error", f"{rs_discharging_data_file} not found.")
                return

            # Retrieve the serial number from self.device_data (assuming it contains a key 'serial_number')
            device_serial_number = self.device_data.get('serial_number')

            # Filter the DataFrame where the serial number matches
            filtered_data = df_rs_discharging[df_rs_discharging['Serial Number'] == device_serial_number]

            # If no matching serial number is found, display a message and return
            if filtered_data.empty:
                messagebox.showwarning("No Data", f"No data found for Serial Number: {device_serial_number}")
                return

            # Clear any existing data in the Treeview
            for item in discharging_table.get_children():
                discharging_table.delete(item)

            # Insert the filtered data into the Treeview
            for _, row in filtered_data.iterrows():
                try:
                    discharging_table.insert('', 'end', values=(
                        row['Hour'], 
                        float(row['Voltage']), 
                        float(row['Current']), 
                        float(row['Temperature']), 
                        float(row['Load Current'])
                    ))
                except ValueError:
                    print(f"Error processing row: {row}")

            print(f"Data for Serial Number {device_serial_number} loaded successfully.")

        # Create Treeview for Charging Info Table
        columns = ('Hour', 'Voltage (V)', 'Current (A)', 'Temperature (°C)', 'Load Current (A)')
        discharging_table = ttk.Treeview(discharge_info_frame, columns=columns, show='headings', height=5)

        discharging_table.heading('Hour', text='Hour')
        discharging_table.heading('Voltage (V)', text='Voltage (V)')
        discharging_table.heading('Current (A)', text='Current (A)')
        discharging_table.heading('Temperature (°C)', text='Temperature (°C)')
        discharging_table.heading('Load Current (A)', text='Load Current (A)')

        discharging_table.column('Hour', width=100, anchor='center')
        discharging_table.column('Voltage (V)', width=100, anchor='center')
        discharging_table.column('Current (A)', width=100, anchor='center')
        discharging_table.column('Temperature (°C)', width=150, anchor='center')
        discharging_table.column('Load Current (A)', width=150, anchor='center')

        # Adding a vertical scrollbar to the discharging table
        scrollbar_discharging = ttk.Scrollbar(discharge_info_frame, orient="vertical", command=discharging_table.yview)
        discharging_table.configure(yscrollcommand=scrollbar_discharging.set)
        discharging_table.grid(row=1, column=0, sticky='nsew')
        scrollbar_discharging.grid(row=1, column=1, sticky='ns')

        discharge_info_frame.grid_rowconfigure(1, weight=1)
        discharge_info_frame.grid_columnconfigure(0, weight=1)

        def load_rs_data():
            """
            Load the CAN charging and discharging data from an Excel file and display it.
            """
            # Define the path to the CAN data Excel file inside the AppData directory
            rs_data_file = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database", "rs_data.xlsx")

            # Read the Excel file using pandas
            try:
                df_rs_data = pd.read_excel(rs_data_file)
            except FileNotFoundError:
                messagebox.showerror("Error", f"{rs_data_file} not found.")
                return

            # Retrieve the serial number from self.device_data
            device_serial_number = self.device_data.get('serial_number')

            # Filter the DataFrame where the serial number matches
            filtered_data = df_rs_data[df_rs_data['Serial Number'] == device_serial_number]

            # If no matching serial number is found, display a message and return
            if filtered_data.empty:
                messagebox.showwarning("No Data", f"No data found for Serial Number: {device_serial_number}")
                return

            # Assuming there's only one row per serial number, use the first row for display
            first_row = filtered_data.iloc[0]

            # Populate individual fields from the data
            device_name_entry.delete(0, tk.END)
            device_name_entry.insert(0, str(first_row.get("Device Name", "")))

            serial_number_entry.delete(0, tk.END)
            serial_number_entry.insert(0, str(first_row.get("Serial Number", "")))

            cycle_count_entry.delete(0, tk.END)
            cycle_count_entry.insert(0, str(first_row.get("Cycle Count", "")))

            capacity_entry.config(state="normal")  # Enable field for update
            capacity_entry.delete(0, tk.END)
            capacity_entry.insert(0, str(first_row.get("Full Charge Capacity", 0)))
            capacity_entry.config(state="readonly")  # Disable field after update

            # Charging Info fields
            charging_ocv_entry.delete(0, tk.END)
            charging_ocv_entry.insert(0, float(first_row.get("OCV Before Charging", 0)))

            charging_date_entry.delete(0, tk.END)
            charging_date_entry.insert(0, str(first_row.get("Charging Date", '')))
            
            # Discharging Info fields
            discharging_ocv_entry.delete(0, tk.END)
            discharging_ocv_entry.insert(0, float(first_row.get("OCV Before Discharging", 0)))

            discharging_date_entry.delete(0, tk.END)
            discharging_date_entry.insert(0, str(first_row.get("Discharging Date", '')))

        load_rs_discharging_data()
        load_rs_data()

        main_frame.grid_rowconfigure(3, weight=1)

        # Row 5: Download Report Button at the bottom
        download_button = ttk.Button(main_frame, text="Download Report", command=lambda:save_values())
        download_button.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        # Ensure the last row is expanded to push the button to the bottom
        main_frame.grid_rowconfigure(4, weight=1)

        # Make the content_frame expand
        self.content_frame.pack(fill="both", expand=True)

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

        if selected_mode == "Testing":
            # Section: Testing Mode
            testing_frame = ttk.Labelframe(load_frame, text="Testing", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
            testing_frame.grid(row=1, column=0, columnspan=6, padx=10, pady=10, sticky="nsew")

            test_button = ttk.Button(testing_frame, text="Set 50A and Turn ON Load", command=set_l1_50a_and_turn_on)
            test_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

            # Turn Off Load Button
            turn_off_button = ttk.Button(testing_frame, text="Turn OFF Load", command=turn_load_off)
            turn_off_button.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

        elif selected_mode == "Maintenance":
            # Section: Maintenance Mode
            maintenance_frame = ttk.Labelframe(load_frame, text="Maintenance", padding=10, borderwidth=10, relief="solid", bootstyle='dark')
            maintenance_frame.grid(row=1, column=0, columnspan=6, padx=10, pady=10, sticky="nsew")

            # Add a checkbox for dynamic fields
            dynamic_fields_checkbox = ttk.Checkbutton(
                maintenance_frame,
                text="Dynamic",
                variable=self.show_dynamic_fields_var,
                command=self.toggle_dynamic_fields  # Method to show/hide fields
            )
            dynamic_fields_checkbox.grid(row=0, column=0, padx=10, pady=10, sticky="w")

            # The "Set 100A and Turn ON Load" button
            self.maintenance_button = ttk.Button(maintenance_frame, text="Set 100A and Turn ON Load", command=set_l1_100a_and_turn_on)
            self.maintenance_button.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

            # The "Turn OFF Load" button
            self.turn_off_button = ttk.Button(maintenance_frame, text="Turn OFF Load", command=turn_load_off)
            self.turn_off_button.grid(row=1, column=2, padx=10, pady=10, sticky="ew")

            # Frame for dynamic fields (L1, L2, T1, T2, Repeat, Save button)
            self.dynamic_fields_frame = ttk.Frame(maintenance_frame)
            self.dynamic_fields_frame.grid(row=1, column=0, columnspan=6, padx=10, pady=10, sticky="nsew")

            # Initially hide the dynamic fields frame
            self.dynamic_fields_frame.grid_remove()

            # L1 Entry
            ttk.Label(self.dynamic_fields_frame, text="L1 :").grid(row=0, column=0, padx=10, pady=10, sticky="e")
            self.l1_entry = ttk.Entry(self.dynamic_fields_frame, width=5)
            self.l1_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

            # L2 Entry
            ttk.Label(self.dynamic_fields_frame, text="L2 :").grid(row=0, column=2, padx=10, pady=10, sticky="e")
            self.l2_entry = ttk.Entry(self.dynamic_fields_frame, width=5)
            self.l2_entry.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

            # T1 Entry
            ttk.Label(self.dynamic_fields_frame, text="T1 :").grid(row=0, column=4, padx=10, pady=10, sticky="e")
            self.t1_entry = ttk.Entry(self.dynamic_fields_frame, width=5)
            self.t1_entry.grid(row=0, column=5, padx=10, pady=10, sticky="ew")

            # T2 Entry
            ttk.Label(self.dynamic_fields_frame, text="T2 :").grid(row=0, column=6, padx=10, pady=10, sticky="e")
            self.t2_entry = ttk.Entry(self.dynamic_fields_frame, width=5)
            self.t2_entry.grid(row=0, column=7, padx=10, pady=10, sticky="ew")

            # Repeat Entry
            ttk.Label(self.dynamic_fields_frame, text="Repeat :").grid(row=0, column=8, padx=10, pady=10, sticky="e")
            self.repeat_entry = ttk.Entry(self.dynamic_fields_frame, width=5)
            self.repeat_entry.grid(row=0, column=9, padx=10, pady=10, sticky="ew")

            # Save Button
            save_button = ttk.Button(self.dynamic_fields_frame, text="Save", command=self.save_dynamic_fields)
            save_button.grid(row=0, column=10, padx=10, pady=10, sticky="ew")

            # Turn ON Load Button (in the same row as dynamic fields)
            turn_on_load_button = ttk.Button(self.dynamic_fields_frame, text="Turn ON Load")
            turn_on_load_button.grid(row=0, column=11, padx=10, pady=10, sticky="ew")

            # Expand columns for even distribution
            for i in range(11):
                self.dynamic_fields_frame.grid_columnconfigure(i, weight=1)

    
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
                turn_load_off
                toggle_button.config(text="Turn ON Load", bootstyle="success")  # Change to green when OFF
                self.load_status.set(False)  # Update state to OFF
            else:
                turn_load_on  # Call the Turn ON function
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
        turn_load_off

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

        # if self.main_window.rs_battery_info is not None:
        #     del self.main_window.rs_battery_info
        #     self.main_window.rs_battery_info = None
    
        self.main_frame.pack_forget()
        self.main_window.show_main_window()

    def show_info(self, event=None):
        self.clear_content_frame()

        style = ttk.Style()

        # Configure the Labelframe title font to be bold
        style.configure("TLabelframe.Label", font=("Helvetica", 10, "bold"))

        # Determine if we are in Testing Mode or Maintenance Mode
        selected_mode = self.mode_var.get()
        self.limited = selected_mode == "Testing"

        if self.rs232_flag:
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
    
    def update_config_values(self):
        """Update the configuration values and save them to the JSON config file."""

        def determine_number_type(value):
            """Determine if the value should be an int or float."""
            try:
                # First, try to convert it to an int
                if '.' in value:
                    return float(value)
                return int(value)
            except ValueError:
                # If conversion fails, return None
                return None

        # Retrieve the updated values from the entry fields
        updated_values = {}
        for config_name, entry in self.config_entries.items():
            updated_values[config_name] = entry.get()

        # Update the configuration values in the config file with dynamic type handling
        config.config_values['rs_config']['logging_time'] = determine_number_type(updated_values['Logging Time (ms)'])
        config.config_values['rs_config']['discharge_cutoff_volt'] = determine_number_type(updated_values['Discharge Cutoff Voltage (V)'])
        config.config_values['rs_config']['charging_cutoff_curr'] = determine_number_type(updated_values['Charge Cutoff Current (A)'])
        config.config_values['rs_config']['charging_cutoff_volt'] = determine_number_type(updated_values['Charge Cutoff Voltage (V)'])
        config.config_values['rs_config']['charging_cutoff_capacity'] = determine_number_type(updated_values['Charge Cutoff Capacity (%)'])
        config.config_values['rs_config']['temperature_caution'] = determine_number_type(updated_values['Temperature Caution (°C)'])
        config.config_values['rs_config']['temperature_alarm'] = determine_number_type(updated_values['Temperature Alarm (°C)'])
        config.config_values['rs_config']['temperature_critical'] = determine_number_type(updated_values['Temperature Critical (°C)'])

        # Save the updated configuration to the config file
        config.save_config()

        # Optionally show a message to the user indicating that the config has been updated
        messagebox.showinfo("Success", "Configuration updated and saved successfully.")
    
    def add_project(self):
        """Add a new project to the CAN project list using a custom input window."""

        # Create a new top-level window for input
        input_window = tk.Toplevel(self.master)
        input_window.title("Add Project")

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
        label = ttk.Label(input_window, text="Enter New Project Name:", font=("Helvetica", 12))
        label.pack(pady=10)

        project_name_var = tk.StringVar()  # Use StringVar to store the project name
        entry = ttk.Entry(input_window, textvariable=project_name_var)
        entry.pack(pady=5)

        # Function to handle OK button click
        def handle_ok():
            new_project_name = project_name_var.get().strip()
            if new_project_name and new_project_name not in config.config_values['rs_config']['projects']:
                config.config_values['rs_config']['projects'].append(new_project_name)
                config.save_config()  # Save the config with the new project
                self.project_combobox['values'] = config.config_values['rs_config']['projects']  # Update the combobox options
                self.project_combobox.set(new_project_name)  # Select the newly added project
                input_window.destroy()  # Close the input window
            elif new_project_name in config.config_values['rs_config']['projects']:
                messagebox.showwarning("Duplicate Project", f"Project '{new_project_name}' already exists.")
            else:
                messagebox.showwarning("Invalid Input", "Project name cannot be empty.")

        # OK Button
        ok_button = ttk.Button(input_window, text="OK", command=handle_ok)
        ok_button.pack(pady=10)

        # Focus the input window and entry
        input_window.focus()
        entry.focus_set()

        # Bind the Return key to trigger the OK button
        input_window.bind('<Return>', lambda event: handle_ok())

    def edit_project(self):
        """Edit the selected project name using a custom input window."""
        selected_project = self.project_combobox.get()

        # Check if a project is selected
        if not selected_project or selected_project == "Select a Project":
            messagebox.showwarning("No Project Selected", "Please select a project to edit.")
            return

        # Create a new top-level window for input
        input_window = tk.Toplevel(self.master)
        input_window.title("Edit Project")

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
        label = ttk.Label(input_window, text=f"Edit Project Name ({selected_project}):", font=("Helvetica", 12))
        label.pack(pady=10)

        project_name_var = tk.StringVar(value=selected_project)  # Pre-fill with current project name
        entry = ttk.Entry(input_window, textvariable=project_name_var)
        entry.pack(pady=5)

        # Function to handle OK button click
        def handle_ok():
            new_project_name = project_name_var.get().strip()
            if new_project_name and new_project_name not in config.config_values['rs_config']['projects']:
                index = config.config_values['rs_config']['projects'].index(selected_project)
                config.config_values['rs_config']['projects'][index] = new_project_name.strip()
                config.save_config()  # Save the config with the edited project
                self.project_combobox['values'] = config.config_values['rs_config']['projects']  # Update the combobox options
                self.project_combobox.set(new_project_name.strip())  # Set the newly edited project as selected
                input_window.destroy()  # Close the input window
            elif new_project_name in config.config_values['rs_config']['projects']:
                messagebox.showwarning("Duplicate Project", f"Project '{new_project_name}' already exists.")
            else:
                messagebox.showwarning("Invalid Input", "Project name cannot be empty.")

        # OK Button
        ok_button = ttk.Button(input_window, text="OK", command=handle_ok)
        ok_button.pack(pady=10)

        # Focus the input window and entry
        input_window.focus()
        entry.focus_set()

        # Bind the Return key to trigger the OK button
        input_window.bind('<Return>', lambda event: handle_ok())
    
    def delete_project(self):
        """Delete the selected project from the CAN project list."""
        selected_project = self.project_combobox.get()

        # Check if a project is selected
        if not selected_project or selected_project == "Select a Project":
            messagebox.showwarning("No Project Selected", "Please select a project to delete.")
            return

        # Proceed with deleting if a project is selected
        if selected_project in config.config_values['rs_config']['projects']:
            config.config_values['rs_config']['projects'].remove(selected_project)
            config.save_config()  # Save the config with the removed project
            self.project_combobox['values'] = config.config_values['rs_config']['projects']  # Update the combobox options
            self.project_combobox.set('Select a Project')  # Reset the combobox selection
            messagebox.showinfo("Deleted project", f"Project {selected_project} deleted successfully")

    def toggle_dynamic_fields(self):
        if self.show_dynamic_fields_var.get():
            self.dynamic_fields_frame.grid()  # Show the dynamic fields
            self.maintenance_button.grid_remove()  # Hide the "Set 100A and Turn ON Load" button
            self.turn_off_button.grid_remove()  # Hide the "Turn OFF Load" button
        else:
            self.dynamic_fields_frame.grid_remove()  # Hide the dynamic fields
            self.maintenance_button.grid()  # Show the "Set 100A and Turn ON Load" button
            self.turn_off_button.grid()  # Show the "Turn OFF Load" button
    
    def save_dynamic_fields(self):
        """Save the dynamic field values entered in the L1, L2, T1, T2, and Repeat fields."""
        # Fetch the values from the entry fields
        l1_value = self.l1_entry.get()
        l2_value = self.l2_entry.get()
        t1_value = self.t1_entry.get()
        t2_value = self.t2_entry.get()
        repeat_value = self.repeat_entry.get()

        # Process the values (for now, we will print them as a placeholder)
        print(f"L1: {l1_value}, L2: {l2_value}, T1: {t1_value}, T2: {t2_value}, Repeat: {repeat_value}")
