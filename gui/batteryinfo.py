#can_battery_info.py

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
from pcan_api.custom_pcan_methods import device_data, unit_mapping, name_mapping, update_device_data, pcan_write_control, pcan_write_read

class CanBatteryInfo(tk.Frame):
    def __init__(self, master, main_window, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.main_window = main_window
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button

        # Set initial window size and update it
        self.master.geometry("1200x600")
        self.master.update_idletasks()  # Update the window to get accurate dimensions

        # Calculate one-fourth of the main window width for the side menu
        self.side_menu_width_ratio = 0.20  # 20% for side menu
        self.update_frame_sizes()  # Set initial size
        self.auto_refresh_var = tk.BooleanVar()
        self.logging_active = False
        self.charge_fet_status = False
        self.discharge_fet_status = False


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
        self.home_icon = self.load_icon("assets/images/home.png")
        self.controls_icon = self.load_icon("assets/images/settings.png")
        self.info_icon = self.load_icon("assets/images/info.png")
        self.help_icon = self.load_icon("assets/images/help.png")
        self.dashboard_icon = self.load_icon("assets/images/menu.png")

        self.side_menu_heading = ttk.Label(
            self.side_menu_frame,
            text="ADE BMS",
            font=("Courier New", 18, "bold"),
            bootstyle="inverse-info",  # Optional: change color style
            anchor="center"
        )
        self.side_menu_heading.pack(fill="x", pady=(10, 20))

        # # Add buttons to the side menu with icons
        self.home_button =  ttk.Button(
            self.side_menu_frame,
            text=" Home         ",
            command=lambda: self.select_button(self.home_button, self.main_window.show_home),
            image=self.home_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info",
        )
        self.home_button.pack(fill="x", pady=5)
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
        self.control_button = ttk.Button(
            self.side_menu_frame,
            text=" Controls      ",
            command=lambda: self.select_button(self.control_button, self.show_controls),
            image=self.controls_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.control_button.pack(fill="x", pady=5)
        self.help_button = ttk.Button(
            self.side_menu_frame,
            text=" Help            ",
            command=lambda: self.select_button(self.help_button),
            image=self.help_icon,
            compound="left",  # Place the icon to the left of the text
            bootstyle="info"
        )
        self.help_button.pack(fill="x", pady=5)
        # asyncio.run(getdata())
        # Initially display the home content
        self.show_dashboard()

        # Bind the window resize event to adjust frame sizes dynamically
        self.master.bind('<Configure>', self.on_window_resize)

    def load_icon(self, path, size=(24, 24)):
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
        # Reset the previous selected button
        if self.selected_button:
            self.selected_button.config(bootstyle="info")

        # Highlight the selected button
        button.config(bootstyle="primary")
        self.selected_button = button

        # Execute the button's command
        if command:
            command()

    def clear_content_frame(self):
        # Clear the content frame before displaying new content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def show_info(self):
        self.clear_content_frame()

        # Title label for the Info Page
        # label = ttk.Label(self.content_frame, text="Battery Info", font=("Helvetica", 16))
        # label.pack(pady=20)
        bms_control_label = ttk.Label(self.content_frame, text="Battery Info", font=("Helvetica", 16, "bold"))
        bms_control_label.pack(pady=(5, 5))

        self.refresh_icon = self.load_icon("assets/images/refresh.png")
        self.start_icon = self.load_icon("assets/images/start.png")
        self.stop_icon = self.load_icon("assets/images/stop.png")
        self.file_icon = self.load_icon("assets/images/folder.png")

        # Create a Frame to hold the checkbox and buttons in one row with a white background
        control_frame = ttk.Frame(self.content_frame)
        control_frame.pack(pady=10, padx=20, fill="x")

        # Initialize the checkbox and link it to the auto-refresh function
        auto_refresh_checkbox = ttk.Checkbutton(
            control_frame, 
            text="Auto Refresh", 
            variable=self.auto_refresh_var, 
            width=12, 
            command=self.auto_refresh  # Link the checkbox command to start/stop auto-refresh
        )
        auto_refresh_checkbox.pack(side="left", padx=5)

        # Buttons for Refresh, Start Logging, and Stop Logging with consistent style
        self.refresh_button = ttk.Button(control_frame, text="Refresh", image=self.refresh_icon, compound="left", command=self.refresh_info, width=8, bootstyle="info")
        self.refresh_button.pack(side="left", padx=5)

        timer_value_label = ttk.Label(control_frame, text="Timer")
        timer_value_label.pack(side="left", padx=5)

        self.timer_value = tk.StringVar(value=5)
        self.timer = ttk.Spinbox(
            control_frame,
            from_=1,
            to=30,
            width=5,
            values=(1, 5, 10, 15, 20, 30),
            textvariable=self.timer_value,
            )
        self.timer.pack(side="left", padx=5)

        self.start_logging_button = ttk.Button(control_frame, text="Start Log", image=self.start_icon, compound="left", command=self.start_logging, bootstyle="success")
        self.start_logging_button.pack(side="left", padx=5)

        self.stop_logging_button = ttk.Button(control_frame, text="Stop Log", image=self.stop_icon, compound="left", command=self.stop_logging, bootstyle="danger", state=tk.DISABLED)
        self.stop_logging_button.pack(side="left", padx=5)

        # Set button states based on whether logging is active
        if self.logging_active:
            self.start_logging_button.config(state=tk.DISABLED)
            self.stop_logging_button.config(state=tk.NORMAL)
        else:
            self.start_logging_button.config(state=tk.NORMAL)
            self.stop_logging_button.config(state=tk.DISABLED)

        file_button = ttk.Button(control_frame, image=self.file_icon, compound="left", command=self.folder_open, bootstyle="success")
        file_button.pack(side="left", padx=5)

        # Frame to hold the Treeview and Scrollbar
        table_frame = ttk.Frame(self.content_frame)
        table_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Create a vertical scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        # Table for showing labels, values, and units
        columns = ('Name', 'Value', 'Units')
        self.info_table = ttk.Treeview(table_frame, columns=columns, show='headings', style="Custom.Treeview", yscrollcommand=scrollbar.set)

        # Configure the scrollbar to work with the Treeview
        scrollbar.config(command=self.info_table.yview)

        # Configure columns to be center-aligned with custom widths and text color
        self.info_table.column('Name', anchor='center', width=200)
        self.info_table.column('Value', anchor='center', width=200)
        self.info_table.column('Units', anchor='center', width=100)

        self.info_table.heading('Name', text='Name', anchor='center')
        self.info_table.heading('Value', text='Value', anchor='center')
        self.info_table.heading('Units', text='Units', anchor='center')

        # Apply custom style for the Treeview with black heading text
        style = ttk.Style()
        style.configure("Custom.Treeview.Heading", font=("Helvetica", 16, "bold"), background="#e6e6e6", padding=5)
        style.configure("Custom.Treeview", rowheight=30, font=("Helvetica", 14), background="#f0f0f0", fieldbackground="#ffffff")

        # Simulate borders by adding padding and using contrasting background colors
        style.configure("Custom.Treeview", padding=(1, 1), borderwidth=1, relief="solid")
        style.map("Custom.Treeview", background=[("selected", "#007bff")], foreground=[("selected", "white")])
        style.configure("Custom.Treeview", background=["#e6e6e6"], fieldbackground="#ffffff")

        # Set up the layout for the table
        self.info_table.pack(fill="both", expand=True)

        # Insert data
        for index, (key, value) in enumerate(device_data.items()):
            name = name_mapping.get(key, key)
            unit = unit_mapping.get(key, key)
            self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))

        self.select_button(self.info_button)

    def auto_refresh(self):
        if self.auto_refresh_var.get():
            asyncio.run(update_device_data())
            # Clear the existing table rows
            for item in self.info_table.get_children():
                self.info_table.delete(item)
    
            # Repopulate the table with updated data
            for index, (key, value) in enumerate(device_data.items()):
                name = name_mapping.get(key, key)
                unit = unit_mapping.get(key, key)
                self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
            self.master.after(5000, self.auto_refresh)

    def refresh_info(self):
        asyncio.run(update_device_data())
        # Clear the existing table rows
        for item in self.info_table.get_children():
            self.info_table.delete(item)

        # Repopulate the table with updated data
        for index, (key, value) in enumerate(device_data.items()):
            name = name_mapping.get(key, key)
            unit = unit_mapping.get(key, key)
            self.info_table.insert('', 'end', values=(name, value, unit), tags=('evenrow' if index % 2 == 0 else 'oddrow'))
    
    def start_logging(self):
        self.logging_active = True
        # Disable the Start Logging button and enable the Stop Logging button
        self.start_logging_button.config(state=tk.DISABLED)
        self.stop_logging_button.config(state=tk.NORMAL)

        # Create the folder path
        folder_path = os.path.join(os.path.expanduser("~"), "Documents", "ADE Log Folder")
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Create a timestamped file name
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"CAN_Connector_Log_{timestamp}.xlsx"
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

            try:
                self.workbook.save(self.file_path)
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

            self.log_interval = int(self.timer_value.get()) * 1000

            # Schedule the next log
            self.master.after(self.log_interval, self.log_data)
    
    def stop_logging(self):
        # Stop logging process and update button states
        self.logging_active = False
        # Disable the Stop Logging button and enable the Start Logging button
        self.stop_logging_button.config(state=tk.DISABLED)
        self.start_logging_button.config(state=tk.NORMAL)

        self.logging_active = False  # Stop the logging loop
        if hasattr(self, 'workbook'):
            try:
                self.workbook.save(self.file_path)
            finally:
                self.workbook.close()
        folder_path = os.path.dirname(self.file_path)
        os.startfile(folder_path)

    def folder_open(self):
        folder_path = os.path.join(os.path.expanduser("~"), "Documents", "ADE Log Folder")
        if os.path.exists(folder_path):
            os.startfile(folder_path)
        else:
            messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")

    def get_bootstyle(self, value):
        if value < 20:
            return "success"
        elif value < 50:
            return "info"
        elif value < 80:
            return "warning"
        else:
            return "danger"
    
    def show_dashboard(self):
        self.clear_content_frame()

        # Top row frame (for Temperature and Current meters)
        top_row_frame = ttk.Frame(self.content_frame)
        top_row_frame.pack(fill="x", padx=5, pady=5)

        # Temperature Meter (Top Left)
        temp_frame = ttk.Frame(top_row_frame, padding=10)
        temp_frame.pack(side="left", fill="x", expand=True, padx=10)
        temp_label = ttk.Label(temp_frame, text="Temperature", font=("Helvetica", 16, "bold"))
        temp_label.pack(pady=(5, 5))
        temp_meter = ttk.Meter(
            master=temp_frame,
            metersize=200,
            amountused=device_data['temperature'],
            meterthickness=10,
            metertype="semi",
            subtext="Temperature",
            textright="°C",
            amounttotal=100,
            bootstyle="danger",
            stripethickness=0,
            subtextfont='-size 10'
        )
        temp_meter.pack()

        # Current Meter (Top Right)
        current_frame = ttk.Frame(top_row_frame, padding=10)
        current_frame.pack(side="left", fill="x", expand=True, padx=10)
        current_label = ttk.Label(current_frame, text="Current", font=("Helvetica", 16, "bold"))
        current_label.pack(pady=(5, 5))
        current_meter = ttk.Meter(
            master=current_frame,
            metersize=200,
            amountused=device_data['current'],
            meterthickness=10,
            metertype="semi",
            subtext="Current",
            textright="mA",
            amounttotal=100,
            bootstyle="info",
            stripethickness=10,
            subtextfont='-size 10'
        )
        current_meter.pack()

        # Bottom row frame (for Capacity and Voltage meters)
        bottom_row_frame = ttk.Frame(self.content_frame)
        bottom_row_frame.pack(fill="x", padx=5, pady=5)

        # Capacity Meter (Bottom Left)
        capacity_frame = ttk.Frame(bottom_row_frame, padding=10)
        capacity_frame.pack(side="left", fill="x", expand=True, padx=10)
        capacity_label = ttk.Label(capacity_frame, text="Capacity", font=("Helvetica", 16, "bold"))
        capacity_label.pack(pady=(10, 10))
        # Calculate amountused, with a check to avoid division by zero
        design_capacity = device_data.get('design_capacity', 0)
        remaining_capacity = device_data.get('remaining_capacity', 0)

        if design_capacity == 0:
            amountused = 0
        else:
            amountused = (remaining_capacity / design_capacity) * 100

        capacity_meter = ttk.Meter(
            master=capacity_frame,
            metersize=200,
            amountused=round(amountused, 1),
            meterthickness=10,
            metertype="semi",
            subtext="Capacity",
            textright="%",
            amounttotal=100,
            bootstyle="warning",
            stripethickness=8,
            subtextfont='-size 10'
        )
        capacity_meter.pack()

        # Voltage Meter (Bottom Right)
        voltage_frame = ttk.Frame(bottom_row_frame, padding=10)
        voltage_frame.pack(side="left", fill="x", expand=True, padx=10)
        voltage_label = ttk.Label(voltage_frame, text="Voltage", font=("Helvetica", 16, "bold"))
        voltage_label.pack(pady=(10, 10))
        voltage_meter = ttk.Meter(
            master=voltage_frame,
            metersize=200,
            amountused=device_data['voltage'],
            meterthickness=10,
            metertype="semi",
            subtext="Voltage",
            textright="V",
            amounttotal=300,
            bootstyle="dark",
            stripethickness=5,
            subtextfont='-size 10'
        )
        voltage_meter.pack()

        # Pack the content_frame itself (if necessary, but typically this frame is already packed)
        self.content_frame.pack(fill="both", expand=True)
        self.select_button(self.dashboard_button)

    def show_help(self):
        self.clear_content_frame()
        label = ttk.Label(self.content_frame, text="Help", font=("Helvetica", 16))
        label.pack(pady=20)
        self.select_button(self.help_button)

    def show_controls(self):
        self.clear_content_frame()

        # label = ttk.Label(self.content_frame, text="Controls", font=("Helvetica", 16))
        # label.pack(pady=20, anchor="center")
        self.reset_icon = self.load_icon("assets/images/reset.png")
        
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

        self.bms_reset_button = ttk.Button(self.content_frame, text="Reset", image=self.reset_icon, compound="left", command=self.bmsreset, width=12, bootstyle="danger")
        self.bms_reset_button.pack(pady=(5, 5))

        activate_heater_label = ttk.Label(self.content_frame, text="Activate Heater", font=("Helvetica", 16, "bold"))
        activate_heater_label.pack(pady=(5, 5))

        self.activate_heater_button = ttk.Button(self.content_frame, text="OFF", compound="left", command=self.activate_heater, width=12, bootstyle="danger")
        self.activate_heater_button.pack(pady=(5, 5))

        # Pack the content_frame itself (if necessary, but typically this frame is already packed)
        self.content_frame.pack(fill="both", expand=True)

        # Highlight the control button in the side menu
        self.select_button(self.control_button)

    def toggle_button_style(self, var, button, control_type):
            charge_on = self.check_var1.get()
            discharge_on = self.check_var2.get()

            if control_type == 'charge':
                self.charge_fet_status = charge_on
                if charge_on:
                    if discharge_on:
                        pcan_write_control('both_on')
                    else:
                        pcan_write_control('charge_on')
                else:
                    if discharge_on:
                        pcan_write_control('discharge_on')
                    else:
                        pcan_write_control('both_off')
            elif control_type == 'discharge':
                self.discharge_fet_status = discharge_on
                if discharge_on:
                    if charge_on:
                        pcan_write_control('both_on')
                    else:
                        pcan_write_control('discharge_on')
                else:
                    if charge_on:
                        pcan_write_control('charge_on')
                    else:
                        pcan_write_control('both_off')

            # Update button styles
            button.config(bootstyle="success round-toggle" if var.get() else "danger round-toggle")

    def bmsreset(self):
        pcan_write_control('bms_reset')
        
    def activate_heater(self):
        if self.activate_heater_button.cget('text') == "OFF":
            self.activate_heater_button.configure(text="ON", bootstyle="success")
        else:
            self.activate_heater_button.configure(text="OFF", bootstyle="danger")

