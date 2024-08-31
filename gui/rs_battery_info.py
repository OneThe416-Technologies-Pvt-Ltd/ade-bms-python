import tkinter as tk
import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk
from ctypes import create_string_buffer
from tkinter import messagebox
import serial.tools.list_ports
from pcomm_api.pcomm import sio_open, sio_ioctl, sio_write, sio_read, sio_close, BAUD_RATE, PARITY_NONE, DATA_BITS_8, STOP_BITS_1, BUFFER_SIZE

class RSBatteryInfo:
    def __init__(self, master, main_window=None):
        self.master = master
        self.main_window = main_window
        self.selected_button = None  # Track the currently selected button

        # Set initial window size and update it
        self.master.geometry("1100x800")
        self.master.update_idletasks()  # Update the window to get accurate dimensions

        # Main frame for the entire layout
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Top frame for COM port selection and connection
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # COM Port selection
        ttk.Label(self.top_frame, text="Select COM Port:", font=("Helvetica", 12)).grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.com_ports = ttk.Combobox(self.top_frame, values=self.get_com_ports(), font=("Helvetica", 12))
        self.com_ports.set("Select COM Port")
        self.com_ports.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Connect Button
        self.connect_button = ttk.Button(self.top_frame, text="Connect", command=self.on_connect, bootstyle="primary")
        self.connect_button.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Battery Info
        ttk.Label(self.top_frame, text="Battery ID:", font=("Helvetica", 12)).grid(row=0, column=3, padx=20, pady=5, sticky="e")
        self.battery_id = ttk.Label(self.top_frame, text="", font=("Helvetica", 12))
        self.battery_id.grid(row=0, column=4, padx=10, pady=5, sticky="w")

        ttk.Label(self.top_frame, text="Battery S/N:", font=("Helvetica", 12)).grid(row=0, column=5, padx=20, pady=5, sticky="e")
        self.battery_sn = ttk.Label(self.top_frame, text="", font=("Helvetica", 12))
        self.battery_sn.grid(row=0, column=6, padx=10, pady=5, sticky="w")

        # Middle frame for the grid layout of the battery info
        self.middle_frame = ttk.Frame(self.main_frame)
        self.middle_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
        self.middle_frame.grid_columnconfigure(0, weight=1)
        self.middle_frame.grid_columnconfigure(1, weight=1)
        self.middle_frame.grid_columnconfigure(2, weight=1)

        # Battery Parameters Frame
        battery_frame = ttk.Labelframe(self.middle_frame, text="Battery Parameters", padding=(10, 5))
        battery_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Battery parameters (Voltage, Temperature, Status)
        for i in range(7):  # Assuming 7 batteries
            ttk.Label(battery_frame, text=f"BAT{i+1}", font=("Helvetica", 10)).grid(row=0, column=i+1, padx=5)
            ttk.Label(battery_frame, text="Voltage", font=("Helvetica", 10)).grid(row=1, column=0, padx=5)
            ttk.Label(battery_frame, text="Temperature", font=("Helvetica", 10)).grid(row=2, column=0, padx=5)
            ttk.Label(battery_frame, text="Status", font=("Helvetica", 10)).grid(row=3, column=0, padx=5)

            # Placeholder labels for values
            ttk.Label(battery_frame, text="00.00").grid(row=1, column=i+1, padx=5)
            ttk.Label(battery_frame, text="00.00").grid(row=2, column=i+1, padx=5)
            ttk.Label(battery_frame, text="●", foreground="green").grid(row=3, column=i+1, padx=5)

        # IC Temperature
        ttk.Label(battery_frame, text="IC Temperature", font=("Helvetica", 10)).grid(row=4, column=0, padx=5)
        ttk.Label(battery_frame, text="00.00").grid(row=4, column=1, padx=5)

        # Charge/Discharge Level and Energy Discharged
        ttk.Label(battery_frame, text="Charge/Discharge Level", font=("Helvetica", 10)).grid(row=5, column=0, padx=5)
        ttk.Label(battery_frame, text="00.00").grid(row=5, column=1, padx=5)

        ttk.Label(battery_frame, text="Energy Discharged (Life Time)", font=("Helvetica", 10)).grid(row=6, column=0, padx=5)
        ttk.Label(battery_frame, text="00.00").grid(row=6, column=1, padx=5)

        # Charger Frame
        charger_frame = ttk.Labelframe(self.middle_frame, text="Charger", padding=(10, 5))
        charger_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # Charger control buttons and indicators
        ttk.Label(charger_frame, text="Charger On/Off", font=("Helvetica", 10)).grid(row=0, column=0, padx=5)
        ttk.Button(charger_frame, text="OFF", bootstyle="danger").grid(row=0, column=1, padx=5)

        ttk.Label(charger_frame, text="Output Relay On/Off", font=("Helvetica", 10)).grid(row=1, column=0, padx=5)
        ttk.Button(charger_frame, text="OFF", bootstyle="danger").grid(row=1, column=1, padx=5)

        # Input and Output
        ttk.Label(charger_frame, text="Input Voltage", font=("Helvetica", 10)).grid(row=2, column=0, padx=5)
        ttk.Label(charger_frame, text="00.00").grid(row=2, column=1, padx=5)

        ttk.Label(charger_frame, text="Output Current", font=("Helvetica", 10)).grid(row=3, column=0, padx=5)
        ttk.Label(charger_frame, text="00.00").grid(row=3, column=1, padx=5)

        # BUS Frame
        bus_frame = ttk.Labelframe(self.middle_frame, text="BUS", padding=(10, 5))
        bus_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

        # BUS ON/OFF Controls
        ttk.Label(bus_frame, text="ON/OFF Control", font=("Helvetica", 10)).grid(row=0, column=0, padx=5)
        ttk.Button(bus_frame, text="OFF", bootstyle="danger").grid(row=0, column=1, padx=5)
        ttk.Button(bus_frame, text="OFF", bootstyle="danger").grid(row=0, column=2, padx=5)

        # BUS Status Indicators
        ttk.Label(bus_frame, text="ON/OFF Status", font=("Helvetica", 10)).grid(row=1, column=0, padx=5)
        ttk.Label(bus_frame, text="●", foreground="green").grid(row=1, column=1, padx=5)

        # Bottom frame for Exit button
        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.grid(row=2, column=0, columnspan=3, pady=(20, 0))

        # Exit button
        self.exit_button = ttk.Button(self.bottom_frame, text="Exit", command=self.master.quit, bootstyle="secondary")
        self.exit_button.pack(side="right")

    def get_com_ports(self):
        """Get a list of available COM ports."""
        ports = list(serial.tools.list_ports.comports())
        return [port.device for port in ports]

    def on_connect(self):
        selected_port = self.com_ports.get()
        if selected_port:
            result = sio_open(self.get_com_ports().index(selected_port) + 1)
            if result == 0:
                config_result = sio_ioctl(self.get_com_ports().index(selected_port) + 1, BAUD_RATE, PARITY_NONE | DATA_BITS_8 | STOP_BITS_1)
                if config_result == 0:
                    messagebox.showinfo("Connection Successful", f"Connected to {selected_port}")
                    # Update battery info, etc.
                else:
                    messagebox.showerror("Connection Failed", "Failed to configure COM port.")
            else:
                messagebox.showerror("Connection Failed", "Failed to open COM port.")
        else:
            messagebox.showerror("Connection Error", "Please select a COM port.")

if __name__ == "__main__":
    root = tk.Tk()
    app = RSBatteryInfo(root)
    root.mainloop()
