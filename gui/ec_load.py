import pyvisa
import tkinter as tk
from tkinter import ttk
import threading
import time

class ElectronicLoadApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Electronic Load Control")
        self.geometry("600x400")

        # PyVISA resource manager
        self.rm = pyvisa.ResourceManager()

        # Store the Chroma device
        self.device = None

        # Create frames for different sections
        self.create_frames()

    def create_frames(self):
        # Frame 1: Connection, Current, Time, and Load Control
        self.frame1 = ttk.Labelframe(self, text="Connect & Set Load", padding=10)
        self.frame1.pack(padx=10, pady=10, fill='x')

        self.create_connect_buttons(self.frame1)
        self.create_set_load_controls(self.frame1)

        # Frame 2: Manual Control of Load
        self.frame2 = ttk.Labelframe(self, text="Manual Control", padding=10)
        self.frame2.pack(padx=10, pady=10, fill='x')
        self.create_manual_controls(self.frame2)

    def create_connect_buttons(self, frame):
        """Creates the connection buttons for the VISA communication."""
        # Scan Devices Button
        self.scan_button = ttk.Button(frame, text="Find Devices", command=self.find_devices)
        self.scan_button.grid(row=0, column=0, padx=5, pady=5)

        # Connect Button
        self.connect_button = ttk.Button(frame, text="Connect to Chroma", command=self.connect_to_chroma)
        self.connect_button.grid(row=1, column=0, padx=5, pady=5)

        # Status Label
        self.status_label = ttk.Label(frame, text="Status: Disconnected")
        self.status_label.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

    def create_set_load_controls(self, frame):
        """Creates the controls for setting current, time, and turning on the load."""
        # Input for Current
        self.current_label = ttk.Label(frame, text="Set Current (A):")
        self.current_label.grid(row=3, column=0, padx=5, pady=5)

        self.current_entry = ttk.Entry(frame)
        self.current_entry.grid(row=3, column=1, padx=5, pady=5)

        # Input for Time (in seconds)
        self.time_label = ttk.Label(frame, text="Set Time (s):")
        self.time_label.grid(row=4, column=0, padx=5, pady=5)

        self.time_entry = ttk.Entry(frame)
        self.time_entry.grid(row=4, column=1, padx=5, pady=5)

        # Button to Set Load
        self.set_load_button = ttk.Button(frame, text="Set Load and Start Timer", command=self.set_load_and_start_timer)
        self.set_load_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

    def create_manual_controls(self, frame):
        """Creates the manual controls for turning the load on and off."""
        # Input for Manual Current
        self.manual_current_label = ttk.Label(frame, text="Manual Current (A):")
        self.manual_current_label.grid(row=0, column=0, padx=5, pady=5)

        self.manual_current_entry = ttk.Entry(frame)
        self.manual_current_entry.grid(row=0, column=1, padx=5, pady=5)

        # Manual Turn On/Off Buttons
        self.manual_on_button = ttk.Button(frame, text="Turn Load ON", command=self.manual_turn_on)
        self.manual_on_button.grid(row=1, column=0, padx=5, pady=5)

        self.manual_off_button = ttk.Button(frame, text="Turn Load OFF", command=self.manual_turn_off)
        self.manual_off_button.grid(row=1, column=1, padx=5, pady=5)

    def find_devices(self):
        """Find all VISA devices and identify the Chroma device."""
        devices = self.rm.list_resources()
        chroma_device = None
        for device in devices:
            try:
                instrument = self.rm.open_resource(device)
                response = instrument.query("*IDN?")
                if "Chroma" in response:
                    chroma_device = device
                    self.status_label.config(text=f"Found Chroma Device: {device}")
                    break
            except Exception as e:
                print(f"Error querying device {device}: {e}")
        if chroma_device is None:
            self.status_label.config(text="Chroma device not found")

    def connect_to_chroma(self):
        """Connect to the Chroma device if found."""
        devices = self.rm.list_resources()
        for device in devices:
            try:
                instrument = self.rm.open_resource(device)
                response = instrument.query("*IDN?")
                if "Chroma" in response:
                    self.device = instrument
                    self.status_label.config(text=f"Connected to Chroma: {device}")
                    break
            except Exception as e:
                self.status_label.config(text=f"Error connecting to device: {e}")

    def set_load_and_start_timer(self):
        """Set the current and time, turn on load, and turn off after the time ends."""
        if not self.device:
            self.status_label.config(text="Error: Not connected to device")
            return

        current = self.current_entry.get()
        time_duration = int(self.time_entry.get())

        try:
            # Set the current
            self.device.write(f":CURR {current}\n")

            # Turn on the load after a timer
            self.device.write(":LOAD ON\n")
            self.status_label.config(text=f"Load ON: {current}A for {time_duration}s")

            # Run a separate thread to turn off the load after the specified time
            threading.Thread(target=self.turn_off_load_after_time, args=(time_duration,)).start()

        except Exception as e:
            self.status_label.config(text=f"Error: {e}")

    def turn_off_load_after_time(self, time_duration):
        """Turn off the load after the specified time in seconds."""
        time.sleep(time_duration)
        self.device.write(":LOAD OFF\n")
        self.status_label.config(text="Load OFF (Timer Ended)")

    def manual_turn_on(self):
        """Manually turn on the load with specified current."""
        if not self.device:
            self.status_label.config(text="Error: Not connected to device")
            return

        current = self.manual_current_entry.get()
        try:
            # Set the current
            self.device.write(f":CURR {current}\n")
            # Turn on the load
            self.device.write(":LOAD ON\n")
            self.status_label.config(text=f"Load Manually ON: {current}A")
        except Exception as e:
            self.status_label.config(text=f"Error: {e}")

    def manual_turn_off(self):
        """Manually turn off the load."""
        if not self.device:
            self.status_label.config(text="Error: Not connected to device")
            return
        try:
            self.device.write(":LOAD OFF\n")
            self.status_label.config(text="Load Manually OFF")
        except Exception as e:
            self.status_label.config(text=f"Error: {e}")

# Run the application
if __name__ == "__main__":
    app = ElectronicLoadApp()
    app.mainloop()
