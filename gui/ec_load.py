import tkinter as tk
from tkinter import ttk
import serial

class ElectronicLoadApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Electronic Load Control")
        self.geometry("400x200")
        
        self.port = None
        
        # Create the connection interface
        self.create_interface()

    def create_interface(self):
        # Serial Port Label and Entry
        self.port_label = ttk.Label(self, text="USB Port:")
        self.port_label.pack(pady=5)
        
        self.port_entry = ttk.Entry(self)
        self.port_entry.pack(pady=5)
        self.port_entry.insert(0, "COM3")  # Default to COM3, change as needed
        
        # Button to Connect to Device
        self.connect_button = ttk.Button(self, text="Connect", command=self.connect_to_device)
        self.connect_button.pack(pady=10)
        
        # Button to Set Load to 2A and Turn On
        self.update_button = ttk.Button(self, text="Set Load to 2A & Turn On", command=self.set_load_and_turn_on)
        self.update_button.pack(pady=10)
        
        # Status Label
        self.status_label = ttk.Label(self, text="Status: Disconnected")
        self.status_label.pack(pady=10)

    def connect_to_device(self):
        """Connect to the electronic load via USB."""
        port_name = self.port_entry.get()
        try:
            # Set up the USB connection
            self.port = serial.Serial(port=port_name, baudrate=9600, timeout=1)
            self.status_label.config(text="Status: Connected")
        except serial.SerialException as e:
            self.status_label.config(text=f"Error: {e}")
            self.port = None

    def set_load_and_turn_on(self):
        """Send commands to set load to 2A and turn on the load."""
        if self.port:
            try:
                # Set the current to 2A
                self.port.write(b":CURR 2\n")
                
                # Turn on the load
                self.port.write(b":LOAD ON\n")
                
                self.status_label.config(text="Load set to 2A and turned ON")
            except serial.SerialException as e:
                self.status_label.config(text=f"Error: {e}")
        else:
            self.status_label.config(text="Error: Not connected to device")

# Run the application
if __name__ == "__main__":
    app = ElectronicLoadApp()
    app.mainloop()
