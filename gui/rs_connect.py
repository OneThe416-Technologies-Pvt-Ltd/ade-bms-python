import tkinter as tk
import winreg
import customtkinter as ctk
import serial.tools.list_ports
from tkinter import messagebox
from pcomm_api.pcomm import connect_to_serial_port,set_active_protocol

import re

class RSConnection(tk.Frame):
    rs_connected = False

    def __init__(self, master, main_window=None):
        super().__init__(master)
        self.main_window = main_window

        # Check if Moxa UPort 1650-8 is connected
        if self.is_moxa_connected():
            # If connected, create the widgets for COM port selection and connection
            self.create_widgets()
        else:
            self.create_widgets()

    def is_moxa_connected(self):
        """Check if Moxa UPort 1650-8 is connected."""
        ports = list(serial.tools.list_ports.comports())
        for port in ports:
            if 'MOXA' in port.description or '1650-8' in port.description:
                return True
        return False

    def create_widgets(self):
        # RS232/RS422 heading label
        tk.Label(self, text="RS232 / RS422 Connector", font=("Helvetica", 16, "bold")).grid(row=0, columnspan=2, pady=10)
        
        # Separator line
        separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
        separator.grid(row=1, columnspan=2, sticky="we", padx=20, pady=(0, 10))

        # Label and dropdown for available COM ports
        tk.Label(self, text="Select COM Port:").grid(row=2, column=0, padx=20, sticky=tk.W)
        self.com_ports = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
        self.com_ports.set("Select a COM Port")  # Set the default value shown in the dropdown
        self.com_ports.grid(row=2, column=1, padx=20, pady=5)

        # Label and dropdown for selecting RS232/RS422
        tk.Label(self, text="Select RS-232/RS-422 Mode:").grid(row=3, column=0, padx=20, sticky=tk.W)
        self.rs232_422_modes = ctk.CTkComboBox(self, values=["RS-232", "RS-422"])
        self.rs232_422_modes.set("Select a Mode")  # Set the default value shown in the dropdown
        self.rs232_422_modes.grid(row=3, column=1, padx=20, pady=5)

        # Connect button
        self.btnConnect = ctk.CTkButton(self, text="Connect", command=self.on_connect, fg_color="green", hover_color='green')
        self.btnConnect.grid(row=4, columnspan=2, pady=10)

        # Center the frame within parent
        self.grid_rowconfigure(5, weight=1)  # Ensure row 5 expands to center vertically
        self.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center horizontally
        self.pack(expand=True, fill='both')

    def set_interface_mode(self,com_port, mode):
        """
        Set the interface mode for the given COM port (RS-232 or RS-422).

        Parameters:
        com_port: The COM port (e.g., 'COM3')
        mode: The desired mode ('RS232' or 'RS422')
        """
        try:
            # Open the registry key where the Moxa configuration is stored (Modify as per your hardware's documentation)
            registry_path = r"SYSTEM\CurrentControlSet\Services\MoxaUsbSerial\Parameters\%s" % com_port

            # Open the registry key for modification
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_SET_VALUE)

            # Set the interface value: "RS232" or "RS422"
            if mode.upper() == "RS-232":
                winreg.SetValueEx(key, "Interface", 0, winreg.REG_SZ, "RS-232")
            elif mode.upper() == "RS-422":
                winreg.SetValueEx(key, "Interface", 0, winreg.REG_SZ, "RS-422")
            else:
                raise ValueError("Invalid mode selected. Choose 'RS232' or 'RS422'.")

            winreg.CloseKey(key)
            print(f"Successfully set {com_port} to {mode} mode.")
            messagebox.showinfo("Success", f"{com_port} set to {mode} mode.")

        except Exception as e:
            print(f"Failed to set {com_port} to {mode} mode. Error: {e}")
            messagebox.showerror("Error", f"Failed to set {com_port} to {mode} mode. Error: {e}")

    def get_com_ports(self):
        """Get a list of available COM ports for Moxa devices and display 'Port X connected to COMY'."""
        ports = list(serial.tools.list_ports.comports())
        moxa_ports = []

        for port in ports:
            # Check if the port belongs to the Moxa hub
            if 'MOXA' in port.description or '1650-8' in port.description:
                # Extract the port number from the description (e.g., "Port 1", "Port 2")
                match = re.search(r'Port\s*(\d+)', port.description)
                if match:
                    port_number = int(match.group(1))  # Extract the port number as an integer for sorting
                    com_port = port.device  # Get the COM port (e.g., "COM3")
                    moxa_ports.append((port_number, f"Port {port_number} ({com_port})"))

        # Sort by the port number (the first element in the tuple)
        moxa_ports.sort(key=lambda x: x[0])

        # Return only the formatted string part
        return [port[1] for port in moxa_ports]

    def on_connect(self):
        selected_port = self.com_ports.get()  # Get the selected COM port from the dropdown
        selected_mode = self.rs232_422_modes.get()
        set_active_protocol(selected_flag=selected_mode)
        match = re.search(r"\((COM\d+)\)", selected_port)
        if match:
            com_port_name = match.group(1)  # Extract the COM port (e.g., "COM7")
            print(f"Connecting to {com_port_name}...")
            # self.set_interface_mode(com_port_name, mode=selected_mode)
            ser=connect_to_serial_port(com_port_name,selected_mode)
            if ser:
                    self.main_window.show_rs_battery_info()
        else:
            print("Invalid port selected.")
            messagebox.showwarning("Selection Error", "Please select a valid COM port.")

    def set_rs232_interface(self, com_port):
        """Set the interface to RS232 using Windows Registry."""
        try:
            # Open the registry key for the Moxa UPort settings (modify this key as needed)
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Serial', 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, 'Interface', 0, winreg.REG_SZ, 'RS232')
            winreg.CloseKey(key)
            print(f"Port {com_port} set to RS232 mode.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set {com_port} to RS232 mode. Error: {e}")

    def set_rs422_interface(self, com_port):
        """Set the interface to RS422 using Windows Registry."""
        try:
            # Open the registry key for the Moxa UPort settings (modify this key as needed)
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Serial', 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, 'Interface', 0, winreg.REG_SZ, 'RS422')
            winreg.CloseKey(key)
            print(f"Port {com_port} set to RS422 mode.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set {com_port} to RS422 mode. Error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RSConnection(root)
    app.pack()

    root.mainloop()
