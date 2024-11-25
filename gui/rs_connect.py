import tkinter as tk
import winreg
import customtkinter as ctk
import serial.tools.list_ports
from tkinter import messagebox
from pcomm_api.pcomm import *
import re
from helpers.logger import logger  # Import the logger
import traceback

class RSConnection(tk.Frame):
    rs_connected = False

    def __init__(self, master, main_window=None):
        """
        Initializes the RSConnection frame for managing serial communication settings.

        Parameters:
        master (tk.Tk): The master window to which this frame is attached.
        main_window (tk.Tk): Reference to the main window to call other functions like showing battery info.
        """
        super().__init__(master)
        self.main_window = main_window

        try:
            # Check if Moxa UPort 1650-8 is connected
            if self.is_moxa_connected():
                # If connected, create the widgets for COM port selection and connection
                self.create_widgets()
            else:
                # If not connected, still create the widgets but inform the user
                self.create_widgets()
        except Exception as e:
            # If any exception occurs during initialization, display an error message
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            logger.error(f"Error initializing RSConnection: {e}")
            messagebox.showerror("Initialization Error", f"Error initializing RSConnection: {e}")

    def is_moxa_connected(self):
        """Check if Moxa UPort 1650-8 is connected."""
        try:
            # List all available COM ports
            ports = list(serial.tools.list_ports.comports())
            for port in ports:
                if 'MOXA' in port.description or '1650-8' in port.description:
                    return True
            return False
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            logger.error(f"Error checking Moxa connection: {e}")
            messagebox.showerror("Connection Error", f"Error checking Moxa connection: {e}")
            return False

    def create_widgets(self):
        """Create the widgets for the RS232/RS422 connection interface."""
        try:
            # RS232/RS422 heading label
            tk.Label(self, text="RS232 / RS422 Connector", font=("Helvetica", 16, "bold")).grid(row=0, columnspan=2, pady=10)
            
            # Separator line
            separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
            separator.grid(row=1, columnspan=2, sticky="we", padx=20, pady=(0, 10))

            # Label and dropdown for selecting RS232/RS422
            tk.Label(self, text="Select RS-232/RS-422 Mode:").grid(row=3, column=0, padx=20, sticky=tk.W)
            self.rs232_422_modes = ctk.CTkComboBox(self, values=["RS-232", "RS-422"])
            self.rs232_422_modes.set("RS-232")  # Set the default value shown in the dropdown
            self.rs232_422_modes.grid(row=3, column=1, padx=20, pady=5)

            # Connect button
            self.btnConnect = ctk.CTkButton(self, text="Connect", command=self.on_connect, fg_color="green", hover_color='green')
            self.btnConnect.grid(row=4, columnspan=2, pady=10)

            # Center the frame within parent
            self.grid_rowconfigure(5, weight=1)  # Ensure row 5 expands to center vertically
            self.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center horizontally
            self.pack(expand=True, fill='both')
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            logger.error(f"Error creating widgets: {e}")
            messagebox.showerror("Widget Creation Error", f"Error creating widgets: {e}")

    def set_interface_mode(self, com_port, mode):
        """
        Set the interface mode for the given COM port (RS-232 or RS-422).

        Parameters:
        com_port (str): The COM port (e.g., 'COM3')
        mode (str): The desired mode ('RS232' or 'RS422')
        """
        try:
            # Open the registry key where the Moxa configuration is stored
            registry_path = r"SYSTEM\CurrentControlSet\Services\MoxaUsbSerial\Parameters\%s" % com_port
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_SET_VALUE)

            # Set the interface value: "RS232" or "RS422"
            if mode.upper() == "RS-232":
                winreg.SetValueEx(key, "Interface", 0, winreg.REG_SZ, "RS-232")
            elif mode.upper() == "RS-422":
                winreg.SetValueEx(key, "Interface", 0, winreg.REG_SZ, "RS-422")
            else:
                raise ValueError("Invalid mode selected. Choose 'RS232' or 'RS422'.")

            winreg.CloseKey(key)
            logger.info(f"Successfully set {com_port} to {mode} mode.")
            messagebox.showinfo("Success", f"{com_port} set to {mode} mode.")
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            logger.error(f"Failed to set {com_port} to {mode} mode. Error: {e}")
            messagebox.showerror("Error", f"Failed to set {com_port} to {mode} mode. Error: {e}")

    def get_com_ports(self):
        """Get a list of available COM ports for Moxa devices and display 'Port X connected to COMY'."""
        try:
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
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            logger.error(f"Error retrieving COM ports: {e}")
            messagebox.showerror("Error", f"Error retrieving COM ports: {e}")
            return []

    def on_connect(self):
        """Handle the connect button click event."""
        def update_serial_numbers_and_proceed():
            try:
                # Retrieve serial numbers from user input
                battery_1_serial = battery_1_entry.get().strip()
                battery_2_serial = battery_2_entry.get().strip()

                if not battery_1_serial or not battery_2_serial:
                    messagebox.showwarning("Warning", "Serial numbers cannot be empty!")
                    return

                # Update serial numbers
                update_serial_number(battery_1_serial, battery_2_serial)

                # Close the pop-up and show the appropriate battery info page
                popup.destroy()
                selected_mode = self.rs232_422_modes.get()
                connect_to_serial_port(selected_mode)

                if selected_mode == "RS-232":
                    self.main_window.show_rs_battery_info()
                elif selected_mode == "RS-422":
                    self.main_window.show_rs_battery_info()
            except Exception as e:
                error_details = traceback.format_exc()
                logger.error(f"Failed to update serial numbers: {e}\n{error_details}")
                messagebox.showerror("Error", f"Failed to update serial numbers: {e}\nCheck logs for details.")

        try:
            selected_mode = self.rs232_422_modes.get()
            set_active_protocol(selected_flag=selected_mode)

            if selected_mode in ["RS-232", "RS-422"]:
                # Create a pop-up dialog for updating serial numbers
                popup = tk.Toplevel(self.main_window.master)
                popup.title("Update Battery Serial Numbers")

                # Set dimensions of the pop-up
                popup_width = 400
                popup_height = 200

                # Get dimensions of the main window
                main_width = self.main_window.master.winfo_width()
                main_height = self.main_window.master.winfo_height()
                main_x = self.main_window.master.winfo_x()
                main_y = self.main_window.master.winfo_y()

                # Calculate the position to center the pop-up relative to the main window
                center_x = main_x + (main_width - popup_width) // 2
                center_y = main_y + (main_height - popup_height) // 2

                # Set the geometry of the pop-up to make it centered
                popup.geometry(f"{popup_width}x{popup_height}+{center_x}+{center_y}")
                popup.transient(self.main_window.master)
                popup.grab_set()  # Make the pop-up modal

                # Function to validate numeric input
                def validate_numeric_input(char):
                    return char.isdigit() or char == ""  # Allow only digits or empty input
                
                # Register the validation function
                validate_command = popup.register(validate_numeric_input)
                
                tk.Label(popup, text="Enter Serial Number for Battery 1:").pack(pady=5)
                battery_1_entry = tk.Entry(popup, width=30, validate="key", validatecommand=(validate_command, "%P"))
                battery_1_entry.pack(pady=5)
                
                tk.Label(popup, text="Enter Serial Number for Battery 2:").pack(pady=5)
                battery_2_entry = tk.Entry(popup, width=30, validate="key", validatecommand=(validate_command, "%P"))
                battery_2_entry.pack(pady=5)
                
                tk.Button(popup, text="Update and Continue", command=lambda:update_serial_numbers_and_proceed()).pack(pady=10)
            else:
                logger.info("Invalid Mode selected.")
                messagebox.showwarning("Selection Error", "Please select a valid Mode.")
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error during connection: {e}\n{error_details}")
            messagebox.showerror("Connection Error", f"Error during connection: {e}")

    def set_rs232_interface(self, com_port):
        """Set the interface to RS232 using Windows Registry."""
        try:
            # Open the registry key for the Moxa UPort settings
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Serial', 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, 'Interface', 0, winreg.REG_SZ, 'RS232')
            winreg.CloseKey(key)
            logger.info(f"Port {com_port} set to RS232 mode.")
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            messagebox.showerror("Error", f"Failed to set {com_port} to RS232 mode. Error: {e}")

    def set_rs422_interface(self, com_port):
        """Set the interface to RS422 using Windows Registry."""
        try:
            # Open the registry key for the Moxa UPort settings
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services\Serial', 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, 'Interface', 0, winreg.REG_SZ, 'RS422')
            winreg.CloseKey(key)
            logger.info(f"Port {com_port} set to RS422 mode.")
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error Failed to set Port: {e}\n{error_details}")
            messagebox.showerror("Error", f"Failed to set {com_port} to RS422 mode. Error: {e}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = RSConnection(root)
        app.pack()
        root.mainloop()
    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Error starting the application: {e}")
        messagebox.showerror("Startup Error", f"Error starting the application: {e}")
