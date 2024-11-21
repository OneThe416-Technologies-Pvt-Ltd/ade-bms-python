import tkinter as tk
import customtkinter as ctk
from pcan_api.custom_pcan_methods import *
import helpers.config as config
import serial.tools.list_ports
from tkinter import messagebox
from pcomm_api.pcomm import connect_to_serial_port,set_active_protocol
import re
from helpers.logger import logger  # Import the logger
import traceback

class Settings(tk.Frame):

    def __init__(self, master, main_window=None):
        super().__init__(master)
        try:
            self.main_window = main_window
            self.rs_config = config.config_values.get("rs_config")
            # Initialize all the dictionaries for dropdown options
            self.m_NonPnPHandles = {
                'PCAN_ISABUS1': PCAN_TYPE_ISA, 'PCAN_ISABUS2': PCAN_TYPE_ISA_SJA, 'PCAN_ISABUS3': PCAN_TYPE_ISA_PHYTEC,
                'PCAN_ISABUS4': PCAN_TYPE_DNG, 'PCAN_ISABUS5': PCAN_TYPE_DNG_EPP, 'PCAN_ISABUS6': PCAN_TYPE_DNG_SJA,
                'PCAN_ISABUS7': PCAN_TYPE_DNG_SJA_EPP
            }

            self.m_BAUDRATES = {
                '1 MBit/sec': PCAN_BAUD_1M, '800 kBit/sec': PCAN_BAUD_800K, '500 kBit/sec': PCAN_BAUD_500K,
                '250 kBit/sec': PCAN_BAUD_250K, '125 kBit/sec': PCAN_BAUD_125K, '100 kBit/sec': PCAN_BAUD_100K,
                '95,238 kBit/sec': PCAN_BAUD_95K, '83,333 kBit/sec': PCAN_BAUD_83K, '50 kBit/sec': PCAN_BAUD_50K,
                '47,619 kBit/sec': PCAN_BAUD_47K, '33,333 kBit/sec': PCAN_BAUD_33K, '20 kBit/sec': PCAN_BAUD_20K,
                '10 kBit/sec': PCAN_BAUD_10K, '5 kBit/sec': PCAN_BAUD_5K
            }

            self.m_HWTYPES = {
                'ISA-82C200': PCAN_TYPE_ISA, 'ISA-SJA1000': PCAN_TYPE_ISA_SJA, 'ISA-PHYTEC': PCAN_TYPE_ISA_PHYTEC,
                'DNG-82C200': PCAN_TYPE_DNG, 'DNG-82C200 EPP': PCAN_TYPE_DNG_EPP, 'DNG-SJA1000': PCAN_TYPE_DNG_SJA,
                'DNG-SJA1000 EPP': PCAN_TYPE_DNG_SJA_EPP
            }

            self.m_IOPORTS = {
                '0100':0x100, '0120':0x120, '0140':0x140, '0200':0x200, '0220':0x220, '0240':0x240, 
                '0260':0x260, '0278':0x278, '0280':0x280, '02A0':0x2A0, '02C0':0x2C0, '02E0':0x2E0, '02E8':0x2E8,
                '02F8':0x2F8, '0300':0x300, '0320':0x320, '0340':0x340, '0360':0x360, '0378':0x378, '0380':0x380, 
                '03BC':0x3BC, '03E0':0x3E0, '03E8':0x3E8, '03F8':0x3F8
            }

            self.m_INTERRUPTS = {
                '3':3, '4':4, '5':5, '7':7, '9':9, '10':10, '11':11, '12':12, '15':15
            }

            # Create GUI elements
            self.create_widgets()
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")


    def load_saved_config(self):
        # Retrieve saved values from the config
        try:
            self.saved_hwtype = config.config_values['can_config'].get('hardware_type', None)
            self.saved_baudrate = config.config_values['can_config'].get('baudrate', None)
            self.saved_ioport = config.config_values['can_config'].get('input_output_port', None)
            self.saved_interrupt = config.config_values['can_config'].get('interrupt', None)
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")

    def create_widgets(self):
        try:
            # Load saved configuration
            self.load_saved_config()
            # Define the StringVar and set up the dropdown
            self.connection_type_var = tk.StringVar(value="CAN")

            # Create label and dropdown
            self.connection_type_selection_label = tk.Label(self, text="Select Connection Type:")
            self.connection_type_dropdown = ctk.CTkComboBox(
                self,
                values=["CAN", "RS232", "RS422"],
                variable=self.connection_type_var  # Use StringVar for tracking changes
            )

            # Place the components in the grid
            self.connection_type_selection_label.grid(row=0, column=0, padx=5, pady=10, sticky="e")
            self.connection_type_dropdown.grid(row=0, column=1, padx=5, pady=10, sticky="e")

            # Trace changes in the dropdown selection
            self.connection_type_var.trace_add("write", self.update_ui_based_on_selection)
            self.update_ui_based_on_selection()

            # Center the frame within parent (can_window)
            self.grid_rowconfigure(7, weight=1)  # Ensure row 7 expands to center vertically
            self.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center horizontally
            self.pack(expand=True, fill='both')
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")
    
    def update_ui_based_on_selection(self, *args):
        try:
            logger.info("Dropdown selection changed to: %s", self.connection_type_dropdown.get())
            # Clear existing widgets
            for widget in self.grid_slaves():
                # Check if the widget is in the rows below the dropdown
                if int(widget.grid_info()['row']) >= 1:
                    widget.destroy()

            selected_connection = self.connection_type_dropdown.get()

            if selected_connection == "CAN":
                tk.Label(self, text="CAN Settings", font=("Helvetica", 16, "bold")).grid(row=1, columnspan=2, pady=10)

                # Separator line
                separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
                separator.grid(row=2, columnspan=2, sticky="we", padx=20, pady=(0, 10))

                # Label and dropdown for hardware type
                tk.Label(self, text="Hardware Type:").grid(row=3, column=0, padx=20, sticky=tk.W)
                self.cbbHwType = ctk.CTkComboBox(self, values=list(self.m_HWTYPES.keys()))
                self.cbbHwType.set(self.saved_hwtype if self.saved_hwtype else list(self.m_HWTYPES.keys())[0])
                self.cbbHwType.grid(row=3, column=1, padx=20, pady=5)

                # Baudrate
                tk.Label(self, text="Baudrate:").grid(row=4, column=0, padx=20, sticky=tk.W)
                self.cbbBaudrates = ctk.CTkComboBox(self, values=list(self.m_BAUDRATES.keys()))
                self.cbbBaudrates.set(self.saved_baudrate if self.saved_baudrate else list(self.m_BAUDRATES.keys())[0])
                self.cbbBaudrates.grid(row=4, column=1, padx=20, pady=5)

                # I/O Port
                tk.Label(self, text="I/O Port:").grid(row=5, column=0, padx=20, sticky=tk.W)
                self.cbbIoPort = ctk.CTkComboBox(self, values=list(self.m_IOPORTS.keys()))
                self.cbbIoPort.set(self.saved_ioport if self.saved_ioport else list(self.m_IOPORTS.keys())[0])
                self.cbbIoPort.grid(row=5, column=1, padx=20, pady=5)

                # Interrupt
                tk.Label(self, text="Interrupt:").grid(row=6, column=0, padx=20, sticky=tk.W)
                self.cbbInterrupt = ctk.CTkComboBox(self, values=list(self.m_INTERRUPTS.keys()))
                self.cbbInterrupt.set(self.saved_interrupt if self.saved_interrupt else list(self.m_INTERRUPTS.keys())[0])
                self.cbbInterrupt.grid(row=6, column=1, padx=20, pady=5)

                # Save button
                self.btnSave = ctk.CTkButton(self, text="Save", command=self.on_save, fg_color="green", hover_color='green')
                self.btnSave.grid(row=7, columnspan=2, pady=10)

            elif selected_connection == "RS232":
                # Add your RS232/RS422 settings here
                tk.Label(self, text="RS232 Settings", font=("Helvetica", 16, "bold")).grid(row=1, columnspan=2, pady=10)
                 # Separator line
                separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
                separator.grid(row=2, columnspan=2, sticky="we", padx=20, pady=(0, 10))

                tk.Label(self, text="Battery 1").grid(row=3, column=0, padx=20, sticky=tk.W)

                # Example for additional fields
                tk.Label(self, text="Select COM Port:").grid(row=4, column=0, padx=20, sticky=tk.W)
                # self.battery_1_com_232 = ctk.CTkComboBox(self, values=["COM 1", "COM 2", "COM 3", "COM 4", "COM 5", "COM 6"])
                self.battery_1_com_232 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_1_com_232.set("Select a COM Port")  # Set the default value shown in the dropdown
                self.battery_1_com_232.grid(row=4, column=1, padx=20, pady=5)


                tk.Label(self, text="Battery 2").grid(row=5, column=0, padx=20, sticky=tk.W)

                # Example for additional fields
                tk.Label(self, text="Select COM Port:").grid(row=6, column=0, padx=20, sticky=tk.W)
                # self.battery_2_com_232 = ctk.CTkComboBox(self, values=["COM 1", "COM 2", "COM 3", "COM 4", "COM 5", "COM 6"])
                self.battery_2_com_232 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_2_com_232.set("Select a COM Port")  # Set the default value shown in the dropdown
                self.battery_2_com_232.grid(row=6, column=1, padx=20, pady=5)

                # Save button
                self.btnSave = ctk.CTkButton(self, text="Save", command=self.on_save, fg_color="green", hover_color='green')
                self.btnSave.grid(row=7, columnspan=2, pady=10)
            elif selected_connection == "RS422":
                # Add your RS232/RS422 settings here
                tk.Label(self, text="RS422 Settings", font=("Helvetica", 16, "bold")).grid(row=1, columnspan=2, pady=10)
                 # Separator line
                separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
                separator.grid(row=2, columnspan=2, sticky="we", padx=20, pady=(0, 10))

                tk.Label(self, text="Battery 1").grid(row=3, column=0, padx=20, sticky=tk.W)

                # Example for additional fields
                tk.Label(self, text="Select RS422 Channel 1 Port:").grid(row=4, column=0, padx=20, sticky=tk.W)
                # self.battery_1_com_422_c_1 = ctk.CTkComboBox(self, values=["COM1", "COM2", "COM3"])
                self.battery_1_com_422_c_1 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_1_com_422_c_1.set("Select RS422 Channel 1 Port:")  # Set the default value shown in the dropdown
                self.battery_1_com_422_c_1.grid(row=4, column=1, padx=20, pady=5)

                # Example for additional fields
                tk.Label(self, text="Select RS422 Channel 2 Port:").grid(row=5, column=0, padx=20, sticky=tk.W)
                # self.battery_1_com_422_c_2 = ctk.CTkComboBox(self, values=["COM1", "COM2", "COM3"])
                self.battery_1_com_422_c_2 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_1_com_422_c_2.set("Select RS422 Channel 2 Port:")  # Set the default value shown in the dropdown
                self.battery_1_com_422_c_2.grid(row=5, column=1, padx=20, pady=5)

                tk.Label(self, text="Battery 2").grid(row=6, column=0, padx=20, sticky=tk.W)

                # Example for additional fields
                tk.Label(self, text="Select RS422 Channel 1 Port:").grid(row=7, column=0, padx=20, sticky=tk.W)
                # self.battery_2_com_422_c_1 = ctk.CTkComboBox(self, values=["COM1", "COM2", "COM3"])
                self.battery_2_com_422_c_1 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_2_com_422_c_1.set("Select RS422 Channel 1 Port:")  # Set the default value shown in the dropdown
                self.battery_2_com_422_c_1.grid(row=7, column=1, padx=20, pady=5)

                # Example for additional fields
                tk.Label(self, text="Select RS422 Channel 2 Port:").grid(row=8, column=0, padx=20, sticky=tk.W)
                # self.battery_2_com_422_c_2 = ctk.CTkComboBox(self, values=["COM1", "COM2", "COM3"])
                self.battery_2_com_422_c_2 = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
                self.battery_2_com_422_c_2.set("Select RS422 Channel 2 Port:")  # Set the default value shown in the dropdown
                self.battery_2_com_422_c_2.grid(row=8, column=1, padx=20, pady=5)

                # Save button
                self.btnSave = ctk.CTkButton(self, text="Save", command=self.on_save, fg_color="green", hover_color='green')
                self.btnSave.grid(row=9, columnspan=2, pady=10)
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")

    def get_com_ports(self):
        """Get a list of available COM ports for Moxa devices and display 'Port X connected to COMY'."""
        try:
            allocated_ports = set()
            # Loop through all key-value pairs in rs_config
            for key, value in self.rs_config.items():
                # Check if the value is a string and matches the COM port pattern (e.g., starts with "COM")
                if ("rs232" in key or "rs422" in key) and isinstance(value, str) and value.startswith("COM"):
                    allocated_ports.add(value)

            logger.info(f"allocated_ports: {allocated_ports}")
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
                        if com_port not in allocated_ports:
                            moxa_ports.append((port_number, f"Port {port_number} ({com_port})"))

            # Sort by the port number (the first element in the tuple)
            moxa_ports.sort(key=lambda x: x[0])
            logger.info(f"free_ports: {[port[1] for port in moxa_ports]}")
            # Return only the formatted string part
            return [port[1] for port in moxa_ports]
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")

    def on_save(self):
        try:
            selected_connection = self.connection_type_dropdown.get()
            if selected_connection == "CAN":
                # Retrieve selected values and convert them to strings for JSON serialization
                selected_hwtype = self.cbbHwType.get()
                hwtype_value = str(self.m_HWTYPES[selected_hwtype])
                logger.info(f"Selected Hardware Type: {selected_hwtype}, Value: {hwtype_value}")
                config.config_values['can_config']['hardware_type'] = selected_hwtype

                selected_baudrate = self.cbbBaudrates.get()
                baudrate_value = str(self.m_BAUDRATES[selected_baudrate])
                logger.info(f"Selected Baudrate: {selected_baudrate}, Value: {baudrate_value}")
                config.config_values['can_config']['baudrate'] = selected_baudrate

                selected_ioport = self.cbbIoPort.get()
                ioport_value = str(self.m_IOPORTS[selected_ioport])
                logger.info(f"Selected I/O Port: {selected_ioport}, Value: {ioport_value}")
                config.config_values['can_config']['input_output_port'] = selected_ioport

                selected_interrupt = self.cbbInterrupt.get()
                interrupt_value = str(self.m_INTERRUPTS[selected_interrupt])
                logger.info(f"Selected Interrupt: {selected_interrupt}, Value: {interrupt_value}")
                config.config_values['can_config']['interrupt'] = selected_interrupt
            elif selected_connection == "RS232":
                selected_battery_1_com_232 = self.battery_1_com_232.get()
                config.config_values['rs_config']['battery_1_rs232'] = selected_battery_1_com_232

                selected_battery_2_com_232 = self.battery_2_com_232.get()
                config.config_values['rs_config']['battery_2_rs232'] = selected_battery_2_com_232
            elif selected_connection == "RS422":
                selected_battery_1_com_422_c_1 = self.battery_1_com_422_c_1.get()
                config.config_values['rs_config']['battery_1_rs422_c_1'] = selected_battery_1_com_422_c_1

                selected_battery_1_com_422_c_2 = self.battery_1_com_422_c_2.get()
                config.config_values['rs_config']['battery_1_rs422_c_2'] = selected_battery_1_com_422_c_2

                selected_battery_2_com_422_c_1 = self.battery_2_com_422_c_1.get()
                config.config_values['rs_config']['battery_2_rs422_c_1'] = selected_battery_2_com_422_c_1

                selected_battery_2_com_422_c_2 = self.battery_2_com_422_c_2.get()
                config.config_values['rs_config']['battery_2_rs422_c_2'] = selected_battery_2_com_422_c_2
            # Save the updated configuration
            config.save_config()

            # Optionally show a message to the user indicating that the config has been updated
            messagebox.showinfo("Success", "Connection settings saved successfully.")
        except Exception as e:
            # Log any unexpected error in the settings function
            error_details = traceback.format_exc()
            logger.error(f"An unexpected error occurred: {e}\n{error_details}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}\nCheck logs for details.")

if __name__ == "__main__":
    root = tk.Tk()
    app = Settings(root)
    app.pack()

    root.mainloop()

