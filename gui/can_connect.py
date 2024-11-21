import tkinter as tk
import customtkinter as ctk
from pcan_api.custom_pcan_methods import *
import helpers.config as config
from helpers.logger import logger  # Import the logger
import traceback

class CanConnection(tk.Frame):
    can_connected = False  # Track CAN connection status

    def __init__(self, master, main_window=None):
        """
        Initialize the CAN Connection frame.
        
        Parameters:
        master (tk.Tk): The master window to which this frame is attached.
        main_window (tk.Tk): Reference to the main application window for managing transitions.
        """
        super().__init__(master)
        self.main_window = main_window
        
        try:
            self.create_widgets()  # Create GUI elements
            self.initialize_options()  # Initialize CAN options
            self.update_displayed_values()  # Update labels with saved values
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing CanConnection: {e}\n{error_details}")
            tk.messagebox.showerror("Initialization Error", f"Error initializing CanConnection: {e}")

    def initialize_options(self):
        """
        Initialize dictionaries for hardware types, baud rates, I/O ports, and interrupts.
        These dictionaries map user-friendly options to values used in the CAN API.
        """
        try:
            self.m_NonPnPHandles = {
                'PCAN_ISABUS1': PCAN_TYPE_ISA, 'PCAN_ISABUS2': PCAN_TYPE_ISA_SJA, 
                'PCAN_ISABUS3': PCAN_TYPE_ISA_PHYTEC, 'PCAN_ISABUS4': PCAN_TYPE_DNG, 
                'PCAN_ISABUS5': PCAN_TYPE_DNG_EPP, 'PCAN_ISABUS6': PCAN_TYPE_DNG_SJA, 
                'PCAN_ISABUS7': PCAN_TYPE_DNG_SJA_EPP
            }

            self.m_BAUDRATES = {
                '1 MBit/sec': PCAN_BAUD_1M, '800 kBit/sec': PCAN_BAUD_800K, 
                '500 kBit/sec': PCAN_BAUD_500K, '250 kBit/sec': PCAN_BAUD_250K,
                '125 kBit/sec': PCAN_BAUD_125K, '100 kBit/sec': PCAN_BAUD_100K,
                '95,238 kBit/sec': PCAN_BAUD_95K, '83,333 kBit/sec': PCAN_BAUD_83K,
                '50 kBit/sec': PCAN_BAUD_50K, '47,619 kBit/sec': PCAN_BAUD_47K,
                '33,333 kBit/sec': PCAN_BAUD_33K, '20 kBit/sec': PCAN_BAUD_20K,
                '10 kBit/sec': PCAN_BAUD_10K, '5 kBit/sec': PCAN_BAUD_5K
            }

            self.m_HWTYPES = {
                'ISA-82C200': PCAN_TYPE_ISA, 'ISA-SJA1000': PCAN_TYPE_ISA_SJA,
                'ISA-PHYTEC': PCAN_TYPE_ISA_PHYTEC, 'DNG-82C200': PCAN_TYPE_DNG,
                'DNG-82C200 EPP': PCAN_TYPE_DNG_EPP, 'DNG-SJA1000': PCAN_TYPE_DNG_SJA,
                'DNG-SJA1000 EPP': PCAN_TYPE_DNG_SJA_EPP
            }

            self.m_IOPORTS = {
                '0100': 0x100, '0120': 0x120, '0140': 0x140, '0200': 0x200,
                '0220': 0x220, '0240': 0x240, '0260': 0x260, '0278': 0x278,
                '0280': 0x280, '02A0': 0x2A0, '02C0': 0x2C0, '02E0': 0x2E0,
                '02E8': 0x2E8, '02F8': 0x2F8, '0300': 0x300, '0320': 0x320,
                '0340': 0x340, '0360': 0x360, '0378': 0x378, '0380': 0x380,
                '03BC': 0x3BC, '03E0': 0x3E0, '03E8': 0x3E8, '03F8': 0x3F8
            }

            self.m_INTERRUPTS = {
                '3': 3, '4': 4, '5': 5, '7': 7, '9': 9, '10': 10, 
                '11': 11, '12': 12, '15': 15
            }
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error initializing options: {e}\n{error_details}")
            tk.messagebox.showerror("Options Error", f"Error initializing CAN options: {e}")

    def load_saved_values(self):
        """Load saved CAN configuration values from the config file."""
        try:
            self.saved_hwtype = config.config_values['can_config'].get('hardware_type', 'ISA-82C200')
            self.saved_baudrate = config.config_values['can_config'].get('baudrate', '250 kBit/sec')
            self.saved_ioport = config.config_values['can_config'].get('input_output_port', '0100')
            self.saved_interrupt = config.config_values['can_config'].get('interrupt', '3')

        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error loading saved values: {e}\n{error_details}")
            tk.messagebox.showerror("Configuration Error", f"Error loading saved values: {e}")

    def create_widgets(self):
        """Create the widgets for the CAN connection interface."""
        try:
            tk.Label(self, text="CAN Connection", font=("Helvetica", 16, "bold")).grid(row=0, columnspan=2, pady=10)

            separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
            separator.grid(row=1, columnspan=2, sticky="we", padx=20, pady=(0, 10))

            self.hwtype_label = tk.Label(self, text="")
            self.baudrate_label = tk.Label(self, text="")
            self.ioport_label = tk.Label(self, text="")
            self.interrupt_label = tk.Label(self, text="")

            tk.Label(self, text="Hardware Type:").grid(row=2, column=0, padx=20, sticky=tk.W)
            self.hwtype_label.grid(row=2, column=1, padx=20, pady=5)

            tk.Label(self, text="Baudrate:").grid(row=3, column=0, padx=20, sticky=tk.W)
            self.baudrate_label.grid(row=3, column=1, padx=20, pady=5)

            tk.Label(self, text="I/O Port:").grid(row=4, column=0, padx=20, sticky=tk.W)
            self.ioport_label.grid(row=4, column=1, padx=20, pady=5)

            tk.Label(self, text="Interrupt:").grid(row=5, column=0, padx=20, sticky=tk.W)
            self.interrupt_label.grid(row=5, column=1, padx=20, pady=5)

            self.btnConnect = ctk.CTkButton(self, text="Connect", command=self.on_connect, fg_color="green", hover_color='green')
            self.btnConnect.grid(row=6, columnspan=2, pady=10)

            self.grid_rowconfigure(7, weight=1)
            self.grid_columnconfigure(0, weight=1)
            self.pack(expand=True, fill='both')
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error creating widgets: {e}\n{error_details}")
            tk.messagebox.showerror("Widget Error", f"Error creating widgets: {e}")

    def update_displayed_values(self):
        """Update labels to show the current configuration values."""
        try:
            self.load_saved_values()  # Load saved configuration values
            self.hwtype_label.config(text=self.saved_hwtype)
            self.baudrate_label.config(text=self.saved_baudrate)
            self.ioport_label.config(text=self.saved_ioport)
            self.interrupt_label.config(text=self.saved_interrupt)
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error updating displayed values: {e}\n{error_details}")
            tk.messagebox.showerror("Display Update Error", f"Error updating displayed values: {e}")

    def on_connect(self):
        """Handle the connect button click event."""
        try:
            selected_hwtype = self.m_HWTYPES.get(self.saved_hwtype)
            selected_baudrate = self.m_BAUDRATES.get(self.saved_baudrate)
            selected_ioport = int(self.saved_ioport, 16)
            selected_interrupt = int(self.saved_interrupt)

            logger.info(f"Connecting with: HWType={selected_hwtype}, Baudrate={selected_baudrate}, IOPort={selected_ioport}, Interrupt={selected_interrupt}")
            
            result = pcan_initialize(selected_baudrate, selected_hwtype, selected_ioport, selected_interrupt)
            
            if result:  # Assume `result` indicates success
                logger.info("CAN connection established successfully.")
                tk.messagebox.showinfo("Connection Status", "CAN connection established successfully!")
                self.update_displayed_values()

                if self.main_window:
                    self.main_window.show_can_battery_info()
            else:
                raise Exception("Failed to establish CAN connection. Please check the settings.")
        except Exception as e:
            error_details = traceback.format_exc()
            logger.error(f"Error during connection: {e}\n{error_details}")
            tk.messagebox.showerror("Connection Error", f"Error connecting to CAN: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CanConnection(root)
    app.pack()
    root.mainloop()
