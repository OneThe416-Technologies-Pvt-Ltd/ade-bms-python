import tkinter as tk
import customtkinter as ctk
from PCAN_API.custom_pcan_methods import *

class CanConnection(tk.Frame):
    can_connected = False

    def __init__(self, parent):
        super().__init__(parent)

        # Initialize instance variables to store selected values
        self.selected_hwtype = None
        self.selected_baudrate = None
        self.selected_ioport = None
        self.selected_interrupt = None

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
            '0100': 0x100, '0120': 0x120, '0140': 0x140, '0200': 0x200, '0220': 0x220, '0240': 0x240, 
            '0260': 0x260, '0278': 0x278, '0280': 0x280, '02A0': 0x2A0, '02C0': 0x2C0, '02E0': 0x2E0, '02E8': 0x2E8,
            '02F8': 0x2F8, '0300': 0x300, '0320': 0x320, '0340': 0x340, '0360': 0x360, '0378': 0x378, '0380': 0x380, 
            '03BC': 0x3BC, '03E0': 0x3E0, '03E8': 0x3E8, '03F8': 0x3F8
        }

        self.m_INTERRUPTS = {
            '3': 3, '4': 4, '5': 5, '7': 7, '9': 9, '10': 10, '11': 11, '12': 12, '15': 15
        }

        # Create GUI elements
        self.create_widgets()

    def create_widgets(self):
        # Label and dropdown for hardware type
        tk.Label(self, text="Hardware Type:").grid(row=0, column=0, padx=20, sticky=tk.W)
        self.cbbHwType = ctk.CTkComboBox(self, values=list(self.m_HWTYPES.keys()))
        self.cbbHwType.grid(row=0, column=1, padx=20, pady=5)

        # Baudrate
        tk.Label(self, text="Baudrate:").grid(row=1, column=0, padx=20, sticky=tk.W)
        self.cbbBaudrates = ctk.CTkComboBox(self, values=list(self.m_BAUDRATES.keys()))
        self.cbbBaudrates.grid(row=1, column=1, padx=20, pady=5)

        # I/O Port
        tk.Label(self, text="I/O Port:").grid(row=2, column=0, padx=20, sticky=tk.W)
        self.cbbIoPort = ctk.CTkComboBox(self, values=list(self.m_IOPORTS.keys()))
        self.cbbIoPort.grid(row=2, column=1, padx=20, pady=5)

        # Interrupt
        tk.Label(self, text="Interrupt:").grid(row=3, column=0, padx=20, sticky=tk.W)
        self.cbbInterrupt = ctk.CTkComboBox(self, values=list(self.m_INTERRUPTS.keys()))
        self.cbbInterrupt.grid(row=3, column=1, padx=20, pady=5)

        # Connect button
        self.btnConnect = ctk.CTkButton(self, text="Connect", command=self.on_connect, fg_color="green", hover_color='green')
        self.btnConnect.grid(row=4, columnspan=2, pady=10)

        # Disconnect button (initially disabled)
        self.btnDisconnect = ctk.CTkButton(self, text="Disconnect", command=self.on_disconnect, fg_color="red", hover_color='red', state=tk.DISABLED)
        self.btnDisconnect.grid(row=5, columnspan=2, pady=10)

        # Center the frame within parent (can_window)
        self.grid_rowconfigure(6, weight=1)  # Ensure row 6 expands to center vertically
        self.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center horizontally
        self.pack(expand=True, fill='both') 

        self.update_widgets()

    def update_widgets(self):
        if CanConnection.can_connected:
            self.btnConnect.configure(state=tk.DISABLED)
            self.btnDisconnect.configure(state=tk.NORMAL)

            # If previously selected values exist, set them
            if self.selected_hwtype is not None:
                self.cbbHwType.set(self.get_key_from_value(self.m_HWTYPES, self.selected_hwtype))
            if self.selected_baudrate is not None:
                self.cbbBaudrates.set(self.get_key_from_value(self.m_BAUDRATES, self.selected_baudrate))
            if self.selected_ioport is not None:
                self.cbbIoPort.set(self.get_key_from_value(self.m_IOPORTS, self.selected_ioport))
            if self.selected_interrupt is not None:
                self.cbbInterrupt.set(self.get_key_from_value(self.m_INTERRUPTS, self.selected_interrupt))

            self.cbbHwType.configure(state=tk.DISABLED)
            self.cbbBaudrates.configure(state=tk.DISABLED)
            self.cbbIoPort.configure(state=tk.DISABLED)
            self.cbbInterrupt.configure(state=tk.DISABLED)
        else:
            self.btnConnect.configure(state=tk.NORMAL)
            self.btnDisconnect.configure(state=tk.DISABLED)
            self.cbbHwType.configure(state=tk.NORMAL)
            self.cbbBaudrates.configure(state=tk.NORMAL)
            self.cbbIoPort.configure(state=tk.NORMAL)
            self.cbbInterrupt.configure(state=tk.NORMAL)

    def get_key_from_value(self, dictionary, value):
        # Helper method to get the key from the value in a dictionary
        return next(key for key, val in dictionary.items() if val == value)

    def on_connect(self):
        # Retrieve selected values and store them in instance variables
        self.selected_hwtype = self.m_HWTYPES[self.cbbHwType.get()]
        self.selected_baudrate = self.m_BAUDRATES[self.cbbBaudrates.get()]
        self.selected_ioport = int(self.cbbIoPort.get(), 16)
        self.selected_interrupt = int(self.cbbInterrupt.get())

        # Initialize the CAN connection
        result = pcan_initialize(self.selected_baudrate, self.selected_hwtype, self.selected_ioport, self.selected_interrupt)
        if result:
            CanConnection.can_connected = True
            self.update_widgets()

    def on_disconnect(self):
        pcan_uninitialize()
        print("Disconnecting...")
        CanConnection.can_connected = False
        self.clear_selected_values()  # Clear the selected values when disconnecting
        self.update_widgets()

    def clear_selected_values(self):
        # Clear instance variables when disconnecting
        self.selected_hwtype = None
        self.selected_baudrate = None
        self.selected_ioport = None
        self.selected_interrupt = None

    def get_connection_status(self):
        return CanConnection.can_connected   

if __name__ == "__main__":
    root = tk.Tk()
    app = CanConnection(root)
    app.pack()

    root.mainloop()
