import tkinter as tk
import customtkinter as ctk
from pcan_api.custom_pcan_methods import *
import helpers.config as config

class Settings(tk.Frame):

    def __init__(self, master, main_window=None):
        super().__init__(master)
        self.main_window = main_window
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

    def create_widgets(self):
        # Battery selection dropdown (positioned on the right side)
        self.battery_selection_label = tk.Label(self, text="Select Connection Type:")
        self.battery_dropdown = ctk.CTkComboBox(
            self,
            values=["CAN", "RS232", "RS422"],
        )
        self.battery_dropdown.set("CAN")

        # Place the battery selection label and dropdown to the right side of the control panel frame
        self.battery_selection_label.grid(row=0, column=0, padx=5, pady=10, sticky="e")  # Adjust column for right-side placement
        self.battery_dropdown.grid(row=0, column=1, padx=5, pady=10, sticky="e")  # Align dropdown to the right of the label
        # RS232/RS422 heading label
        tk.Label(self, text="CAN Settings", font=("Helvetica", 16, "bold")).grid(row=1, columnspan=2, pady=10)

        # Separator line
        separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
        separator.grid(row=2, columnspan=2, sticky="we", padx=20, pady=(0, 10))

        # Load saved configuration
        self.load_saved_config()

        # Label and dropdown for hardware type
        tk.Label(self, text="Hardware Type:").grid(row=3, column=0, padx=20, sticky=tk.W)
        self.cbbHwType = ctk.CTkComboBox(self, values=list(self.m_HWTYPES.keys()))
        self.cbbHwType.set(self.saved_hwtype if self.saved_hwtype else list(self.m_HWTYPES.keys())[0])  # Set default or saved value
        self.cbbHwType.grid(row=3, column=1, padx=20, pady=5)

        # Baudrate
        tk.Label(self, text="Baudrate:").grid(row=4, column=0, padx=20, sticky=tk.W)
        self.cbbBaudrates = ctk.CTkComboBox(self, values=list(self.m_BAUDRATES.keys()))
        self.cbbBaudrates.set(self.saved_baudrate if self.saved_baudrate else list(self.m_BAUDRATES.keys())[3])  # Set default or saved value
        self.cbbBaudrates.grid(row=4, column=1, padx=20, pady=5)

        # I/O Port
        tk.Label(self, text="I/O Port:").grid(row=5, column=0, padx=20, sticky=tk.W)
        self.cbbIoPort = ctk.CTkComboBox(self, values=list(self.m_IOPORTS.keys()))
        self.cbbIoPort.set(self.saved_ioport if self.saved_ioport else list(self.m_IOPORTS.keys())[0])  # Set default or saved value
        self.cbbIoPort.grid(row=5, column=1, padx=20, pady=5)

        # Interrupt
        tk.Label(self, text="Interrupt:").grid(row=6, column=0, padx=20, sticky=tk.W)
        self.cbbInterrupt = ctk.CTkComboBox(self, values=list(self.m_INTERRUPTS.keys()))
        self.cbbInterrupt.set(self.saved_interrupt if self.saved_interrupt else list(self.m_INTERRUPTS.keys())[0])  # Set default or saved value
        self.cbbInterrupt.grid(row=6, column=1, padx=20, pady=5)

        # Save button
        self.btnSave = ctk.CTkButton(self, text="Save", command=self.on_save, fg_color="green", hover_color='green')
        self.btnSave.grid(row=7, columnspan=2, pady=10)

        # Center the frame within the parent window
        self.grid_rowconfigure(8, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.pack(expand=True, fill='both')

    def load_saved_config(self):
        # Retrieve saved values from the config
        self.saved_hwtype = config.config_values['can_config'].get('hardware_type', None)
        self.saved_baudrate = config.config_values['can_config'].get('baudrate', None)
        self.saved_ioport = config.config_values['can_config'].get('input_output_port', None)
        self.saved_interrupt = config.config_values['can_config'].get('interrupt', None)

    def on_save(self):
        # Retrieve selected values and convert them to strings for JSON serialization
        selected_hwtype = self.cbbHwType.get()
        hwtype_value = str(self.m_HWTYPES[selected_hwtype])
        print(f"Selected Hardware Type: {selected_hwtype}, Value: {hwtype_value}")
        config.config_values['can_config']['hardware_type'] = selected_hwtype

        selected_baudrate = self.cbbBaudrates.get()
        baudrate_value = str(self.m_BAUDRATES[selected_baudrate])
        print(f"Selected Baudrate: {selected_baudrate}, Value: {baudrate_value}")
        config.config_values['can_config']['baudrate'] = selected_baudrate
        
        selected_ioport = self.cbbIoPort.get()
        ioport_value = str(self.m_IOPORTS[selected_ioport])
        print(f"Selected I/O Port: {selected_ioport}, Value: {ioport_value}")
        config.config_values['can_config']['input_output_port'] = selected_ioport
        
        selected_interrupt = self.cbbInterrupt.get()
        interrupt_value = str(self.m_INTERRUPTS[selected_interrupt])
        print(f"Selected Interrupt: {selected_interrupt}, Value: {interrupt_value}")
        config.config_values['can_config']['interrupt'] = selected_interrupt

        # Save the updated configuration
        config.save_config()
        
        # Optionally show a message to the user indicating that the config has been updated
        messagebox.showinfo("Success", "Connection settings saved successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = Settings(root)
    app.pack()

    root.mainloop()

