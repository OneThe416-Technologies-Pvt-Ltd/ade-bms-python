import serial
import time
import threading
import os
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

class MoxaApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Moxa Serial Communication")
        self.serial_port = serial.Serial()
        self.rs232_flag = False
        self.rs422_flag = False

        self.buffer = bytearray(64)  # 64-byte buffer
        self.charger_status_display = ""
        self.status_information_display = ""

        # GUI setup (replace these with your actual design)
        self.port_combobox = Listbox(self.master, selectmode=SINGLE)
        self.port_combobox.pack()

        # Connect and Disconnect Buttons
        self.connect_button = Button(self.master, text="Connect", command=self.connect)
        self.connect_button.pack()

        self.disconnect_button = Button(self.master, text="Disconnect", command=self.disconnect)
        self.disconnect_button.pack()

        self.charger_on_button = Button(self.master, text="Charger ON/OFF", command=self.toggle_charger)
        self.charger_on_button.pack()

        self.rs232_button = Button(self.master, text="RS232 Mode", command=self.rs232_mode)
        self.rs232_button.pack()

        self.rs422_button = Button(self.master, text="RS422 Mode", command=self.rs422_mode)
        self.rs422_button.pack()

        # Populate available COM ports
        self.populate_com_ports()

    def populate_com_ports(self):
        ports = self.list_available_com_ports()
        for port in ports:
            self.port_combobox.insert(END, port)

    def list_available_com_ports(self):
        """List all available COM ports."""
        if os.name == 'nt':  # Windows
            available_ports = ['COM%s' % (i + 1) for i in range(256)]
        else:  # Unix-based systems
            available_ports = ['/dev/ttyS%s' % i for i in range(256)]
        return available_ports

    def connect(self):
        """Connect to the selected COM port."""
        selected_port = self.port_combobox.get(ACTIVE)
        try:
            if self.serial_port.is_open:
                self.serial_port.close()
            self.serial_port.port = selected_port
            self.serial_port.baudrate = 19200
            self.serial_port.parity = serial.PARITY_NONE
            self.serial_port.stopbits = serial.STOPBITS_ONE
            self.serial_port.bytesize = serial.EIGHTBITS
            self.serial_port.open()

            messagebox.showinfo("Info", f"Connected to {self.serial_port.port}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect: {e}")

    def disconnect(self):
        """Disconnect the serial port."""
        if self.serial_port.is_open:
            self.serial_port.close()
            messagebox.showinfo("Info", f"Disconnected from {self.serial_port.port}")

    def toggle_charger(self):
        """Toggle charger ON/OFF and send command."""
        if not self.serial_port.is_open:
            messagebox.showerror("Error", "Serial port not open.")
            return

        charger_on = True  # Replace this with actual charger status logic
        new_message_to_send = bytearray([0xAA, 0x00, 0x00, 0x00, 0x00, 0x00, 0x55])

        # Update message based on charger state
        if charger_on:
            new_message_to_send[1] = 0x01
            self.serial_port.write(new_message_to_send)
        else:
            new_message_to_send[1] = 0x00
            self.serial_port.write(new_message_to_send)

        charger_on = not charger_on

    def rs232_mode(self):
        """Switch to RS232 mode."""
        self.rs232_flag = True
        self.rs422_flag = False
        self.rs232_button.config(bg="yellowgreen")
        self.rs422_button.config(bg="red")

    def rs422_mode(self):
        """Switch to RS422 mode."""
        self.rs232_flag = False
        self.rs422_flag = True
        self.rs232_button.config(bg="red")
        self.rs422_button.config(bg="yellowgreen")

    def serial_data_received(self):
        """Handle data received from the serial port."""
        while self.serial_port.is_open:
            try:
                if self.rs232_flag:
                    data = self.serial_port.read(64)  # Read 64 bytes
                    self.process_rs232_data(data)
                elif self.rs422_flag:
                    data = self.serial_port.read(18)  # Read 18 bytes for RS422
                    self.process_rs422_data(data)
            except Exception as e:
                print(f"Error receiving data: {e}")

    def process_rs232_data(self, data):
        """Process RS232 data."""
        if len(data) == 64:
            pattern1 = 0x66
            pattern2 = 0xAA
            index1 = data.find(pattern1.to_bytes(1, 'big'))
            index2 = data.find(pattern2.to_bytes(1, 'big'))

            if index1 != -1 and index2 != -1 and index2 - index1 == 1:
                # Process relevant data here (use data[index1 + x] for values)
                print(f"Charger status: {data[index1 + 2]}")
                # Update TextBoxes or UI elements here

    def process_rs422_data(self, data):
        """Process RS422 data."""
        if len(data) == 18:
            print(f"RS422 data received: {data.hex()}")
            # Update TextBoxes or UI elements here

    def send_reset_message(self):
        """Send a reset message to the serial port."""
        if not self.serial_port.is_open:
            messagebox.showerror("Error", "Serial port not open.")
            return

        reset_message = bytearray([0xAA, 0x00, 0x23, 0x00, 0x00, 0x00, 0x55])
        reset_message[5] = self.xor_bytes(reset_message)
        self.serial_port.write(reset_message)

    def xor_bytes(self, byte_array):
        """XOR all bytes in the array (1-4) and return the result."""
        result = 0
        for i in range(1, 5):
            result ^= byte_array[i]
        return result


if __name__ == "__main__":
    root = Tk()
    app = MoxaApp(root)
    root.mainloop()
