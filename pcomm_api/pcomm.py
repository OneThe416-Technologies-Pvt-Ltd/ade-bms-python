import serial
import threading
import struct
from openpyxl import Workbook, load_workbook
import datetime
import os
from tkinter import messagebox
import pandas as pd
import helpers.pdf_generator as pdf_generator  # Config helper

battery_1=None
battery_2=None
rs_232_flag=False
rs_422_flag=False
rs232_interval_event = None
rs422_interval_event = None

rs232_device_0_data ={
            'battery_id': 1,
            'serial_number': 1,
            'hw_version': 1,
            'sw_version': 1,
            'cell_1_voltage': 2,
            'cell_2_voltage': 2,
            'cell_3_voltage': 2,
            'cell_4_voltage': 2,
            'cell_5_voltage': 2,
            'cell_6_voltage': 2,
            'cell_7_voltage': 2,
            'cell_1_temp': 1,
            'cell_2_temp': 1,
            'cell_3_temp': 1,
            'cell_4_temp': 1,
            'cell_5_temp': 1,
            'cell_6_temp': 1,
            'cell_7_temp': 1,
            'ic_temp': 25,
            'bus_1_voltage_before_diode': 0,
            'bus_2_voltage_before_diode': 0,
            'bus_1_voltage_after_diode': 0,
            'bus_2_voltage_after_diode': 0,
            'bus_1_current_sensor1': 0,
            'bus_2_current_sensor1': 0,
            'charger_input_current': 0,
            'charger_output_current': 0,
            'charger_output_voltage': 0,
            'charging_on_off_status': 1,
            'charger_relay_status': 1,
            'constant_voltage_mode': 1,
            'constant_current_mode': 1,
            'input_under_voltage': 1,
            'output_over_current': 1,
            'bus1_status': 1,
            'bus2_status': 1,
            'heater_pad': 1,
}

rs422_device_data ={
            'eb_1_relay_status': 0,
            'eb_2_relay_status': 0,
            'heater_pad_charger_relay_status': 0,
            'charger_status': 0,
            'voltage': 0,
            'eb_1_current': 0,
            'eb_2_current': 0,
            'charge_current': 0,
            'temperature': 0,
            'channel_selected': 0
}
rs232_device_1_data ={
            'battery_id': 1,
            'serial_number': 1,
            'hw_version': 1,
            'sw_version': 1,
            'cell_1_voltage': 2,
            'cell_2_voltage': 2,
            'cell_3_voltage': 2,
            'cell_4_voltage': 2,
            'cell_5_voltage': 2,
            'cell_6_voltage': 2,
            'cell_7_voltage': 2,
            'cell_1_temp': 1,
            'cell_2_temp': 1,
            'cell_3_temp': 1,
            'cell_4_temp': 1,
            'cell_5_temp': 1,
            'cell_6_temp': 1,
            'cell_7_temp': 1,
            'ic_temp': 25,
            'bus_1_voltage_before_diode': 0,
            'bus_2_voltage_before_diode': 0,
            'bus_1_voltage_after_diode': 0,
            'bus_2_voltage_after_diode': 0,
            'bus_1_current_sensor1': 0,
            'bus_2_current_sensor1': 0,
            'charger_input_current': 0,
            'charger_output_current': 0,
            'charger_output_voltage': 0,
            'charging_on_off_status': 1,
            'charger_relay_status': 1,
            'constant_voltage_mode': 1,
            'constant_current_mode': 1,
            'input_under_voltage': 1,
            'output_over_current': 1,
            'bus1_status': 1,
            'bus2_status': 1,
            'heater_pad': 1,
}

rs422_device_1_data ={
            'eb_1_relay_status': 0,
            'eb_2_relay_status': 0,
            'heater_pad_charger_relay_status': 0,
            'charger_status': 0,
            'voltage': 0,
            'eb_1_current': 0,
            'eb_2_current': 0,
            'charge_current': 0,
            'temperature': 0,
            'channel_selected': 0
}

rs232_write = {
    'header' : 0xAA,
    'cmd_byte_1' : 0x00,
    'cmd_byte_2' : 0x00,
    'cmd_byte_3' : 0x00,
    'cmd_byte_4' : 0x00,
}

rs422_write = {
    'header' : 0x55,
    'cmd_byte_1' : 0xC0,
    'cmd_byte_2' : 0x08,
    'cmd_byte_3' : 0x00,
}

# Custom setInterval function using threading
def set_interval_async(func, interval):
    """Repeat a func every `interval` seconds using threading and an event to stop."""
    stopped = threading.Event()

    def wrapper():
        """Wrap the function to ensure it stops when the stop event is set."""
        while not stopped.is_set():
            func()  # Call the function synchronously
            stopped.wait(interval)  # Wait for the interval or stop event

    thread = threading.Thread(target=wrapper)
    thread.daemon = True  # Ensure it exits when the main program exits
    thread.start()

    return stopped


def connect_to_serial_port(port_name, flag):
    global battery_1, rs_232_flag, rs_422_flag
    """Connect to the serial port and configure it with default settings."""
    baud_rate = 19200 if flag == "RS-232" else 9600
    data_bits = serial.EIGHTBITS
    parity = serial.PARITY_NONE
    stop_bits = serial.STOPBITS_ONE

    try:
        battery_1 = serial.Serial(
            port=port_name,
            baudrate=baud_rate,
            bytesize=data_bits,
            parity=parity,
            stopbits=stop_bits,
            timeout=1  # 1-second timeout
        )
        if battery_1.is_open:
            print(f"Connected to {port_name} with baud rate {baud_rate}.")
            # Once connected, start periodic send/read loop
            start_communication()
            return battery_1
        else:
            print(f"Failed to open {port_name}.")
            return None
    except serial.SerialException as e:
        print(f"Error opening the serial port: {e}")
        return None

# Method to send RS232 or RS422 data
def send_data():
    while battery_1 and battery_1.is_open:
        if rs_232_flag:
            send_rs232_data()
        elif rs_422_flag:
            send_rs422_data()
        # battery_1.flush()  # Flush the buffer
        threading.Event().wait(0.16)  # Wait for 160ms between each send

# Method to read RS232 or RS422 data
def read_data():
    while battery_1 and battery_1.is_open:
        if rs_232_flag:
            read_rs232_data()
        elif rs_422_flag:
            read_rs422_data()
        threading.Event().wait(0.16)  # Adjust the read interval if necessary

# Separate thread for sending data
def start_sending():
    send_thread = threading.Thread(target=send_data, daemon=True)
    send_thread.start()

# Separate thread for reading data
def start_reading():
    read_thread = threading.Thread(target=read_data, daemon=True)
    read_thread.start()

def send_rs232_data():
    """Send RS232 data."""
    # if battery_1.is_open:
    data = create_rs232_packet()
    print(f"{data} data send")
    battery_1.write(data)
    print(f"Sent RS232 data: {data.hex()}")

def send_rs422_data():
    """Send RS422 data."""
    # if battery_1.is_open:
    data = create_rs422_packet()
    battery_1.write(data)
    print(f"Sent RS422 data: {data.hex()}")

def read_rs232_data():
    """Read 256 bytes of data, find the valid 64-byte block, and update RS232 data."""
    if battery_1.is_open:
        received_data = battery_1.read(256)  # Read 256 bytes of data
        if len(received_data) == 256:
            # Log the received data in hex format for debugging purposes
            received_hex = received_data.hex()
            print(f"RS232 Data received: {received_hex}")

            # Search for the start (0x66AA) and end (0x55BB) of the 64-byte data block
            for i in range(0, 192):  # Search up to byte 192 to have room for 64 bytes
                # Check if the current byte and the next one match 0x66AA
                if received_data[i] == 0x66 and received_data[i+1] == 0xAA:
                    # Check if the 62nd and 63rd bytes after this match 0x55BB
                    if received_data[i+62] == 0x55 and received_data[i+63] == 0xBB:
                        # Extract the 64-byte block
                        valid_block = received_data[i:i+64]
                        print(f"RS232 vaild Data received: {valid_block.hex()}")
                        print(f"Valid RS232 data block found from index {i} to {i+63}.")
                        # Pass the valid block to update the device data
                        update_rs232_device_0_data(valid_block)
                        log_rs_data(rs232_device_0_data)
                        return  # Exit after finding and processing the valid block
            print("No valid RS232 data block found in the received 256 bytes.")
        else:
            print("Received data is not 256 bytes long. Discarding data.")

def read_rs422_data():
    """Read 20 bytes, validate the first 18-byte RS422 data block, and update the device data."""
    if battery_1.is_open:
        received_data = battery_1.read(40)  # Read 40 bytes of data
        if len(received_data) == 40:
            print(f"RS422 Data received (40 bytes): {received_data.hex()}")

            # Search for 0x55 within the received data
            start_index = -1  # Initialize with an invalid index
            for i in range(len(received_data)):
                if received_data[i] == 0x55:
                    start_index = i
                    break

            if start_index != -1:
                print(f"Found 0x55 at index {start_index}. Proceeding with validation.")

                # Ensure we have enough data from this start index (need 40 bytes)
                if start_index + 17 <= len(received_data):
                    block = received_data[start_index:start_index + 19]

                    # Calculate checksum for the first 39 bytes
                    calculated_checksum = 0
                    for byte in block[0:17]:
                        calculated_checksum ^= byte  # XOR each byte

                    # Validate checksum against the 40th byte
                    if calculated_checksum == block[18]:
                        print(f"Valid RS422 data block found.{block.hex()}")
                        # Pass the valid 40-byte block to update the device data
                        update_rs422_device_data(block)
                    else:
                        print(f"Checksum mismatch. Calculated: {calculated_checksum}, Received: {block[18]}. Discarding block.")
                else:
                    print(f"Not enough data to process a full 40-byte block starting from index {start_index}.")
            else:
                print("0x55 not found in the received data. Discarding block.")
        else:
            print("Received data is not 40 bytes long. Discarding data.")

def create_rs232_packet():
    """Create an RS232 packet to send."""
    packet = [rs232_write['header'], rs232_write['cmd_byte_1'], rs232_write['cmd_byte_2'], rs232_write['cmd_byte_3'], rs232_write['cmd_byte_4']]
    checksum = calculate_checksum(packet[0:4])
    packet.append(checksum)
    packet.append(0x55)  # Footer
    # Convert the packet to a hex string, making sure all letters are uppercase
    hex_string = ''.join([f'{byte:02X}' for byte in packet])
    
    # Manually add the AA and footer in uppercase, then convert back to bytes
    final_packet = bytes.fromhex(hex_string)
    
    return final_packet

def create_rs422_packet():
    """Create an RS422 packet to send."""
    packet = [rs422_write['header'], rs422_write['cmd_byte_1'], rs422_write['cmd_byte_2']]
    pmu_heartbeat = 0x0000
    packet.extend(struct.pack('<H', pmu_heartbeat))
    checksum = calculate_checksum(packet)
    packet.append(checksum)
    return bytearray(packet)

def update_rs232_device_0_data(data_hex):
    """Update rs232_device_0_data with the received bytes."""
    
    # Helper function to calculate and validate cell voltage
    def calculate_cell_voltage(raw_data):
        # Raw data + 2 to adjust the value
        cell_voltage = raw_data * 0.01 + 2
        # Ensure the voltage is in the range of 2 to 4.5V
        if 2.0 <= cell_voltage <= 4.5:
            return round(cell_voltage,2)
        else:
            print(f"Warning: Cell voltage {cell_voltage}V out of range.")
            return None
    
    # Byte 2: Battery ID
    rs232_device_0_data['battery_id'] = data_hex[2]

    # Byte 3: Battery Serial Number (1-255)
    rs232_device_0_data['serial_number'] = data_hex[3]

    # Byte 4: Hardware Version (1-9)
    rs232_device_0_data['hw_version'] = data_hex[4]

    # Byte 5: Software Version (1-9)
    rs232_device_0_data['sw_version'] = data_hex[5]

    # Process each cell voltage (adding 2 to raw value and ensuring it's within range)
    # Bytes 6-7: Cell-1 Voltage (Resolution 0.01V)
    cell_1_voltage_raw = data_hex[6]
    rs232_device_0_data['cell_1_voltage'] = calculate_cell_voltage(cell_1_voltage_raw)
    
    # Bytes 8-9: Cell-2 Voltage (Resolution 0.01V)
    cell_2_voltage_raw = data_hex[7]
    rs232_device_0_data['cell_2_voltage'] = calculate_cell_voltage(cell_2_voltage_raw)
    
    # Bytes 10-11: Cell-3 Voltage (Resolution 0.01V)
    cell_3_voltage_raw = data_hex[8]
    rs232_device_0_data['cell_3_voltage'] = calculate_cell_voltage(cell_3_voltage_raw)
    
    # Bytes 12-13: Cell-4 Voltage (Resolution 0.01V)
    cell_4_voltage_raw = data_hex[9]
    rs232_device_0_data['cell_4_voltage'] = calculate_cell_voltage(cell_4_voltage_raw)
    
    # Bytes 14-15: Cell-5 Voltage (Resolution 0.01V)
    cell_5_voltage_raw = data_hex[10]
    rs232_device_0_data['cell_5_voltage'] = calculate_cell_voltage(cell_5_voltage_raw)
    
    # Bytes 16-17: Cell-6 Voltage (Resolution 0.01V)
    cell_6_voltage_raw = data_hex[11]
    rs232_device_0_data['cell_6_voltage'] = calculate_cell_voltage(cell_6_voltage_raw)
    
    # Bytes 18-19: Cell-7 Voltage (Resolution 0.01V)
    cell_7_voltage_raw = data_hex[12]
    rs232_device_0_data['cell_7_voltage'] = calculate_cell_voltage(cell_7_voltage_raw)

    # Byte 20: Cell-1 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_1_temp'] = data_hex[13] - 40

    # Byte 21: Cell-2 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_2_temp'] = data_hex[14] - 40

    # Byte 22: Cell-3 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_3_temp'] = data_hex[15] - 40

    # Byte 23: Cell-4 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_4_temp'] = data_hex[16] - 40

    # Byte 24: Cell-5 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_5_temp'] = data_hex[17] - 40

    # Byte 25: Cell-6 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_6_temp'] = data_hex[18] - 40

    # Byte 26: Cell-7 Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['cell_7_temp'] = data_hex[19] - 40

    cell_status_byte = data_hex[20]  # Byte 21 in the data (index 20)
    binary_status = format(cell_status_byte, '08b')  # Convert to 8-bit binary string

    # Get the status of each cell from the corresponding bit
    rs232_device_0_data['cell_1_status'] = (cell_status_byte & 0x01)  # Bit 0: Cell 1 Status
    rs232_device_0_data['cell_2_status'] = (cell_status_byte & 0x02) >> 1  # Bit 1: Cell 2 Status
    rs232_device_0_data['cell_3_status'] = (cell_status_byte & 0x04) >> 2  # Bit 2: Cell 3 Status
    rs232_device_0_data['cell_4_status'] = (cell_status_byte & 0x08) >> 3  # Bit 3: Cell 4 Status
    rs232_device_0_data['cell_5_status'] = (cell_status_byte & 0x10) >> 4  # Bit 4: Cell 5 Status
    rs232_device_0_data['cell_6_status'] = (cell_status_byte & 0x20) >> 5  # Bit 5: Cell 6 Status
    rs232_device_0_data['cell_7_status'] = (cell_status_byte & 0x40) >> 6  # Bit 6: Cell 7 Status


    # Byte 27: Monitor IC Temperature (-40 to +125, subtract 40)
    rs232_device_0_data['ic_temp'] = data_hex[22] - 40

    # Bytes 28-29: Bus 1 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_0_data['bus_1_voltage_before_diode'] = round(data_hex[23] * 0.06,2)

    # Bytes 30-31: Bus 2 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_0_data['bus_2_voltage_before_diode'] = round(data_hex[24] * 0.06,2)

    # Bytes 32-33: Bus 1 Voltage after Diode (Resolution 0.06V)
    rs232_device_0_data['bus_1_voltage_after_diode'] = round(data_hex[25] * 0.06,2)

    # Bytes 34-35: Bus 2 Voltage after Diode (Resolution 0.06V)
    rs232_device_0_data['bus_2_voltage_after_diode'] = round(data_hex[26] * 0.06,2)

    # Bytes 36-37: Current / Bus 1 Sensor 1 (0-63.75A, Resolution 0.25A)
    rs232_device_0_data['bus_1_current_sensor1'] = round((data_hex[27]-128) * 0.25,2)

    # Bytes 38-39: Current / Bus 2 Sensor 2 (0-63.75A, Resolution 0.25A)
    rs232_device_0_data['bus_2_current_sensor1'] = round((data_hex[29]-128) * 0.25,2)

    # Byte 40-41: Charger Input Current (0-15A, Resolution 0.1A)
    rs232_device_0_data['charger_input_current'] = round(data_hex[31] * 0.1,1)
    
    # Bytes 32-33: Charger Output Current (Resolution 0.1A)
    rs232_device_0_data['charger_output_current'] = round(data_hex[32] * 0.1,2)

    # Bytes 34-35: Charger Output Voltage (Resolution 0.1V)
    rs232_device_0_data['charger_output_voltage'] = round(data_hex[33] * 0.06,2)

    charger_status_byte = data_hex[35]  # Byte 35 in the data

    # Get the charging on/off status (0th bit)
    rs232_device_0_data['charging_on_off_status'] = (charger_status_byte & 0x01)  # Bit 0: Charging On/Off Status

    # Byte 37: Charger Relay Status
    rs232_device_0_data['charger_relay_status'] = (charger_status_byte & 0x02) >> 1

    # Byte 38: Constant Voltage Mode (1 byte)
    rs232_device_0_data['constant_voltage_mode'] = (charger_status_byte & 0x04) >> 2

    # Byte 39: Constant Current Mode (1 byte)
    rs232_device_0_data['constant_current_mode'] = (charger_status_byte & 0x08) >> 3

    # Byte 40: Input Under Voltage (1 byte)
    rs232_device_0_data['input_under_voltage'] = (charger_status_byte & 0x10) >> 4

    # Byte 41: Output Over Current (1 byte)
    rs232_device_0_data['output_over_current'] = (charger_status_byte & 0x20) >> 5

    info_status_byte = data_hex[35]  # Byte 35 in the data

    # Byte 42: Bus1 Status (1 byte)
    rs232_device_0_data['bus1_status'] = (info_status_byte & 0x01)

    # Byte 43: Bus2 Status (1 byte)
    rs232_device_0_data['bus2_status'] = (info_status_byte & 0x02) >> 1

    # Byte 44: Heater Pad Status (1 byte)
    rs232_device_0_data['heater_pad'] = (info_status_byte & 0x04) >> 2
          
    print(f"Updated RS232 device data: {rs232_device_0_data}")

def update_rs422_device_data(data_bytes):
    # Example of parsing each byte using the appropriate resolution or bitwise operations:

    raw_status_byte = data_bytes[1]  # First byte (Byte 1)

    # Convert the first byte to binary and extract individual bits using bitwise operations
    rs422_device_data['eb_1_relay_status'] = (raw_status_byte & 0x01)  # Bit 0: EB1 Relay Status
    rs422_device_data['eb_2_relay_status'] = (raw_status_byte & 0x02) >> 1  # Bit 1: EB2 Relay Status
    rs422_device_data['heater_pad_status'] = (raw_status_byte & 0x04) >> 2  # Bit 2: Heater Pad Status
    rs422_device_data['charger_output_relay_status'] = (raw_status_byte & 0x08) >> 3  # Bit 3: Charger Output Relay Status

    channel_status_byte = data_bytes[3]

    # Extract the packet count for channel 1 (Bit 3:0)
    channel_1_count = channel_status_byte & 0x0F  # Mask with 0x0F to get bits 3:0
    
    # Extract the packet count for channel 2 (Bit 7:4)
    channel_2_count = (channel_status_byte & 0xF0) >> 4  # Mask with 0xF0 and shift right by 4 to get bits 7:4

    # Determine if the channel 1 is healthy
    if 0 < channel_1_count < 15:
        channel_status = 1  # Healthy
    else:
        channel_status = 0  # Not Healthy
    
    # Determine if the channel 2 is healthy
    if 0 < channel_2_count < 15:
        channel_status = 2  # Healthy
    else:
        channel_status = 0  # Not Healthy

    rs422_device_data['channel_selected'] = channel_status

    raw_voltage = data_bytes[4]
    rs422_device_data['voltage'] = raw_voltage * 0.15

    # Byte 7-8: EB-1 Current (2-byte, 0.125A resolution, multiply raw value by 0.125)
    raw_eb1_current = data_bytes[5]
    rs422_device_data['eb_1_current'] = (raw_eb1_current * 0.5)-63.75

    # Byte 9-10: EB-2 Current (2-byte, 0.125A resolution, multiply raw value by 0.125)
    raw_eb2_current = data_bytes[6]
    rs422_device_data['eb_2_current'] = (raw_eb2_current * 0.5)-63.75

    # Byte 11-12: Charge Current (2-byte, 0.25A resolution, multiply raw value by 0.25)
    raw_charge_current = data_bytes[7]
    rs422_device_data['charge_current'] = raw_charge_current * 0.05

    # Byte 13-14: Battery Temperature (2-byte, 0.5°C resolution, multiply raw value by 0.5)
    raw_temperature = data_bytes[8]
    rs422_device_data['temperature'] = raw_temperature-40

    # Byte 15-16: State of Charge (2-byte, 0-100%, no resolution conversion needed)
    raw_state_of_charge = data_bytes[9]
    rs422_device_data['state_oc_charge'] = raw_state_of_charge

    print(f"Updated RS422 device data: {rs422_device_data}")

def calculate_checksum(data):
    """Calculate XOR checksum for a given byte array."""
    checksum = 0x00
    for byte in data:
        checksum ^= byte
    return checksum

def start_communication():
    start_sending()
    start_reading()

def stop_communication():
    """Stop communication."""
    global rs_232_flag, rs_422_flag, rs232_interval_event, rs422_interval_event

    # Stop RS232 or RS422 intervals if they are running
    if rs232_interval_event:
        rs232_interval_event.set()  # Stop RS232 interval
    if rs422_interval_event:
        rs422_interval_event.set()  # Stop RS422 interval

    rs_232_flag = False
    rs_422_flag = False

    if battery_1 and battery_1.is_open:
        battery_1.close()
        print("Serial connection closed.")

def get_active_protocol():
    global rs_422_flag, rs_232_flag
    if rs_232_flag:
        return "RS-232"
    elif rs_422_flag:
        return "RS-422"


def set_active_protocol(selected_flag):
    global rs_422_flag, rs_232_flag
    if selected_flag == "RS-232":
        rs_232_flag = True
    elif selected_flag == "RS-422":
        rs_422_flag = True




def log_rs_data(update_rs_data):
    """
    Log RS232 data to an Excel file. If the serial number exists, do nothing;
    if it doesn't exist, append a new row with default values.
    """
    try:
        # Define the path for the RS232 data Excel file inside the AppData directory
        folder_path = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database")
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Define the log file name
        file_name = "rs_data.xlsx"
        file_path = os.path.join(folder_path, file_name)

        # Create or load workbook
        if os.path.exists(file_path):
            try:
                workbook = load_workbook(file_path)
                sheet = workbook.active
            except Exception as e:
                print(f"Error loading the Excel file: {str(e)}")
                messagebox.showerror("Error", f"Failed to load the Excel file: {str(e)}. Recreating the file.")
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "RS232 Data"
                # Add headers for RS232 data
                headers = [
                    "SI No", "Date", "Time", "Project", "Device Name", "Manufacturer Name", 
                    "Serial Number", "Cycle Count", "Full Charge Capacity", "Charging Date", 
                    "OCV Before Charging", "Discharging Date", "OCV Before Discharging"
                ]
                sheet.append(headers)
        else:
            # File does not exist, so create it
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "RS232 Data"
            # Add headers for RS232 data
            headers = [
                "SI No", "Date", "Time", "Project", "Device Name", "Manufacturer Name", 
                "Serial Number", "Cycle Count", "Full Charge Capacity", "Charging Date", 
                "OCV Before Charging", "Discharging Date", "OCV Before Discharging"
            ]
            sheet.append(headers)

        # Log the current RS232 data
        current_datetime = datetime.datetime.now()
        date = current_datetime.strftime("%Y-%m-%d")
        time = current_datetime.strftime("%H:%M:%S")
        serial_number = update_rs_data.get('serial_number', 'N/A')

        # Check if the serial number exists in the file
        serial_number_column = 7  # Assuming Serial Number is in column G (7th column)
        serial_found = False

        for row in range(2, sheet.max_row + 1):  # Start at 2 to skip the header
            if sheet.cell(row=row, column=serial_number_column).value == serial_number:
                serial_found = True
                print(f"Serial number {serial_number} already exists. No action taken.")
                break

        if not serial_found:
            # If serial number is not found, add a new row with default values
            next_row = sheet.max_row + 1
            rs_data = [
                next_row - 1,  # SI No
                date,  # Date
                time,  # Time
                update_rs_data.get('project', ''),  # Project
                update_rs_data.get('device_name', 'BT-70939APH'),  # Device Name
                update_rs_data.get('manufacturer_name', 'Bren-Tronics'),  # Manufacturer Name
                serial_number,  # Serial Number
                int(update_rs_data.get('cycle_count', 0)),  # Ensure Cycle Count is an int
                103,  # Ensure Full Charge Capacity is a float
                update_rs_data.get('charging_date', date),  # Charging Date
                float(update_rs_data.get('ocv_before_charging', 0.0)),  # Ensure OCV Before Charging is a float
                update_rs_data.get('discharging_date', date),  # Discharging Date
                float(update_rs_data.get('ocv_before_discharging', 0.0))  # Ensure OCV Before Discharging is a float
            ]
            sheet.append(rs_data)

        # Save the Excel file
        workbook.save(file_path)
        workbook.close()
        print(f"RS232 data logged in {file_path}.")

    except Exception as e:
        print(f"An error occurred while updating the Excel file: {str(e)}")
        messagebox.showerror("Error", f"An error occurred while updating the Excel file: {str(e)}")



def get_latest_rs_data(serial_number):
    """
    Retrieves the most recent RS232 data for the specified serial number.
    """
    # Path to the RS232 data Excel file
    rs_data_file = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database", "rs_data.xlsx")

    if not os.path.exists(rs_data_file):
        messagebox.showerror("Error", "No RS232 data file found. Please log RS232 data first.")
        return None

    # Load the existing workbook and access the first sheet
    workbook = load_workbook(rs_data_file)
    sheet = workbook.active

    # Find the serial number column (assuming it's the 7th column, adjust if necessary)
    serial_number_column = 7  # Serial Number is the 7th column (G column)

    # Search for the serial number in the rows and store the data if found
    serial_found = False
    latest_data = {}

    for row in range(2, sheet.max_row + 1):  # Start from the second row to skip headers
        current_serial_number = sheet.cell(row=row, column=serial_number_column).value
        if current_serial_number == serial_number:
            serial_found = True
            print(f"Serial Number {serial_number} found in row {row}")

            # Retrieve the headers from the first row
            headers = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
            # Retrieve the values for the matching row
            values = [sheet.cell(row=row, column=col).value for col in range(1, sheet.max_column + 1)]
            
            # Map headers to values and print for debugging
            latest_data = dict(zip(headers, values))
            for header, value in latest_data.items():
                print(f"{header}: {value}")  # Debug print to ensure correct values are fetched
            break

    if not serial_found:
        messagebox.showwarning("Warning", f"Serial number {serial_number} not found in the RS232 data.")
        return None

    return latest_data



def update_cycle_count_in_rs_data(serial_number, new_cycle_count):
    """
    Update the cycle count for a specific device identified by its serial number
    in the RS232 Data Excel file.
    """
    # Define the folder and file path for the RS232 data file
    folder_path = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database")
    file_path = os.path.join(folder_path, "rs_data.xlsx")

    # Check if the file exists
    if not os.path.exists(file_path):
        messagebox.showerror("Error", "RS232 data file not found.")
        return

    # Open the existing Excel file
    workbook = load_workbook(file_path)
    sheet = workbook.active

    # Loop through the rows to find the correct row by Serial Number
    serial_number_column = 7  # Assuming Serial Number is in column G (7th column)
    cycle_count_column = 8    # Assuming Cycle Count is in column H (8th column)

    for row in range(2, sheet.max_row + 1):  # Start at 2 to skip the header
        if sheet.cell(row=row, column=serial_number_column).value == serial_number:
            # Update the Cycle Count value in the corresponding row
            sheet.cell(row=row, column=cycle_count_column).value = new_cycle_count
            break
    else:
        messagebox.showwarning("Warning", f"Serial number {serial_number} not found.")
        return

    # Save the updated Excel file
    workbook.save(file_path)
    messagebox.showinfo("Success", f"Cycle count updated successfully for Serial Number {serial_number}.")


def update_excel_and_download_pdf(data):
    """
    Get the updated values from the form, update the Excel file, and generate a PDF.
    """
    # Define the path to the RS232 data Excel file
    rs_data_file = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database", "rs_data.xlsx")

    # Load the existing data
    try:
        df_rs_data = pd.read_excel(rs_data_file)
    except FileNotFoundError:
        messagebox.showerror("Error", f"{rs_data_file} not found.")
        return
    
    # Assuming the serial number is unique, update the corresponding row
    serial_number = data[2]  # serial_number is the second element in the array
    df_rs_data['Serial Number'] = df_rs_data['Serial Number'].astype(str).str.strip()
    df_rs_data['OCV Before Charging'] = df_rs_data['OCV Before Charging'].astype(float)
    df_rs_data['OCV Before Discharging'] = df_rs_data['OCV Before Discharging'].astype(float)
    index = df_rs_data[df_rs_data['Serial Number'] == serial_number].index
    print(f"{df_rs_data['Serial Number']} test {serial_number}")
    print(f"{index} index")
    if not index.empty:
        # Update the values in the DataFrame
        df_rs_data.loc[index[0], 'Project'] = str(data[0])  # Ensure it's a string
        df_rs_data.loc[index[0], 'Device Name'] = str(data[1])  # Ensure it's a string
        df_rs_data.loc[index[0], 'Manufacturer Name'] = str(data[3])  # Ensure it's a string
        df_rs_data.loc[index[0], 'Serial Number'] = int(data[2])  # Ensure it's a string
        df_rs_data.loc[index[0], 'Cycle Count'] = int(data[4])  # Convert to int
        df_rs_data.loc[index[0], 'Full Charge Capacity'] = float(data[5])  # Convert to float
        df_rs_data.loc[index[0], 'Charging Date'] = str(data[6])  # Ensure it's a string
        df_rs_data.loc[index[0], 'OCV Before Charging'] = float(data[7])  # Convert to float
        df_rs_data.loc[index[0], 'Discharging Date'] = str(data[8])  # Ensure it's a string
        df_rs_data.loc[index[0], 'OCV Before Discharging'] = float(data[9])  # Convert to float

        # Save the updated DataFrame back to the Excel file
        df_rs_data.to_excel(rs_data_file, index=False)
        messagebox.showinfo("Success", "Data updated successfully.")
    else:
        messagebox.showwarning("Warning", "Serial Number not found in the Excel file.")

    # # Generate the PDF with the updated values
    pdf_generator.create_rs_report_pdf(serial_number,"RS232")
