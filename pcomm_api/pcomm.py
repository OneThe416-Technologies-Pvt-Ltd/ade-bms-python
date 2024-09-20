import serial
import threading
import asyncio
import struct
from openpyxl import Workbook, load_workbook
import datetime
from fpdf import FPDF
import os

control=None
rs_232_flag=False
rs_422_flag=False
rs232_interval_event = None
rs422_interval_event = None

rs232_device_data ={
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
    global control, rs_232_flag, rs_422_flag
    """Connect to the serial port and configure it with default settings."""
    baud_rate = 19200 if flag == "RS-232" else 9600
    data_bits = serial.EIGHTBITS
    parity = serial.PARITY_NONE
    stop_bits = serial.STOPBITS_ONE

    try:
        control = serial.Serial(
            port=port_name,
            baudrate=baud_rate,
            bytesize=data_bits,
            parity=parity,
            stopbits=stop_bits,
            timeout=1  # 1-second timeout
        )
        if control.is_open:
            print(f"Connected to {port_name} with baud rate {baud_rate}.")
            # Once connected, start periodic send/read loop
            start_communication()
            return control
        else:
            print(f"Failed to open {port_name}.")
            return None
    except serial.SerialException as e:
        print(f"Error opening the serial port: {e}")
        return None

# Method to send RS232 or RS422 data
def send_data():
    while control and control.is_open:
        if rs_232_flag:
            send_rs232_data()
        elif rs_422_flag:
            send_rs422_data()
        # control.flush()  # Flush the buffer
        threading.Event().wait(0.16)  # Wait for 160ms between each send


# Method to read RS232 or RS422 data
def read_data():
    while control and control.is_open:
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
    # if control.is_open:
    data = create_rs232_packet()
    print(f"{data} data send")
    control.write(data)
    print(f"Sent RS232 data: {data.hex()}")

def send_rs422_data():
    """Send RS422 data."""
    # if control.is_open:
    data = create_rs422_packet()
    control.write(data)
    print(f"Sent RS422 data: {data.hex()}")

def read_rs232_data():
    """Read 256 bytes of data, find the valid 64-byte block, and update RS232 data."""
    if control.is_open:
        received_data = control.read(256)  # Read 256 bytes of data
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
                        update_rs232_device_data(valid_block)
                        return  # Exit after finding and processing the valid block
            print("No valid RS232 data block found in the received 256 bytes.")
        else:
            print("Received data is not 256 bytes long. Discarding data.")

def read_rs422_data():
    """Read 20 bytes, validate the first 18-byte RS422 data block, and update the device data."""
    if control.is_open:
        received_data = control.read(40)  # Read 40 bytes of data
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

def update_rs232_device_data(data_hex):
    """Update rs232_device_data with the received bytes."""
    
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
    rs232_device_data['battery_id'] = data_hex[2]

    # Byte 3: Battery Serial Number (1-255)
    rs232_device_data['serial_number'] = data_hex[3]

    # Byte 4: Hardware Version (1-9)
    rs232_device_data['hw_version'] = data_hex[4]

    # Byte 5: Software Version (1-9)
    rs232_device_data['sw_version'] = data_hex[5]

    # Process each cell voltage (adding 2 to raw value and ensuring it's within range)
    # Bytes 6-7: Cell-1 Voltage (Resolution 0.01V)
    cell_1_voltage_raw = data_hex[6]
    rs232_device_data['cell_1_voltage'] = calculate_cell_voltage(cell_1_voltage_raw)
    
    # Bytes 8-9: Cell-2 Voltage (Resolution 0.01V)
    cell_2_voltage_raw = data_hex[7]
    rs232_device_data['cell_2_voltage'] = calculate_cell_voltage(cell_2_voltage_raw)
    
    # Bytes 10-11: Cell-3 Voltage (Resolution 0.01V)
    cell_3_voltage_raw = data_hex[8]
    rs232_device_data['cell_3_voltage'] = calculate_cell_voltage(cell_3_voltage_raw)
    
    # Bytes 12-13: Cell-4 Voltage (Resolution 0.01V)
    cell_4_voltage_raw = data_hex[9]
    rs232_device_data['cell_4_voltage'] = calculate_cell_voltage(cell_4_voltage_raw)
    
    # Bytes 14-15: Cell-5 Voltage (Resolution 0.01V)
    cell_5_voltage_raw = data_hex[10]
    rs232_device_data['cell_5_voltage'] = calculate_cell_voltage(cell_5_voltage_raw)
    
    # Bytes 16-17: Cell-6 Voltage (Resolution 0.01V)
    cell_6_voltage_raw = data_hex[11]
    rs232_device_data['cell_6_voltage'] = calculate_cell_voltage(cell_6_voltage_raw)
    
    # Bytes 18-19: Cell-7 Voltage (Resolution 0.01V)
    cell_7_voltage_raw = data_hex[12]
    rs232_device_data['cell_7_voltage'] = calculate_cell_voltage(cell_7_voltage_raw)

    # Byte 20: Cell-1 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_1_temp'] = data_hex[13] - 40

    # Byte 21: Cell-2 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_2_temp'] = data_hex[14] - 40

    # Byte 22: Cell-3 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_3_temp'] = data_hex[15] - 40

    # Byte 23: Cell-4 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_4_temp'] = data_hex[16] - 40

    # Byte 24: Cell-5 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_5_temp'] = data_hex[17] - 40

    # Byte 25: Cell-6 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_6_temp'] = data_hex[18] - 40

    # Byte 26: Cell-7 Temperature (-40 to +125, subtract 40)
    rs232_device_data['cell_7_temp'] = data_hex[19] - 40

    cell_status_byte = data_hex[20]  # Byte 21 in the data (index 20)
    binary_status = format(cell_status_byte, '08b')  # Convert to 8-bit binary string

    # Get the status of each cell from the corresponding bit
    rs232_device_data['cell_1_status'] = (cell_status_byte & 0x01)  # Bit 0: Cell 1 Status
    rs232_device_data['cell_2_status'] = (cell_status_byte & 0x02) >> 1  # Bit 1: Cell 2 Status
    rs232_device_data['cell_3_status'] = (cell_status_byte & 0x04) >> 2  # Bit 2: Cell 3 Status
    rs232_device_data['cell_4_status'] = (cell_status_byte & 0x08) >> 3  # Bit 3: Cell 4 Status
    rs232_device_data['cell_5_status'] = (cell_status_byte & 0x10) >> 4  # Bit 4: Cell 5 Status
    rs232_device_data['cell_6_status'] = (cell_status_byte & 0x20) >> 5  # Bit 5: Cell 6 Status
    rs232_device_data['cell_7_status'] = (cell_status_byte & 0x40) >> 6  # Bit 6: Cell 7 Status


    # Byte 27: Monitor IC Temperature (-40 to +125, subtract 40)
    rs232_device_data['ic_temp'] = data_hex[22] - 40

    # Bytes 28-29: Bus 1 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_data['bus_1_voltage_before_diode'] = round(data_hex[23] * 0.06,2)

    # Bytes 30-31: Bus 2 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_data['bus_2_voltage_before_diode'] = round(data_hex[24] * 0.06,2)

    # Bytes 32-33: Bus 1 Voltage after Diode (Resolution 0.06V)
    rs232_device_data['bus_1_voltage_after_diode'] = round(data_hex[25] * 0.06,2)

    # Bytes 34-35: Bus 2 Voltage after Diode (Resolution 0.06V)
    rs232_device_data['bus_2_voltage_after_diode'] = round(data_hex[26] * 0.06,2)

    # Bytes 36-37: Current / Bus 1 Sensor 1 (0-63.75A, Resolution 0.25A)
    rs232_device_data['bus_1_current_sensor1'] = round((data_hex[27]-128) * 0.25,2)

    # Bytes 38-39: Current / Bus 2 Sensor 2 (0-63.75A, Resolution 0.25A)
    rs232_device_data['bus_2_current_sensor1'] = round((data_hex[29]-128) * 0.25,2)

    # Byte 40-41: Charger Input Current (0-15A, Resolution 0.1A)
    rs232_device_data['charger_input_current'] = round(data_hex[31] * 0.1,1)
    
    # Bytes 32-33: Charger Output Current (Resolution 0.1A)
    rs232_device_data['charger_output_current'] = round(data_hex[32] * 0.1,2)

    # Bytes 34-35: Charger Output Voltage (Resolution 0.1V)
    rs232_device_data['charger_output_voltage'] = round(data_hex[33] * 0.06,2)

    charger_status_byte = data_hex[35]  # Byte 35 in the data

    # Get the charging on/off status (0th bit)
    rs232_device_data['charging_on_off_status'] = (charger_status_byte & 0x01)  # Bit 0: Charging On/Off Status

    # Byte 37: Charger Relay Status
    rs232_device_data['charger_relay_status'] = (charger_status_byte & 0x02) >> 1

    # Byte 38: Constant Voltage Mode (1 byte)
    rs232_device_data['constant_voltage_mode'] = (charger_status_byte & 0x04) >> 2

    # Byte 39: Constant Current Mode (1 byte)
    rs232_device_data['constant_current_mode'] = (charger_status_byte & 0x08) >> 3

    # Byte 40: Input Under Voltage (1 byte)
    rs232_device_data['input_under_voltage'] = (charger_status_byte & 0x10) >> 4

    # Byte 41: Output Over Current (1 byte)
    rs232_device_data['output_over_current'] = (charger_status_byte & 0x20) >> 5

    info_status_byte = data_hex[35]  # Byte 35 in the data

    # Byte 42: Bus1 Status (1 byte)
    rs232_device_data['bus1_status'] = (info_status_byte & 0x01)

    # Byte 43: Bus2 Status (1 byte)
    rs232_device_data['bus2_status'] = (info_status_byte & 0x02) >> 1

    # Byte 44: Heater Pad Status (1 byte)
    rs232_device_data['heater_pad'] = (info_status_byte & 0x04) >> 2
          
    print(f"Updated RS232 device data: {rs232_device_data}")

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

    if control and control.is_open:
        control.close()
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
    
# async def fetch_current(battery_no):
#     global charging_start_time, is_fetching_current
#     print(f"Starting to fetch current for battery {battery_no}...")
    
#     # Keep fetching the current value asynchronously
#     while is_fetching_current:
#         try:
#             # Call the pcan_write_read method asynchronously
#             print(f"Calling pcan_write_read for current on battery {battery_no}")
#             await asyncio.to_thread(pcan_write_read, 'current', battery_no)
            
#             # Get the current value based on the battery number
#             current_value = None
#             if battery_no == 1:
#                 current_value = device_data_battery_1['charging_current']
#             elif battery_no == 2:
#                 current_value = device_data_battery_2['charging_current']

#             # Log the current value retrieved
#             print(f"Retrieved current value for battery {battery_no}: {current_value}")
            
#             # If current is greater than 0, start tracking charging time
#             if current_value is not None and current_value > 0:
#                 print(f"Battery {battery_no} is charging with current {current_value}")
                
#                 # If this is the first time fetching, set the start time
#                 if charging_start_time is None:
#                     charging_start_time = datetime.datetime.now()
#                     print(f"Charging started for battery {battery_no} at {charging_start_time}")
            
#             # Stop if current drops to 0 or below, calculate the charging time
#             elif current_value is not None and current_value <= 0:
#                 if charging_start_time is not None:
#                     charging_end_time = datetime.datetime.now()
#                     charging_duration = charging_end_time - charging_start_time
#                     print(f"Charging stopped for battery {battery_no} at {charging_end_time}")
#                     print(f"Total charging duration: {charging_duration}")

#                     # Update charging duration in the CAN log
#                     update_charging_duration_in_can_log(
#                         device_data_battery_1['serial_number'] if battery_no == 1 else device_data_battery_2['serial_number'],
#                         charging_duration
#                     )
#                 else:
#                     print(f"No charging start time recorded for battery {battery_no}")

#                 # Reset the charging start time and stop the loop
#                 charging_start_time = None
#                 print(f"Stopping fetch for battery {battery_no}")
#                 is_fetching_current = False
#                 break

#         except Exception as e:
#             print(f"Error fetching current for battery {battery_no}: {e}")
        
#         # Delay before fetching again
#         print(f"Sleeping for 500ms before next fetch for battery {battery_no}")
#         await asyncio.sleep(0.5)  # Sleep for 500ms asynchronously before trying again


# def start_fetching_current(battery_no):
#     global is_fetching_current
#     # Ensure the asyncio event loop is running
#     start_asyncio_event_loop()
    
#     # Start fetching in the event loop
#     if not is_fetching_current:
#         is_fetching_current = True
#         asyncio.run_coroutine_threadsafe(fetch_current(battery_no), asyncio.get_event_loop())


# def stop_fetching_current():
#     # Stop fetching current
#     global charging_start_time,is_fetching_current   
#     device_data_battery_1['charging_current'] = 0  # Reset the start time
#     if charging_start_time is None:
#         is_fetching_current = False


# def update_charging_duration_in_can_log(serial_number, charging_duration):
#     """
#     Update the charging date and duration for a specific battery identified by its serial number
#     in the CAN Log Excel file (Can_Log.xlsx).
#     """
#     # Define the log file path for Can_Log.xlsx
#     folder_path = os.path.join(os.path.expanduser("~"), "Documents", "Battery_Logs")
#     file_path = os.path.join(folder_path, "Can_Log.xlsx")

#     # Load existing workbook or create a new one if the file does not exist
#     if os.path.exists(file_path):
#         workbook = load_workbook(file_path)
#         sheet = workbook.active
#     else:
#         print("Log file not found.")
#         return

#     # Find the row corresponding to the battery's serial number
#     serial_number_column = 8  # Assuming Serial Number is in the 8th column (H)
#     charging_duration_column = 12  # Charging Duration in the 12th column (L)
#     charging_date_column = 13  # Charging Date in the 13th column (M)

#     found = False

#     for row in range(2, sheet.max_row + 1):
#         if sheet.cell(row=row, column=serial_number_column).value == serial_number:
#             # Found the row, retrieve previous duration and add the new one
#             previous_duration = sheet.cell(row=row, column=charging_duration_column).value or 0
#             total_duration = previous_duration + (charging_duration.total_seconds() / 60)  # Convert to minutes
#             total_duration = round(total_duration, 2)

#             # Update the charging duration and date in the corresponding row
#             sheet.cell(row=row, column=charging_duration_column).value = total_duration
#             sheet.cell(row=row, column=charging_date_column).value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             found = True
#             break

#     if not found:
#         print(f"Serial number {serial_number} not found in the log.")

#     # Save the updated Excel file
#     workbook.save(file_path)
#     print(f"Updated charging duration and date for Serial Number {serial_number} in Can_Log.xlsx.")


# async def fetch_voltage(battery_no,load_value):
#     global discharging_start_time, is_fetching_voltage, ocv
#     if load_value == 50:
#         print(f"Starting to fetch current for battery {battery_no}...")
#         # Keep fetching the current value asynchronously
#         while is_fetching_voltage:
#             try:
#                 # Call the pcan_write_read method asynchronously (commented out, replace with actual call)
#                 print(f"Calling pcan_write_read for current on battery {battery_no}")
#                 await asyncio.to_thread(pcan_write_read, 'voltage', battery_no)

#                 # Get the current voltage based on the battery number
#                 voltage_value = None
#                 if battery_no == 1:
#                     ocv = device_data_battery_1['charging_voltage']
#                     voltage_value = device_data_battery_1['voltage']
#                 elif battery_no == 2:
#                     ocv = device_data_battery_2['charging_voltage']
#                     voltage_value = device_data_battery_2['voltage']

#                 # Log the current value retrieved
#                 print(f"Retrieved current value for battery {battery_no}: {voltage_value}")

#                 # Start tracking discharging process
#                 if voltage_value is not None and voltage_value > 0:
#                     print(f"Battery {battery_no} is discharging with voltage {voltage_value}")
#                     print(f"Discharging started for battery {battery_no} at {discharging_start_time}")
#                     print(f"Open-circuit voltage (OCV): {ocv}")
#                     voltage_value_0s = device_data_battery_1['voltage'] if battery_no == 1 else device_data_battery_2['voltage']
#                     # Capture voltage at 0 seconds
#                     print(f"Voltage at 0 seconds: {voltage_value}")

#                     # Wait for 15 seconds and capture voltage
#                     time.sleep(15)
#                     voltage_value_15s = device_data_battery_1['voltage'] if battery_no == 1 else device_data_battery_2['voltage']
#                     print(f"Voltage at 15 seconds: {voltage_value_15s}")

#                     # Wait for another 15 seconds (total 30 seconds) and capture voltage
#                     time.sleep(15)
#                     voltage_value_30s = device_data_battery_1['voltage'] if battery_no == 1 else device_data_battery_2['voltage']
#                     print(f"Voltage at 30 seconds: {voltage_value_30s}")
        
#                     # Update discharging duration in the CAN log
#                     update_discharging_duration_in_can_log(
#                         device_data_battery_1['serial_number'] if battery_no == 1 else device_data_battery_2['serial_number'],
#                         0, load_value, voltage_value_0s=voltage_value_0s,voltage_value_15s=voltage_value_15s,voltage_value_30s=voltage_value_30s
#                     )
#                     # Reset the discharging start time and stop the loop
#                     discharging_start_time = None
#                     print(f"Stopping fetch for battery {battery_no}")
#                     is_fetching_voltage = False

#             except Exception as e:
#                 print(f"Error fetching current for battery {battery_no}: {e}")
#                 is_fetching_voltage = False
#                 break

#             except Exception as e:
#                 print(f"Error fetching current for battery {battery_no}: {e}")
#     else:
#         print(f"Starting to fetch current for battery {battery_no}...")
#         # Keep fetching the current value asynchronously
#         while is_fetching_voltage:
#             try:
#                 # Call the pcan_write_read method asynchronously
#                 print(f"Calling pcan_write_read for current on battery {battery_no}")
#                 # await asyncio.to_thread(pcan_write_read, 'current', battery_no)
#                 # Get the current value based on the battery number
#                 voltage_value = None
#                 if battery_no == 1:
#                     ocv=device_data_battery_1['charging_voltage']
#                     voltage_value = device_data_battery_1['voltage']
#                 elif battery_no == 2:
#                     voltage_value = device_data_battery_2['voltage']
#                 # Log the current value retrieved
#                 print(f"Retrieved current value for battery {battery_no}: {voltage_value}")
#                 # If current is greater than 0, start tracking charging time
#                 if voltage_value is not None and voltage_value > 0:
#                     print(f"Battery {battery_no} is charging with voltage {voltage_value}")
#                     # If this is the first time fetching, set the start time
#                     if discharging_start_time is None:
#                         discharging_start_time = datetime.datetime.now()
#                         print(f"Charging started for battery {battery_no} at {discharging_start_time}")
#                 # Stop if current drops to 0 or below, calculate the charging time
#                 elif voltage_value is not None and voltage_value <= 21:
#                     if discharging_start_time is not None:
#                         discharging_end_time = datetime.datetime.now()
#                         discharging_duration = discharging_end_time - discharging_start_time
#                         print(f"Charging stopped for battery {battery_no} at {discharging_end_time}")
#                         print(f"Total charging duration: {discharging_duration}")
#                         # Update charging duration in the CAN log
#                         update_discharging_duration_in_can_log(
#                             device_data_battery_1['serial_number'] if battery_no == 1 else device_data_battery_2['serial_number'],
#                             discharging_duration,load_value
#                         )
#                     else:
#                         print(f"No charging start time recorded for battery {battery_no}")
#                     # Reset the charging start time and stop the loop
#                     discharging_start_time = None
#                     print(f"Stopping fetch for battery {battery_no}")
#                     is_fetching_voltage = False
#                     break
#             except Exception as e:
#                 print(f"Error fetching current for battery {battery_no}: {e}")
#             # Delay before fetching again
#             print(f"Sleeping for 500ms before next fetch for battery {battery_no}")
#             await asyncio.sleep(0.5)  # Sleep for 500ms asynchronously before trying again


# def start_fetching_voltage(battery_no,load_value):
#     global is_fetching_voltage
#     # Ensure the asyncio event loop is running
#     start_asyncio_event_loop()
    
#     # Start fetching in the event loop
#     if not is_fetching_voltage:
#         is_fetching_voltage = True
#         asyncio.run_coroutine_threadsafe(fetch_voltage(battery_no,load_value), asyncio.get_event_loop())


# def stop_fetching_voltage():
#     # Stop fetching current
#     global discharging_start_time,is_fetching_voltage   
#     device_data_battery_1['voltage'] = 0  # Reset the start time
#     if discharging_start_time is None:
#         is_fetching_voltage = False


# def update_discharging_duration_in_can_log(serial_number, discharging_duration,load_value,voltage_value_0s=0,voltage_value_15s=0,voltage_value_30s=0):
#     """
#     Update the charging date and duration for a specific battery identified by its serial number
#     in the CAN Log Excel file (Can_Log.xlsx).
#     """
#     # Define the log file path for Can_Log.xlsx
#     folder_path = os.path.join(os.path.expanduser("~"), "Documents", "Battery_Logs")
#     file_path = os.path.join(folder_path, "Can_Log.xlsx")

#     # Load existing workbook or create a new one if the file does not exist
#     if os.path.exists(file_path):
#         workbook = load_workbook(file_path)
#         sheet = workbook.active
#     else:
#         print("Log file not found.")
#         return

#     # Find the row corresponding to the battery's serial number
#     serial_number_column = 8  # Assuming Serial Number is in the 8th column (H)
#     ocv_column = 14
#     current_load_column = 15
#     discharging_duration_column = 16  # Charging Duration in the 12th column (L)
#     voltage_value_0s_column = 17
#     voltage_value_15s_column = 18
#     voltage_value_30s_column = 19
#     discharging_date_column = 20  # Charging Date in the 13th column (M)

#     found = False

#     for row in range(2, sheet.max_row + 1):
#         if sheet.cell(row=row, column=serial_number_column).value == serial_number:
#             # Found the row, retrieve previous duration and add the new one
#             previous_duration = sheet.cell(row=row, column=discharging_duration_column).value or 0
#             if discharging_duration == 0:
#                 total_duration = 0
#             else:
#                 total_duration = previous_duration + (discharging_duration.total_seconds() / 60)  # Convert to minutes
#                 total_duration = round(total_duration, 2)

#             # Update the charging duration and date in the corresponding row
#             sheet.cell(row=row, column=discharging_duration_column).value = total_duration
#             sheet.cell(row=row, column=discharging_date_column).value = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#             sheet.cell(row=row, column=ocv_column).value = ocv
#             sheet.cell(row=row, column=current_load_column).value = load_value
#             sheet.cell(row=row, column=voltage_value_0s_column).value = voltage_value_0s
#             sheet.cell(row=row, column=voltage_value_15s_column).value = voltage_value_15s
#             sheet.cell(row=row, column=voltage_value_30s_column).value = voltage_value_30s
#             found = True
#             break

#     if not found:
#         print(f"Serial number {serial_number} not found in the log.")

#     # Save the updated Excel file
#     workbook.save(file_path)
#     print(f"Updated charging duration and date for Serial Number {serial_number} in Can_Log.xlsx.")
