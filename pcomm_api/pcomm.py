import serial
import time
import struct
import threading

control=None
rs_232_flag=False
rs_422_flag=False

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
            'bus_1_voltage_before_diode': 10,
            'bus_2_voltage_before_diode': 10,
            'bus_1_voltage_after_diode': 10,
            'bus_2_voltage_after_diode': 10,
            'bus_1_current_sensor1': 10,
            'bus_2_current_sensor2': 10,
            'charger_input_current': 10,
            'charger_output_current': 10,
            'charger_output_voltage': 10,
            'charger_status': 1,
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
            'heater_pad_charger_relay_status': 1,
            'charger_status': 1,
            'voltage': 10,
            'eb_1_current': 10,
            'eb_2_current': 10,
            'charge_current': 10,
            'temperature': 10,
            'state_oc_charge': 10
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
    'cmd_byte_1' : 0x00,
    'cmd_byte_2' : 0x08,
    'cmd_byte_3' : 0x00,
}

def connect_to_serial_port(port_name, flag):
    global control, rs_232_flag, rs_422_flag
    """Connect to the serial port and configure it with default settings."""
    baud_rate = 19200 if flag == "RS-232" else 9600
    data_bits = serial.EIGHTBITS
    parity = serial.PARITY_NONE
    stop_bits = serial.STOPBITS_ONE

    if flag == "RS-232":
        rs_232_flag = True
        rs_422_flag = False
    elif flag == "RS-422":
        rs_232_flag = False
        rs_422_flag = True

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
            start_communication()  # Start the communication after successful connection
            return control
        else:
            print(f"Failed to open {port_name}.")
            return None
    except serial.SerialException as e:
        print(f"Error opening the serial port: {e}")
        return None

def periodic_rs232_send_read():
    """Send and read data using RS232 protocol in a loop."""
    global rs232_device_data
    next_write_time = time.perf_counter()  # High-precision timer
    interval = 0.16  # 160 milliseconds

    while rs_232_flag and control.is_open:
        current_time = time.perf_counter()
        
        if current_time >= next_write_time:
            send_rs232_data()
            read_rs232_data()
            next_write_time = current_time + interval

def periodic_rs422_send_read():
    """Send and read data using RS422 protocol in a loop."""
    global rs422_device_data
    next_write_time = time.perf_counter()  # High-precision timer
    interval = 0.16  # 160 milliseconds
    while rs_422_flag and control.is_open:
        current_time = time.perf_counter()
        
        if current_time >= next_write_time:
            send_rs422_data()
            read_rs422_data()
            next_write_time = current_time + interval

def send_rs232_data():
    """Send RS232 data."""
    if control.is_open:
        data = create_rs232_packet()
        print(f"{data} data send")
        control.write(data)
        print(f"Sent RS232 data: {data.hex()}")

def send_rs422_data():
    """Send RS422 data."""
    if control.is_open:
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
                        print(f"Checksum mismatch. Calculated: {calculated_checksum}, Received: {block[39]}. Discarding block.")
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
    # pmu_heartbeat = 0xFFFF
    # packet.extend(struct.pack('<H', pmu_heartbeat))
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
            return cell_voltage
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

    # Byte 27: Monitor IC Temperature (-40 to +125, subtract 40)
    rs232_device_data['ic_temp'] = data_hex[22] - 40

    # Bytes 28-29: Bus 1 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_data['bus_1_voltage_before_diode'] = data_hex[23] * 0.06

    # Bytes 30-31: Bus 2 Voltage (18-33.3V, Resolution 0.06V)
    rs232_device_data['bus_2_voltage_before_diode'] = data_hex[24] * 0.06

    # Bytes 32-33: Bus 1 Voltage after Diode (Resolution 0.06V)
    rs232_device_data['bus_1_voltage_after_diode'] = data_hex[25] * 0.06

    # Bytes 34-35: Bus 2 Voltage after Diode (Resolution 0.06V)
    rs232_device_data['bus_2_voltage_after_diode'] = data_hex[26] * 0.06

    # Bytes 36-37: Current / Bus 1 Sensor 1 (0-63.75A, Resolution 0.25A)
    rs232_device_data['bus_1_current_sensor1'] = data_hex[28] * 0.25

    # Bytes 38-39: Current / Bus 2 Sensor 2 (0-63.75A, Resolution 0.25A)
    rs232_device_data['bus_2_current_sensor2'] = data_hex[30] * 0.25

    # Byte 40-41: Charger Input Current (0-15A, Resolution 0.1A)
    rs232_device_data['charger_input_current'] = round(data_hex[31] * 0.1,1)
    
    # Bytes 32-33: Charger Output Current (Resolution 0.1A)
    rs232_device_data['charger_output_current'] = round(data_hex[32] * 0.1,2)

    # Bytes 34-35: Charger Output Voltage (Resolution 0.1V)
    rs232_device_data['charger_output_voltage'] = data_hex[33] * 0.06

    # # Byte 36: Charger Status
    # rs232_device_data['charger_status'] = data_hex[36]

    # # Byte 37: Charger Relay Status
    # rs232_device_data['charger_relay_status'] = data_hex[37]

    # # Byte 38: Constant Voltage Mode (1 byte)
    # rs232_device_data['constant_voltage_mode'] = data_hex[38]

    # # Byte 39: Constant Current Mode (1 byte)
    # rs232_device_data['constant_current_mode'] = data_hex[39]

    # # Byte 40: Input Under Voltage (1 byte)
    # rs232_device_data['input_under_voltage'] = data_hex[40]

    # # Byte 41: Output Over Current (1 byte)
    # rs232_device_data['output_over_current'] = data_hex[41]

    # # Byte 42: Bus1 Status (1 byte)
    # rs232_device_data['bus1_status'] = data_hex[42]

    # # Byte 43: Bus2 Status (1 byte)
    # rs232_device_data['bus2_status'] = data_hex[43]

    # # Byte 44: Heater Pad Status (1 byte)
    # rs232_device_data['heater_pad'] = data_hex[44]
          
    print(f"Updated RS232 device data: {rs232_device_data}")

def update_rs422_device_data(data_bytes):
    # Example of parsing each byte using the appropriate resolution or bitwise operations:
    
    # Byte 0: Header (Fixed)
    # Byte 1: EB-1 Relay Status (1-bit, 0: Off, 1: On)

    # Byte 5-6: Battery Voltage (2-byte, 0.25V resolution, multiply raw value by 0.25)
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
    rs422_device_data['charge_current'] = raw_charge_current * 0.25

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
    """Start the communication loop based on the protocol."""
    if rs_232_flag:
        # Start RS232 communication loop
        rs232_thread = threading.Thread(target=periodic_rs232_send_read)
        rs232_thread.daemon = True  # This will ensure the thread exits when the main program exits
        rs232_thread.start()
    elif rs_422_flag:
        # Start RS422 communication loop
        rs422_thread = threading.Thread(target=periodic_rs422_send_read)
        rs422_thread.daemon = True  # Ensures thread stops when the main program exits
        rs422_thread.start()

def stop_communication():
    """Stop the communication."""
    global rs_232_flag, rs_422_flag
    rs_232_flag = False
    rs_422_flag = False
    if control and control.is_open:
        control.close()

def get_active_protocol():
    global rs_422_flag, rs_232_flag
    if rs_232_flag :
        return "RS-232"
    elif rs_422_flag:
        return "RS-422"