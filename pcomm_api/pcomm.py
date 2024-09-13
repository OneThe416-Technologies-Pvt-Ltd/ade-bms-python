import serial
import time
import struct

control=None
rs_232_flag=False
rs_422_flag=False

def connect_to_serial_port(port_name):
    global control
    """Connect to the serial port and configure it with default settings."""
    baud_rate = 19200  # Default baud rate
    data_bits = serial.EIGHTBITS  # Default to 8 data bits
    parity = serial.PARITY_NONE  # Default to no parity
    stop_bits = serial.STOPBITS_ONE  # Default to 1 stop bit

    try:
        # Create a serial object with the port configuration
        control = serial.Serial(
            port=port_name,
            baudrate=baud_rate,
            bytesize=data_bits,
            parity=parity,
            stopbits=stop_bits,
            timeout=1  # 1 second timeout for reads/writes
        )

        if control.is_open:
            print(f"Successfully connected to {port_name} with baud rate {baud_rate}.")
            return control
        else:
            print(f"Failed to open {port_name}.")
            return None

    except serial.SerialException as e:
        print(f"Error opening the serial port: {e}")
        return None



def calculate_checksum(data):
    """Calculate XOR checksum for a given byte array."""
    checksum = 0x00
    for byte in data:
        checksum ^= byte
    return checksum

def create_rs422_packet():
    """Create a packet to send via RS422 (6 bytes)."""
    # Based on the image details for RS422 packet structure
    header = 0x55  # Fixed header
    battery_power_command = 0xFF  # Example: Battery power command byte
    av_status = 0x00  # AV status (bit-based values can be set here)
    
    # PMU heartbeat is a 16-bit frame counter, we set an example value here
    pmu_heartbeat = 0xFFFF

    # Construct the packet
    packet = [header, battery_power_command, av_status]
    packet.extend(struct.pack('<H', pmu_heartbeat))  # Pack PMU heartbeat (2 bytes)
    
    # Calculate checksum (XOR of bytes 0 to 4)
    checksum = calculate_checksum(packet)
    
    # Append checksum
    packet.append(checksum)

    return bytearray(packet)

def create_rs232_packet():
    """Create a packet to send via RS232 (7 bytes)."""
    # Based on the image details for RS232 packet structure
    header = 0xAA  # Fixed header
    cmd_byte_1 = 0x01  # Example: Bus 1 On/Off, Bus 2 On/Off, Charger Output Relay
    cmd_byte_2 = 0xFF  # Battery Parameters Reset Command
    cmd_byte_3 = 0xFF  # Example details to be defined later
    cmd_byte_4 = 0xFF  # Example details to be defined later
    
    # Construct the packet
    packet = [header, cmd_byte_1, cmd_byte_2, cmd_byte_3, cmd_byte_4]
    
    # Calculate checksum (XOR of bytes 1 to 4)
    checksum = calculate_checksum(packet[1:5])
    
    # Append checksum and footer
    packet.append(checksum)
    packet.append(0x55)  # Fixed footer

    return bytearray(packet)

def send_data():
    """Send data to the connected serial port based on the communication protocol."""
    global control
    if control and control.is_open:
        try:
            if rs_422_flag:
                # Create and send RS422 packet
                data = create_rs422_packet()
                control.write(data)  # Send the data
                print(f"RS422 Data sent: {data.hex()}")
            elif rs_232_flag:
                # Create and send RS232 packet
                data = create_rs232_packet()
                control.write(data)  # Send the data
                print(f"RS232 Data sent: {data.hex()}")
        except serial.SerialTimeoutException:
            print("Write timeout occurred.")
    else:
        print("Serial port is not open.")


def read_data():
    """Read data from the connected serial port and split it into variables."""
    global control
    if control and control.is_open:
        try:
            received_data = control.readline()  # Read a line of data from the serial port
            if received_data:
                received_data = received_data.strip()
                print(f"Data received: {received_data.decode()}")

                data_bytes = bytes.fromhex(received_data.decode())  # Convert hex string to bytes
                print(f"READ DATA{data_bytes}")
                if rs_232_flag:
                    # RS232 parsing logic for 64 bytes of data
                    if len(data_bytes) != 64:
                        print("Invalid RS232 data length.")
                        return

                    header = data_bytes[0:2]  # Header (2 bytes)
                    battery_id = data_bytes[2]  # Battery ID (1 byte)
                    battery_serial = data_bytes[3]  # Battery Sl No (1 byte)
                    hw_version = data_bytes[4]  # HW Version (1 byte)
                    sw_version = data_bytes[5]  # SW Version (1 byte)

                    # Cells voltages (each 2 bytes)
                    cell_1_voltage = struct.unpack('<H', data_bytes[6:8])[0] * 0.01  # Cell 1 voltage (resolution 0.01V)
                    # Continue for other cells...

                    # Charger Output Current and Voltage (from RS232 table)
                    charger_output_current = struct.unpack('<H', data_bytes[32:34])[0] * 0.1
                    charger_output_voltage = struct.unpack('<H', data_bytes[33:35])[0] * 0.06
                    present_charge_discharge_status = struct.unpack('<h', data_bytes[34:36])[0]

                    # Lifetime Data (Power On Time, Battery On Time, Lifetime Energy Discharged Data)
                    power_on_time = struct.unpack('<H', data_bytes[49:51])[0] * 5  # Power on time (Resolution 5 sec)
                    battery_on_time = struct.unpack('<I', data_bytes[51:55])[0] * 60  # Battery on time (Resolution 1 min)
                    lifetime_energy_discharge = struct.unpack('<I', data_bytes[55:59])[0]  # Lifetime Energy Discharged (1 AH @ nominal voltage)

                    # Fault codes (1 byte)
                    fault_code = data_bytes[59]

                    # Checksum and footer
                    checksum_received = struct.unpack('<H', data_bytes[61:63])[0]  # Ex-Or Checksum
                    footer = data_bytes[63:]

                    print(f"RS232 Data:")
                    print(f"Header: {header}")
                    print(f"Battery ID: {battery_id}")
                    print(f"HW Version: {hw_version}")
                    print(f"SW Version: {sw_version}")
                    print(f"Cell 1 Voltage: {cell_1_voltage} V")
                    print(f"Charger Output Current: {charger_output_current} A")
                    print(f"Charger Output Voltage: {charger_output_voltage} V")
                    print(f"Present Charge/Discharge Status: {present_charge_discharge_status} AH")
                    print(f"Power On Time: {power_on_time} sec")
                    print(f"Battery On Time: {battery_on_time} min")
                    print(f"Lifetime Energy Discharged: {lifetime_energy_discharge} AH")
                    print(f"Fault Code: {fault_code}")
                    print(f"Checksum: {checksum_received}")
                    print(f"Footer: {footer}")

                elif rs_422_flag:
                    # RS422 parsing logic for 18 bytes of data
                    if len(data_bytes) != 18:
                        print("Invalid RS422 data length.")
                        return

                    header = data_bytes[0]  # Header (1 byte, fixed at 0x55)
                    relay_status = data_bytes[1]  # Relay status (1 byte)
                    cb_status_interlock = data_bytes[2]  # CB Status and Interlock (1 byte)

                    # Relay status bits (extract individual bits)
                    relay_eb1_status = relay_status & 0x01  # Bit 0: Battery EB-1 relay status
                    relay_eb2_status = (relay_status >> 1) & 0x01  # Bit 1: Battery EB-2 relay status
                    heater_pad_status = (relay_status >> 2) & 0x01  # Bit 2: Heater/Charger relay status

                    # CB Status + Interlock (extract individual bits)
                    cb_eb1_status = cb_status_interlock & 0x01  # Bit 0: Battery EB-1 CB status
                    cb_eb2_status = (cb_status_interlock >> 1) & 0x01  # Bit 1: Battery EB-2 CB status
                    safety_interlock_status = (cb_status_interlock >> 4) & 0x01  # Bit 4: Safety interlock

                    # Battery voltage and current (each 2 bytes)
                    battery_voltage_28v = struct.unpack('<H', data_bytes[4:6])[0] * 0.15  # Battery voltage-28V (resolution 0.15V)
                    battery_current_eb1 = struct.unpack('<h', data_bytes[6:8])[0] * 0.5  # Battery current-28V EB1 (resolution 0.5 Amp)
                    battery_current_eb2 = struct.unpack('<h', data_bytes[8:10])[0] * 0.5  # Battery current-28V EB2 (resolution 0.5 Amp)
                    battery_charge_current = struct.unpack('<h', data_bytes[10:12])[0] * 0.05  # Battery charge current (resolution 0.05 Amp)
                    battery_temperature = struct.unpack('<h', data_bytes[12:14])[0]  # Battery temperature (-40 to 100°C)
                    battery_soc = data_bytes[14]  # Battery state of charge (0 to 255 -> 0% to 100%)

                    # Checksum
                    checksum_received = data_bytes[17]
                    checksum_calculated = 0x00
                    for i in range(17):
                        checksum_calculated ^= data_bytes[i]

                    print(f"RS422 Data:")
                    print(f"Relay EB1 Status: {'On' if relay_eb1_status else 'Off'}")
                    print(f"Battery Voltage: {battery_voltage_28v} V")
                    print(f"Battery Current EB1: {battery_current_eb1} A")
                    print(f"Battery Temperature: {battery_temperature} °C")
                    print(f"Battery State of Charge: {battery_soc} %")
                    print(f"Checksum: {'Valid' if checksum_received == checksum_calculated else 'Invalid'}")

            else:
                print("No data received.")
        except serial.SerialTimeoutException:
            print("Read timeout occurred.")
    else:
        print("Serial port is not open.")

def close_serial_port():
    """Close the serial port."""
    global control
    if control and control.is_open:
        control.close()
        print(f"Serial port {control.port} closed.")
        control = None


# Example Usage:
if __name__ == "__main__":
    # COM port settings
    port_name = "COM9"  # Change this to the correct port name (e.g., "COM3", "COM9", etc.)
    baud_rate = 19200
    data_bits = serial.EIGHTBITS
    stop_bits = serial.STOPBITS_ONE
    parity = serial.PARITY_NONE

    # Connect to the serial port
    ser = connect_to_serial_port(port_name, baud_rate, parity, data_bits, stop_bits)

    if ser:
        # Send some data
        send_data(ser, "Test Message")

        # Read response data (if applicable)
        read_data(ser)

        # Close the serial port
        close_serial_port(ser)
