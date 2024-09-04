#pcan_methods.py

from pcan_api.pcan import *
from tkinter import messagebox
import time
import asyncio

m_objPCANBasic = PCANBasic()
m_PcanHandle = 81

battery_status_flags = {
    "overcharged_alarm": 0,  # Bit 15
    "terminate_charge_alarm": 0,  # Bit 14
    "over_temperature_alarm": 0,  # Bit 12
    "terminate_discharge_alarm": 0,  # Bit 11
    "remaining_capacity_alarm": 0,  # Bit 9
    "remaining_time_alarm": 0,  # Bit 8
    "initialization": 0,  # Bit 7
    "charge_fet_test": 0,  # Bit 6
    "fully_charged": 0,  # Bit 5
    "fully_discharged": 0,  # Bit 4
    "error_codes": 0  # Bits 3:0
}

can_connected = False
rs_connected = False
continuous_update_thread = None
stop_continuous_update = False

name_mapping = {
            "device_name": "Device Name",
            "serial_number": "Serial No",
            "firmware_version": "Firmware Version",
            "manufacturer_date": "Manufacturer Date",
            "manufacturer_name": "Manufacturer Name",
            "battery_status": "Battery Status",
            "cycle_count": "Cycle Count",
            "manual_cycle_count": "Manual Cycle Count",
            "design_capacity": "Design Capacity",
            "design_voltage": "Design Voltage",
            "at_rate_ok_text": "At Rate OK",
            "at_rate_time_to_full": "At Rate Time To Full",
            "at_rate_time_to_empty": "At Rate Time To Empty",
            "at_rate": "At Rate",
            "rel_state_of_charge": "Rel State of Charge",
            "abs_state_of_charge": "Absolute State of Charge",
            "run_time_to_empty": "Run Time To Empty",
            "avg_time_to_empty": "Avg Time To Empty",
            "avg_time_to_full": "Avg Time To Full",
            "charging_battery_status": "Charging Status",
            "max_error": "Max Error",
            "temperature": "Temperature",
            "current": "Current",
            "remaining_capacity": "Remaining Capacity",
            "voltage": "Voltage",
            "avg_current": "Avg Current",
            "charging_current": "Charging Current",
            "full_charge_capacity": "Full Charge Capacity",
            "charging_voltage": "Charging Voltage"
        }

unit_mapping = {
    "device_name": "string",
    "firmware_version": "string",
    "serial_number": "number",
    "manufacturer_date": "unsigned int",
    "manufacturer_name": "string",
    "battery_status": "bit flags",
    "cycle_count": "Count",
    "manual_cycle_count": "Count",
    "design_capacity": "Ah",
    "design_voltage": "mV",
    "at_rate_ok_text": "Yes/No",
    "at_rate_time_to_full": "minutes",
    "at_rate_time_to_empty": "minutes",
    "at_rate":"A",
    "rel_state_of_charge": "Percent",
    "abs_state_of_charge": "Percent",
    "run_time_to_empty": "minutes",
    "avg_time_to_empty": "minutes",
    "charging_battery_status": "string",
    "avg_time_to_full": "minutes",
    "max_error": "Percent",
    "temperature": "°C",
    "current": "A",
    "remaining_capacity": "Ah",
    "voltage": "mV",
    "avg_current": "A",
    "charging_current": "A",
    "full_charge_capacity": "Ah",
    "charging_voltage": "mV"
}

device_data = {
            'device_name': "BT-70939APH",
            'serial_number': 0,
            'manufacturer_name': "Bren-Tronics",
            'firmware_version': "",
            'battery_status': "",
            'cycle_count': 0,
            'manual_cycle_count':0,
            'design_capacity': 0,
            'design_voltage': 0,
            'remaining_capacity': 0,
            'temperature': 0,
            'current': 20,
            'voltage': 0,
            'avg_current': 0,
            'charging_current': 0,
            'full_charge_capacity': 0,
            'charging_voltage': 0,
            'at_rate_time_to_full': 0,
            'at_rate_time_to_empty': 0,
            'at_rate_ok_text': "",
            'at_rate':0,
            'charging_battery_status':"Off",
            'rel_state_of_charge': 0,
            'abs_state_of_charge': 0,
            'run_time_to_empty': 0,
            'avg_time_to_empty': 0,
            'avg_time_to_full': 0,
            'max_error': 0
        } 


async def update_device_data():
    data_points = [
        ('serial_number', 'serial_number'),
        ('design_capacity', 'design_capacity'),
        ('design_voltage', 'design_voltage'),
        ('remaining_capacity', 'remaining_capacity'),
        ('temperature', 'temperature'),
        ('current', 'current'),
        ('voltage', 'voltage'),
        ('battery_status', 'battery_status'),
        ('cycle_count', 'cycle_count'),
        ('avg_current', 'avg_current'),
        ('charging_current', 'charging_current'),
        ('full_charge_capacity', 'full_charge_capacity'),
        ('charging_voltage', 'charging_voltage'),
        ('at_rate_time_to_full', 'at_rate_time_to_full'),
        ('at_rate_time_to_empty', 'at_rate_time_to_empty'),
        ('at_rate_ok_text', 'at_rate_ok_text'),
        ('at_rate', 'at_rate'),
        ('rel_state_of_charge', 'rel_state_of_charge'),
        ('abs_state_of_charge', 'abs_state_of_charge'),
        ('run_time_to_empty', 'run_time_to_empty'),
        ('avg_time_to_empty', 'avg_time_to_empty'),
        ('avg_time_to_full', 'avg_time_to_full'),
        ('max_error', 'max_error')
    ]

    # Create a list of tasks to be run concurrently
    tasks = []
    for call_name, key in data_points:
        task = asyncio.create_task(fetch_and_store_data(call_name, key))
        tasks.append(task)

    # Wait for all tasks to complete
    await asyncio.gather(*tasks)


async def fetch_and_store_data(call_name, key):
    await asyncio.to_thread(pcan_write_read, call_name)


async def process_data_points():
    data_points = [
        ('temperature', 'temperature'),
        ('manufacturer_name', 'manufacturer_name'),
        ('device_name', 'device_name'),
        ('current', 'current'),
        ('remaining_capacity', 'remaining_capacity'),
        ('voltage', 'voltage'),
        ('design_capacity', 'design_capacity'),
        ('design_voltage', 'design_voltage'),
        ('manufacturer_date', 'manufacturer_date'),
        ('serial_number', 'serial_number')
    ]

    tasks = []
    for call_name, key in data_points:
        task = asyncio.create_task(fetch_and_store_data(call_name, key))
        tasks.append(task)

    # Wait for all tasks to complete
    # await asyncio.gather(*tasks)

async def getdata():
    await process_data_points()

def pcan_initialize(baudrate, hwtype, ioport, interrupt):
    result = m_objPCANBasic.Initialize(m_PcanHandle, baudrate, hwtype, ioport, interrupt)
    if result != PCAN_ERROR_OK:
        if result == 5120:
            result = 512
        messagebox.showerror("Error!", GetFormatedError(result))
        return True
    else:
        pcan_write_read('serial_number')
        if device_data['serial_number'] != 0:
            pcan_write_read('temperature')
            pcan_write_read('firmware_version')       
            pcan_write_read('voltage')       
            pcan_write_read('cycle_count')       
            pcan_write_read('current')       
            pcan_write_read('charging_current')       
            pcan_write_read('remaining_capacity')       
            pcan_write_read('full_charge_capacity')
            messagebox.showinfo("Info!", "Connection established!")
            return True
        else:
           pcan_write_read('serial_number')
           if device_data['serial_number'] != 0:
               pcan_write_read('temperature')
               pcan_write_read('firmware_version')  
               pcan_write_read('charging_current')       
               pcan_write_read('voltage')   
               pcan_write_read('cycle_count')     
               pcan_write_read('current')       
               pcan_write_read('remaining_capacity')       
               pcan_write_read('full_charge_capacity')
               messagebox.showinfo("Info!", "Connection established!")
               return True
        messagebox.showinfo("Error!", "Connection Failed!")
        m_objPCANBasic.Uninitialize(m_PcanHandle)
        return False


#PCAN Uninitialize API Call
def pcan_uninitialize(): 
        result =  m_objPCANBasic.Uninitialize(m_PcanHandle)
        if result != PCAN_ERROR_OK:
            update_device_data_to_default()
            update_battery_status_flags_to_default()
            messagebox.showerror("Error!", GetFormatedError(result))
            return False
        else:
            # Reset device_data and battery_status_flags using the update methods
            update_device_data_to_default()
            update_battery_status_flags_to_default()
            messagebox.showinfo("Info!", "Connection Disconnect!")
            return True


#PCAN Write API Call
def pcan_write_read(call_name):
        CANMsg = TPCANMsg()
        CANMsg.ID = int('18EFC0D0',16)
        CANMsg.LEN = int(8)
        CANMsg.MSGTYPE = PCAN_MESSAGE_EXTENDED
        if call_name == 'firmware_version':
            CANMsg.DATA[0] = int('00',16)
        else:
            CANMsg.DATA[0] = int('01',16)
        CANMsg.DATA[1] = int('00',16)
        if call_name == 'serial_number':
            CANMsg.DATA[2] = int('1c',16)
        elif call_name == 'at_rate':
            CANMsg.DATA[2] = int('04',16)
        elif call_name == 'at_rate_time_to_full':
            CANMsg.DATA[2] = int('05',16)
        elif call_name == 'at_rate_time_to_empty':
            CANMsg.DATA[2] = int('06',16)
        elif call_name == 'at_rate_ok_text':
            CANMsg.DATA[2] = int('07',16)
        elif call_name == 'temperature':
            CANMsg.DATA[2] = int('08',16)
        elif call_name == 'voltage':
            CANMsg.DATA[2] = int('09',16)
        elif call_name == 'current':
            CANMsg.DATA[2] = int('0a',16)
        elif call_name == 'avg_current':
            CANMsg.DATA[2] = int('0b',16)
        elif call_name == 'max_error':
            CANMsg.DATA[2] = int('0c',16)
        elif call_name == 'rel_state_of_charge':
            CANMsg.DATA[2] = int('0d',16)
        elif call_name == 'abs_state_of_charge':
            CANMsg.DATA[2] = int('0e',16)
        elif call_name == 'remaining_capacity':
            CANMsg.DATA[2] = int('0f',16)
        elif call_name == 'full_charge_capacity':
            CANMsg.DATA[2] = int('10',16)
        elif call_name == 'run_time_to_empty':
            CANMsg.DATA[2] = int('11',16)
        elif call_name == 'avg_time_to_empty':
            CANMsg.DATA[2] = int('12',16)
        elif call_name == 'avg_time_to_full':
            CANMsg.DATA[2] = int('13',16)
        elif call_name == 'charging_current':
            CANMsg.DATA[2] = int('14',16)
        elif call_name == 'charging_voltage':
            CANMsg.DATA[2] = int('15',16)
        elif call_name == 'battery_status':
            CANMsg.DATA[2] = int('16',16)
        elif call_name == 'cycle_count':
            CANMsg.DATA[2] = int('17',16)
        elif call_name == 'design_capacity':
            CANMsg.DATA[2] = int('18',16)
        elif call_name == 'design_voltage':
            CANMsg.DATA[2] = int('19',16)
        elif call_name == 'manufacturer_date':
            CANMsg.DATA[2] = int('1b',16)
        elif call_name == 'manufacturer_name':
            CANMsg.DATA[2] = int('20',16)
        elif call_name == 'device_name':
            CANMsg.DATA[2] = int('21',16)
        elif call_name == 'firmware_version':
            CANMsg.DATA[2] = int('00',16)
        else:
            messagebox.showinfo("Error!", "Write operation not found!")   
        if call_name == 'manufacturer_name':
            CANMsg.DATA[3] = int('02',16)
        elif call_name == 'device_name':
            CANMsg.DATA[3] = int('02',16)
        else:
            CANMsg.DATA[3] = int('01',16)  
        CANMsg.DATA[4] = int('08',16)
        CANMsg.DATA[5] = int('02',16)
        CANMsg.DATA[6] = int('00',16)
        CANMsg.DATA[7] = int('00',16)
        result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
        print(f"{call_name}:{result}")
        if result == PCAN_ERROR_OK:
            messagebox.showerror(f"Error! {call_name}", GetFormatedError(result))
            return -2
        else:
            time.sleep(0.1)
            # result_code = retry_pcan_read(call_name)
            result_code = pcan_read(call_name)
            return result_code

def retry_pcan_read(call_name, retries=1, delay=0.1):
    for _ in range(retries):
        result_code = pcan_read(call_name)
        if result_code == PCAN_ERROR_OK:
            return result_code
        time.sleep(delay)
    return result_code


#PCAN Read API Call
def pcan_read(call_name):
    result = m_objPCANBasic.Read(m_PcanHandle)
    if result[0] != PCAN_ERROR_OK:
        print("battery_status_flags before UI update:", battery_status_flags)

        # Update the battery_status_flags with the new values
        battery_status_flags.update({
            "overcharged_alarm": 1,  # Bit 15
            "terminate_charge_alarm": 1,  # Bit 14
            "over_temperature_alarm": 1,  # Bit 12
            "terminate_discharge_alarm": 1,  # Bit 11
            "remaining_capacity_alarm": 1,  # Bit 9
            "remaining_time_alarm": 1,  # Bit 8
            "initialization": 1,  # Bit 7
            "charge_fet_test": 1,  # Bit 6
            "fully_charged": 1,  # Bit 5
            "fully_discharged": 0,  # Bit 4
            "error_codes": 1  # Bits 3:0, convert remaining bits to an integer
        })

        print("battery_status_flags after data update:", battery_status_flags)
        # messagebox.showerror(f"Error!{call_name}", GetFormatedError(result[0]))
        return result[0]
    else:
        args = result[1:]
        theMsg = args[0]
        newMsg = TPCANMsgFD()
        newMsg.ID = theMsg.ID
        newMsg.DLC = theMsg.LEN
        for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN):
            newMsg.DATA[i] = theMsg.DATA[i]
        

        resulthex = [hex(theMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]

        print(f"{call_name} ID: {hex(theMsg.ID)}")
        print(f"{call_name} Result Hex: {resulthex}")
        
        command_code = newMsg.DATA[4]
        first_byte = newMsg.DATA[0]
        second_byte = newMsg.DATA[1]

        # Swap the bytes to create the correct hex value
        swapped_hex = (second_byte << 8) | first_byte
        print(f"{call_name} Result Swapped Hex: {swapped_hex}")
        decimal_value = int(swapped_hex)
        print(f"{call_name} Result decimal_value: {decimal_value}")

        if newMsg.ID == 0x1CECFFC0 or newMsg.ID == 0x1CEBFFC0:
            if newMsg.ID == 0x1CECFFC0:
                hex_string1 = ''
                hex_string2 = ''
                result = m_objPCANBasic.Read(m_PcanHandle)
                if result[0] != PCAN_ERROR_OK:
                    return result[0]
                else:
                    args = result[1:]
                    theMsg = args[0]
                    byte_values1 = [hex(theMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]
                    hex_string1 = ''.join(format(int(h, 16), '02X') for h in byte_values1)
                    print(f"{call_name} Result hex_string1: {hex_string1}")

                result = m_objPCANBasic.Read(m_PcanHandle)
                if result[0] != PCAN_ERROR_OK:
                    return result[0]
                else:
                    m_objPCANBasic.Read(m_PcanHandle)
                    args = result[1:]
                    theMsg2 = args[0]
                    byte_values2 = [hex(theMsg2.DATA[i]) for i in range(8 if (theMsg2.LEN > 8) else theMsg2.LEN)]
                    hex_string2 = ''.join(format(int(h, 16), '02X') for h in byte_values2)
                    print(f"{call_name} Result hex_string2: {hex_string2}")

                # Convert hex strings to ASCII and filter out non-printable characters
                ascii_string1 = ''.join(chr(b) for b in bytes.fromhex(hex_string1) if 32 <= b <= 126)
                ascii_string2 = ''.join(chr(b) for b in bytes.fromhex(hex_string2) if 32 <= b <= 126)

                print(f"{call_name} Result ascii_string1: {ascii_string1}")
                print(f"{call_name} Result ascii_string2: {ascii_string2}")
                # Join the two ASCII strings
                final_string = (ascii_string1 + ascii_string2).rstrip()

                if call_name == 'manufacturer_name':
                    convert_data('manufacturer_name', final_string)
                    print(f"{call_name} Result converted value: {device_data[call_name]}")
                elif call_name == 'device_name':
                    convert_data('device_name', final_string)
                    print(f"{call_name} Result converted value: {device_data[call_name]}")
            elif newMsg.ID == 0x1CEBFFC0:
                hex_string1=''
                hex_string2=''
                result = m_objPCANBasic.Read(m_PcanHandle)
                if result[0] != PCAN_ERROR_OK:
                    return result[0]
                else:
                    args = result[1:]
                    theMsg1 = args[0]
                    byte_values1 = [hex(newMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]
                    hex_string1 = ''.join(format(int(h, 16), '02X') for h in byte_values1)
                    byte_values2 = [hex(theMsg1.DATA[i]) for i in range(8 if (theMsg1.LEN > 8) else theMsg1.LEN)]
                    hex_string2 = ''.join(format(int(h, 16), '02X') for h in byte_values2)
                    print(f"{call_name} Result hex_string1: {hex_string1}")
                    print(f"{call_name} Result hex_string2: {hex_string2}")
                ascii_string1 = ''.join(chr(b) for b in bytes.fromhex(hex_string1) if 32 <= b <= 126)
                ascii_string2 = ''.join(chr(b) for b in bytes.fromhex(hex_string2) if 32 <= b <= 126)
                print(f"{call_name} Result ascii_string1: {ascii_string1}")
                print(f"{call_name} Result ascii_string2: {ascii_string2}")
                # Join the two ASCII strings
                final_string = ascii_string1 + ascii_string2
                if call_name == 'manufacturer_name':
                    convert_data('manufacturer_name', final_string)
                elif call_name == 'device_name':
                    convert_data('device_name', final_string)
        elif call_name == 'firmware_version':
            data_packet = [int(hex(newMsg.DATA[i]), 16) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]

            # Extract version information
            major_version = data_packet[0]
            minor_version = data_packet[1]
            patch_number = (data_packet[3] << 8) | data_packet[2]
            build_number = (data_packet[5] << 8) | data_packet[4]

            version_string = f"{major_version}.{minor_version}.{patch_number}.{build_number}"
            convert_data('firmware_version', version_string)
            print(f"{call_name} Result converted value: {device_data[call_name]}")
        else:
            convert_data(command_code, decimal_value)
            print(f"{call_name} Result converted value: {device_data[call_name]}")

        print(f"Call Name : {call_name}")
        return result[0]


def convert_data(command_code, decimal_value):
    global battery_status_flags
    # Conversion rules
    if command_code == 0x04:  # AtRate: mA / 40 unsigned
        device_data['at_rate'] = (decimal_value*40)/1000
    elif command_code == 0x05:  # AtRateTimeToFull: minutes unsigned
        device_data['at_rate_time_to_full'] = round((decimal_value / 1000),1)
    elif command_code == 0x06:  # AtRateTimeToEmpty: minutes unsigned
        device_data['at_rate_time_to_empty'] = round((decimal_value / 1000),1)
    elif command_code == 0x07:  # AtRateOK: Boolean
        device_data['at_rate_ok_text'] = "Yes" if decimal_value != 0 else "No" 
    elif command_code == 0x08: # Temperature: Boolean
        temperature_k = decimal_value / 10.0
        temperature_c = temperature_k - 273.15
        device_data['temperature'] = round(temperature_c,1)
    elif command_code == 0x09:  # Voltage: mV unsigned
        device_data['voltage'] = round((decimal_value / 1000),1)
    elif command_code == 0x0a:  # Current: mA / 40 signed
        if decimal_value == 0:
            device_data['current'] = decimal_value
            device_data['charging_battery_status'] = "Idle"
        else:
            if decimal_value > 32767:
                decimal_value -= 65536
            currentmA = decimal_value*40
            currentA = currentmA / 1000
            if currentA > 0:
                device_data['charging_battery_status'] = "Charging"
            elif currentA < 0:
                device_data['charging_battery_status'] = "Discharging"
            else:
                device_data['charging_battery_status'] = "Idle"
            device_data['current'] = abs(currentA)
    elif command_code == 0x0b:  # Avg Current: mA / 40 signed
        if decimal_value == 0:
            device_data['avg_current'] = decimal_value
        else:
            if decimal_value > 32767:
                decimal_value -= 65536
            currentmA = decimal_value*40
            currentA = currentmA / 1000
            device_data['avg_current'] = currentA
    elif command_code == 0x0c:  # MaxError: Percent unsigned
        device_data['max_error'] = decimal_value
    elif command_code == 0x0d:  # RelStateofCharge: Percent unsigned
        device_data['rel_state_of_charge'] = decimal_value
    elif command_code == 0x0e:  # AbsoluteStateofCharge: Percent unsigned
        device_data['abs_state_of_charge'] = decimal_value
    elif command_code == 0x0f:  # RemainingCapacity: mAh / 40 unsigned
        device_data['remaining_capacity'] = (decimal_value*40)/1000
    elif command_code == 0x10:  # FullChargeCapacity: mAh / 40 unsigned
        device_data['full_charge_capacity'] = (decimal_value*40)/1000
    elif command_code == 0x11:  # RunTimeToEmpty: minutes unsigned
        device_data['run_time_to_empty'] = round((decimal_value / 10),1)
    elif command_code == 0x12:  # AvgTimeToEmpty: minutes unsigned
        device_data['avg_time_to_empty'] = round((decimal_value / 10),1)
    elif command_code == 0x13:  # AvgTimeToFull: minutes unsigned
        device_data['avg_time_to_full'] = round((decimal_value / 1000),1)
    elif command_code == 0x14:  # ChargingCurrent: mA / 40 unsigned
        print(f'charging:{decimal_value}')
        if decimal_value == 0:
            device_data['charging_current'] = decimal_value
        else:
            if decimal_value > 32767:
                decimal_value -= 65536
            currentmA = decimal_value*40
            print(f'charging currentmA:{currentmA}')
            currentA = currentmA / 1000
            print(f'charging currentA:{currentA}')
            device_data['charging_current'] = currentA
    elif command_code == 0x15:  # ChargingVoltage: mV unsigned
        device_data['charging_voltage'] = round((decimal_value / 1000),1)
    elif command_code == 0x16:  # BatteryStatus: bit flags unsigned
        binary_value = format(decimal_value, '016b')  # Convert to 16-bit binary string
        
        # Swap the bytes before interpreting the bits
        swapped_value = binary_value[8:] + binary_value[:8]
        device_data['battery_status'] = swapped_value
        
        # Update the battery_status_flags dictionary with the interpreted flags
        battery_status_flags = {
            "Overcharged Alarm": int(swapped_value[0]),  # Bit 15
            "Terminate Charge Alarm": int(swapped_value[1]),  # Bit 14
            "over_temperature_alarm": int(swapped_value[3]),  # Bit 12
            "Terminate Discharge Alarm": int(swapped_value[4]),  # Bit 11
            "Remaining Capacity Alarm": int(swapped_value[6]),  # Bit 9
            "Remaining Time Alarm": int(swapped_value[7]),  # Bit 8
            "Initialization": int(swapped_value[8]),  # Bit 7
            "Charge FET Test": int(swapped_value[9]),  # Bit 6
            "Fully Charged": int(swapped_value[10]),  # Bit 5
            "Fully Discharged": int(swapped_value[11]),  # Bit 4
            "Error Codes": int(swapped_value[12:], 2)  # Bits 3:0, convert remaining bits to an integer
        }

        # Example usage
        print(battery_status_flags)
    elif command_code == 0x17:  # CycleCount: Count unsigned
        device_data['cycle_count'] = decimal_value
    elif command_code == 0x18:  # DesignCapacity: mAh / 40 unsigned
        device_data['design_capacity'] = (decimal_value*40)/1000
    elif command_code == 0x19:  # DesignVoltage: mV unsigned
        device_data['design_voltage'] = round((decimal_value / 1000),1)
    elif command_code == 0x1b:  # ManufactureDate: unsigned int unsigned
        device_data['manufacturer_date'] = decimal_value
    elif command_code == 0x1c:  # SerialNumber: number unsigned
        device_data['serial_number'] = decimal_value
    elif command_code == 'manufacturer_name':  # Manufacturer Name: string
        device_data['manufacturer_name'] = decimal_value 
    elif command_code == 'device_name':  # DeviceName: string
        device_data['device_name'] = decimal_value
    elif command_code == 'firmware_version':  # Firmware Version: string
        device_data['firmware_version'] = decimal_value
    else:

        print(f"Result Command Code Not Found : {hex(command_code)}")


def GetFormatedError(error):
        # Gets the text using the GetErrorText API function
        # If the function success, the translated error is returned. If it fails,
        # a text describing the current error is returned.
        #
        stsReturn = m_objPCANBasic.GetErrorText(error, 0)
        if stsReturn[0] != PCAN_ERROR_OK:
            return "An error occurred. Error-code's text ({0:X}h) couldn't be retrieved".format(error)
        else:
            return stsReturn[1]


def pcan_write_control(call_name):
        print("test control")
        CANMsg = TPCANMsg()
        CANMsg.ID = int('18EFC0D0',16)
        CANMsg.LEN = int(8)
        CANMsg.MSGTYPE = PCAN_MESSAGE_EXTENDED
        CANMsg.DATA[0] = int('03',16)
        CANMsg.DATA[1] = int('03',16)
        CANMsg.DATA[2] = int('37',16)   
        CANMsg.DATA[3] = int('30',16)
        CANMsg.DATA[4] = int('39',16)
        CANMsg.DATA[5] = int('33',16)
        CANMsg.DATA[6] = int('39',16)
        CANMsg.DATA[7] = int('00',16)
        result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
        if result == PCAN_ERROR_OK:
            print("Enter Diag State : Success")
            CANMsg = TPCANMsg()
            CANMsg.ID = int('18EFC0D0',16)
            CANMsg.LEN = int(8)
            CANMsg.MSGTYPE = PCAN_MESSAGE_EXTENDED
            CANMsg.DATA[0] = int('03',16)
            if call_name == 'both_off':
                CANMsg.DATA[1] = int('01',16)
            elif call_name == 'heater_on':
                CANMsg.DATA[1] = int('05',16)
            else:
                CANMsg.DATA[1] = int('00',16)
            if call_name == 'both_off':
                CANMsg.DATA[2] = int('00',16)
            elif call_name == 'charge_on':
                CANMsg.DATA[2] = int('01',16)
            elif call_name == 'discharge_on':
                CANMsg.DATA[2] = int('02',16)
            elif call_name == 'both_on':
                CANMsg.DATA[2] = int('03',16)
            elif call_name == 'bms_reset':
                CANMsg.DATA[2] = int('01',16)
            elif call_name == 'heater_on':
                CANMsg.DATA[2] = int('00',16)
            else:
                messagebox.showinfo("Error!", "Write operation not found!")     
            CANMsg.DATA[3] = int('01',16)
            CANMsg.DATA[4] = int('08',16)
            CANMsg.DATA[5] = int('02',16)
            CANMsg.DATA[6] = int('00',16)
            CANMsg.DATA[7] = int('00',16)
            # m_objPCANBasic.Write(m_PcanHandle, CANMsg)
            result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
            if result == PCAN_ERROR_OK:
                print("FET Control State : Success")
                CANMsg = TPCANMsg()
                CANMsg.ID = int('18EFC0D0',16)
                CANMsg.LEN = int(8)
                CANMsg.MSGTYPE = PCAN_MESSAGE_EXTENDED
                CANMsg.DATA[0] = int('03',16)
                CANMsg.DATA[1] = int('04',16)
                CANMsg.DATA[2] = int('00',16)   
                CANMsg.DATA[3] = int('00',16)
                CANMsg.DATA[4] = int('00',16)
                CANMsg.DATA[5] = int('00',16)
                CANMsg.DATA[6] = int('00',16)
                CANMsg.DATA[7] = int('00',16)
                result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
                if result == PCAN_ERROR_OK:
                    print("Exit the Diag State : Success")
                    return 0
                else:
                    print("Exit the Diag State : Failed")
                    return 0
            else:
                print("FET Control State : Failed")
                return 0
        else:
            print("Enter Diag State : Failed")
            return 0


def update_device_data_to_default():
    for key in device_data.keys():
        if key == 'device_name':
            device_data[key] = "BT-70939APH"  # Example: Set a specific default device name
        elif key == 'manufacturer_name':
            device_data[key] = "Bren-Tronics"  # Example: Set a specific manufacturer name
        elif key == 'charging_battery_status':
            device_data[key] = "Off"  # Example: Set a specific charging status
        elif isinstance(device_data[key], str):
            device_data[key] = ""  # Reset other string values to empty strings
        elif isinstance(device_data[key], (int, float)):
            device_data[key] = 0  # Reset numeric values to zero


def update_battery_status_flags_to_default():
    for key in battery_status_flags.keys():
        battery_status_flags[key] = 0
