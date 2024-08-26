#pcan_methods.py

from pcan_api.pcan import *
from tkinter import messagebox
import time
import asyncio

m_objPCANBasic = PCANBasic()
m_PcanHandle = 81

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
    "design_capacity": "mAh / 40",
    "design_voltage": "mV",
    "at_rate_ok_text": "yes/no",
    "at_rate_time_to_full": "minutes",
    "at_rate_time_to_empty": "minutes",
    "at_rate":"mA / 40 ",
    "rel_state_of_charge": "Percent",
    "abs_state_of_charge": "Percent",
    "run_time_to_empty": "minutes",
    "avg_time_to_empty": "minutes",
    "avg_time_to_full": "minutes",
    "max_error": "Percent",
    "temperature": "0.1°K",
    "current": "mA / 40",
    "remaining_capacity": "mAh / 40",
    "voltage": "mV",
    "avg_current": "mA / 40",
    "charging_current": "mA / 40",
    "full_charge_capacity": "mAh / 40",
    "charging_voltage": "mV"
}

device_data = {
            'device_name': "",
            'serial_number': 0,
            'manufacturer_date': 0,
            'manufacturer_name': "",
            'firmware_version': "",
            'battery_status': "",
            'cycle_count': 0,
            'design_capacity': 0,
            'design_voltage': 0,
            'remaining_capacity': 0,
            'temperature': 0,
            'current': 0,
            'voltage': 0,
            'avg_current': 0,
            'charging_current': 0,
            'full_charge_capacity': 0,
            'charging_voltage': 0,
            'at_rate_time_to_full': 0,
            'at_rate_time_to_empty': 0,
            'at_rate_ok_text': "",
            'at_rate':0,
            'rel_state_of_charge': 0,
            'abs_state_of_charge': 0,
            'run_time_to_empty': 0,
            'avg_time_to_empty': 0,
            'avg_time_to_full': 0,
            'max_error': 0
        } 


# async def fetch_and_store_data(call_name, key):
#     value = await pcan_write_read(call_name)  # Assuming pcan_write_read is async
#     device_data[key] = value

async def update_device_data():
    data_points = [
        ('serial_number', 'serial_number'),
        # ('firmware_version', 'firmware_version'),
        ('design_capacity', 'design_capacity'),
        ('design_voltage', 'design_voltage'),
        ('manufacturer_date', 'manufacturer_date'),
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
        return False
    else:
        pcan_write_read('serial_number')
        if device_data['serial_number'] != 0:
            pcan_write_read('temperature')
            # pcan_write_read('firmware_version')       
            pcan_write_read('voltage')       
            pcan_write_read('current')       
            pcan_write_read('remaining_capacity')       
            pcan_write_read('full_charge_capacity')
            pcan_write_read('device_name') 
            pcan_write_read('serial_number')
            pcan_write_read('manufacturer_name')
            messagebox.showinfo("Info!", "Connection established!")
        else:
           pcan_write_read('serial_number')
           if device_data['serial_number'] != 0:
               pcan_write_read('temperature')
               # pcan_write_read('firmware_version')       
               pcan_write_read('voltage')       
               pcan_write_read('current')       
               pcan_write_read('remaining_capacity')       
               pcan_write_read('full_charge_capacity')
               pcan_write_read('device_name') 
               pcan_write_read('serial_number')
               pcan_write_read('manufacturer_name')
               messagebox.showinfo("Info!", "Connection established!")
        
        return True


#PCAN Uninitialize API Call
def pcan_uninitialize(): 
        result =  m_objPCANBasic.Uninitialize(m_PcanHandle)
        if result != PCAN_ERROR_OK:
            messagebox.showerror("Error!", GetFormatedError(result))
            return False
        else:
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
        if result != PCAN_ERROR_OK:
            messagebox.showerror(f"Error! {call_name}", GetFormatedError(result))
            return -2
        else:
            time.sleep(0.1)
            result_code = retry_pcan_read(call_name)
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
        
        if call_name == 'current':
            current = [hex(theMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]
            print(f"current: {current}")
        
        command_code = newMsg.DATA[4]
        first_byte = newMsg.DATA[0]
        second_byte = newMsg.DATA[1]

        # Swap the bytes to create the correct hex value
        swapped_hex = (second_byte << 8) | first_byte
        decimal_value = int(swapped_hex)

        if newMsg.ID == 0x1CECFFC0 or newMsg.ID == 0x1CEBFFC0:
            if newMsg.ID == 0x1CECFFC0:
                hex_string1=''
                hex_string2=''
                result = m_objPCANBasic.Read(m_PcanHandle)
                if result[0] != PCAN_ERROR_OK:
                    return result[0]
                else:
                    args = result[1:]
                    theMsg = args[0]
                    byte_values1 = [hex(theMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]
                    hex_string1 = ''.join(format(int(h, 16), '02X') for h in byte_values1)
                result = m_objPCANBasic.Read(m_PcanHandle)
                if result[0] != PCAN_ERROR_OK:
                    return result[0]
                else:
                    args = result[1:]
                    theMsg2 = args[0]
                    byte_values2 = [hex(theMsg2.DATA[i]) for i in range(8 if (theMsg2.LEN > 8) else theMsg2.LEN)]
                    hex_string2 = ''.join(format(int(h, 16), '02X') for h in byte_values2)
                ascii_string1 = ''.join(chr(b) for b in bytes.fromhex(hex_string1) if 32 <= b <= 126)
                ascii_string2 = ''.join(chr(b) for b in bytes.fromhex(hex_string2) if 32 <= b <= 126)

                # Join the two ASCII strings
                final_string = ascii_string1 + ascii_string2
                if call_name == 'manufacturer_name':
                    convert_data('manufacturer_name', final_string)
                elif call_name == 'device_name':
                    convert_data('device_name', final_string)
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
                ascii_string1 = ''.join(chr(b) for b in bytes.fromhex(hex_string1) if 32 <= b <= 126)
                ascii_string2 = ''.join(chr(b) for b in bytes.fromhex(hex_string2) if 32 <= b <= 126)
                # Join the two ASCII strings
                final_string = ascii_string1 + ascii_string2
                if call_name == 'manufacturer_name':
                    convert_data('manufacturer_name', final_string)
                elif call_name == 'device_name':
                    convert_data('device_name', final_string)
        elif call_name == 'firmware_version':
            data_packet = [hex(newMsg.DATA[i]) for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN)]

            # Extract version information
            major_version = data_packet[0]
            minor_version = data_packet[1]
            patch_number = (data_packet[3] << 8) | data_packet[2]
            build_number = (data_packet[5] << 8) | data_packet[4]

            version_string = f"{major_version}.{minor_version}.{patch_number}.{build_number}"
            convert_data('firmware_version', version_string)
        else:
            convert_data(command_code, decimal_value)

        return result[0]


def convert_data(command_code, decimal_value):
    # Conversion rules
    if command_code == 0x04:  # AtRate: mA / 40 unsigned
        device_data['at_rate'] = decimal_value
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
        else:
            if decimal_value > 32767:
                decimal_value -= 65536
            currentmA = decimal_value*40
            currentA = currentmA / 1000
            device_data['current'] = currentA
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
        device_data['remaining_capacity'] = decimal_value
    elif command_code == 0x10:  # FullChargeCapacity: mAh / 40 unsigned
        device_data['full_charge_capacity'] = decimal_value
    elif command_code == 0x11:  # RunTimeToEmpty: minutes unsigned
        device_data['run_time_to_empty'] = round((decimal_value / 10),1)
    elif command_code == 0x12:  # AvgTimeToEmpty: minutes unsigned
        device_data['avg_time_to_empty'] = round((decimal_value / 10),1)
    elif command_code == 0x13:  # AvgTimeToFull: minutes unsigned
        device_data['avg_time_to_full'] = round((decimal_value / 1000),1)
    elif command_code == 0x14:  # ChargingCurrent: mA / 40 unsigned
        if decimal_value == 0:
            device_data['charging_current'] = decimal_value
        else:
            if decimal_value > 32767:
                decimal_value -= 65536
            currentmA = decimal_value*40
            currentA = currentmA / 1000
            device_data['charging_current'] = currentA
    elif command_code == 0x15:  # ChargingVoltage: mV unsigned
        device_data['charging_voltage'] = round((decimal_value / 1000),1)
    elif command_code == 0x16:  # BatteryStatus: bit flags unsigned
        device_data['battery_status'] = hex(decimal_value)
    elif command_code == 0x17:  # CycleCount: Count unsigned
        device_data['cycle_count'] = decimal_value
    elif command_code == 0x18:  # DesignCapacity: mAh / 40 unsigned
        device_data['design_capacity'] = decimal_value
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
        print(f"Command Code Not Found : {hex(command_code)}")


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

