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
            'design_capacity': 0,
            'design_voltage': 0,
            'remaining_capacity': 0,
            'temperature': 0,
            'current': 0,
            'voltage': 0,
            'battery_status': "",
            'cycle_count': 0,
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
    value = await asyncio.to_thread(pcan_write_read, call_name)  # Run in a separate thread
    device_data[key] = value

async def process_data_points():
    data_points = [
        ('temperature', 'temperature'),
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
    await asyncio.gather(*tasks)

async def pcan_initialize(baudrate, hwtype, ioport, interrupt):
    result = m_objPCANBasic.Initialize(m_PcanHandle, baudrate, hwtype, ioport, interrupt)
    if result != PCAN_ERROR_OK:
        if result == 5120:
            result = 512
        messagebox.showerror("Error!", GetFormatedError(result))
        return False
    else:
        messagebox.showinfo("Info!", "Connection established!")
        await process_data_points()
        
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
        else:
            messagebox.showinfo("Error!", "Write operation not found!")     
        CANMsg.DATA[3] = int('01',16)
        CANMsg.DATA[4] = int('08',16)
        CANMsg.DATA[5] = int('02',16)
        CANMsg.DATA[6] = int('00',16)
        CANMsg.DATA[7] = int('00',16)
        result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
        # result = 0
        # print(f"call name: {call_name}")
        if result != PCAN_ERROR_OK:
            messagebox.showerror(f"Error! {call_name}", GetFormatedError(result))
            return -2
        else:
            time.sleep(0.1)
            read_result = retry_pcan_read(call_name)
            return read_result


def retry_pcan_read(call_name, retries=1, delay=0.1):
    for _ in range(retries):
        result_code, read_result = pcan_read(call_name)
        if result_code == PCAN_ERROR_OK:
            return read_result
        time.sleep(delay)
    return read_result


#PCAN Read API Call
def pcan_read(call_name):
    result = m_objPCANBasic.Read(m_PcanHandle)
    # result = 1
    # print("can result 0 value",result[0])
    # print("can result 1 value",result[1])
    if result[0] != PCAN_ERROR_OK:
        # messagebox.showerror(f"Error!{call_name}", GetFormatedError(result[0]))
        return result[0], 25
    else:
        args = result[1:]
        # print("args",args[0])
        theMsg = args[0]
        # print("theMsg",theMsg)
        newMsg = TPCANMsgFD()
        newMsg.ID = theMsg.ID
        newMsg.DLC = theMsg.LEN
        # print("theMsg.LEN",theMsg.LEN)
        for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN):
            newMsg.DATA[i] = theMsg.DATA[i]
        
        command_code = newMsg.DATA[2]
        first_byte = newMsg.DATA[0]
        second_byte = newMsg.DATA[1]

        # Swap the bytes to create the correct hex value
        swapped_hex = (second_byte << 8) | first_byte
        decimal_value = int(swapped_hex)

        display_value = convert_data(command_code, decimal_value)

        print(f"Display value {call_name}: {display_value}")

        return result[0], display_value


def convert_data(command_code, decimal_value):
    # Conversion rules
    if command_code == 0x04:  # AtRate: mA / 40 unsigned
        return decimal_value * 40
    elif command_code == 0x05:  # AtRateTimeToFull: minutes unsigned
        return decimal_value
    elif command_code == 0x06:  # AtRateTimeToEmpty: minutes unsigned
        return decimal_value
    elif command_code == 0x07:  # AtRateOK: Boolean
        return "Yes" if decimal_value != 0 else "No" 
    elif command_code == 0x08:  # Temperature: 0.1⁰K signed
        return (decimal_value/1000)-273.15
    elif command_code == 0x09:  # Voltage: mV unsigned
        return decimal_value
    elif command_code == 0x0c:  # MaxError: Percent unsigned
        return decimal_value
    elif command_code == 0x0d:  # RelStateofCharge: Percent unsigned
        return decimal_value
    elif command_code == 0x0e:  # AbsoluteStateofCharge: Percent unsigned
        return decimal_value
    elif command_code == 0x0f:  # RemainingCapacity: mAh / 40 unsigned
        return decimal_value * 40
    elif command_code == 0x10:  # FullChargeCapacity: mAh / 40 unsigned
        return decimal_value * 40
    elif command_code == 0x11:  # RunTimeToEmpty: minutes unsigned
        return decimal_value
    elif command_code == 0x12:  # AvgTimeToEmpty: minutes unsigned
        return decimal_value
    elif command_code == 0x13:  # AvgTimeToFull: minutes unsigned
        return decimal_value
    elif command_code == 0x14:  # ChargingCurrent: mA / 40 unsigned
        return decimal_value * 40
    elif command_code == 0x15:  # ChargingVoltage: mV unsigned
        return decimal_value
    elif command_code == 0x16:  # BatteryStatus4: bit flags unsigned
        return bin(decimal_value)
    elif command_code == 0x17:  # CycleCount: Count unsigned
        return decimal_value
    elif command_code == 0x18:  # DesignCapacity: mAh / 40 unsigned
        return decimal_value * 40
    elif command_code == 0x19:  # DesignVoltage: mV unsigned
        return decimal_value
    elif command_code == 0x1b:  # ManufactureDate: unsigned int unsigned
        return decimal_value
    elif command_code == 0x1c:  # SerialNumber: number unsigned
        return decimal_value
    elif command_code in [0x20, 0x21]:  # ManufacturerName and DeviceName: string
        return hex(decimal_value)  # Assuming string data is processed differently

    return decimal_value


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
        CANMsg.DATA[1] = int('00',16)
        if call_name == 'both_off':
            CANMsg.DATA[2] = int('00',16)
        elif call_name == 'charge_on':
            CANMsg.DATA[2] = int('01',16)
        elif call_name == 'discharge_on':
            CANMsg.DATA[2] = int('02',16)
        elif call_name == 'both_on':
            CANMsg.DATA[2] = int('03',16)
        else:
            messagebox.showinfo("Error!", "Write operation not found!")     
        CANMsg.DATA[3] = int('01',16)
        CANMsg.DATA[4] = int('08',16)
        CANMsg.DATA[5] = int('02',16)
        CANMsg.DATA[6] = int('00',16)
        CANMsg.DATA[7] = int('00',16)
        result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)
        # result = 0
        if result == PCAN_ERROR_OK:
            print(f"call name: {call_name}")
            return 0
        else:
            # messagebox.showerror("Error",f"{call_name}")
            return 0

