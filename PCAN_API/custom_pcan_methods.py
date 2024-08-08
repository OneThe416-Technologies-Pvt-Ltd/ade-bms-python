#pcan_methods.py
from PCAN_API.PCANBasic import *
from tkinter import messagebox
import time

m_objPCANBasic = PCANBasic()
m_PcanHandle = 81
can_connected = False
rs_connected = False

#PCAN Initialize API Call
def pcan_initialize(baudrate,hwtype,ioport,interrupt):
        result =  m_objPCANBasic.Initialize(m_PcanHandle,baudrate,hwtype,ioport,interrupt)
        if result != PCAN_ERROR_OK:
            if result == 5120:
                 result = 512
            messagebox.showerror("Error!", GetFormatedError(result))
            return True
        else:
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
def pcan_write(call_name):
        CANMsg = TPCANMsg()
        CANMsg.ID = int('18EFC0D0',16)
        CANMsg.LEN = int(8)
        CANMsg.MSGTYPE = PCAN_MESSAGE_EXTENDED
        CANMsg.DATA[0] = int('01',16)
        CANMsg.DATA[1] = int('00',16)
        if call_name == 'serial_no':
            CANMsg.DATA[2] = int('1c',16)
        elif call_name == 'at_rate':
            CANMsg.DATA[2] = int('04',16)
        elif call_name == 'at_rate_time_to_full':
            CANMsg.DATA[2] = int('05',16)
        elif call_name == 'at_rate_time_to_empty':
            CANMsg.DATA[2] = int('06',16)
        elif call_name == 'at_rate_ok':
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
        elif call_name == 'absolute_state_of_charge':
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
        elif call_name == 'desgin_voltage':
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
        print(f"call name: {call_name}")
        if result != PCAN_ERROR_OK:
            messagebox.showerror("Error!", GetFormatedError(result))
            return 0
        else:
            time.sleep(0.1)
            read_result = retry_pcan_read()
            return read_result
        
def retry_pcan_read(retries=5, delay=0.1):
    for _ in range(retries):
        read_result = pcan_read()
        if read_result != 0:
            return read_result
        time.sleep(delay)
    return 0
             
#PCAN Read API Call
def pcan_read():
    result = m_objPCANBasic.Read(m_PcanHandle)
    # print("can result 0 value",result[0])
    # print("can result 1 value",result[1])
    if result[0] != PCAN_ERROR_OK:
        # messagebox.showerror("Error!", GetFormatedError(result))
        return 250
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

        print(f"Display value: {display_value}")

        return display_value

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
        return decimal_value / 10
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


