#pcan_methods.py
from PCAN_API.PCANBasic import *
from tkinter import messagebox

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
        elif call_name == 'manufacturer_ate':
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
        print(f"Can write: {CANMsg.DATA[2]}")
        result = m_objPCANBasic.Write(m_PcanHandle, CANMsg)  
        if result != PCAN_ERROR_OK:
            messagebox.showerror("Error!", GetFormatedError(result))
            return 0, False
        else:
            read_result = pcan_read()
            return read_result, True
             
#PCAN Read API Call
def pcan_read():
        second_byte = int('73', 16)
        first_byte = int('E2', 16)

        # Swap the bytes and create the hexadecimal number
        swapped_hex = (second_byte << 8) | first_byte

        # Convert to decimal
        result = int(swapped_hex)
        first_four_digits = int(str(result)[:4])
        
        # Divide the extracted number by 100 and round to two decimal places
        read_result = round(first_four_digits / 100, 2)

        print(f"Original first byte: {first_byte:02X}")
        print(f"Original second byte: {second_byte:02X}")
        print(f"Swapped hex: {second_byte:02X}{first_byte:02X}")
        print(f"Decimal value: {read_result}") 
        return read_result
        # result = m_objPCANBasic.Read(m_PcanHandle)
        # if result == PCAN_ERROR_OK:
        #     args = result[1:]
        #     theMsg = args[0][0]
        #     newMsg = TPCANMsg()
        #     for i in range(8 if (theMsg.LEN > 8) else theMsg.LEN):
        #         newMsg.DATA[i] = theMsg.DATA[i]
        #     # first_byte = newMsg.DATA[0]
        #     # second_byte = newMsg.DATA[1]
        #     first_byte = 'E2'
        #     second_byte = '73'

        #     # Swap the bytes and create the hexadecimal number
        #     swapped_hex = (second_byte << 8) | first_byte

        #     # Convert to decimal
        #     decimal_value = int(swapped_hex)
        #     print(f"Original first byte: {first_byte:02X}")
        #     print(f"Original second byte: {second_byte:02X}")
        #     print(f"Swapped hex: {second_byte:02X}{first_byte:02X}")
        #     print(f"Decimal value: {decimal_value}")
        #     return decimal_value
        # else:
        #     messagebox.showinfo("Connect Confirmation", "Connection established!")

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


