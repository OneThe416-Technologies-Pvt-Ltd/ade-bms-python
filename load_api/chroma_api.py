import pyvisa
import time
import threading
from tkinter import messagebox
from openpyxl import Workbook
import os
import datetime


rm = pyvisa.ResourceManager()
device = None
is_test=False
is_maintains = False
is_turn_on=False
is_connected = False

def find_and_connect(self):
    """Find and connect to the Chroma device."""
    devices = rm.list_resources()
    for device in devices:
        try:
            instrument = rm.open_resource(device)
            response = instrument.query("*IDN?")
            if "Chroma" in response:
                device = instrument
                device.write(f"CONF:VOLT:RANG HIGH\n")
                device.write(f"MODE CCH\n")
                is_connected = True   
                return f"Connected to Chroma: {device}"
        except Exception as e:
            messagebox.showwarning("Error Occured","Turn OFF Load and Turn ON again to Connect")
    messagebox.showwarning("Device Not Found","Chroma device not found")
def set_current(current):
    """Set the load current on the Chroma device."""
    if device:
        try:
            status = device.query("STAT:QUES:COND?")
            print(f"Device Status: {status}")
            device.write(f"CURR:STAT:L1 {current}\n")
            return f"Current set to {current}A"
        except Exception as e:
            return f"Error setting current: {e}"
    else:
        return "Device not connected"
def turn_load_on(self):
    """Turn on the load."""
    if device:
        try:
            device.write(":LOAD ON\n")
            return "Load turned ON"
        except Exception as e:
            return f"Error turning on load: {e}"
    else:
        return "Device not connected"
def turn_load_off(self):
    """Turn off the load."""
    if device:
        try:
            device.write(":LOAD OFF\n")
            return "Load turned OFF"
        except Exception as e:
            return f"Error turning off load: {e}"
    else:
        return "Device not connected"
def set_load_with_timer(current, time_duration):
    """Set the load with a timer."""
    result = set_current(current)
    if "Error" in result:
        return result
    turn_load_on()
    # Create Excel log before starting the load
    # create_excel_log()
    # # Log data at different intervals
    # threading.Thread(target=_log_and_turn_off, args=(time_duration,)).start()
    return f"Load set to {current}A for {time_duration} seconds"
# def create_excel_log(self):
#     """Create an Excel file to log data."""
#     folder_path = os.path.join(os.path.expanduser("~"), "Documents", "Chroma_Logs")
#     if not os.path.exists(folder_path):
#         os.makedirs(folder_path)
#     # Define the log file name
#     file_name = f"Chroma_Log_{serial_number}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
#     file_path = os.path.join(folder_path, file_name)
#     # Create a new Excel workbook
#     workbook = Workbook()
#     sheet = workbook.active
#     sheet.title = "Chroma Data Log"
    
#     # Add headers
#     headers = ["Serial Number", "Timestamp", "Charging Voltage", "Discharging Voltage", "Interval"]
#     sheet.append(headers)
#     log_file_path = file_path
#     print(f"Excel log created at: {file_path}")
# def log_to_excel(interval, voltage, discharging_voltage):
#     """Log data to the Excel file."""
#     timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
#     sheet.append([serial_number, timestamp, voltage, discharging_voltage, interval])
# def save_excel(self):
#     """Save the Excel log to disk."""
#     workbook.save(log_file_path)
#     print(f"Data saved to Excel: {log_file_path}")
# def _log_and_turn_off(time_duration):
#     """Helper function to log data at intervals and turn off the load after a delay."""
#     # Log immediately before starting the load
#     voltage_before = device.query("MEAS:VOLT?")  # Assuming this command reads voltage
#     discharging_voltage = "0"  # Assuming 0 before load starts
#     log_to_excel("Before Load ON", voltage_before, discharging_voltage)
#     time.sleep(2)
#     # Log after 2 seconds
#     voltage_after_2s = device.query("MEAS:VOLT?")
#     log_to_excel("2 Seconds", voltage_after_2s, discharging_voltage)
#     time.sleep(3)
#     # Log after 5 seconds
#     voltage_after_5s = device.query("MEAS:VOLT?")
#     log_to_excel("5 Seconds", voltage_after_5s, discharging_voltage)
#     time.sleep(25)
#     # Log after 30 seconds
#     voltage_after_30s = device.query("MEAS:VOLT?")
#     log_to_excel("30 Seconds", voltage_after_30s, discharging_voltage)
#     # Turn off the load after 30 seconds
#     turn_load_off()
#     save_excel()
#     print("Load turned off after 30 seconds and data logged.")
# Method 1: Set L1 to 50A and turn on the load
def set_l1_50a_and_turn_on(self):
    is_test=True
    is_maintains = False
    result = set_current(50)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 50A and turned on.")
    else:
        print(result)
# Method 2: Set L1 to 100A and turn on the load
def set_l1_100a_and_turn_on(self):
    is_test=False
    is_maintains = True
    result = set_current(100)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 100A and turned on.")
    else:
        print(result)
# Method 3: Set custom L1 value and turn load on or off separately
def set_custom_l1_value(value):
    print(value,"test")
    result = set_current(value)
    if "Error" not in result:
        print(f"Load set to {value}A.")
    else:
        print(result)
def custom_turn_on(self):
    """Turn on the load after setting a custom value."""
    result = turn_load_on()
    print(result)
def custom_turn_off(self):
    """Turn off the load after setting a custom value."""
    result = turn_load_off()
    print(result)
 # Method 1: Set L1 to 50A and turn on the load
def set_l1_25a_and_turn_on(self):
    is_test=True
    is_maintains = False
    result = set_current(25)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 50A and turned on.")
    else:
        print(result)
# Method 2: Set L1 to 100A and turn on the load
def set_l1_50a_and_turn_on(self):
    is_test=False
    is_maintains = True
    result = set_current(50)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 100A and turned on.")
    else:
        print(result)