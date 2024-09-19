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

def find_and_connect():
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
def turn_load_on():
    """Turn on the load."""
    if device:
        try:
            device.write(":LOAD ON\n")
            return "Load turned ON"
        except Exception as e:
            return f"Error turning on load: {e}"
    else:
        return "Device not connected"
def turn_load_off():
    """Turn off the load."""
    if device:
        try:
            device.write(":LOAD OFF\n")
            return "Load turned OFF"
        except Exception as e:
            return f"Error turning off load: {e}"
    else:
        return "Device not connected"

def set_l1_50a_and_turn_on():
    is_test=True
    is_maintains = False
    result = set_current(50)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 50A and turned on.")
    else:
        print(result)
# Method 2: Set L1 to 100A and turn on the load
def set_l1_100a_and_turn_on():
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
def custom_turn_on():
    """Turn on the load after setting a custom value."""
    result = turn_load_on()
    print(result)
def custom_turn_off():
    """Turn off the load after setting a custom value."""
    result = turn_load_off()
    print(result)
 # Method 1: Set L1 to 50A and turn on the load
def set_l1_25a_and_turn_on():
    is_test=True
    is_maintains = False
    result = set_current(25)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 25A and turned on.")
    else:
        print(result)
# Method 2: Set L1 to 100A and turn on the load
def set_l1_50a_and_turn_on():
    is_test=False
    is_maintains = True
    result = set_current(50)
    if "Error" not in result:
        turn_load_on()
        print("Load set to 50A and turned on.")
    else:
        print(result)