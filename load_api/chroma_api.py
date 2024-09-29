import pyvisa
import time
import threading
from tkinter import messagebox
from openpyxl import Workbook
import os
import datetime


rm = pyvisa.ResourceManager()
chroma = None
is_connected = False

def find_and_connect():
    """Find and connect to the Chroma device."""
    global chroma
    devices = rm.list_resources()
    for device in devices:
        try:
            instrument = rm.open_resource(device)
            response = instrument.query("*IDN?")
            if "Chroma" in response:
                chroma = instrument
                chroma.write(f"CONF:VOLT:RANG HIGH\n")
                chroma.write(f"MODE CCH\n")
                is_connected = True   
                return f"Connected to Chroma: {chroma}"
        except Exception as e:
            pass
    messagebox.showwarning("Device Not Found","Chroma device not found")
def set_current(current):
    """Set the load current on the Chroma device."""
    global chroma
    if chroma:
        try:
            status = chroma.query("STAT:QUES:COND?")
            print(f"Device Status: {status}")
            chroma.write(f"CURR:STAT:L1 {current}\n")
            return f"Current set to {current}A"
        except Exception as e:
            return f"Error setting current: {e}"
    else:
        return "Device not connected"
    
def turn_load_on():
    """Turn on the load."""
    global chroma
    if chroma:
        chroma.write(":LOAD ON\n")

def turn_load_off():
    """Turn off the load."""
    global chroma
    if chroma:
        chroma.write(":LOAD OFF\n")

def set_l1_50a_and_turn_on():
    global chroma
    if chroma:
        chroma.write("CURR:STAT:L1 50")
        turn_load_on()
        print("Load set to 50A and turned on.")

# Method 2: Set L1 to 100A and turn on the load
def set_l1_100a_and_turn_on():
    global chroma
    if chroma:
        chroma.write("CURR:STAT:L1 100")
        turn_load_on()
        print("Load set to 100A and turned on.")

# Method 3: Set custom L1 value and turn load on or off separately
def set_custom_l1_value(value):
    global chroma
    print(value,"test")
    if chroma:
        chroma.write(f"CURR:STAT:L1 {value}\n")
        print(f"Load set to {value}A.")


def set_l1_25a_and_turn_on():
    global chroma
    if chroma:
        chroma.write("CURR:STAT:L1 25")
        turn_load_on()
        print("Load set to 25A and turned on.")

# Method 2: Set L1 to 100A and turn on the load
def set_l1_50a_and_turn_on():
    global chroma
    if chroma:
        chroma.write("CURR:STAT:L1 50")
        turn_load_on()
        print("Load set to 50A and turned on.")

def set_dynamic_values(l1, l2, t1, t2, repeat):
    global chroma
    if chroma:
        chroma.write(f"MODE CCDH\n")
        chroma.write(f"CURR:DYN:L1 {l1}\n")
        chroma.write(f"CURR:DYN:L2 {l2}\n")
        chroma.write(f"CURR:DYN:T1 {t1}\n")
        chroma.write(f"CURR:DYN:T2 {t2}\n")
        chroma.write(f"CURR:DYN:REP {repeat}\n")
        print("Load set to L1,L2,T1,T2,REP ")