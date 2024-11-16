import pyvisa
from tkinter import messagebox
from helpers.logger import logger  # Import the logger

# Initialize variables
rm = pyvisa.ResourceManager()  # Create a ResourceManager instance to communicate with instruments
chroma = None  # To hold the Chroma device connection
load_current = 0  # Default load current
is_connected = False  # Track if the Chroma device is connected

def find_and_connect():
    """Find and connect to the Chroma device."""
    global chroma, is_connected
    try:
        # List available devices
        devices = rm.list_resources()
        # Search through each device to find the Chroma device
        for device in devices:
            try:
                instrument = rm.open_resource(device)
                response = instrument.query("*IDN?")  # Query device ID
                # Check if the device is Chroma
                if "Chroma" in response:
                    chroma = instrument  # Store the Chroma device object
                    chroma.write(f"CONF:VOLT:RANG HIGH\n")  # Set voltage range
                    chroma.write(f"MODE CCH\n")  # Set mode to Constant Current/Voltage
                    is_connected = True  # Flag as connected
                    logger.info(f"Connected to Chroma: {chroma}")
                    return f"Connected to Chroma: {chroma}"  # Return success message
            except Exception as e:
                logger.error(f"Error connecting to device {device}: {e}")
        # If no Chroma device found, show a warning message
        messagebox.showwarning("Device Not Found", "Chroma device not found")
    except Exception as e:
        logger.error(f"Error during device search: {e}")
        messagebox.showerror("Connection Error", "An error occurred while searching for devices.")

def turn_load_on():
    """Turn on the load."""
    global chroma
    try:
        if chroma:
            chroma.write(":LOAD ON\n")  # Send command to turn load on
            logger.info("Load turned on.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error turning load on: {e}")

def turn_load_off():
    """Turn off the load."""
    global chroma
    try:
        if chroma:
            chroma.write(":LOAD OFF\n")  # Send command to turn load off
            logger.info("Load turned off.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error turning load off: {e}")

# Set load to 100A and turn it on
def set_l1_100a_and_turn_on():
    global chroma, load_current
    try:
        load_current = 100
        if chroma:
            chroma.write("CURR:STAT:L1 100")  # Set current to 100A for L1
            turn_load_on()  # Turn the load on
            logger.info("Load set to 100A and turned on.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error setting load to 100A: {e}")

# Set custom L1 value and turn load on or off separately
def set_custom_l1_value(value):
    global chroma, load_current
    try:
        load_current = value
        if chroma:
            chroma.write(f"CURR:STAT:L1 {value}\n")  # Set current to custom value
            logger.info(f"Load set to {value}A.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error setting custom load current: {e}")

# Set load to 25A and turn it on
def set_l1_25a_and_turn_on():
    global chroma, load_current
    try:
        load_current = 25
        if chroma:
            chroma.write("CURR:STAT:L1 25")  # Set current to 25A for L1
            turn_load_on()  # Turn the load on
            logger.info("Load set to 25A and turned on.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error setting load to 25A: {e}")

# Set load to 50A and turn it on
def set_l1_50a_and_turn_on():
    global chroma, load_current
    try:
        load_current = 50
        if chroma:
            chroma.write("CURR:STAT:L1 50")  # Set current to 50A for L1
            turn_load_on()  # Turn the load on
            logger.info("Load set to 50A and turned on.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error setting load to 50A: {e}")

# Set dynamic values for current, time, and repeat
def set_dynamic_values(l1, l2, t1, t2, repeat):
    global chroma
    try:
        if chroma:
            chroma.write(f"MODE CCDH\n")  # Set mode to constant current with dynamic features
            chroma.write(f"CURR:DYN:L1 {l1}\n")  # Set L1 dynamic current
            chroma.write(f"CURR:DYN:L2 {l2}\n")  # Set L2 dynamic current
            chroma.write(f"CURR:DYN:T1 {t1}\n")  # Set time for L1
            chroma.write(f"CURR:DYN:T2 {t2}\n")  # Set time for L2
            chroma.write(f"CURR:DYN:REP {repeat}\n")  # Set repeat value
            logger.info(f"Dynamic load set to L1={l1}, L2={l2}, T1={t1}, T2={t2}, Repeat={repeat}.")
        else:
            logger.warning("Chroma device is not connected.")
    except Exception as e:
        logger.error(f"Error setting dynamic values: {e}")

def get_load_current():
    """Get the current load value."""
    global load_current
    return load_current
