# pcomm.py

from ctypes import *
import platform

# Constants for the port and configuration
PORT = 2  # COM2
BAUD_RATE = 19200
PARITY_NONE = 0x00
DATA_BITS_8 = 0x08
STOP_BITS_1 = 0x01
BUFFER_SIZE = 256

# Load the PComm library
if platform.system() == 'Windows':
    pcomm_dll = windll.LoadLibrary("PCOMM")
else:
    raise Exception("PComm library is only supported on Windows")

# Define PComm functions
sio_open = pcomm_dll.sio_open
sio_open.restype = c_int
sio_open.argtypes = [c_int]

sio_ioctl = pcomm_dll.sio_ioctl
sio_ioctl.restype = c_int
sio_ioctl.argtypes = [c_int, c_int, c_int]

sio_write = pcomm_dll.sio_write
sio_write.restype = c_int
sio_write.argtypes = [c_int, c_char_p, c_int]

sio_read = pcomm_dll.sio_read
sio_read.restype = c_int
sio_read.argtypes = [c_int, c_char_p, c_int]

sio_close = pcomm_dll.sio_close
sio_close.restype = c_int
sio_close.argtypes = [c_int]

# Port Control - Enable the port
ret = sio_open(PORT)

if ret == 0:  # SIO_OK
    # Port Control - Set baud, parity, data bits, and stop bits
    ret = sio_ioctl(PORT, BAUD_RATE, PARITY_NONE | DATA_BITS_8 | STOP_BITS_1)
    
    if ret == 0:  # SIO_OK
        # Output Data function - Write data to the port
        data_to_write = b"ABCDE"
        sio_write(PORT, data_to_write, len(data_to_write))
        
        # Input Data function - Read data from the port
        buffer = create_string_buffer(BUFFER_SIZE)
        sio_read(PORT, buffer, BUFFER_SIZE)
        print("Data received:", buffer.value)
    
    # Port Control - Disable the port
    sio_close(PORT)
else:
    print("Failed to open COM port")
