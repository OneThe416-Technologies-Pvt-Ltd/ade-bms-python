import pyvisa
import time
import threading

class ChromaAPI:
    def __init__(self):
        # Initialize VISA resource manager
        self.rm = pyvisa.ResourceManager()
        self.device = None

    def find_and_connect(self):
        """Find and connect to the Chroma device."""
        devices = self.rm.list_resources()
        for device in devices:
            try:
                instrument = self.rm.open_resource(device)
                response = instrument.query("*IDN?")
                if "Chroma" in response:
                    self.device = instrument
                    return f"Connected to Chroma: {device}"
            except Exception as e:
                return f"Error connecting to device: {e}"
        return "Chroma device not found"

    def set_current(self, current):
        """Set the load current on the Chroma device."""
        if self.device:
            try:
                self.device.write(f":CURR {current}\n")
                return f"Current set to {current}A"
            except Exception as e:
                return f"Error setting current: {e}"
        else:
            return "Device not connected"

    def turn_load_on(self):
        """Turn on the load."""
        if self.device:
            try:
                self.device.write(":LOAD ON\n")
                return "Load turned ON"
            except Exception as e:
                return f"Error turning on load: {e}"
        else:
            return "Device not connected"

    def turn_load_off(self):
        """Turn off the load."""
        if self.device:
            try:
                self.device.write(":LOAD OFF\n")
                return "Load turned OFF"
            except Exception as e:
                return f"Error turning off load: {e}"
        else:
            return "Device not connected"

    def set_load_with_timer(self, current, time_duration):
        """Set the load with a timer."""
        result = self.set_current(current)
        if "Error" in result:
            return result
        self.turn_load_on()

        # Turn off load after the specified time
        threading.Thread(target=self._turn_off_after_delay, args=(time_duration,)).start()
        return f"Load set to {current}A for {time_duration} seconds"

    def _turn_off_after_delay(self, time_duration):
        """Helper function to turn off the load after a delay."""
        time.sleep(time_duration)
        self.turn_load_off()
