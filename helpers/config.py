import os
import json
import pandas as pd  # Ensure you have pandas installed for Excel file handling

# Path to store the configuration file in the user's AppData/Local/ade bms/config directory
CONFIG_DIR = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "config")
CONFIG_FILE_PATH = os.path.join(CONFIG_DIR, "config.json")

# Path for the database folder inside ADE BMS
DATABASE_DIR = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database")

# Excel files to be created inside the database folder
EXCEL_FILES = {
    "can_data.xlsx": [
        "Si No", "Date", "Time", "Project", "Device Name", "Manufacturer Name", 
        "Serial Number", "Cycle Count", "Full Charge Capacity", "Charging Date", 
        "OCV Before Charging", "Discharging Date", "OCV Before Discharging"
    ],
    "rs_data.xlsx": [
        "Si No", "Date", "Time", "Project", "Device Name", "Manufacturer Name", 
        "Serial Number", "Cycle Count", "Full Charge Capacity", "Charging Date", 
        "OCV Before Charging", "Discharging Date", "OCV Before Discharging"
    ],
    "can_charging_data.xlsx": [
        "Date", "Serial Number", "Hour", "Voltage", "Current", 
        "Temperature", "State of Charge"
    ],
    "rs_charging_data.xlsx": [
        "Date", "Serial Number", "Hour", "Voltage", "Current", 
        "Temperature", "State of Charge"
    ],
    "can_discharging_data.xlsx": [
        "Date", "Serial Number", "Hour", "Voltage", "Current", 
        "Temperature", "Load Current"
    ],
    "rs_discharging_data.xlsx": [
        "Date", "Serial Number", "Hour", "Voltage", "Current", 
        "Temperature", "Load Current"
    ]
}

# Default configuration values for CAN and RS232/RS422
default_config = {
    "can_config": {
        "discharging_current_max": 450,
        "logging_time": 5,
        "discharge_cutoff_volt": 21,
        "charging_cutoff_curr": 1,
        "charging_cutoff_volt": 28.4,
        "charging_cutoff_capacity": 95,
        "temperature_caution": 60,
        "temperature_alarm": 80,
        "temperature_critical": 100,
        "projects": ["TAPAS", "ARCHER", "SETB"]  # Default projects for CAN
    },
    "rs_config": {
        "logging_time": 5,
        "discharge_cutoff_volt": 21,
        "charging_cutoff_curr": 450,
        "charging_cutoff_volt": 5000,
        "charging_cutoff_capacity": 95,
        "temperature_caution": 60,
        "temperature_alarm": 80,
        "temperature_critical": 100,
        "projects": ["TAPAS", "ARCHER", "SETB"]  # Default projects for RS
    }
}

# Global variable to hold the current config values
config_values = {}

def ensure_config_dir_exists():
    """Ensure that the configuration directory exists in AppData/Local/ade bms/config."""
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
        print(f"Created directory {CONFIG_DIR} for configuration files.")

def ensure_database_dir_exists():
    """Ensure that the database directory and Excel files exist in AppData/Local/ade bms/database."""
    if not os.path.exists(DATABASE_DIR):
        os.makedirs(DATABASE_DIR)
        print(f"Created directory {DATABASE_DIR} for database files.")

    # Create Excel files with headers if they don't exist
    for file_name, headers in EXCEL_FILES.items():
        file_path = os.path.join(DATABASE_DIR, file_name)
        if not os.path.exists(file_path):
            df = pd.DataFrame(columns=headers)  # Create DataFrame with the specified headers
            df.to_excel(file_path, index=False)  # Save the DataFrame to the Excel file
            print(f"Created Excel file {file_path} with headers: {headers}")
        else:
            print(f"Excel file {file_path} already exists.")

def load_config():
    """Load configuration values from a JSON file, or create the file with default values if it doesn't exist."""
    global config_values
    ensure_config_dir_exists()  # Ensure the config directory exists
    ensure_database_dir_exists()  # Ensure the database directory and Excel files exist
    if os.path.exists(CONFIG_FILE_PATH):
        with open(CONFIG_FILE_PATH, 'r') as config_file:
            config_values = json.load(config_file)
        print(f"Configuration loaded from {CONFIG_FILE_PATH}")
    else:
        # If the config file doesn't exist, create it with default values
        config_values = default_config
        save_config()  # Create the config.json file with default values
        print(f"Configuration file not found. Created default configuration at {CONFIG_FILE_PATH}.")

def save_config():
    """Save the current configuration values to a JSON file."""
    ensure_config_dir_exists()  # Ensure directory exists before saving
    with open(CONFIG_FILE_PATH, 'w') as config_file:
        json.dump(config_values, config_file, indent=4)
    print(f"Configuration saved to {CONFIG_FILE_PATH}")

# The rest of the update methods (update_can_config, update_rs_config, etc.) remain unchanged.
def update_can_config(discharging_current_max=None, logging_time=None, discharge_cutoff_volt=None, 
                      charging_cutoff_curr=None, charging_cutoff_volt=None, charging_cutoff_capacity=None,
                      temperature_caution=None, temperature_alarm=None, temperature_critical=None):
    """Update CAN configuration values and save them to the file."""
    if discharging_current_max is not None:
        config_values['can_config']['discharging_current_max'] = discharging_current_max
    if logging_time is not None:
        config_values['can_config']['logging_time'] = logging_time
    if discharge_cutoff_volt is not None:
        config_values['can_config']['discharge_cutoff_volt'] = discharge_cutoff_volt
    if charging_cutoff_curr is not None:
        config_values['can_config']['charging_cutoff_curr'] = charging_cutoff_curr
    if charging_cutoff_volt is not None:
        config_values['can_config']['charging_cutoff_volt'] = charging_cutoff_volt
    if charging_cutoff_capacity is not None:
        config_values['can_config']['charging_cutoff_capacity'] = charging_cutoff_capacity
    if temperature_caution is not None:
        config_values['can_config']['temperature_caution'] = temperature_caution
    if temperature_alarm is not None:
        config_values['can_config']['temperature_alarm'] = temperature_alarm
    if temperature_critical is not None:
        config_values['can_config']['temperature_critical'] = temperature_critical

    save_config()


def update_rs_config(logging_time=None, discharge_cutoff_volt=None, charging_cutoff_curr=None, 
                     charging_cutoff_volt=None, charging_cutoff_capacity=None, temperature_caution=None,
                     temperature_alarm=None, temperature_critical=None):
    """Update RS configuration values and save them to the file."""
    if logging_time is not None:
        config_values['rs_config']['logging_time'] = logging_time
    if discharge_cutoff_volt is not None:
        config_values['rs_config']['discharge_cutoff_volt'] = discharge_cutoff_volt
    if charging_cutoff_curr is not None:
        config_values['rs_config']['charging_cutoff_curr'] = charging_cutoff_curr
    if charging_cutoff_volt is not None:
        config_values['rs_config']['charging_cutoff_volt'] = charging_cutoff_volt
    if charging_cutoff_capacity is not None:
        config_values['rs_config']['charging_cutoff_capacity'] = charging_cutoff_capacity
    if temperature_caution is not None:
        config_values['rs_config']['temperature_caution'] = temperature_caution
    if temperature_alarm is not None:
        config_values['rs_config']['temperature_alarm'] = temperature_alarm
    if temperature_critical is not None:
        config_values['rs_config']['temperature_critical'] = temperature_critical

    save_config()

# Project management methods
def add_project_to_can(project_name):
    """Add a new project to the CAN project list and save it."""
    if project_name not in config_values['can_config']['projects']:
        config_values['can_config']['projects'].append(project_name)
        save_config()
        print(f"Added project '{project_name}' to CAN projects.")
    else:
        print(f"Project '{project_name}' already exists in CAN projects.")

def add_project_to_rs(project_name):
    """Add a new project to the RS project list and save it."""
    if project_name not in config_values['rs_config']['projects']:
        config_values['rs_config']['projects'].append(project_name)
        save_config()
        print(f"Added project '{project_name}' to RS projects.")
    else:
        print(f"Project '{project_name}' already exists in RS projects.")

def delete_can_project(project_name):
    """Delete a project from the CAN project list."""
    if project_name in config_values['can_config']['projects']:
        config_values['can_config']['projects'].remove(project_name)
        save_config()
        print(f"Deleted project '{project_name}' from CAN projects.")
    else:
        print(f"Project '{project_name}' not found in CAN projects.")

def delete_rs_project(project_name):
    """Delete a project from the RS project list."""
    if project_name in config_values['rs_config']['projects']:
        config_values['rs_config']['projects'].remove(project_name)
        save_config()
        print(f"Deleted project '{project_name}' from RS projects.")
    else:
        print(f"Project '{project_name}' not found in RS projects.")

def edit_can_project(old_name, new_name):
    """Edit the name of a CAN project."""
    if old_name in config_values['can_config']['projects']:
        index = config_values['can_config']['projects'].index(old_name)
        config_values['can_config']['projects'][index] = new_name
        save_config()
        print(f"Renamed CAN project '{old_name}' to '{new_name}'.")
    else:
        print(f"Project '{old_name}' not found in CAN projects.")

def edit_rs_project(old_name, new_name):
    """Edit the name of an RS project."""
    if old_name in config_values['rs_config']['projects']:
        index = config_values['rs_config']['projects'].index(old_name)
        config_values['rs_config']['projects'][index] = new_name
        save_config()
        print(f"Renamed RS project '{old_name}' to '{new_name}'.")
    else:
        print(f"Project '{old_name}' not found in RS projects.")