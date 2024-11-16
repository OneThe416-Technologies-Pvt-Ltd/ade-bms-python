import pandas as pd
from fpdf import FPDF
import os
import datetime
from tkinter import messagebox, filedialog
from helpers.logger import logger  # Import the logger

def create_can_report_pdf(serial_number, protocol):
    try:
        # Define the output directory for saving reports
        documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "Battery Reports")
        # Define the input folder path where data files are stored
        folder_path = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database")

        # Determine file paths based on the selected protocol (CAN or RS232)
        if protocol == "CAN":
            logger.info(f"Creating CAN report for serial number {serial_number}")
            data_file = os.path.join(folder_path, "can_data.xlsx")
            charging_file = os.path.join(folder_path, "can_charging_data.xlsx")
            discharging_file = os.path.join(folder_path, "can_discharging_data.xlsx")
        elif protocol == "RS232":
            logger.info(f"Creating RS232 report for serial number {serial_number}")
            data_file = os.path.join(folder_path, "rs_data.xlsx")
            charging_file = os.path.join(folder_path, "rs_charging_data.xlsx")
            discharging_file = os.path.join(folder_path, "rs_discharging_data.xlsx")
        else:
            # Show error if protocol is invalid
            messagebox.showerror("Error", "Invalid protocol. Please use 'CAN' or 'RS'.")
            return

        # Attempt to load data from Excel files, log and notify if files are missing
        try:
            general_data = pd.read_excel(data_file)
            charging_data = pd.read_excel(charging_file)
            discharging_data = pd.read_excel(discharging_file)
        except FileNotFoundError as e:
            messagebox.showerror("Error", f"File not found: {e}")
            logger.error(f"File not found: {e}")
            return

        # Filter the data to include only rows for the specified serial number
        general_data = general_data[general_data["Serial Number"] == int(serial_number)]
        charging_data = charging_data[charging_data["Serial Number"] == int(serial_number)]
        discharging_data = discharging_data[discharging_data["Serial Number"] == int(serial_number)]

        # If no data is found for the serial number, log and notify the user
        if general_data.empty or charging_data.empty or discharging_data.empty:
            logger.info(f"No data found for Serial Number {serial_number}")
            messagebox.showinfo("No Data", "No data found for the given serial number.")
            return

        # Create a project-specific folder in the documents directory if it doesn't exist
        project_name = str(general_data.iloc[0]["Project"])
        project_folder = os.path.join(documents_folder, project_name)
        if not os.path.exists(project_folder):
            os.makedirs(project_folder)

        # Retrieve cycle count for use in the PDF filename
        cycle_count = str(general_data.iloc[0]["Cycle Count"])

        # Define the PDF save path including cycle count and current date
        save_path = os.path.join(
            project_folder, 
            f"Battery_Report_{serial_number}_Cycle_{cycle_count}_{datetime.datetime.now().strftime('%Y-%m-%d')}.pdf"
        )

        # Retrieve general battery information such as charge and discharge dates and OCV values
        charging_date = general_data['Charging Date'].iloc[0] if not general_data.empty else 'N/A'
        discharging_date = general_data['Discharging Date'].iloc[0] if not general_data.empty else 'N/A'
        ocv_before_charging = general_data.get('OCV Before Charging', 'N/A')
        ocv_before_discharging = general_data.get('OCV Before Discharging', 'N/A')

        # Initialize the PDF and set up its layout (auto page break, margins)
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        # Load and insert the company logo, handle exceptions if the logo is missing
        base_path = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "../assets/logo/")
        logo_path = os.path.join(assets_path, "ade_pdf_logo.jpg")
        try:
            pdf.image(logo_path, x=10, y=10, w=190)
        except Exception as e:
            logger.error(f"Error loading logo image: {e}")
        
        pdf.ln(50)  # Space after logo

        # Add title and headers for the report
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Aeronautical Development Establishment, Bangalore", ln=True, align='C')
        pdf.cell(0, 10, "APS", ln=True, align='C')
        pdf.cell(0, 10, "BATTERY REPORT", ln=True, align='C')
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align='C')
        pdf.ln(10)

        # General Battery Information section
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Battery Information", ln=True)

        # Add specific details about the battery (project, serial number, manufacturer, etc.)
        pdf.set_font("Arial", "", 10)
        pdf.cell(50, 8, "Project:", border=1)
        pdf.cell(0, 8, str(general_data.iloc[0]["Project"]), ln=True, border=1)
        pdf.cell(50, 8, "Device Name:", border=1)
        pdf.cell(0, 8, str(general_data.iloc[0]["Device Name"]), ln=True, border=1)
        pdf.cell(50, 8, "Serial Number:", border=1)
        pdf.cell(0, 8, str(serial_number), ln=True, border=1)
        pdf.cell(50, 8, "Manufacturer:", border=1)
        pdf.cell(0, 8, str(general_data.iloc[0]["Manufacturer Name"]), ln=True, border=1)
        pdf.cell(50, 8, "Cycle Count:", border=1)
        pdf.cell(0, 8, str(general_data.iloc[0]["Cycle Count"]), ln=True, border=1)
        pdf.cell(50, 8, "Full Charge Capacity (Ah):", border=1)
        pdf.cell(0, 8, str(general_data.iloc[0]["Full Charge Capacity"]), ln=True, border=1)
        pdf.ln(10)

        # Charging Information Table Header
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Charging Information", ln=True)
        pdf.set_font("Arial", 'I', 12)
        pdf.cell(100, 10, f'Date: {charging_date}', ln=False)
        pdf.cell(0, 10, f'OCV Before Charging: {ocv_before_charging}', ln=True)
        pdf.ln(5)

        # Add charging data table headers and data rows
        pdf.set_font("Arial", "", 10)
        pdf.cell(25, 8, "Mins", border=1)
        pdf.cell(26, 8, "Voltage (V)", border=1)
        pdf.cell(26, 8, "Current (A)", border=1)
        pdf.cell(26, 8, "Temp (°C)", border=1)
        pdf.cell(35, 8, "State of Charge (%)", border=1)
        pdf.cell(40, 8, "Remarks", border=1, ln=True)
        for index, row in charging_data.iterrows():
            pdf.cell(25, 8, str(row["Hour"]), border=1)
            pdf.cell(26, 8, str(row["Voltage"]), border=1)
            pdf.cell(26, 8, str(row["Current"]), border=1)
            pdf.cell(26, 8, str(row["Temperature"]), border=1)
            pdf.cell(35, 8, str(row["State of Charge"]), border=1)
            pdf.cell(40, 8, "", border=1, ln=True)
        pdf.ln(10)

        # Discharging Information Table
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Discharging Information", ln=True)
        pdf.set_font("Arial", 'I', 12)
        pdf.cell(100, 10, f'Date: {discharging_date}', ln=False)
        pdf.cell(0, 10, f'OCV Before Discharging: {ocv_before_discharging}', ln=True)
        pdf.ln(5)

        # Add discharging data table headers and data rows
        pdf.set_font("Arial", "", 10)
        pdf.cell(25, 8, "Mins", border=1)
        pdf.cell(26, 8, "Voltage (V)", border=1)
        pdf.cell(26, 8, "Current (A)", border=1)
        pdf.cell(26, 8, "Temp (°C)", border=1)
        pdf.cell(35, 8, "State of Charge (%)", border=1)
        pdf.cell(40, 8, "Remarks", border=1, ln=True)
        for index, row in discharging_data.iterrows():
            pdf.cell(25, 8, str(row["Hour"]), border=1)
            pdf.cell(26, 8, str(row["Voltage"]), border=1)
            pdf.cell(26, 8, str(row["Current"]), border=1)
            pdf.cell(26, 8, str(row["Temperature"]), border=1)
            pdf.cell(35, 8, str(row["State of Charge"]), border=1)
            pdf.cell(40, 8, "", border=1, ln=True)

        # Attempt to save the PDF file and notify the user
        try:
            pdf.output(save_path)
            messagebox.showinfo("Report Generated", f"Report saved as {save_path}")
            logger.info(f"Report saved successfully: {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save the PDF file: {e}")
            logger.error(f"Failed to save PDF: {e}")
    except Exception as e:
        # Log any unexpected error in the main function
        logger.error(f"An unexpected error occurred: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
