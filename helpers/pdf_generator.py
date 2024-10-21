import pandas as pd
from fpdf import FPDF
import os
import datetime
from tkinter import messagebox, filedialog

def create_can_report_pdf(serial_number, protocol):
    # Get the path to the Documents folder
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", "Battery Reports")
    # Define the folder path
    folder_path = os.path.join(os.getenv('LOCALAPPDATA'), "ADE BMS", "database")
    
    # Select file paths based on the protocol (CAN or RS)
    if protocol == "CAN":
        print(f"{serial_number} test1 pdf")
        data_file = os.path.join(folder_path, "can_data.xlsx")
        charging_file = os.path.join(folder_path, "can_charging_data.xlsx")
        discharging_file = os.path.join(folder_path, "can_discharging_data.xlsx")
    elif protocol == "RS232":
        print(f"{serial_number} test2 pdf")
        data_file = os.path.join(folder_path, "rs_data.xlsx")
        charging_file = os.path.join(folder_path, "rs_charging_data.xlsx")
        discharging_file = os.path.join(folder_path, "rs_discharging_data.xlsx")
    else:
        messagebox.showerror("Error", "Invalid protocol. Please use 'CAN' or 'RS'.")
        return

    # Load the data from the Excel files
    try:
        general_data = pd.read_excel(data_file)
        charging_data = pd.read_excel(charging_file)
        discharging_data = pd.read_excel(discharging_file)
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"File not found: {e}")
        return

    # Filter the data by Serial Number
    general_data = general_data[general_data["Serial Number"] == int(serial_number)]
    charging_data = charging_data[charging_data["Serial Number"] == int(serial_number)]
    discharging_data = discharging_data[discharging_data["Serial Number"] == int(serial_number)]

    if general_data.empty or charging_data.empty or discharging_data.empty:
        print(f"{general_data} general_data pdf")
        print(f"{charging_data} charging_data pdf")
        print(f"{discharging_data} discharging_data pdf")
        messagebox.showinfo("No Data", "No data found for the given serial number.")
        return

    # Create project folder based on the "Project" name
    project_name = str(general_data.iloc[0]["Project"])
    project_folder = os.path.join(documents_folder, project_name)
    if not os.path.exists(project_folder):
        os.makedirs(project_folder)

    # Retrieve the cycle count from the general data
    cycle_count = str(general_data.iloc[0]["Cycle Count"])

    # Define the PDF save path including the cycle count in the filename
    save_path = os.path.join(
        project_folder, 
        f"Battery_Report_{serial_number}_Cycle_{cycle_count}_{datetime.datetime.now().strftime('%Y-%m-%d')}.pdf"
    )


    charging_date = general_data['Charging Date'].iloc[0] if not general_data.empty else 'N/A'
    discharging_date = general_data['Discharging Date'].iloc[0] if not general_data.empty else 'N/A'
    ocv_before_charging = general_data['OCV Before Charging'].iloc[0] if 'OCV Before Charging' in general_data.columns else 'N/A'
    ocv_before_discharging = general_data['OCV Before Discharging'].iloc[0] if 'OCV Before Discharging' in general_data.columns else 'N/A'

    if not save_path:
        messagebox.showinfo("Cancelled", "File saving cancelled.")
        return

    # Initialize the PDF and set up margins and font
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Margins
    left_margin = 15
    right_margin = 15
    hour_width = 25
    voltage_width = 26
    current_width = 26
    temperature_width = 26
    soc_width = 35  # State of Charge or Load Current
    remarks_width = 40
    pdf.set_left_margin(left_margin)
    pdf.set_right_margin(right_margin)

    # Logo
    base_path = os.path.dirname(os.path.abspath(__file__))
    assets_path = os.path.join(base_path, "../assets/logo/")
    logo_path = os.path.join(assets_path,"ade_pdf_logo.jpg")
    pdf.image(logo_path, x=10, y=10, w=190)
    pdf.ln(50)  # Space after the logo

    # Title and Headers
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Aeronautical Development Establishment, Bangalore", ln=True, align='C')
    pdf.cell(0, 10, "APS", ln=True, align='C')
    pdf.cell(0, 10, "BATTERY REPORT", ln=True, align='C')
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align='C')
    pdf.ln(10)

    # General Battery Information
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Battery Information", ln=True)

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

    pdf.ln(10)  # Space before the next section

    # Charging Information Table
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Charging Information", ln=True)

    pdf.set_font("Arial", 'I', 12)
    pdf.cell(100, 10, f'Date: {charging_date}', ln=False)
    pdf.cell(0, 10, f'OCV Before Charging: {ocv_before_charging}', ln=True)
    pdf.ln(5)  # Line break

    pdf.set_font("Arial", "", 10)
    pdf.cell(hour_width, 8, "Mins", border=1)
    pdf.cell(voltage_width, 8, "Voltage (V)", border=1)
    pdf.cell(current_width, 8, "Current (A)", border=1)
    pdf.cell(temperature_width, 8, "Temp (°C)", border=1)
    pdf.cell(soc_width, 8, "State of Charge (%)", border=1)
    pdf.cell(remarks_width, 8, "Remarks", border=1, ln=True)  # Extra column for Remarks

    for index, row in charging_data.iterrows():
        pdf.cell(hour_width, 8, str(row["Hour"]), border=1)
        pdf.cell(voltage_width, 8, str(row["Voltage"]), border=1)
        pdf.cell(current_width, 8, str(row["Current"]), border=1)
        pdf.cell(temperature_width, 8, str(row["Temperature"]), border=1)
        pdf.cell(soc_width, 8, str(row["State of Charge"]), border=1)
        pdf.cell(remarks_width, 8, "", border=1, ln=True)

    pdf.ln(10)

    # Discharging Information Table
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Discharging Information", ln=True)

    pdf.set_font("Arial", 'I', 12)
    pdf.cell(100, 10, f'Date: {discharging_date}', ln=False)
    pdf.cell(0, 10, f'OCV Before Discharging: {ocv_before_discharging}', ln=True)
    pdf.ln(5)  # Line break

    pdf.set_font("Arial", "", 10)
    pdf.cell(hour_width, 8, "Mins", border=1)
    pdf.cell(voltage_width, 8, "Voltage (V)", border=1)
    pdf.cell(current_width, 8, "Current (A)", border=1)
    pdf.cell(temperature_width, 8, "Temp (°C)", border=1)
    pdf.cell(soc_width, 8, "Load Current (A)", border=1)
    pdf.cell(remarks_width, 8, "Remarks", border=1, ln=True)  # Extra column for Remarks

    for index, row in discharging_data.iterrows():
        pdf.cell(hour_width, 8, str(row["Hour"]), border=1)
        pdf.cell(voltage_width, 8, str(row["Voltage"]), border=1)
        pdf.cell(current_width, 8, str(row["Current"]), border=1)
        pdf.cell(temperature_width, 8, str(row["Temperature"]), border=1)
        pdf.cell(soc_width, 8, str(row["Load Current"]), border=1)
        pdf.cell(remarks_width, 8, "", border=1, ln=True)

    pdf.ln(20)

    # Define column widths
    prepared_by_width = 50
    checked_by_width = 50
    approved_by_width = 70
    signature_width = 50

    # Prepared By, Checked By, Approved By
    pdf.set_font("Arial", "B", 12)
    total_width = prepared_by_width + checked_by_width + approved_by_width

    # Aligning the headers with equal spacing
    pdf.cell(prepared_by_width, 10, "Prepared By:", 0, 0)
    pdf.cell(checked_by_width, 10, "Checked By:", 0, 0, "C")
    pdf.cell(approved_by_width, 10, "Approved By:", 0, 1, "C")

    pdf.ln(10)

    # Signature section
    # Row for Signature
    pdf.cell(50, 10, "Signature:", 0, 0)
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Signature space
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Signature space
    pdf.cell(50, 10, "", 0, 1)  # Move to next line

    # Row for Name
    pdf.cell(50, 10, "Name:", 0, 0)
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Name space
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Name space
    pdf.cell(50, 10, "", 0, 1)  # Move to next line

    # Row for Designation
    pdf.cell(50, 10, "Designation:", 0, 0)
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Designation space
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Designation space
    pdf.cell(50, 10, "", 0, 1)  # Move to next line

    # Row for Date
    pdf.cell(50, 10, "Date:", 0, 0)
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Date space
    pdf.cell(50, 10, "", 0, 0, "C")  # Placeholder for Date space
    pdf.cell(50, 10, "", 0, 1)  # Move to next line

    # Save the PDF
    pdf.output(save_path)

    # Notify user about the saved file
    print(f"PDF saved successfully to {save_path}")

    # Open the PDF file automatically after creation (if on Windows)
    if os.name == "nt":
        os.startfile(save_path)
