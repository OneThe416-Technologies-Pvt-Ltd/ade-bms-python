from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        # Add a logo
        self.image('assets/logo/ade_logo.png', 10, 8, 33)  # Assuming 'logo.png' is your logo file
        self.set_font('Arial', 'B', 12)
        self.cell(80)
        # Title
        self.cell(30, 10, 'Battery Testing Report', 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'I', 8)
        # Page number
        self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'C')
    
    def chapter_title(self, title):
        # Arial 12
        self.set_font('Arial', 'B', 12)
        # Background color
        self.set_fill_color(200, 220, 255)
        # Title
        self.cell(0, 10, '%s' % title, 0, 1, 'L', 1)
        # Line break
        self.ln(4)
    
    def chapter_body(self, body):
        # Read text file
        self.set_font('Arial', '', 12)
        # Output justified text
        self.multi_cell(0, 10, body)
        # Line break
        self.ln()
    
    def add_device_info(self, info):
        self.add_page()
        self.chapter_title('Battery Information')
        for key, value in info.items():
            self.cell(0, 10, f'{key}: {value}', 0, 1)
        self.ln()

pdf = PDF()

# Sample data
device_info = {
    "Device Name": "Battery Tester XYZ",
    "Manufacturer Name": "Battery Co.",
    "Serial Number": "123456789",
    "Capacity": "1000mAh",
    "Voltage": "21V",
    "Start Time": "2024-08-31 19:50:25",
    "End Time": "2024-08-31 23:35:25"
}

pdf.set_title('Battery Testing Report')
pdf.set_author('Battery Testing System')

# Adding device info to the PDF
pdf.add_device_info(device_info)

# Save the PDF to a file
pdf.output('Battery_Testing_Report.pdf')
