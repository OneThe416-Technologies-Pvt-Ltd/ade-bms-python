import customtkinter as ctk
from PIL import Image
Image.CUBIC = Image.BICUBIC
import ttkbootstrap as ttk
#from ttkbootstrap.constants import SUCCESS, WARNING, INFO

class BatteryInfo(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ADE BMS")
        self.geometry("1000x600")
        self.resizable(False, False) 
        # Set custom icon
        self.iconbitmap('asserts/logo/drdo_icon.ico')

        self.main_container = ctk.CTkFrame(self, fg_color="#e0e0e0")
        self.main_container.pack(fill="both", expand=True)

        self.side_menu = ctk.CTkFrame(self.main_container, corner_radius=10, fg_color="#e0e0e0")
        self.side_menu.pack(side="left", fill="y", padx=30, pady=20)

        self.separator = ctk.CTkFrame(self.main_container, width=2, fg_color="#d0d0d0")
        self.separator.pack(side="left", fill="y", padx=5)

        self.content_area = ctk.CTkFrame(self.main_container, corner_radius=10, fg_color="#ffffff")
        self.content_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.create_side_menu_options()

        self.active_button = None
        self.show_content("Device Info")  # Show Device Info by default

    def create_side_menu_options(self):
        self.buttons = []

        # Device Info button
        device_info_button = self.create_custom_button("Device Info", lambda: self.show_content("Device Info"))
        self.buttons.append(("Device Info", device_info_button))
        device_info_button.pack(pady=10, padx=10, fill="x")

        options = [
            ("Temperature", lambda: self.show_content("Temperature")),
            ("Current", lambda: self.show_content("Current")),
            ("Pressure", lambda: self.show_content("Pressure")),
            ("Capacity", lambda: self.show_content("Capacity")),
            ("Voltage", lambda: self.show_content("Voltage")),
        ]
        for text, command in options:
            button = self.create_custom_button(text, command)
            self.buttons.append((text, button))
            button.pack(pady=10, padx=10, fill="x")

    def create_initial_content(self):
        self.content_frame = ctk.CTkFrame(self.content_area, corner_radius=10, fg_color="#ffffff")
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        title_label = ctk.CTkLabel(
            self.content_frame,
            text="BATTERY - INFORMATION",
            font=("Palatino Linotype", 24, "bold"),
            text_color="#333333"
        )
        title_label.pack(pady=20)

        info_frame = ctk.CTkFrame(self.content_frame, corner_radius=10, fg_color="#f8f8f8", border_color="#d0d0d0", border_width=2)
        info_frame.pack(padx=40, pady=40, anchor="n")
        
        labels = [
            "Device Name: Example Device",
            "Serial Number: 123456789",
            "Manufacture Date: 2023-01-01",
            "Manufacturer Name: Example Manufacturer",
            "Battery Status: Good",
            "Cycle Count: 150",
            "Hardware Version: 1.0",
            "Software Version: 2.0"
        ]

        for label_text in labels:
            label = ctk.CTkLabel(info_frame, text=label_text, font=("Palatino Linotype", 18), text_color="#333333")
            label.pack(padx=10, pady=10, anchor="w")

    def show_content(self, content_name):
        for widget in self.content_area.winfo_children():
            widget.destroy()

        for text, button in self.buttons:
            if text == content_name:
                button.configure(fg_color="#003f6b", hover_color="#001f3b")
                self.active_button = button
            else:
                button.configure(fg_color="#007acc", hover_color="#005fa3")

        if content_name == "Device Info":
            self.create_initial_content()
        else:
            content = ContentTemplate(self.content_area, title=f"{content_name} Information", content_name=content_name)
            content.pack(fill="both", expand=True)

    def create_custom_button(self, text, command):
        button = ctk.CTkButton(self.side_menu, text=text, command=lambda: self.on_button_click(text, command), corner_radius=10, fg_color="#007acc", hover_color="#005fa3")
        return button

    def on_button_click(self, text, command):
        if self.active_button:
            self.active_button.configure(fg_color="#007acc", hover_color="#005fa3")

        for t, button in self.buttons:
            if t == text:
                button.configure(fg_color="#003f6b", hover_color="#001f3b")
                self.active_button = button

        command()

class ContentTemplate(ctk.CTkFrame):
    def __init__(self, parent, title, content_name):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        title_label = ctk.CTkLabel(self, text=title, font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(20, 10))
        meter = ttk.Meter(
            master=self,
            metersize=180,
            padding=5,
            amountused=25,
            metertype="semi",
            subtext="miles per hour",
            bootstyle="info"
        )
        meter.pack()

        # update the amount used directly
        meter.configure(amountused = 50)

        # update the subtext
        meter.configure(subtext="Pressure")

        # if content_name == "Temperature":
        #     self.add_meter("Temperature")
        # elif content_name == "Current":
        #     self.add_meter("Current")
        # elif content_name == "Pressure":
        #     self.add_meter("Pressure")
        # elif content_name == "Capacity":
        #     self.add_meter("Capacity")
        # elif content_name == "Voltage":
        #     self.add_meter("Voltage")

    # def add_meter(self, subtext):
    #     meter = self.Meter(
    #         metersize=180,
    #         padding=5,
    #         amountused=25,
    #         metertype="semi",
    #         subtext="miles per hour",
    #     )
    #     meter.pack()

    #     # update the amount used directly
    #     meter.configure(amountused = 50)

    #     # update the subtext
    #     meter.configure(subtext="Pressure")

if __name__ == "__main__":
    app = BatteryInfo()
    app.mainloop()
