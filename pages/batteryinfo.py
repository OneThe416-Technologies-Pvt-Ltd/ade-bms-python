import tkinter as tk
import customtkinter as ctk
from PCAN_API.custom_pcan_methods import *
from components.gauges.thermometer import CircularGauge
from components.gauges.gauge import full_circle_gauge, semi_circle_gauge, quarter_circle_gauge, custom_arc_gauge
from ttkbootstrap import Style

class BatteryInfo(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ADE BMS")
        self.geometry("1000x600")
        self.resizable(False, False)  # Disable maximize

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
            ("Temperature", lambda: self.show_content("content1")),
            ("Current", lambda: self.show_content("content2")),
            ("Pressure", lambda: self.show_content("content3")),
            ("Capacity", lambda: self.show_content("content4")),
            ("Voltage", lambda: self.show_content("content5")),
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
        elif content_name == "content1":
            content1 = Content1(self.content_area)
            content1.pack(fill="both", expand=True)
        elif content_name == "content2":
            content2 = Content2(self.content_area)
            content2.pack(fill="both", expand=True)
        elif content_name == "content3":
            content3 = Content3(self.content_area)
            content3.pack(fill="both", expand=True)
        elif content_name == "content4":
            content4 = Content4(self.content_area)
            content4.pack(fill="both", expand=True)
        elif content_name == "content5":
            content5 = Content5(self.content_area)
            content5.pack(fill="both", expand=True)

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

class Content1(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")

        thermometer = CircularGauge(self, min_temp=0, max_temp=100)
        thermometer.pack(pady=20)
        thermometer.set_temperature(10)
        
        # Example of adding the new gauges
        full_gauge = full_circle_gauge(self)
        full_gauge.pack(pady=10)

        semi_gauge = semi_circle_gauge(self)
        semi_gauge.pack(pady=10)

        quarter_gauge = quarter_circle_gauge(self)
        quarter_gauge.pack(pady=10)

        custom_gauge = custom_arc_gauge(self)
        custom_gauge.pack(pady=10)

class Content2(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Current.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content3(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Pressure.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content4(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Capacity.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content5(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Voltage.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

if __name__ == "__main__":
    style = Style(theme='superhero')
    app = BatteryInfo()
    app.mainloop()
