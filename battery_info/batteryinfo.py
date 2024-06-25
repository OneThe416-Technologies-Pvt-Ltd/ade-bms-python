import tkinter as tk
import customtkinter as ctk

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Side Menu")
        self.geometry("800x600")
        self.configure(bg="#f0f0f0")
        self.panel_expanded = True

        self.main_container = ctk.CTkFrame(self, fg_color="#e0e0e0")
        self.main_container.pack(fill="both", expand=True)

        self.side_menu = ctk.CTkFrame(self.main_container, corner_radius=10, fg_color="#e0e0e0")
        self.side_menu.pack(side="left", fill="y", padx=30, pady=20)

        self.separator = ctk.CTkFrame(self.main_container, width=2, fg_color="#d0d0d0")
        self.separator.pack(side="left", fill="y", padx=5)

        self.content_area = ctk.CTkFrame(self.main_container, corner_radius=10, fg_color="#ffffff")
        self.content_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.create_side_menu_options()
        self.create_initial_content()

        self.toggle_button = ctk.CTkButton(self, text="Collapse", command=self.toggle_side_menu, corner_radius=10, fg_color="#007acc", hover_color="#005fa3")
        self.toggle_button.pack(side="top", pady=10)

    def create_side_menu_options(self):
        options = [
            ("Temperature", lambda: self.show_content("content1")),
            ("Current", lambda: self.show_content("content2")),
            ("Pressure", lambda: self.show_content("content3")),
            ("Capacity", lambda: self.show_content("content4")),
            ("Voltage", lambda: self.show_content("content5")),
            ("Device Info", lambda: self.show_content("content6")),
        ]
        for text, command in options:
            button = self.create_custom_button(text, command)
            button.pack(pady=10)

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

        # Create a frame to hold the labels
        info_frame = ctk.CTkFrame(self.content_frame, corner_radius=10, fg_color="#f8f8f8", border_color="#d0d0d0", border_width=2)
        info_frame.pack(padx=60, pady=60, anchor="n")

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
            label.pack(padx=10,pady=10, anchor="w")

    def show_content(self, content_name):
        for widget in self.content_area.winfo_children():
            widget.destroy()

        if content_name == "content1":
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
        elif content_name == "content6":
            content6 = Content6(self.content_area)
            content6.pack(fill="both", expand=True)

    def create_custom_button(self, text, command):
        button = ctk.CTkButton(self.side_menu, text=text, command=command, corner_radius=10, fg_color="#007acc", hover_color="#005fa3")
        return button

    def toggle_side_menu(self):
        if self.panel_expanded:
            self.side_menu.pack_forget()
            self.separator.pack_forget()
            self.toggle_button.configure(text="Expand")
            self.panel_expanded = False
        else:
            self.side_menu.pack(side="left", fill="y", padx=10, pady=10)
            self.separator.pack(side="left", fill="y", padx=5)
            self.toggle_button.configure(text="Collapse")
            self.panel_expanded = True

class Content1(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 1.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content2(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 2.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content3(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 3.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content4(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 4.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content5(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 5.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

class Content6(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, corner_radius=10, fg_color="#ffffff")
        label = ctk.CTkLabel(self, text="This is the content for Option 6.", font=("Palatino Linotype", 14), text_color="#333333")
        label.pack(pady=20)

if __name__ == "__main__":
    ctk.set_appearance_mode("light")  # Modes: system (default), light, dark
    ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
    app = App()
    app.mainloop()
