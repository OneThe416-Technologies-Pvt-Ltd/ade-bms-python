import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk

class BMS_UI:
    def __init__(self, root):
        self.root = root
        self.root.title("Battery Management System")
        self.create_widgets()

    def create_widgets(self):
        # Create a frame for the title and tabs
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(expand=True, fill="both")

        # Title Label
        overview_label = ctk.CTkLabel(main_frame, text="Battery Management System", font=('Helvetica', 18))
        overview_label.pack(pady=20)

        # Create a frame for the tab buttons and content
        content_frame = ctk.CTkFrame(main_frame)
        content_frame.pack(expand=True, fill="both", pady=10)

        # Create tab buttons
        tab_button_frame = ctk.CTkFrame(content_frame)
        tab_button_frame.pack(side="left", fill="y", padx=10)

        can_button = ctk.CTkButton(tab_button_frame, text="CAN Connect", command=lambda: self.show_frame("CAN Connect"))
        can_button.pack(pady=10)

        rs232_button = ctk.CTkButton(tab_button_frame, text="RS232/RS422 Connect", command=lambda: self.show_frame("RS232/RS422 Connect"))
        rs232_button.pack(pady=10)

        # Create a frame for the tab content
        self.tab_content = ctk.CTkFrame(content_frame)
        self.tab_content.pack(side="left", expand=True, fill="both")

        # CAN Interface Tab
        self.frames = {}
        self.can_bg_img_original = Image.open("assets/images/can_diagram.png")
        can_tab = ctk.CTkFrame(self.tab_content)
        self.can_bg_label = tk.Label(can_tab)
        self.can_bg_label.pack(fill="both", expand=True)
        can_tab.bind("<Configure>", self.resize_can_image)
        self.frames["CAN Connect"] = can_tab

        # RS232/RS422 Interface Tab
        self.rs232_bg_img_original = Image.open("assets/images/moxa_diagram.png")
        rs232_tab = ctk.CTkFrame(self.tab_content)
        self.rs232_bg_label = tk.Label(rs232_tab)
        self.rs232_bg_label.pack(fill="both", expand=True)
        rs232_tab.bind("<Configure>", self.resize_rs232_image)
        self.frames["RS232/RS422 Connect"] = rs232_tab

        # Show the default frame
        self.show_frame("CAN Connect")

    def resize_can_image(self, event):
        width = event.width
        height = event.height
        resized_image = self.can_bg_img_original.resize((width, height), Image.ANTIALIAS)
        self.can_bg_img = ImageTk.PhotoImage(resized_image)
        self.can_bg_label.config(image=self.can_bg_img)
        self.can_bg_label.image = self.can_bg_img

    def resize_rs232_image(self, event):
        width = event.width
        height = event.height
        resized_image = self.rs232_bg_img_original.resize((width, height), Image.ANTIALIAS)
        self.rs232_bg_img = ImageTk.PhotoImage(resized_image)
        self.rs232_bg_label.config(image=self.rs232_bg_img)
        self.rs232_bg_label.image = self.rs232_bg_img

    def show_frame(self, name):
        frame = self.frames[name]
        frame.pack(expand=True, fill="both")
        frame.tkraise()

if __name__ == "__main__":
    root = ctk.CTk()
    app = BMS_UI(root)
    root.mainloop()
