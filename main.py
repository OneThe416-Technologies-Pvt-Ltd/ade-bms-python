
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from splash_screen.splash_screen import SplashScreen
import os
import customtkinter
from dashboard.dashboard import Dashboard, RSFrame

customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue")

class CustomImageButton(tk.Frame):
    def __init__(self, parent, image_path, width, height, command=None, **kwargs):
        super().__init__(parent, bg="#ffffff")  # White background frame
        self.command = command
        
        self.load_image(image_path, width, height)
        
        self.button = ttk.Button(self, command=self.command, image=self.photo, style='Primary.TButton', compound='center')
        self.button.image = self.photo  # Keep a reference to avoid garbage collection
        self.button.pack(expand=True)

    def load_image(self, image_path, width, height):
        image = Image.open(image_path)
        image = image.resize((width, height), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(image)

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("ADE-BATTERY MANAGEMENT")
        self.geometry(f"{1100}x580")

        # create main frame
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self.main_frame, width=140, corner_radius=0)
        self.sidebar_frame.pack(side="left", fill="y")

        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="BMS-ADE", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.pack(padx=20, pady=(20, 10))

        welcome_message = """\
        Welcome to the ADE Battery Management System!
        Our state-of-the-art system ensures optimal 
        performance and safety for your battery operations.
        Monitor, control, and maintain your battery 
        systems efficiently and effectively.
        Navigate through CAN and RS232 interfaces with ease.
        Enhance battery performance and longevity 
        at your fingertips.
        """

        self.welcome_label = customtkinter.CTkLabel(self.sidebar_frame, text=welcome_message, font=customtkinter.CTkFont(size=12))
        self.welcome_label.pack(padx=20, pady=(10, 20), anchor="w")

        # create slider and progressbar frame
        self.slider_progressbar_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        self.slider_progressbar_frame.pack(side="top", fill="both", expand=True)
        '''
        # Option menu for appearance mode
        self.appearance_mode_optionmenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionmenu.pack(padx=20, pady=(0, 20), side="bottom", anchor="center")

        self.appearance_mode_optionmenu.set("Light")

        # "Appearance Mode" label
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:")
        self.appearance_mode_label.pack(padx=20, pady=(10, 20), side="bottom", anchor="center")
        '''

        # Define the base path for your images
        base_path = os.path.join(os.path.dirname(__file__), "asserts", "images")

        # Load images for buttons
        can_image_path = os.path.join(base_path, "canmy.png")
        rs_image_path = os.path.join(base_path, "rsmy.png")

        # Create CAN button
        self.button_can_right = CustomImageButton(self.slider_progressbar_frame, can_image_path, width=400, height=150, command=self.show_can)
        self.button_can_right.pack(side="top",anchor="center", fill="both", expand=True)

        # Create RS button
        self.button_rs_right = CustomImageButton(self.slider_progressbar_frame, rs_image_path, width=400, height=150, command=self.show_rs)
        self.button_rs_right.pack(side="top",anchor="center", fill="both", expand=True)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def show_can(self):
        # Hide the main frame and show the dashboard frame
        self.main_frame.pack_forget()
        self.dashboard_frame = Dashboard(self, self)
        self.dashboard_frame.pack(fill="both", expand=True)

    def show_rs(self):
        # Hide the main frame and show the RS frame
        self.main_frame.pack_forget()
        self.rs_frame = RSFrame(self, self)
        self.rs_frame.pack(fill="both", expand=True)

if __name__ == "__main__":
    # Create and run the splash screen
    splash = SplashScreen()
    splash.run()

    app = App()
    app.mainloop()
