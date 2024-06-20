# dashboard.py

import customtkinter

class Dashboard(customtkinter.CTkFrame):
    def __init__(self, parent, main_app):
        super().__init__(parent,fg_color="#ffffff")

        self.main_app = main_app

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
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

        # Label to indicate CAN selection
        self.can_label = customtkinter.CTkLabel(self, text="CAN selected", font=customtkinter.CTkFont(size=20))
        self.can_label.pack(padx=20, pady=(40, 20))

        # Back button to navigate back to the main frame
        self.back_button = customtkinter.CTkButton(self, text="Back", command=self.go_back)
        self.back_button.pack(padx=20, pady=(20, 20))

    def go_back(self):
        self.pack_forget()
        self.main_app.main_frame.pack(fill="both", expand=True)

class RSFrame(customtkinter.CTkFrame):
    def __init__(self, parent, main_app):
        super().__init__(parent,fg_color="#ffffff")

        self.main_app = main_app

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
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

        # Label to indicate RS selection
        self.rs_label = customtkinter.CTkLabel(self, text="RS selected", font=customtkinter.CTkFont(size=20))
        self.rs_label.pack(padx=20, pady=(40, 20))

        # Back button to navigate back to the main frame
        self.back_button = customtkinter.CTkButton(self, text="Back", command=self.go_back)
        self.back_button.pack(padx=20, pady=(20, 20))

        
    def go_back(self):
        self.pack_forget()
        self.main_app.main_frame.pack(fill="both", expand=True)
