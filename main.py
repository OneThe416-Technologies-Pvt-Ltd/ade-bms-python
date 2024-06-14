import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from splash_screen.splash_screen import SplashScreen
from formreg.formreg import UserDetailsForm
from dashboard.dashboard import create_dashboard, load_can, load_rs
import os

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Desktop App - ADE Defense")
        self.root.configure(bg="#267dff")  # Dark blue background for the main window

        self.setup_styles()  # Setup custom styles

        self.show_home()

    def setup_styles(self):
        # Create custom style for primary buttons (CAN and RS buttons)
        style = ttk.Style()
        style.configure('Primary.TButton', font=('Arial', 15, 'bold'), foreground="white", background="#ffffff", borderwidth=0)
        style.map('Primary.TButton', bordercolor=[('pressed', '#000000'), ('active', '#000000')])

    def show_home(self):
        self.clear_frame()
        self.home_frame = tk.Frame(self.root, bg="#f0f0f0")  # Light grey background for the frame
        self.home_frame.pack(fill="both", expand=True)

        # Set the grid layout for the home_frame
        self.home_frame.grid_columnconfigure(0, weight=1)
        self.home_frame.grid_columnconfigure(1, weight=1)
        self.home_frame.grid_rowconfigure(0, weight=1)

        # Left frame for welcome message and buttons (dark blue)
        left_frame = tk.Frame(self.home_frame, bg="#267dff")  # Dark blue background
        left_frame.grid(row=0, column=0, sticky="nsew")

        # Right frame for CAN and RS buttons (white)
        right_frame = tk.Frame(self.home_frame, bg="#ffffff")  # White background
        right_frame.grid(row=0, column=1, sticky="nsew")

        # Add welcome message to the left frame
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
        welcome_label = tk.Label(
            left_frame, text=welcome_message, bg="#267dff", fg="white", 
            font=("Palatino Linotype", 14, "bold"), justify=tk.LEFT, padx=20, pady=20
        )
        welcome_label.pack(expand=True)

        # Define the base path for your images
        base_path = os.path.join(os.path.dirname(__file__), "images2")

        # Load and resize images for buttons
        can_image_path = os.path.join(base_path, "canmy.png")
        rs_image_path = os.path.join(base_path, "rsmy.png")
        
        can_image = Image.open(can_image_path)  # Replace with your CAN image path
        rs_image = Image.open(rs_image_path)   # Replace with your RS image path
        button_size = (250, 100)  # Adjusted size
        can_image = can_image.resize(button_size, Image.LANCZOS)
        rs_image = rs_image.resize(button_size, Image.LANCZOS)

        can_photo = ImageTk.PhotoImage(can_image)
        rs_photo = ImageTk.PhotoImage(rs_image)

        # Create a container frame to center buttons vertically
        center_frame = tk.Frame(right_frame, bg="#ffffff")  # White background
        center_frame.pack(expand=True)

        # Center the buttons vertically and horizontally
        center_frame.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center widgets
        center_frame.grid_rowconfigure((0, 1), weight=1)  # Ensure rows expand to center vertically

        # Create CAN button in the right frame
        button_can_right = ttk.Button(center_frame, command=self.show_can, image=can_photo, style='Primary.TButton')
        button_can_right.image = can_photo  # Keep a reference to avoid garbage collection
        button_can_right.grid(row=0, column=0, padx=20, pady=20)

        # Create RS button in the right frame
        button_rs_right = ttk.Button(center_frame, command=self.show_rs, image=rs_photo, style='Primary.TButton')
        button_rs_right.image = rs_photo  # Keep a reference to avoid garbage collection
        button_rs_right.grid(row=1, column=0, padx=20, pady=20)

    def show_dashboard(self):
        self.clear_frame()
        create_dashboard(self.root, self.show_home, load_can, load_rs)  # Pass functions for CAN and RS232 loading

    def show_can(self):
        self.clear_frame()
        load_can(self.root, self.show_home)

    def show_rs(self):
        self.clear_frame()
        load_rs(self.root, self.show_home)

    def show_form(self):
        self.clear_frame()
        self.form_frame = tk.Frame(self.root, bg="#f0f0f0")  # Light grey background for the frame
        self.form_frame.pack(fill="both", expand=True)

        UserDetailsForm(self.form_frame, self.show_home)  # Pass show_home as the back_command

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    # Create and run the splash screen
    splash = SplashScreen()
    splash.run()

    # After the splash screen closes, create the main app window
    root = tk.Tk()
    root.title("Python Desktop App - ADE Defense")
    root.configure(bg="#267dff")  # Dark blue background

    # Get the screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Set the geometry to fill the entire screen
    root.geometry(f"{screen_width}x{screen_height}+0+0")

    # Maximize the window
    root.state('zoomed')

    app = MainApp(root)
    root.mainloop()
