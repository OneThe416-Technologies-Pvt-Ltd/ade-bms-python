import tkinter as tk
from tkinter import messagebox  # Import messagebox to show error messages
from PIL import Image, ImageTk
import os

# Define base directory for assets
base_dir = os.path.dirname(__file__)

class SplashScreen(tk.Toplevel):
    def __init__(self, master=None):
        """Initialize the SplashScreen, set up window properties, background, and text."""
        try:
            super().__init__(master)
            self.title("Welcome to: Battery Management System")
            self.overrideredirect(True)  # Hide window decorations for a clean splash screen

            # Get the screen width and height to center the splash screen
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()

            # Set desired splash screen dimensions
            splash_width = 500
            splash_height = 400

            # Calculate the position to center the splash screen on the screen
            x_pos = (screen_width - splash_width) // 2
            y_pos = (screen_height - splash_height) // 2

            # Set the geometry of the splash screen window
            self.geometry(f"{splash_width}x{splash_height}+{x_pos}+{y_pos}")

            # Load background image for splash screen
            background_image_path = os.path.join(base_dir, '../assets/images/splash_bg.png')
            self.background_image = ImageTk.PhotoImage(Image.open(background_image_path))
            
            # Create a canvas to hold the background image
            self.canvas = tk.Canvas(self, width=splash_width, height=splash_height)
            self.canvas.pack(fill="both", expand=True)
            self.canvas.create_image(splash_width // 2, splash_height // 2, anchor="center", image=self.background_image)

            # Display the text "BATTERY MANAGEMENT SYSTEM" in the center
            text_font = ('Palatino Linotype', int(24 * 0.75), 'bold')  # 25% smaller font size
            text_color = 'white'
            text_width = self.canvas.create_text(splash_width // 2, splash_height // 2 - 20, text="BATTERY MANAGEMENT SYSTEM", font=text_font, fill=text_color)

            # Check if the text exceeds canvas width and adjust position if necessary
            text_bbox = self.canvas.bbox(text_width)
            if text_bbox[2] > splash_width:
                self.canvas.coords(text_width, splash_width // 2, splash_height // 2 - 20)

            # Load and display logo at the bottom of the splash screen
            logo_image_path = os.path.join(base_dir, '../assets/logo/ade_logo.png')
            self.logo_image = ImageTk.PhotoImage(Image.open(logo_image_path))
            logo_width = self.logo_image.width()
            logo_height = self.logo_image.height()
            self.canvas.create_image(splash_width // 2, splash_height - logo_height // 2 - 20, anchor="center", image=self.logo_image)

            # Set the splash screen to close after 5 seconds
            self.after(5000, self.close_splash)  # Close splash screen after 5 seconds

        except Exception as e:
            # Handle any exceptions that occur during initialization
            print(f"Error initializing splash screen: {e}")
            messagebox.showerror("Error", f"Failed to load splash screen: {e}")

    def close_splash(self):
        """Close the splash screen after the timeout."""
        try:
            self.destroy()  # Close the splash screen
        except Exception as e:
            print(f"Error closing splash screen: {e}")
            messagebox.showerror("Error", f"Failed to close splash screen: {e}")

    def run(self):
        """Start the main event loop for the splash screen."""
        try:
            self.mainloop()  # Start the Tkinter event loop
        except Exception as e:
            # Handle any exceptions during the main event loop
            print(f"Error in mainloop: {e}")
            messagebox.showerror("Error", f"Error running the splash screen: {e}")
