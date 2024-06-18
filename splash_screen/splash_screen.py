import tkinter as tk
from PIL import Image, ImageTk
import os

class SplashScreen:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Welcome to: Battery Management System")
        self.root.overrideredirect(True)  # Hides the window decorations

        # Get the screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Set the width and height of the splash screen
        splash_width = 500
        splash_height = 400

        # Calculate the position of the splash screen window
        x_pos = (screen_width - splash_width) // 2
        y_pos = (screen_height - splash_height) // 2

        # Set the geometry of the splash screen window
        self.root.geometry(f"{splash_width}x{splash_height}+{x_pos}+{y_pos}")

        # Define the base path for your images
        base_path = os.path.join(os.path.dirname(__file__), "../asserts", "images")
        base_path2 = os.path.join(os.path.dirname(__file__), "../asserts")

        # Load and display the background image
        splash_image_path = os.path.join(base_path, "splash_bg.png")
        self.background_image = ImageTk.PhotoImage(Image.open(splash_image_path))
        self.canvas = tk.Canvas(self.root, width=splash_width, height=splash_height)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(splash_width // 2, splash_height // 2, anchor="center", image=self.background_image)

        # Display "BATTERY MANAGEMENT SYSTEM" in the center
        text_font = ('Palatino Linotype', int(24 * 0.75), 'bold')  # 25% smaller font size
        text_color = 'white'
        text_width = self.canvas.create_text(splash_width // 2, splash_height // 2 - 20, text="BATTERY MANAGEMENT SYSTEM", font=text_font, fill=text_color)

        # Adjust text position if it exceeds canvas width
        text_bbox = self.canvas.bbox(text_width)
        if text_bbox[2] > splash_width:
            self.canvas.coords(text_width, splash_width // 2, splash_height // 2 - 20)

        # Load and display your logo at the bottom (adjust path accordingly)
        logo_image_path = os.path.join(base_path2, "logo", "ade_logo.png")
        self.logo_image = ImageTk.PhotoImage(Image.open(logo_image_path))
        logo_width = self.logo_image.width()
        logo_height = self.logo_image.height()
        self.canvas.create_image(splash_width // 2, splash_height - logo_height // 2 - 20, anchor="center", image=self.logo_image)

        self.root.after(5000, self.close_splash)  # Close splash screen after 5 seconds

    def close_splash(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()

# Example usage
if __name__ == "__main__":
    splash_screen = SplashScreen()
    splash_screen.run()
