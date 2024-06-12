import tkinter as tk

class TransparentLabel(tk.Canvas):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(highlightthickness=0, borderwidth=0)  # Remove border and highlight

    def set_text(self, text, font=None, **kwargs):
        self.delete("all")  # Clear any existing text
        self.create_text(self.winfo_width() // 2, self.winfo_height() // 2, text=text, font=font, **kwargs)

class SplashScreen:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Welcome to: Start My Apps!")
        self.root.overrideredirect(True)  # Hides the window decorations

        # Get the screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Set the width and height of the splash screen
        splash_width = 500
        splash_height = 400  # Increased height to accommodate the image

        # Calculate the position of the splash screen window
        x_pos = screen_width - splash_width
        y_pos = (screen_height - splash_height) // 2

        # Set the geometry of the splash screen window
        self.root.geometry(f"{splash_width}x{splash_height}+{x_pos}+{y_pos}")

        # Load and display the background image
        self.background_image = tk.PhotoImage(file="D:/Projects/ADE Project/ade-bms-python/asserts/images/5623076.png")
        self.canvas = tk.Canvas(self.root, width=splash_width, height=splash_height)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(0, 0, anchor="center", image=self.background_image)

        self.root.after(5000, self.close_splash)  # Close splash screen after 5 seconds

    def close_splash(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()

