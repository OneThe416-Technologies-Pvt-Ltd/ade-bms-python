import tkinter as tk
from splash_screen.splash_screen import SplashScreen
from formreg.formreg import UserDetailsForm

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Desktop App")
        self.root.configure(bg="#090b33")  # Set background color for the main window

        self.show_home()

    def show_home(self):
        self.clear_frame()
        self.home_frame = tk.Frame(self.root, bg="#090b33")  # Set background color for the frame
        self.home_frame.pack(fill="both", expand=True)

        open_form_button = tk.Button(self.home_frame, text="Open Form", command=self.show_form)
        open_form_button.pack(pady=20)

    def show_form(self):
        self.clear_frame()
        self.form_frame = tk.Frame(self.root, bg="#090b33")  # Set background color for the frame
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
    root.title("Python Desktop App")
    root.configure(bg="#090b33")  # Set background color for the main window
    app = MainApp(root)
    root.mainloop()
