import tkinter as tk
from splash_screen.splash_screen import SplashScreen

def main():
    # Create and display splash screen
    splash = SplashScreen()
    splash.run()
    
    #aravind
    # Create main application window
    app = tk.Tk()
    app.title("Python Desktop App")
    app.geometry("400x300")

    label = tk.Label(app, text="Enter your name:")
    label.pack()

    app.mainloop()

if __name__ == "__main__":
    main()
