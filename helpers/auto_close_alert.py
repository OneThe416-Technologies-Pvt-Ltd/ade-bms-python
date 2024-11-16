import tkinter as tk
from tkinter import ttk, messagebox  # Import messagebox from tkinter

class AutoCloseMessageBox:
    def __init__(self, root, title="Message", message="This is a message", timeout=10, message_type="info"):
        """
        Initialize the AutoCloseMessageBox with a title, message, auto-close timeout, and message type.
        
        Parameters:
            root (tk.Tk or tk.Toplevel): The root window or parent window for the message box.
            title (str): Title of the message box.
            message (str): Message text to display in the message box.
            timeout (int): Time in seconds before the message box auto-closes.
            message_type (str): Type of message ("info", "warning", "error") for color styling.
        """
        try:
            self.root = root
            self.title = title
            self.message = message
            self.timeout = timeout
            self.message_type = message_type

            self.create_message_box()
        except Exception as e:
            print(f"Error initializing AutoCloseMessageBox: {e}")

    def create_message_box(self):
        """Creates a custom message box that closes automatically after the specified timeout and is centered on the screen."""
        try:
            # Create a top-level window for the message box
            self.message_box = tk.Toplevel(self.root)
            self.message_box.title(self.title)
            self.message_box.geometry("300x100")
            self.message_box.resizable(False, False)

            # Set background and text colors based on message type
            if self.message_type == "info":
                bg_color = "#DFF0D8"  # Light green for info
                fg_color = "#3C763D"  # Dark green text
            elif self.message_type == "warning":
                bg_color = "#FCF8E3"  # Light yellow for warning
                fg_color = "#8A6D3B"  # Dark yellow/brown text
            elif self.message_type == "error":
                bg_color = "#F2DEDE"  # Light red for error
                fg_color = "#A94442"  # Dark red text
            else:
                bg_color = "#FFFFFF"  # Default to white
                fg_color = "#000000"  # Black text

            self.message_box.configure(bg=bg_color)

            # Center the message box on the main window
            self.root.update_idletasks()  # Ensure all window sizes are calculated
            window_width = 300
            window_height = 100
            screen_width = self.root.winfo_width()
            screen_height = self.root.winfo_height()
            screen_x = self.root.winfo_x()
            screen_y = self.root.winfo_y()
            position_x = screen_x + (screen_width // 2) - (window_width // 2)
            position_y = screen_y + (screen_height // 2) - (window_height // 2)
            self.message_box.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

            # Label to display the message
            label = ttk.Label(self.message_box, text=self.message, font=("Helvetica", 10, "bold"), wraplength=250, foreground=fg_color, background=bg_color)
            label.pack(pady=20)

            # Automatically close the message box after the specified timeout
            self.message_box.after(self.timeout * 1000, self.message_box.destroy)

            # Optional close button for manual dismissal
            close_button = ttk.Button(self.message_box, text="Close", command=self.message_box.destroy)
            close_button.pack(pady=5)

        except Exception as e:
            print(f"Error creating message box: {e}")
            messagebox.showerror("Error", "An error occurred while creating the message box.")

# Example usage:
# try:
#     AutoCloseMessageBox(root, title="Info", message="This is an informational message.", timeout=10, message_type="info")
#     AutoCloseMessageBox(root, title="Warning", message="This is a warning message.", timeout=10, message_type="warning")
#     AutoCloseMessageBox(root, title="Error", message="This is an error message.", timeout=10, message_type="error")
# except Exception as e:
#     print(f"Error during message box creation: {e}")
