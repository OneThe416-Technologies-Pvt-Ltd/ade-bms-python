import tkinter as tk
import customtkinter as ctk
import serial.tools.list_ports
from tkinter import messagebox

class RSConnection(tk.Frame):
    rs_connected = False

    def __init__(self, master, main_window=None):
        super().__init__(master)
        self.main_window = main_window

        # Check if Moxa UPort 1650-8 is connected
        if self.is_moxa_connected():
            # If connected, create the widgets for COM port selection and connection
            self.create_widgets()
        else:
            self.create_widgets()
            # Show a messagebox if Moxa is not connected
            # messagebox.showerror("Connection Error", "Moxa UPort 1650-8 not connected.")

    def is_moxa_connected(self):
        """Check if Moxa UPort 1650-8 is connected."""
        ports = list(serial.tools.list_ports.comports())
        for port in ports:
            if 'MOXA' in port.description or '1650-8' in port.description:
                return True
        return False

    def create_widgets(self):
         # RS232/RS422 heading label
        tk.Label(self, text="RS232 / RS422 Connector", font=("Helvetica", 16, "bold")).grid(row=0, columnspan=2, pady=10)
        
        # Separator line
        separator = tk.Frame(self, height=2, bd=1, relief=tk.SUNKEN)
        separator.grid(row=1, columnspan=2, sticky="we", padx=20, pady=(0, 10))

        # Label and dropdown for available COM ports
        tk.Label(self, text="Select COM Port:").grid(row=2, column=0, padx=20, sticky=tk.W)
        self.com_ports = ctk.CTkComboBox(self, values=list(self.get_com_ports()))
        self.com_ports.set("Select a COM Port")  # Set the default value shown in the dropdown
        self.com_ports.grid(row=2, column=1, padx=20, pady=5)

        # Connect button
        self.btnConnect = ctk.CTkButton(self, text="Connect", command=self.on_connect, fg_color="green", hover_color='green')
        self.btnConnect.grid(row=3, columnspan=2, pady=10)

        # # Disconnect button (initially disabled)
        # self.btnDisconnect = ctk.CTkButton(self, text="Disconnect", command=self.on_disconnect, fg_color="red", hover_color='red', state=tk.DISABLED)
        # self.btnDisconnect.grid(row=4, columnspan=2, pady=10)

        # Center the frame within parent
        self.grid_rowconfigure(5, weight=1)  # Ensure row 5 expands to center vertically
        self.grid_columnconfigure(0, weight=1)  # Ensure column 0 expands to center horizontally
        self.pack(expand=True, fill='both')

    def get_com_ports(self):
        """Get a list of available COM ports."""
        ports = list(serial.tools.list_ports.comports())
        return [port.device for port in ports if 'MOXA' in port.description or '1650-8' in port.description]

    def on_connect(self):
        selected_port = self.com_ports.get()
        # if selected_port:
        #     # # Perform connection logic here
        #     print(f"Connecting to {selected_port}...")
        if self.main_window:
            self.main_window.show_rs_battery_info()
            # self.rs_connected = True
            # self.update_widgets()

    # def on_disconnect(self):
    #     # Perform disconnection logic here
    #     print("Disconnecting...")
    #     self.rs_connected = False
    #     self.update_widgets()

    # def update_widgets(self):
    #     if self.rs_connected:
    #         self.btnConnect.configure(state=tk.DISABLED)
    #         self.btnDisconnect.configure(state=tk.NORMAL)
    #         self.com_ports.configure(state=tk.DISABLED)
    #     else:
    #         self.btnConnect.configure(state=tk.NORMAL)
    #         self.btnDisconnect.configure(state=tk.DISABLED)
    #         self.com_ports.configure(state=tk.NORMAL)
    
    # def get_connection_status(self):
    #     return RSConnection.rs_connected 

if __name__ == "__main__":
    root = tk.Tk()
    app = RSConnection(root)
    app.pack()

    root.mainloop()
