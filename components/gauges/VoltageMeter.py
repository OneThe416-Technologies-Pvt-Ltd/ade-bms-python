import tkinter as tk
from tkinter import ttk
import math

class ModernVoltageMeterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Modern Voltage Meter")
        
        self.create_widgets()
    
    def create_widgets(self):
        style = ttk.Style()
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white', font=('Arial', 14))
        style.configure('TButton', font=('Arial', 14))

        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=20, pady=20, fill='both', expand=True)

        meter_frame = ttk.Frame(main_frame)
        meter_frame.pack(padx=10, pady=10, fill='both', expand=True)

        self.canvas = tk.Canvas(meter_frame, width=400, height=400, bg='white')
        self.canvas.pack(padx=10, pady=10)
        
        self.create_meter()
        
        entry_frame = ttk.Frame(main_frame)
        entry_frame.pack(padx=10, pady=10, fill='both', expand=True)

        entry_label = ttk.Label(entry_frame, text="Enter Voltage:")
        entry_label.pack(side='left', padx=5)

        self.entry = ttk.Entry(entry_frame, width=10, font=('Arial', 14))
        self.entry.pack(side='left', padx=5)

        update_button = ttk.Button(entry_frame, text="Update Meter", command=self.update_meter)
        update_button.pack(side='left', padx=5)

        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(padx=10, pady=10)

        # Label to display the voltage value
        self.value_label = ttk.Label(meter_frame, text="", font=('Arial', 20))
        self.value_label.place(x=200, y=200, anchor='center')

    def create_meter(self):
        # Create the background arc for the meter
        self.canvas.create_arc(50, 50, 350, 350, start=0, extent=180, style=tk.ARC, outline="black", width=2)
        
        # Create the tick marks and labels
        for i in range(11):
            angle = 180 - (i * 18)
            x_start = 200 + 150 * math.cos(math.radians(angle))
            y_start = 200 - 150 * math.sin(math.radians(angle))
            x_end = 200 + 140 * math.cos(math.radians(angle))
            y_end = 200 - 140 * math.sin(math.radians(angle))
            self.canvas.create_line(x_start, y_start, x_end, y_end, width=2)
            label_x = 200 + 120 * math.cos(math.radians(angle))
            label_y = 200 - 120 * math.sin(math.radians(angle))
            self.canvas.create_text(label_x, label_y, text=f"{i * 10}", font=('Arial', 10))
        
        # Create the arrow
        self.arrow = self.canvas.create_line(200, 200, 200, 100, arrow=tk.LAST, fill='red', width=4)
    
    def update_meter(self):
        try:
            voltage_value = float(self.entry.get())
            if 0 <= voltage_value <= 100:
                angle = 180 - (voltage_value * 1.8)
                x_end = 200 + 150 * math.cos(math.radians(angle))
                y_end = 200 - 150 * math.sin(math.radians(angle))
                self.canvas.coords(self.arrow, 200, 200, x_end, y_end)
                self.value_label.config(text=f"{voltage_value:.1f} V")
                self.status_label.config(text="Voltage updated successfully", foreground="green")
            else:
                self.status_label.config(text="Enter a value between 0 and 100", foreground="red")
        except ValueError:
            self.status_label.config(text="Invalid input. Please enter a valid number.", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernVoltageMeterApp(root)
    root.mainloop()
