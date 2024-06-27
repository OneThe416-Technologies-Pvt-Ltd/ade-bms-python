import tkinter as tk
from tkinter import ttk
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

# Print Matplotlib version
#print("Matplotlib Version: {}".format(matplotlib.__version__))

colors = ['#FF0000', '#FF4500', '#FFA500', '#FFFF00', '#ADFF2F', '#228B22', '#006400']
values = [100, 80, 60, 40, 20, 0, -20, -40]
x_axis_vals = [0, 0.44, 0.88, 1.32, 1.76, 2.2, 2.64]

def create_plot():
    fig = plt.figure(figsize=(8, 8))  # Adjusted size to fit better in Tkinter window
    ax = fig.add_subplot(projection="polar")

    ax.bar(
        x=x_axis_vals,
        width=0.5,
        height=0.5,
        bottom=2,
        linewidth=3,
        edgecolor="white",
        color=colors,
        align="edge"
    )

    plt.annotate("Overheated", xy=(0.16, 2.1), rotation=-75, color="white", fontweight="bold")
    plt.annotate("Hot", xy=(0.65, 2.08), rotation=-55, color="white", fontweight="bold")
    plt.annotate("Warm", xy=(1.14, 2.1), rotation=-32, color="white", fontweight="bold")
    plt.annotate("Optimal", xy=(1.62, 2.2), color="white", fontweight="bold")
    plt.annotate("Cool", xy=(2.08, 2.25), rotation=20, color="white", fontweight="bold")
    plt.annotate("Cold", xy=(2.46, 2.25), rotation=45, color="white", fontweight="bold")
    plt.annotate("Subzero", xy=(3.0, 2.25), rotation=75, color="white", fontweight="bold")

    for loc, val in zip([0, 0.44, 0.88, 1.32, 1.76, 2.2, 2.64, 3.14], values):
        plt.annotate(val, xy=(loc, 2.5), ha="right" if val <= 20 else "left")

    plt.annotate(
        "50",
        xytext=(0, 0),
        xy=(1.1, 2.0),
        arrowprops=dict(arrowstyle="wedge,tail_width=0.5", color="black", shrinkA=0),
        bbox=dict(boxstyle="circle", facecolor="black", linewidth=2.0),
        fontsize=45,
        color="white",
        ha="center"
    )

    plt.title("Temperature Gauge Chart", loc="center", pad=20, fontsize=35, fontweight="bold")
    ax.set_axis_off()

    return fig

# Initialize Tkinter
root = tk.Tk()
root.title("Matplotlib in Tkinter")

# Create a plot
fig = create_plot()

# Embed the plot in the Tkinter window
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.draw()
canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

# Add a quit button
quit_button = ttk.Button(root, text="Quit", command=root.quit)
quit_button.pack(side=tk.BOTTOM)

# Run the Tkinter main loop
root.mainloop()
