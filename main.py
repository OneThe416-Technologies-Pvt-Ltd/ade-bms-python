import tkinter as tk

def on_submit():
    pass

app = tk.Tk()
app.title("Python Desktop App")
app.geometry("400x300")

label = tk.Label(app, text="Enter your name:")
label.pack()

entry = tk.Entry(app)
entry.pack()

submit_button = tk.Button(app, text="Submit", command=on_submit)
submit_button.pack()

app.mainloop()