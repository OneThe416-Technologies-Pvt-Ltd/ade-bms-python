import tkinter as tk
from tkinter import ttk, messagebox

class UserDetailsForm:
    def __init__(self, master, back_command):
        self.master = master
        self.back_command = back_command

        # Add frame with grey background inside master frame
        self.border_frame = tk.Frame(master, bg='grey', highlightbackground='black', highlightthickness=1)
        self.border_frame.pack(fill="both", expand=True)  # Make the frame fill the entire master window

        self.label_personal_details = ttk.Label(self.border_frame, text="Personal Details", font=('Palatino Linotype', 16, 'bold'), foreground='black', background='grey')
        self.label_personal_details.grid(row=0, column=0, columnspan=2, pady=10, sticky='w')

        self.label_name = ttk.Label(self.border_frame, text="Name:", font=('Palatino Linotype', 12, 'bold'), foreground='black', background='grey')
        self.label_name.grid(row=1, column=0, pady=5, sticky='w')
        self.entry_name = ttk.Entry(self.border_frame, font=('Palatino Linotype', 12))
        self.entry_name.grid(row=1, column=1, pady=5, padx=10, sticky='w')

        self.label_email = ttk.Label(self.border_frame, text="Email:", font=('Palatino Linotype', 12, 'bold'), foreground='black', background='grey')
        self.label_email.grid(row=2, column=0, pady=5, sticky='w')
        self.entry_email = ttk.Entry(self.border_frame, font=('Palatino Linotype', 12))
        self.entry_email.grid(row=2, column=1, pady=5, padx=10, sticky='w')

        self.label_age = ttk.Label(self.border_frame, text="Age:", font=('Palatino Linotype', 12, 'bold'), foreground='black', background='grey')
        self.label_age.grid(row=3, column=0, pady=5, sticky='w')
        self.entry_age = ttk.Entry(self.border_frame, font=('Palatino Linotype', 12))
        self.entry_age.grid(row=3, column=1, pady=5, padx=10, sticky='w')

        self.label_gender = ttk.Label(self.border_frame, text="Gender:", font=('Palatino Linotype', 12, 'bold'), foreground='black', background='grey')
        self.label_gender.grid(row=4, column=0, pady=5, sticky='w')
        self.gender_var = tk.StringVar()
        self.gender_combobox = ttk.Combobox(self.border_frame, textvariable=self.gender_var, values=["Male", "Female", "Other"], font=('Palatino Linotype', 12), state='readonly')
        self.gender_combobox.grid(row=4, column=1, pady=5, padx=10, sticky='w')

        self.label_country = ttk.Label(self.border_frame, text="Country:", font=('Palatino Linotype', 12, 'bold'), foreground='black', background='grey')
        self.label_country.grid(row=5, column=0, pady=5, sticky='w')
        self.country_var = tk.StringVar()
        self.country_combobox = ttk.Combobox(self.border_frame, textvariable=self.country_var, values=["USA", "UK", "Canada", "Australia", "Other"], font=('Palatino Linotype', 12), state='readonly')
        self.country_combobox.grid(row=5, column=1, pady=5, padx=10, sticky='w')

        self.btn_submit = ttk.Button(self.border_frame, text="Submit", command=self.submit_form, style='Submit.TButton')
        self.btn_submit.grid(row=6, column=0, columnspan=2, pady=20)

        self.back_button = ttk.Button(self.border_frame, text="Back to Home", command=self.back_command)
        self.back_button.grid(row=7, column=0, columnspan=2, pady=10)

    def submit_form(self):
        name = self.entry_name.get()
        email = self.entry_email.get()
        age = self.entry_age.get()
        gender = self.gender_var.get()
        country = self.country_var.get()

        messagebox.showinfo("Form Submitted", f"Name: {name}\nEmail: {email}\nAge: {age}\nGender: {gender}\nCountry: {country}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UserDetailsForm(root, lambda: print("Back button pressed"))
    root.mainloop()
