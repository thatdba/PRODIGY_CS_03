import tkinter as tk
import pandas as pd
from openpyxl import Workbook
from tkinter import messagebox
import re


def save_to_excel():
    # Retrieve form entries
    name = name_entry.get()
    email = email_entry.get()
    password = password_entry.get()
    password2 = password2_entry.get()

    # Check if any field is empty
    if not (name and email and password and password2):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Check if passwords match
    if password != password2:
        messagebox.showerror("Error", "Passwords do not match.")
        return

    # Check password length to be more than 8 characters
    if len(password) < 8:
        messagebox.showerror("Error", "Password must be at least 8 characters long.")
        return
    
    # Check password length to be less than or equal to 16 characters
    if len(password) > 16: 
        messagebox.showerror("Error", "length should be not be more than 16 characters long.") 
        return

    # Check for special characters
    if not re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
        messagebox.showerror("Error", "Password must contain at least one special character.")
        return

    #check if password has Uppercase letters
    if not any(char.islower() for char in password): 
        messagebox.showerror("Error", "Password should have at least one lower case letter.")
        return

      #check if password has Uppercase letters
    if not any(char.isupper() for char in password): 
        messagebox.showerror("Error", "Password should have at least one upper case letter.")
        return

    #check if password has Uppercase letters
    if not any(char.isdigit() for char in password): 
        messagebox.showerror("Error", "Password should have at least one number.")
        return

    # Create a new Workbook or load existing one
    try:
        existing_data = pd.read_excel('user_registration.xlsx')
    except FileNotFoundError:
        existing_data = pd.DataFrame(columns=['Name', 'Email', 'Password', 'Password2'])

    # Append user data to the DataFrame
    new_entry = pd.DataFrame({'Name': [name], 'Email': [email], 'Password': [password], 'Password2': [password2]})
    updated_data = pd.concat([existing_data, new_entry], ignore_index=True)

    # Save DataFrame to Excel
    try:
        updated_data.to_excel('user_registration.xlsx', index=False)
        messagebox.showinfo("Success", "User data saved to Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the GUI window
root = tk.Tk()
root.title("User Registration Form")
#provide size to window
root.geometry("450x250")

# Create form fields
tk.Label(root, text="Name:").grid(row=0, column=0, sticky="e")
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1)

tk.Label(root, text="Email:").grid(row=1, column=0, sticky="e")
email_entry = tk.Entry(root)
email_entry.grid(row=1, column=1)

tk.Label(root, text="Password:").grid(row=2, column=0, sticky="e")
password_entry = tk.Entry(root, show="*")
password_entry.grid(row=2, column=1)

tk.Label(root, text="Re-enter Password:").grid(row=3, column=0, sticky="e")
password2_entry = tk.Entry(root, show="*")
password2_entry.grid(row=3, column=1)

# Button to save form entries
save_button = tk.Button(root, text="Register", command=save_to_excel)
save_button.grid(row=4, columnspan=2)

# Run the GUI main loop
root.mainloop()
