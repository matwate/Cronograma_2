import os
import subprocess
import customtkinter as tk
from main import createCronogram
# Create the main window
window = tk.CTk()
def on_button_click():
    # Get the values from the name inputs and store them in an array
    names = [name_input.get() for name_input in name_inputs]

    # Get the value from the year input and convert it to an integer
    year = int(year_input.get())

    # Get the value from the checkbox
    is_leap_year = leap_var.get()
    
    createCronogram(year, names, is_leap_year)
    
    subprocess.Popen(f'explorer "{os.path.realpath("./excel/result.xlsx")}"')
# Create the text inputs for names
name_inputs = []
for i in range(5):
    name_label = tk.CTkLabel(window, text=f"Nombre {i+1}:")
    name_label.grid(row=i, column=0, padx=10, pady=10)  # Add padding and place in column 0
    name_input = tk.CTkEntry(window)
    name_input.grid(row=i, column=1, padx=10, pady=10)  # Add padding and place in column 1
    name_inputs.append(name_input)

# Create the text input for the year
year_label = tk.CTkLabel(window, text="Año:")
year_label.grid(row=5, column=0, padx=10, pady=10)
year_input = tk.CTkEntry(window)
year_input.grid(row=5, column=1, padx=10, pady=10)

# Create the checkmark for setting the year as leap or not
leap_var = tk.BooleanVar(value=True)
leap_checkbutton = tk.CTkCheckBox(window, text="Es Año Bisiesto?", variable=leap_var)
leap_checkbutton.grid(row=6, column=0, padx=10, pady=10)

# Create a button
button = tk.CTkButton(window, text="Generar", command=on_button_click )
button.grid(row=6, column=1, padx=10, pady=10)




# Start the main event loop
window.mainloop()
