import tkinter as tk
from tkinter import filedialog, messagebox
import os
from .FinalBigSpread import verify_file, setup_logging


def generate_output_filename(input_file):
    # Generate the output file name by appending "_verified" before the file extension
    base, ext = os.path.splitext(input_file)
    return f"{base}_verified{ext}"

def select_file():
    # Open a file dialog for the user to select an Excel file
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    
    if not file_path:
        return  # User canceled the dialog
    
    # Run the verification process
    try:
        # Set up logging before verifying the file
        log_file = setup_logging(file_path)
        print(f"Log file created: {log_file}")

        # Generate output file name
        output_file = generate_output_filename(file_path)

        # Run the file verification
        verify_file(file_path, output_file)

        # Notify the user upon success
        messagebox.showinfo("Success", f"Verification complete! Verified file saved as:\n{output_file}\nLog file saved as:\n{log_file}")
    except Exception as e:
        # Show error message if something goes wrong
        messagebox.showerror("Error", f"An error occurred during verification:\n{e}")





# Set up the UI
root = tk.Tk()
root.title("Excel Validator")
root.geometry("400x200")

# Create and place UI elements
label = tk.Label(root, text="Grab your Excel file for validation", font=("Arial", 12), wraplength=300)
label.pack(pady=20)

button = tk.Button(root, text="Select File", command=select_file, font=("Arial", 12), bg="blue", fg="white")
button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
