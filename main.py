import os
from fileFunctions import compare_with_mother_file
from tkinter import Tk, Label, Button, filedialog, messagebox, font
from tkinter import ttk
import time
import threading
import subprocess
import platform


def browse_mother_file():
    file_path = filedialog.askopenfilename(
        title="Select Mother File",
        filetypes=[("Supported Files", "*.pdf;*.docx;*.txt")],
    )
    if file_path:
        mother_file_label.config(text=file_path)


def browse_comparison_folder():
    folder_path = filedialog.askdirectory(title="Select Comparison Folder")
    if folder_path:
        comparison_folder_label.config(text=folder_path)

def animate_progress_bar():
    for i in range(1, 95, 5):  # Move from 0 to 95 slowly
        progress_bar["value"] = i
        root.update_idletasks()
        time.sleep(1)  # Adjust speed as needed

    # Wait for the comparison to finish before finalizing the bar
    while progress_bar["value"] < 100:
        time.sleep(0.2)

    progress_bar.grid_remove()


def open_excels_folder():
    mother_file = mother_file_label.cget("text")
    mother_folder = os.path.dirname(mother_file)
    reports_folder = os.path.join(mother_folder, "Comparison_Reports")

    if os.path.exists(reports_folder) and os.path.isdir(reports_folder):
        os.startfile(reports_folder)  # For Windows
    else:
        messagebox.showerror("Error", "Comparison_Reports folder not found.")


def run_comparison(mother_file, comparison_folder):
    try:
        # Perform the comparison process
        compare_with_mother_file(mother_file, comparison_folder)

        # Finalize the progress bar
        progress_bar["value"] = 100
        root.update_idletasks()
        time.sleep(0.5)  # Small pause for completion effect

        # Clear the progress bar and show success message
        progress_bar.grid_remove()
        status_label.config(text="Comparison Complete!", fg="green")
        messagebox.showinfo("Success", "Comparison process is complete!")

        # Activate the button to open the Excels folder
        open_excels_button.config(state="normal")

    except Exception as e:
        progress_bar.grid_remove()
        status_label.config(text="Error occurred during comparison.", fg="red")
        messagebox.showerror("Error", f"An error occurred: {e}")



def start_comparison():
    mother_file = mother_file_label.cget("text")
    comparison_folder = comparison_folder_label.cget("text")

    if not os.path.isfile(mother_file):
        messagebox.showerror("Error", "Please select a valid Base file.")
        return

    if not os.path.isdir(comparison_folder):
        messagebox.showerror("Error", "Please select a valid Comparison folder.")
        return

    # Reset and show progress bar
    status_label.config(text="Processing...", fg="blue")
    progress_bar.grid()
    progress_bar["value"] = 0
    root.update_idletasks()

    # Start progress bar animation in a separate thread
    threading.Thread(target=animate_progress_bar).start()

    # Start file comparison in a separate thread
    threading.Thread(target=run_comparison, args=(mother_file, comparison_folder)).start()



if __name__ == "__main__":
    root = Tk()
    root.geometry("1000x500")
    root.title("File Comparison Tool")
    root.config(bg="#f0f0f0")  # Light grey background

    # Set default font size and style for all widgets
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(size=12)
    root.option_add("*Font", default_font)
    root.option_add("*Button.Font", "Helvetica 10 bold")

    # Title Label
    Label(
        root,
        text="File Comparison Tool",
        font=("Helvetica", 18, "bold"),
        bg="#f0f0f0"
    ).grid(row=0, column=0, columnspan=3, pady=20)

    # Base File Selection
    Label(root, text="Select Base File:", bg="#f0f0f0").grid(
        row=1, column=0, padx=20, pady=10, sticky="w"
    )
    mother_file_label = Label(
        root,
        text="No file selected",
        width=60,
        anchor="w",
        bg="white",
        relief="groove",
        borderwidth=2
    )
    mother_file_label.grid(row=1, column=1, padx=10, pady=10)
    Button(root, text="Browse", command=browse_mother_file, width=15).grid(
        row=1, column=2, padx=10, pady=10
    )

    # Comparison Folder Selection
    Label(root, text="Select Comparison Folder:", bg="#f0f0f0").grid(
        row=2, column=0, padx=20, pady=10, sticky="w"
    )
    comparison_folder_label = Label(
        root,
        text="No folder selected",
        width=60,
        anchor="w",
        bg="white",
        relief="groove",
        borderwidth=2
    )
    comparison_folder_label.grid(row=2, column=1, padx=10, pady=10)
    Button(root, text="Browse", command=browse_comparison_folder, width=15).grid(
        row=2, column=2, padx=10, pady=10
    )

    # Start Comparison Button
    Button(
        root,
        text="Start Comparison",
        command=start_comparison,
        bg="#4CAF50",
        fg="white",
        activebackground="#45a049",
        relief="raised",
        width=20
    ).grid(row=3, column=1, pady=20)

    # Open Excels Folder Button (Initially Disabled)
    open_excels_button = Button(
        root,
        text="Open Comparison Reports",
        state="disabled",
        command=open_excels_folder,
        width=25,
        bg="#2196F3",
        fg="white",
        activebackground="#1e88e5"
    )
    open_excels_button.grid(row=4, column=1, pady=10)

    # Progress Bar
    progress_bar = ttk.Progressbar(root, mode='determinate', length=500)
    progress_bar.grid(row=5, column=0, columnspan=3, pady=20)
    progress_bar["maximum"] = 100
    progress_bar["value"] = 0
    progress_bar.grid_remove()  # Initially hidden

    # Status Label
    status_label = Label(root, text="", fg="blue", bg="#f0f0f0", font=("Helvetica", 12, "italic"))
    status_label.grid(row=6, column=0, columnspan=3, pady=10)

    root.mainloop()

