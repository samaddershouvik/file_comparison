import os
from fileFunctions import compare_with_mother_file
from tkinter import Tk, Label, Button, filedialog, messagebox


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


def start_comparison():
    mother_file = mother_file_label.cget("text")
    comparison_folder = comparison_folder_label.cget("text")

    if not os.path.isfile(mother_file):
        messagebox.showerror("Error", "Please select a valid mother file.")
        return

    if not os.path.isdir(comparison_folder):
        messagebox.showerror("Error", "Please select a valid comparison folder.")
        return

    # Notify the user that processing has started
    status_label.config(text="Processing...", fg="blue")
    root.update_idletasks()  # Ensures the GUI updates immediately

    try:
        # Perform the comparison process
        compare_with_mother_file(mother_file, comparison_folder)
        status_label.config(text="Comparison Complete!", fg="green")
    except Exception as e:
        status_label.config(text="Error occurred during comparison.", fg="red")
        messagebox.showerror("Error", f"An error occurred: {e}")


if __name__ == "__main__":
    root = Tk()
    root.title("File Comparison Tool")

    # GUI Components
    Label(root, text="Select Base File:").grid(
        row=0, column=0, padx=10, pady=5, sticky="w"
    )
    mother_file_label = Label(
        root, text="No file selected", width=60, anchor="w", bg="white", relief="solid"
    )
    mother_file_label.grid(row=0, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=browse_mother_file).grid(
        row=0, column=2, padx=10, pady=5
    )

    Label(root, text="Select Comparison Folder:").grid(
        row=1, column=0, padx=10, pady=5, sticky="w"
    )
    comparison_folder_label = Label(
        root,
        text="No folder selected",
        width=60,
        anchor="w",
        bg="white",
        relief="solid",
    )
    comparison_folder_label.grid(row=1, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=browse_comparison_folder).grid(
        row=1, column=2, padx=10, pady=5
    )

    Button(root, text="Start Comparison", command=start_comparison).grid(
        row=2, column=1, pady=20
    )

    # Status Label for Progress Updates
    status_label = Label(root, text="", fg="black", anchor="center")
    status_label.grid(row=3, column=0, columnspan=3, pady=10)

    root.mainloop()
