import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image
import pytesseract
import docx

def image_to_word(input_entry, output_entry, status_label):
    """Converts an image to a Word document."""

    input_file = input_entry.get()
    output_file = output_entry.get()

    try:
        img = Image.open(input_file)
        text = pytesseract.image_to_string(img)

        doc = docx.Document()
        doc.add_paragraph(text)
        doc.save(output_file)

        status_label.config(text="Conversion successful!")
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")

def select_input_file(entry):
    """Selects input image file."""
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def select_output_file(entry):
    """Selects output Word file."""
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word documents", "*.docx")])
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def create_gui():
    """Creates the GUI."""

    root = tk.Tk()
    root.title("Image to Word Converter")

    # Input file frame
    input_frame = ttk.Frame(root)
    input_frame.pack(pady=10)

    input_label = ttk.Label(input_frame, text="Input Image:")
    input_label.pack(side="left")
    input_entry = ttk.Entry(input_frame, width=30)
    input_entry.pack(side="left")
    input_button = ttk.Button(input_frame, text="Browse", command=lambda: select_input_file(input_entry))
    input_button.pack(side="left")

    # Output file frame
    output_frame = ttk.Frame(root)
    output_frame.pack(pady=10)

    output_label = ttk.Label(output_frame, text="Output Word:")
    output_label.pack(side="left")
    output_entry = ttk.Entry(output_frame, width=30)
    output_entry.pack(side="left")
    output_button = ttk.Button(output_frame, text="Browse", command=lambda: select_output_file(output_entry))
    output_button.pack(side="left")

    # Conversion button
    convert_button = ttk.Button(root, text="Convert", command=lambda: image_to_word(input_entry, output_entry, status_label))
    convert_button.pack(pady=10)

    # Status label
    status_label = ttk.Label(root, text="")
    status_label.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()
