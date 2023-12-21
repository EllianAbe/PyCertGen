import tkinter as tk
from tkinter import filedialog, messagebox
from . import generator


CERTIFIED_LIST_ENTRY = None
TEMPLATE_ENTRY = None
OUTPUT_ENTRY = None
TO_PDF_METHOD = None
KEY_AUTOGEN = None


def submit():
    certified_list_path = CERTIFIED_LIST_ENTRY.get()
    template_path = TEMPLATE_ENTRY.get()
    output_path = OUTPUT_ENTRY.get()
    key_autogen = KEY_AUTOGEN.get()
    to_pdf_method = TO_PDF_METHOD.get()

    generator.gen(certified_list_path, template_path, output_path,
                  key_autogen, to_pdf_method)

    messagebox.showinfo("Process Finished",
                        "The process finished with success!")


def certified_list_manager(frame):
    global CERTIFIED_LIST_ENTRY

    certified_list_label = tk.Label(frame, text="Certified List Path:")
    certified_list_label.grid(row=0, column=0, sticky='W', pady=2)
    CERTIFIED_LIST_ENTRY = tk.Entry(frame)
    CERTIFIED_LIST_ENTRY.grid(row=0, column=1, pady=2)

    def browse_certified_list():
        filename = filedialog.askopenfilename(
            filetypes=[('Excel Files', '*.xls*')])
        CERTIFIED_LIST_ENTRY.insert(tk.END, filename)

    certified_list_browse_button = tk.Button(
        frame, text="Browse", command=browse_certified_list)
    certified_list_browse_button.grid(row=0, column=2, pady=2)


def template_manager(frame):
    global TEMPLATE_ENTRY

    template_label = tk.Label(frame, text="Template Path:")
    template_label.grid(row=1, column=0, sticky='W', pady=2)
    TEMPLATE_ENTRY = tk.Entry(frame)
    TEMPLATE_ENTRY.grid(row=1, column=1, pady=2)

    def browse_template():
        filename = filedialog.askopenfilename(
            filetypes=[('Powerpoint Files', '*.ppt*')])
        TEMPLATE_ENTRY.insert(tk.END, filename)

    template_browse_button = tk.Button(
        frame, text="Browse", command=browse_template)
    template_browse_button.grid(row=1, column=2, pady=2)


def output_manager(frame):
    global OUTPUT_ENTRY

    output_label = tk.Label(frame, text="Output directory:")
    output_label.grid(row=2, column=0, sticky='W', pady=2)
    OUTPUT_ENTRY = tk.Entry(frame)
    OUTPUT_ENTRY.grid(row=2, column=1, pady=2)

    def browse_template():
        dirname = filedialog.askdirectory()
        OUTPUT_ENTRY.insert(tk.END, dirname)

    output_browse_button = tk.Button(
        frame, text="Browse", command=browse_template)
    output_browse_button.grid(row=2, column=2, pady=2)


def autogen_manager(frame):
    global KEY_AUTOGEN

    KEY_AUTOGEN = tk.BooleanVar()
    key_autogen_choice = tk.Checkbutton(
        frame, text="Key Autogen", variable=KEY_AUTOGEN)
    key_autogen_choice.grid(row=5, column=0, pady=2)

    to_pdf_label = tk.Label(frame, text="To PDF Method:")
    to_pdf_label.grid(row=3, column=0, pady=2)


def to_pdf_method_manager(frame):
    global TO_PDF_METHOD

    TO_PDF_METHOD = tk.StringVar()

    to_pdf_libreoffice = tk.Radiobutton(
        frame, text="LibreOffice", variable=TO_PDF_METHOD, value="libreoffice")
    to_pdf_libreoffice.grid(row=3, column=1, pady=2)

    to_pdf_powerpoint = tk.Radiobutton(
        frame, text="PowerPoint", variable=TO_PDF_METHOD, value="powerpoint")
    to_pdf_powerpoint.grid(row=4, column=1, pady=2)


def run():
    root = tk.Tk()
    root.title("PyCertGen")
    root.geometry("400x280")  # Adjust the width as needed

    frame = tk.Frame(root)
    frame.pack()

    certified_list_manager(frame)
    template_manager(frame)
    output_manager(frame)
    autogen_manager(frame)
    to_pdf_method_manager(frame)

    submit_button = tk.Button(root, text="Generate", command=submit)
    submit_button.pack(pady=10)

    root.mainloop()
