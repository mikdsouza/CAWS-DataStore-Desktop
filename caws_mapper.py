from openpyxl import Workbook, load_workbook

from tkinter import *
from tkinter import filedialog


def browse_file_call_back(import_text):
    filename = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"),))
    import_text.delete(0, END)
    import_text.insert(0, filename)


def import_workbook(filename):
    wb = load_workbook(filename)
    w_sheets = wb.worksheets

    for ws in w_sheets:
        print(w_sheets.name)


def main():
    root = Tk()
    root.title("CAWS Data Mapper")
    root.geometry("500x500")

    frame = Frame(root)
    frame.pack(side=TOP, fill=X)

    Label(frame, text="Import excel file").pack(side=LEFT, padx=2, pady=2)

    import_text = Entry(frame)
    import_text.pack(side=LEFT, padx=2, pady=2)

    Button(frame, text="Browse", command=lambda: browse_file_call_back(import_text)).pack(side=LEFT, padx=2, pady=2)

    Button(root, text="Import", command=lambda: import_workbook(import_text.get())).pack(side=TOP, padx=2, pady=2)

    root.mainloop()


main()
