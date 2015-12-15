from openpyxl import load_workbook

from tkinter import *
from tkinter import filedialog


def browse_file_call_back(import_text):
    filename = filedialog.askopenfilename(filetypes=(("Excel Files", "*.xlsx"),))
    import_text.delete(0, END)
    import_text.insert(0, filename)


def import_workbook(filename):
    wb = load_workbook(filename)
    w_sheets = wb.worksheets

    cols = list_all_cols(w_sheets)
    print(cols)


def list_all_cols(w_sheets):
    cols = []

    for ws in w_sheets:
        col_num = 1

        while True:
            if ws.cell(row=1, column=col_num).value is not None:
                col_name = ws.cell(row=1, column=col_num).value
            else:
                col_name = ws.cell(row=2, column=col_num).value

            if col_name is None or col_name == "":
                break

            if col_name not in cols:
                cols.append(col_name)

            col_num += 1

    return cols


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
