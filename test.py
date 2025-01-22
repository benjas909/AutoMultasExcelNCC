import openpyxl, re
import tkinter as tk
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

def XLSXHandling(filename):
    # Abrir archivo excel de entrada
    workbook = openpyxl.load_workbook(filename=filename)
    sheet = workbook["Item 12"] # Selección de hoja

    contents = []

    
    # Guarda contenidos de hoja 
    for row in sheet.iter_rows():
        contents.append(row)

    i = 1
    for row in contents:
        sheet[f"C{i}"] = str(row[2].value)
        numList = re.findall(r"\d{6}", row[2].value)

        if len(numList) > 0:
            sheet[f"C{i}"] = "2024 " + numList[0]
            sheet[f"K{i}"] = f'=VLOOKUP(C{i},Tickets!$A$2:$X$410,22,FALSE)'
            sheet[f"M{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X$410,23,FALSE)),"Vacío",VLOOKUP(C{i},Tickets!$A$2:$X$410,23,FALSE))'
            sheet[f"N{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X$410,24,FALSE)),"Vacío",VLOOKUP(C{i},Tickets!$A$2:$X$410,24,FALSE))'
            sheet[f"O{i}"] = f'=IF(OR(ISTEXT(M{i}), ISTEXT(K{i})), "No Disponible", M{i}-K{i})'
            sheet[f"P{i}"] = f'=IF(OR(ISTEXT(N{i}), ISTEXT(K{i})), "No Disponible", N{i}-K{i})'

        if len(numList) > 1:
            ATM = row[0].value
            Comuna = row[1].value
            Apertura = row[10].value

            for item in numList[1:]:
                newNum = f"2024 {item}"
                sheet.insert_rows(idx= i + 1)
                i += 1
                sheet[f"A{i}"] = ATM
                sheet[f"B{i}"] = Comuna
                sheet[f"C{i}"] = newNum
                sheet[f"K{i}"] = f'=VLOOKUP(C{i},Tickets!$A$2:$X$410,22,FALSE)'
                sheet[f"M{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X$410,23,FALSE)),"Vacío",VLOOKUP(C{i},Tickets!$A$2:$X$410,23,FALSE))'
                sheet[f"N{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X$410,24,FALSE)),"Vacío",VLOOKUP(C{i},Tickets!$A$2:$X$410,24,FALSE))'
                sheet[f"O{i}"] = f'=IF(OR(ISTEXT(M{i}), ISTEXT(K{i})), "No Disponible", M{i}-K{i})'
                sheet[f"P{i}"] = f'=IF(OR(ISTEXT(N{i}), ISTEXT(K{i})), "No Disponible", N{i}-K{i})'
                
                print(ATM, "|", Comuna, "|", newNum, "|", Apertura)

        i += 1

    for r in sheet[ "A2:V154" ]:
        for cell in r:
            cell.font = Font(name = "Calibri", size = 9)
            cell.alignment = Alignment(horizontal = "center", vertical = "center")

    for r in sheet["J1:N154"]:
        for cell in r:
            cell.number_format = "dd/mm/yyyy h:mm"

    for r in sheet["O1:P154"]:
        for cell in r:
            cell.number_format = "h:mm:ss"

    sheet.column_dimensions["J"].width = 20
    sheet.column_dimensions["K"].width = 20
    sheet.column_dimensions["L"].width = 20
    sheet.column_dimensions["M"].width = 20
    sheet.column_dimensions["N"].width = 20
    sheet.column_dimensions["O"].width = 15
    sheet.column_dimensions["P"].width = 15
    

    # sheet["K"].number_format = 

    workbook.save(filename="output_test.xlsx")


def selectFile():
    filetypes = (
        ('Excel spreadsheets', '.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title = "Abrir archivo",
        initialdir="/",
        filetypes=filetypes
    )

    showinfo(title="Archivo seleccionado", message=filename)

    XLSXHandling(filename)

    showinfo(title="Listo", message="Archivo guardado")

    

def main():
    window = tk.Tk()

    window.resizable(False, False)
    window.geometry("300x300")
    openButton = ttk.Button(window, text="Abrir un archivo", command=selectFile)
    openButton.pack(expand=True)


    window.mainloop()



if __name__ == "__main__":
    main()