import openpyxl, re
import tkinter as tk
from openpyxl.styles import Font, Alignment
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

def XLSXHandling(filename):
    # Abrir archivo excel de entrada
    workbook = openpyxl.load_workbook(filename=filename)
    ticketsSheet = workbook["Tickets"]
    tickMax = ticketsSheet.max_row
    print(tickMax)
    sheet = workbook["Item 12"] # Selección de hoja
    lastTicket = sheet.max_row - 3
    print(lastTicket)

    contents = []

    # Guarda contenidos de hoja 
    for row in sheet.iter_rows():
        contents.append(row)

    i = 1
    addedCells = 0
    for row in contents:

        if (row[0].value is None):
            break

        sheet[f"C{i}"] = str(row[2].value)
        numList = re.findall(r"\d{6}", row[2].value)

        # Si ATM tiene por lo menos un ticket
        if len(numList) > 0:
            sheet[f"C{i}"] = "2024 " + numList[0] # Reescribe número de ticket con formato correcto
            sheet[f"K{i}"] = f'=IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},22,FALSE)), "No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},22,FALSE))'
            sheet[f"M{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)),"Vacío",IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)),"No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)))'
            sheet[f"N{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)),"Vacío",IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)),"No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)))'
            sheet[f"O{i}"] = f'=IF(ISERR(M{i}-K{i}), "No Disponible", M{i}-K{i})'
            sheet[f"P{i}"] = f'=IF(ISERR(N{i}-K{i}), "No Disponible", N{i}-K{i})'

            # Si ATM tiene más de un ticket
            if len(numList) > 1:
                # Guarda datos de ATM actual
                ATM = row[0].value
                Comuna = row[1].value
                Apertura = row[10].value

                # Recorre lista de tickets, desde el segundo ticket
                for item in numList[1:]:
                    newNum = f"2024 {item}"
                    sheet.insert_rows(idx= i + 1)
                    addedCells += 1
                    i += 1
                    sheet[f"A{i}"] = ATM
                    sheet[f"B{i}"] = Comuna
                    sheet[f"C{i}"] = newNum
                    sheet[f"K{i}"] = f'=IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},22,FALSE)), "No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},22,FALSE))'
                    sheet[f"M{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)),"Vacío",IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)),"No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},23,FALSE)))'
                    sheet[f"N{i}"] = f'=IF(ISBLANK(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)),"Vacío",IF(ISNA(VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)),"No Encontrado", VLOOKUP(C{i},Tickets!$A$2:$X${tickMax},24,FALSE)))'
                    sheet[f"O{i}"] = f'=IF(ISERR(M{i}-K{i}), "No Disponible", M{i}-K{i})'
                    sheet[f"P{i}"] = f'=IF(ISERR(N{i}-K{i}), "No Disponible", N{i}-K{i})'
                    
                    print(ATM, "|", Comuna, "|", newNum, "|", Apertura)

        # ATM sin número de ticket
        elif(len(numList) == 0 and i != 1):
            sheet[f"K{i}"] = "No Disponible"
            sheet[f"M{i}"] = "No Disponible"
            sheet[f"N{i}"] = "No Disponible"
            sheet[f"O{i}"] = "No Disponible"
            sheet[f"P{i}"] = "No Disponible"


        i += 1

    sheet[f"I{lastTicket + addedCells + 3}"] = f'=SUM(I2:I{lastTicket + addedCells})'


    # Aplica estilos a la hoja
    for r in sheet[ f"A2:V{lastTicket + addedCells + 3}" ]:
        for cell in r:
            cell.font = Font(name = "Calibri", size = 9)
            cell.alignment = Alignment(horizontal = "center", vertical = "center")

    # Aplica formato de fecha a columnas correspondientes
    for r in sheet[ f"J1:N{lastTicket + addedCells + 3}" ]:
        for cell in r:
            cell.number_format = "dd/mm/yyyy h:mm"

    # Aplica formato de hora a columnas correspondientes
    for r in sheet[ f"O1:P{lastTicket + addedCells + 3}" ]:
        for cell in r:
            cell.number_format = "h:mm:ss"

    # Aplica ancho de columna
    sheet.column_dimensions["J"].width = 20
    sheet.column_dimensions["K"].width = 20
    sheet.column_dimensions["L"].width = 20
    sheet.column_dimensions["M"].width = 20
    sheet.column_dimensions["N"].width = 20
    sheet.column_dimensions["O"].width = 15
    sheet.column_dimensions["P"].width = 15

    workbook.save(filename="Junio2024_test.xlsx")


# Ventana de selección de archivo
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

    # Ventana principal
    window = tk.Tk()
    window.resizable(False, False)
    window.title("AutoMultas")
    window.geometry("300x300")
    openButton = ttk.Button(window, text="Abrir un archivo", command=selectFile)
    openButton.pack(expand=True)

    window.mainloop()


if __name__ == "__main__":
    main()