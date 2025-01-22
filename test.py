import openpyxl, re

def main():
    print("hola")

    workbook = openpyxl.load_workbook(filename="./GTD_Multas_2024-05_redbanc_prueba_edit.xlsx")
    sheet = workbook["Item 12"]


    contents = []
    

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

    # sheet["K"].number_format = 

    workbook.save(filename="output_test.xlsx")


if __name__ == "__main__":
    main()