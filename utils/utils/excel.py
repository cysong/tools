def write_sheet(sheet, header, rows):
    for i in range(0, len(header)):
        sheet.cell(row=1, column=i + 1, value=str(header[i]))

    for i in range(0, len(rows)):
        for j in range(0, len(rows[i])):
            sheet.cell(row=i + 2, column=j + 1, value=rows[i][j])
