
from pathlib import Path
from openpyxl import Workbook,load_workbook
import win32com.client as win32


def main():
    wb = Workbook()
    filePath = str(Path.cwd()) + r'\change_colour_of_all_cells.xlsx'
    path = Path(filePath)
    if path.is_file():
        del path
    wb.worksheets[0].title = 'white_background'
    wb.create_sheet("Sheet2")
    wb.save(filePath)

    xlApp = win32.Dispatch("Excel.Application")
    xlApp.Visible = False
    wb = xlApp.Workbooks.Open(filePath)



    ws = wb.Sheets("white_background")
    #białe tło dla wszystkich komórek
    ws.Cells.Interior.Color = 16777215

    ws = wb.Sheets("Sheet2")
    #białe tło dla wybranych wierszy, kolumn i komórek
    ws.Range("23:1048576").Interior.Color = 16777215
    ws.Range("R:XFD").Interior.Color = 16777215
    ws.Range("A1:C4").Interior.Color = 16777215


    wb.Save()
    wb.Close()
    xlApp.Quit()


if __name__ == "__main__":
    main()
