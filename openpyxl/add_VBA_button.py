
import win32com.client as win32
from pathlib import Path
from openpyxl import Workbook,load_workbook

def createButton(filePath, sheetName):
        xl = win32.Dispatch("Excel.Application")
        wb = xl.Workbooks.Open(filePath)
        ws = wb.Worksheets(sheetName)
        obj = ws.OLEObjects()
        button = obj.Add(ClassType="Forms.CommandButton.1", Link=False, DisplayAsIcon=False).Object
        button.Caption = "Przycisk"
        macro_code = '''Private Sub Przycisk1_Click()
    Sheets("Sheet1").Range("A1").Value = 1
    Sheets("Sheet1").Range("A2").Value = 2
    Sheets("Sheet1").Range("A3").Value = 3
End Sub'''
        button.Name = "Przycisk1"
        button.Left=70
        button.Top=10
        button.Width=100
        button.Height=30
        wb.VBProject.VBComponents("Arkusz1").CodeModule.AddFromString(macro_code)
        wb.Save()
        wb.Close()
        xl.Quit()

def main():
    wb = Workbook()
    filePath = str(Path.cwd()) + r'\add_VBA_button.xlsm'
    path = Path(filePath)
    if path.is_file():
        del path
    wb.worksheets[0].title = 'Sheet1'
    sheetName = wb.worksheets[0].title
    wb.save(filePath)
    wb1 = load_workbook('add_VBA_button.xlsm', keep_vba=True)
    wb1.save('add_VBA_button.xlsm')
    createButton(filePath, sheetName)

if __name__ == "__main__":
    main()
