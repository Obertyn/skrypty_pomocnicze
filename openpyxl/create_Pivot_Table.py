
from pathlib import Path
from openpyxl import Workbook,load_workbook
import random
import win32com.client as win32

def createTable(filePath, sheetName, wb):

    sheet = wb[sheetName]
    sheet[('A1')] = 'Kol1'
    sheet[('B1')] = 'Kol2'
    sheet[('C1')] = 'Kol3'
    sheet[('D1')] = 'Kol4'
    sheet[('E1')] = 'Kol5'

    tab1 =[]
    tab2 =[]
    tab3 =[]
    tab4 =[]
    Rows=30
    for i in range(Rows):
        tab1.append(random.choice(["value1", "value2", "value3"]))
        tab2.append(random.choice("abcdefghijklmnopqrstuvwxyz"))
        tab3.append(random.randint(0, 10))
        tab4.append(random.randint(0, 100))
    
    print(tab1)
    print(tab2)
    print(tab3)
    print(tab4)
    for i in range(Rows):
        sheet[('A'+str(i+2))] = tab1[i]
        sheet[('B'+str(i+2))] = tab2[i]
        sheet[('C'+str(i+2))] = tab3[i]
        sheet[('D'+str(i+2))] = tab4[i]
        sheet[('E'+str(i+2))] = '=D'+str(i+2)+'/sum(D2:D'+str(Rows+1)+')'
    wb.save(filePath)
    return Rows

def createPivotTable(filePath,sheetName, Rows):
    xlApp = win32.Dispatch("Excel.Application")
    xlApp.Visible = False
    wb = xlApp.Workbooks.Open(filePath)
    ws_data = wb.Worksheets(sheetName)
    ws_report = wb.Worksheets("Pivot_Tables")

    cl1 = ws_data.Cells(1,1) #gdzie zaczyna się tabela źródłowa
    cl2 = ws_data.Cells(Rows,5) #gdzie kończy się tabela źródłowa
    PivotSourceRange = ws_data.Range(cl1,cl2)
    pt_cache = wb.PivotCaches().Create(1, SourceData=PivotSourceRange)
    pt = pt_cache.CreatePivotTable(ws_report.Range("A2"), "myreport_summary")
    pt.ColumnGrand = True
    pt.RowGrand = True
    pt.RowAxisLayout(1)
    pt.TableStyle2 = "pivotStyleMedium21"

    def create_pivot_table1(pt):

        #orientation 1 to Wiersze
        #orientation 2 to Kolumny
        #orientation 3 to Filtry
        #orientation 4 to Wartości
        field_rows = {}
        field_rows["Kol1"] = pt.PivotFields("Kol1")
        field_rows["Kol1"].Orientation = 1

        field_values = {}
        field_values["Kol3"] = pt.PivotFields("Kol3")
        field_values["Kol3"].Orientation = 4

        field_values = {}
        field_values["Kol4"] = pt.PivotFields("Kol4")
        field_values["Kol4"].Orientation = 4

        #NumberFormat to formatowanie komórek
        #Przykłady "0%" "0"
        field_values = {}
        field_values["Kol5"] = pt.PivotFields("Kol5")
        field_values["Kol5"].Orientation = 4
        field_values["Kol5"].NumberFormat = "0%"

        #Calculation to 'Pokaż wartości jako'
        #Calculation = 8 to % sumy końcowej
        field_values = {}
        field_values["Kol2"] = pt.PivotFields("Kol2")
        field_values["Kol2"].Orientation = 4
        field_values["Kol2"].Calculation = 8
        field_values["Kol2"].NumberFormat = "0%"

        #Function to 'Podsumuj pole wartości według'
        #Function = 1 to Liczba
        field_values = {}
        field_values["Kol1"] = pt.PivotFields("Kol1")
        field_values["Kol1"].Orientation = 4
        field_values["Kol1"].Function = 1

        #Function = 2 to Średnia
        field_values = {}
        field_values["Kol3"] = pt.PivotFields("Kol3")
        field_values["Kol3"].Orientation = 4
        field_values["Kol3"].Function = 2

    create_pivot_table1(pt)

    def create_pivot_table2(pt):
        field_rows = {} 
        field_rows["Kol2"] = pt.PivotFields("Kol2")
        field_rows["Kol2"].Orientation = 1
        field_values = {}
        field_values["Kol3"] = pt.PivotFields("Kol3")
        field_values["Kol3"].Orientation = 4
        field_values = {}
        field_values["Kol4"] = pt.PivotFields("Kol4")
        field_values["Kol4"].Orientation = 4
        field_values = {}
        field_values["Kol5"] = pt.PivotFields("Kol5")
        field_values["Kol5"].Orientation = 4
        field_values["Kol5"].NumberFormat = "0%"
        field_values = {}
        field_values["Kol2"] = pt.PivotFields("Kol2")
        field_values["Kol2"].Orientation = 4
        field_values["Kol2"].Calculation = 8
        field_values["Kol2"].NumberFormat = "0%"
        field_values = {}
        field_values["Kol1"] = pt.PivotFields("Kol1")
        field_values["Kol1"].Orientation = 4
        field_values["Kol1"].Function = 1
        field_values = {}
        field_values["Kol3"] = pt.PivotFields("Kol3")
        field_values["Kol3"].Orientation = 4
        field_values["Kol3"].Function = 2

    pt = pt_cache.CreatePivotTable(ws_report.Range("A9"), "myreport_summary2")
    pt.ColumnGrand = True
    pt.RowGrand = True
    pt.RowAxisLayout(1)
    pt.TableStyle2 = "pivotStyleMedium21"


    create_pivot_table2(pt)


    wb.Close(True)
    xlApp.Quit()




def main():
    wb = Workbook()
    filePath = str(Path.cwd()) + r'\create_Pivot_Table.xlsx'
    path = Path(filePath)
    if path.is_file():
        del path
    wb.worksheets[0].title = 'Sheet1'
    sheetName = wb.worksheets[0].title
    wb.create_sheet("Pivot_Tables")
    Rows = createTable(filePath, sheetName, wb)
    createPivotTable(filePath, sheetName, Rows)



if __name__ == "__main__":
    main()

