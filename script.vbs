Function ExecutarFiltro(parameters)
    SplitParameters = Split(parameters, "|")
    pathFile = SplitParameters(0)

    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    objExcel.DisplayAlerts = False

    Set objWorkbook = objExcel.Workbooks.Open(pathFile)
    Set objWorksheet = objWorkbook.Sheets("Sheet1")
    Set objRange = objWorksheet.UsedRange

    Set sortCol1 = objWorksheet.Range("D1")

    Const xlAscending = 1
    Const xlDescending = 2


    Const xlGuess = 0
    Const xlYes = 1
    Const xlNo = 2

    objRange.Sort sortCol1, xlAscending, , , , , , xlYes

    objWorkbook.Save
    objWorkbook.Close
    objExcel.Quit
End Function

parameter = "test.xlsx"
Call ExecutarFiltro(parameter)

