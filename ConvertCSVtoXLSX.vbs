'=============================================
' ConvertCSVtoXLSX Function
'=============================================
'
' Description:
'   This VBScript function converts a CSV (Comma-Separated Values) file to an XLSX (Excel Workbook) file.
'
' Usage:
'   ConvertCSVtoXLSX(parameter)
'
' Parameters:
'   parameter - String representing the input and output file paths separated by '|'. Example: "input.csv|output.xlsx"
'
' Returns:
'   "success" - If the conversion is successful.
'   "error" - If the provided parameter format is invalid or missing input/output paths.
'
' Example:
'   ' For Testing: Specify input and output file paths
'   Dim vParameter
'   vParameter = "C:\Path\To\Input\File.csv|C:\Path\To\Output\File.xlsx"
'   ConvertCSVtoXLSX vParameter
'
' Notes:
'   - The function uses Excel automation to perform the conversion.
'   - The Excel application is created in the background (invisible) to execute the conversion.
'   - The function returns a success message or an error message based on the provided parameter.
'   - The XLSX format is used for the output file.
'
'=============================================

Sub ConvertCSVtoXLSX(parameter)
    ' Create Excel object and Split parameter into input and output file paths
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
    objExcel.DisplayAlerts = False
    Dim vParameter
    vParameter = Split(parameter, "|")
    
    ' Check if both input and output paths are provided
    If UBound(vParameter) = 1 Then
        Dim inputFile
        Dim outputFile
        inputFile = vParameter(0)
        outputFile = vParameter(1)

        ' Open CSV file, Save as XLSX file, Close Excel file and Clean up memory
        Dim objWorkbook
        Set objWorkbook = objExcel.Workbooks.Open(inputFile)
        objWorkbook.SaveAs outputFile, 51  ' 51 represents the XLSX format
        objWorkbook.Close
        objExcel.Quit
        Set objWorkbook = Nothing
        Set objExcel = Nothing

        ' Return success message
        ConvertCSVtoXLSX = "success"
    Else
        ' MsgBox "Invalid parameter format. Provide both input and output file paths separated by '|'.", vbExclamation
        ConvertCSVtoXLSX = "error"
    End If
End Sub

' For Testing: Specify input and output file paths
' Dim vParameter
' vParameter = "C:\Users\julio\OneDrive\Ambiente de Trabalho\VBS\Log.csv|C:\Users\julio\OneDrive\Ambiente de Trabalho\VBS\Log.xlsx"
' ConvertCSVtoXLSX vParameter
