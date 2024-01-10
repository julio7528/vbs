'=============================================
' ConvertCSVtoXLSX Function
' https://github.com/julio7528/vbs/blob/master/ConvertCSVtoXLSX.vbs
'=============================================
'
' Description:
'   This VBScript function converts a CSV (Comma-Separated Values) file to an XLSX (Excel Workbook) file.
'
' Usage:
'   result = ConvertCSVtoXLSX(parameter)
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
'   result = ConvertCSVtoXLSX(vParameter)
'   WScript.Echo result
'
' Notes:
'   - The function uses Excel automation to perform the conversion.
'   - The Excel application is created in the background (invisible) to execute the conversion.
'   - The function returns a success message or an error message based on the provided parameter.
'   - The XLSX format is used for the output file.
'
'=============================================

Function ConvertCSVtoXLSX(parameter)
    ' Create Excel object and Split parameter into input and output file paths
    Dim objExcel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
    objExcel.DisplayAlerts = False
    Dim vParameter
    vParameter = Split(parameter, "|")

    ' Check if both input and output paths are provided
    If UBound(vParameter) <> 1 Then
        ' Return error message if the parameter format is invalid
        ConvertCSVtoXLSX = "error: Invalid parameter format. Provide input and output file paths separated by '|'."
        Exit Function
    End If

    Dim inputFile
    Dim outputFile
    inputFile = vParameter(0)
    outputFile = vParameter(1)

    On Error Resume Next

    ' Open CSV file, Save as XLSX file, Close Excel file, and Clean up memory
    Dim objWorkbook
    Set objWorkbook = objExcel.Workbooks.Open(inputFile)
    objWorkbook.SaveAs outputFile, 51  ' 51 represents the XLSX format
    objWorkbook.Close
    objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing

    If Err.Number = 0 Then
        ' Return success message
        ConvertCSVtoXLSX = "success"
    Else
        ' Return error message if an error occurred during the conversion
        ConvertCSVtoXLSX = "error: " & Err.Description
    End If

    On Error GoTo 0
End Function

' For Testing: Specify input and output file paths
'Dim vParameter
'vParameter = "C:\Users\julio\OneDrive\Ambiente de Trabalho\VBS\Log.csv|C:\Users\julio\OneDrive\Ambiente de Trabalho\VBS\Log.xlsx"
'result = ConvertCSVtoXLSX(vParameter)
'WScript.Echo result
