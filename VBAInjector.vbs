function InjectMacro(parameter)
    'Criado por Júlio Gomes
    'Injeta o código do VBA dentro da planilha do Excel e salva no formato xlsm

    'Split parameter delimidef by |
    Dim vParametros
    vParametros = Split(parameter, "|")
    xlsxExt = vParametros(0)
    xlsmExt = vParametros(1)

    Dim objExcel, objWorkbook, objSheet
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    objExcel.DisplayAlerts = False
    Set objWorkbook = objExcel.Workbooks.Open(xlsxExt, 0, False, , , , True, , , True)

    ' Adiciona um novo código VBA ao livro de trabalho
    objWorkbook.VBProject.VBComponents.Add 1 ' Módulo VBA
    objWorkbook.VBProject.VBComponents(1).CodeModule.AddFromString _

        "Sub valorUnAdm2()" & vbCrLf & _
        "msgbox ""teste"" & ""teste""" & vbCrLf & _
        "End Sub"

    ' Salva o livro de trabalho como plan.xlsm
    objWorkbook.SaveAs xlsmExt, 52 ' 52 = xlOpenXMLWorkbookMacroEnabled

    objWorkbook.Close
    objExcel.Quit
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End function

vArquOrigem = "Arquivo.xlsm"
vArquDestino = "Arquivo.xlsm"

Call InjectMacro(vArquOrigem & "|" & vArquDestino)
