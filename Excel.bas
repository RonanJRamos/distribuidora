Attribute VB_Name = "Excel"
Sub ExportaDados()
Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    'Abrir o arquivo do Excel
    Set xlw = xl.Workbooks.Open("c:\teste\teste.xls")

    ' definir qual a planilha de trabalho
    xlw.Sheets("Plan1").Select

    'Exibe o conte�do da c�lula na posi��o 2,3

   ' variavel = xlw.Application.Cells(2, 3).Value
    MsgBox xlw.Application.Cells(2, 3).Value

 

    ' Fechar a planilha sem salvar altera��es
    ' Para salvar mude False para True

    xlw.Close False

    ' Liberamos a mem�ria

    Set xlw = Nothing
    Set xl = Nothing

End Sub
