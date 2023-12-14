Attribute VB_Name = "ControleExcel"
Public Type Tipo
    Valor As String
    CelulasReculo As Integer
    Negrito As Boolean
    ColunasAPintar As Integer
    Cor As Integer
End Type
Public Mt() As Tipo
Sub gera(NomeRelatorio As String)
On Error GoTo ErroGeraExcel
Dim x               As Integer
Dim Y               As Integer
Dim AplicacaoExcel  As Excel.Application
Dim ArquivoExcel    As Excel.Workbook
Dim PlanilhaExecel  As Excel.Worksheet
Dim Valores()       As String
Dim PosicaoRange    As String
Dim PosicaoColuna   As String
Dim NomeArquivo     As String


Screen.MousePointer = 11
Set AplicacaoExcel = CreateObject("excel.application")

Set ArquivoExcel = AplicacaoExcel.Workbooks.Add
Set PlanilhaExecel = AplicacaoExcel.Worksheets.Add

'==> Mudando o Nome da Palnilha
PlanilhaExecel.Name = NomeRelatorio

DoEvents
For x = 0 To UBound(Mt)
   Valores = Split(Mt(x).Valor, "|")
   '==> preenche as celulas
   For Y = 0 To UBound(Valores)
       PlanilhaExecel.Cells(x + 1, Y + 1 + Mt(x).CelulasReculo) = Valores(Y)
       '==> Se for a primeira linha, vamos colocar em negrito
       '==> Acerta largura das colunas
       PosicaoColuna = PosicaoLetra(Y + 1 + Mt(x).CelulasReculo)
       PlanilhaExecel.Columns(PosicaoColuna & ":" & PosicaoColuna).EntireColumn.AutoFit
   Next
   'If x = 0 Then
   If Mt(x).Negrito Then
      PosicaoRange = PosicaoLetra(Mt(x).CelulasReculo + 1) & x + 1 & ":" & PosicaoLetra(Mt(x).CelulasReculo + Mt(x).ColunasAPintar) & x + 1
      PlanilhaExecel.Range(PosicaoRange).Font.Bold = True
      AcertaCor PlanilhaExecel, PosicaoRange, Mt(x).Cor
   End If
   
Next

DoEvents

'==> solicitando a caixa para o nome do arquivo
With Relatorios.CommonDialog1
    .InitDir = App.Path
    .Filter = "*.xls"
    .FileName = NomeRelatorio & ".xls"
    .ShowOpen
    
    NomeArquivo = .FileName
End With
PlanilhaExecel.SaveAs NomeArquivo

Screen.MousePointer = 0
AplicacaoExcel.Visible = True

Set AplicacaoExcel = Nothing
Set ArquivoExcel = Nothing
Set PlanilhaExecel = Nothing

Exit Sub
ErroGeraExcel:
Screen.MousePointer = 0

If err.Number = 1004 Then
   If MsgBox("O Arquivo " & NomeArquivo & " Esta aberto por outra aplicação." & Chr(13) & "Feche antes de salvar." & Chr(13) & "Tenta salvar novamente?", vbYesNo, "Erro salvando arquivo.") = vbYes Then
      Resume 0
   End If
End If

MsgBox "Ocorreu o seguinte erro salvando o arquivo " & Chr(13) & err.Description, 64, "N:" & err.Number
Resume 0
End Sub
Function PosicaoLetra(NumeroPosicao As Integer) As String
On Error Resume Next
Dim Posicao As String

Select Case NumeroPosicao
    Case Is = 1
        Posicao = "A"
    Case Is = 2
        Posicao = "B"
    Case Is = 3
        Posicao = "C"
    Case Is = 4
        Posicao = "D"
    Case Is = 5
        Posicao = "E"
    Case Is = 6
        Posicao = "F"
    Case Is = 7
        Posicao = "G"
    Case Is = 8
        Posicao = "H"
    Case Is = 9
        Posicao = "I"
    Case Is = 10
        Posicao = "J"
    Case Is = 11
        Posicao = "K"
    Case Is = 12
        Posicao = "L"
    Case Is = 13
        Posicao = "M"
    Case Is = 14
        Posicao = "N"
    Case Is = 15
        Posicao = "O"
    Case Is = 16
        Posicao = "P"
    Case Is = 17
        Posicao = "Q"
    Case Is = 18
        Posicao = "R"
    Case Is = 19
        Posicao = "S"
    Case Is = 20
        Posicao = "T"
    Case Is = 21
        Posicao = "U"
    Case Is = 22
        Posicao = "V"
    Case Is = 23
        Posicao = "W"
    Case Is = 24
        Posicao = "X"
    Case Is = 25
        Posicao = "Y"
    Case Is = 26
        Posicao = "Z"
    Case Is = 27
        Posicao = "AA"
    Case Is = 28
        Posicao = "AB"
    Case Is = 29
        Posicao = "AC"
    
End Select
PosicaoLetra = Posicao

End Function
Sub AcertaCor(Planilha As Excel.Worksheet, PosicaoRange As String, Cor As Integer)
On Error Resume Next
    With Planilha.Range(PosicaoRange).Interior
        .ColorIndex = Cor
        .Pattern = xlSolid
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
      '  .LineStyle = xlContinuous
        .Weight = xlThin
   '     .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
    '    .LineStyle = xlContinuous
        .Weight = xlThin
     '   .ColorIndex = xlAutomatic
    End With
End Sub
