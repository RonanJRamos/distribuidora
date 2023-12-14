Attribute VB_Name = "Controle"
Sub Main()
On Error Resume Next
IniciaDll
'BaseDeDados = LeIni("CFG", "PathRpt", App.Path & CfgFile)
'LancamentoBalanco.Show
Principal.Show
End Sub
Private Sub IniciaDll()
Dim tempo As String
Dim NomeArquivo As String
Dim X As Boolean
'Dim objprotecao As New protecao.Controlador
NomeDoArquivo = App.EXEName
abreconexao

End Sub
