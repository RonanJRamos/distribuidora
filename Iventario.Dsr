VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} iventario 
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   OleObjectBlob   =   "Iventario.dsx":0000
End
Attribute VB_Name = "iventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
On Error Resume Next
Dim a As Integer
Posicao = 0
ReDim Preserve Mt(Posicao)

lcgrupo2 = Text9.Text & "|" & Text8.Text & "|" & Text7.Text & "|" & Text12.Text & "|" & Text6.Text & "|" & Text5.Text & "|" & Text4.Text & "|" & Text3.Text & "|" & Text2.Text & "|" & Text10.Text

Mt(Posicao).valor = lcgrupo2
Mt(Posicao).CelulasReculo = 0
Mt(Posicao).Negrito = True
Mt(Posicao).Cor = 15
Mt(Posicao).ColunasAPintar = 10
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
On Error Resume Next
Dim a As Integer
Dim Posicao As Integer
a = UBound(Mt)
If err = 0 Then
   Posicao = UBound(Mt) + 1
Else
   Posicao = 0
End If
ReDim Preserve Mt(Posicao)


 lcgrupo2 = Field9.Value & "|" & Field10.Value & "|" & Format(Field11.Value, "mm/dd/yy") & "|" & Format(Field21.Value, "mm/dd/yy") & "|" & Field12.Value & "|" & Format(Field13.Value, "mm/dd/yy") & "|" & Format(Field14.Value, "mm/dd/yy") & "|" & Field15.Value & "|" & Field16.Value & "|" & Field17.Value

Mt(Posicao).valor = lcgrupo2
Mt(Posicao).CelulasReculo = 0
Mt(Posicao).Negrito = False
Mt(Posicao).Cor = 15
Mt(Posicao).ColunasAPintar = 10
End Sub

Private Sub Section7_Format(ByVal pFormattingInfo As Object)
Field22.Value = BuscarSaldoAnterior(Field14.Value)
Field23.Value = CDbl(BuscarSaldoUltimo(Field14.Value))

End Sub
