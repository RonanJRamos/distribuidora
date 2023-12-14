VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form DetalhaEtiqueta 
   BackColor       =   &H00C5FEE1&
   Caption         =   "Seleciona Clientes Para Imprimir Etiquetas"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Altera Todos Para &Sim F4"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Altera Todos Para Não F3"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Exibe 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Para Alterar Imprimir entre Sim e Não, Utilize a Barra de Espaço ou Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   7560
   End
End
Attribute VB_Name = "DetalhaEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a, b As Integer
Private Sub Command1_Click()
On Error Resume Next

AbreBase
Dim RsEtiquetaAnt As Recordset
Set RsEtiquetaAnt = Dbbase.OpenRecordset("etiquetasAnteriores", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption

For a = 1 To Exibe.Rows - 1
    Me.Caption = "Alterando Informações da Etiqueta " & a & " de " & LcTamanhoEtiqueta
    For b = 0 To LcTamanhoEtiqueta
        
        If MtEtiqueta(b).Codigo = Exibe.TextMatrix(a, 0) Then
           If Exibe.TextMatrix(a, 2) = "Sim" Then
              MtEtiqueta(b).Imprime = True
           Else
              MtEtiqueta(b).Imprime = False
           End If
           Exit For
        End If
    Next
Next
Do Until RsEtiquetaAnt.EOF
    RsEtiquetaAnt.Delete
    RsEtiquetaAnt.MoveNext
Loop

For a = 0 To LcTamanhoEtiqueta
       Me.Caption = "Gravando Etiqueta " & a & " de " & LcTamanhoEtiqueta
       RsEtiquetaAnt.AddNew
       RsEtiquetaAnt("Nome") = MtEtiqueta(a).Nome & ""
       RsEtiquetaAnt("End") = MtEtiqueta(a).Endereco & ""
       RsEtiquetaAnt("bairro") = MtEtiqueta(a).Bairro & ""
       RsEtiquetaAnt("cidade") = MtEtiqueta(a).Cidade & ""
       RsEtiquetaAnt("cep") = MtEtiqueta(a).Cep & ""
       RsEtiquetaAnt("Estado") = MtEtiqueta(a).UF & ""
       RsEtiquetaAnt("imprime") = MtEtiqueta(a).Imprime
       RsEtiquetaAnt.Update
    
Next
Me.Caption = LcCap
RsEtiquetaAnt.Close
Set RsEtiquetaAnt = Nothing
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 114 Then SendKeys "%{A}"
If KeyCode = 115 Then SendKeys "%{S}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 114 Then SendKeys "%{A}"
If KeyCode = 115 Then SendKeys "%{S}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
b = 1
For a = 0 To LcTamanhoEtiqueta
     Exibe.Rows = b + 1
     Exibe.TextMatrix(b, 0) = MtEtiqueta(a).Codigo
     Exibe.TextMatrix(b, 1) = MtEtiqueta(a).Nome
     Exibe.TextMatrix(b, 2) = "Não"
     b = b + 1
Next
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 114 Then SendKeys "%{A}"
If KeyCode = 115 Then SendKeys "%{S}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command4_Click()
On Error Resume Next
For a = 0 To LcTamanhoEtiqueta
     Exibe.Rows = b + 1
     Exibe.TextMatrix(b, 0) = MtEtiqueta(a).Codigo
     Exibe.TextMatrix(b, 1) = MtEtiqueta(a).Nome
     Exibe.TextMatrix(b, 2) = "Sim"
     b = b + 1
Next
End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 114 Then SendKeys "%{A}"
If KeyCode = 115 Then SendKeys "%{S}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Exibe_DblClick()
LcLinha = Exibe.Row
If Exibe.TextMatrix(LcLinha, 2) = "Sim" Then
   Exibe.TextMatrix(LcLinha, 2) = "Não"
Else
  Exibe.TextMatrix(LcLinha, 2) = "Sim"
End If

End Sub

Private Sub Exibe_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 32 Then
   LcLinha = Exibe.Row
   If Exibe.TextMatrix(LcLinha, 2) = "Sim" Then
      Exibe.TextMatrix(LcLinha, 2) = "Não"
   Else
      Exibe.TextMatrix(LcLinha, 2) = "Sim"
   End If
End If

If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 114 Then SendKeys "%{A}"
If KeyCode = 115 Then SendKeys "%{S}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Form_Load()
GeraGrid
ExibePesquisa
End Sub
Function GeraGrid()
Exibe.ColAlignment(0) = 7
Exibe.ColAlignment(1) = 1
Exibe.ColAlignment(2) = 1


Exibe.ColWidth(0) = 1000
Exibe.ColWidth(1) = 4900
Exibe.ColWidth(2) = 900


Exibe.TextMatrix(0, 0) = "Código"
Exibe.TextMatrix(0, 1) = "Nome"
Exibe.TextMatrix(0, 2) = "Imprime"

LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
b = 1
For a = 0 To LcTamanhoEtiqueta
     Exibe.Rows = b + 1
     Exibe.TextMatrix(b, 0) = MtEtiqueta(a).Codigo
     Exibe.TextMatrix(b, 1) = MtEtiqueta(a).Nome
     If MtEtiqueta(a).Imprime Then
        Exibe.TextMatrix(b, 2) = "Sim"
     Else
        Exibe.TextMatrix(b, 2) = "Não"
     End If
     b = b + 1
Next

  



Exit Function
errorExibeCli:
If err = 5 Then Resume Next
Resume Next


End Function
