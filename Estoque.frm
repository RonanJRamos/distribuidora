VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Estoque 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque Disponivel"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MostraCliente 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
GeraGrid
ExibePesquisa
End Sub
Function GeraGrid()
MostraCliente.ColAlignment(0) = 7
MostraCliente.ColAlignment(1) = 1
MostraCliente.ColWidth(0) = 3000
MostraCliente.ColWidth(1) = 800
MostraCliente.TextMatrix(0, 0) = "Galpao"
MostraCliente.TextMatrix(0, 1) = "Estoque"
LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
On Error GoTo errorExibeCli
Dim rsCliente As Recordset, RsCliente1 As Recordset, RsCliente2 As Recordset
Dim LcCriSql, LcCriSql1, LcCriSql2 As String
Dim LcTamanho, a As Long
Dim LcAchou As Integer
LcCriSql = "select * From alid013 where item='" & FrmSaidaProduto.Txt(1).Text & "'"

'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
AbreBase
Set rsCliente = Dbbase.OpenRecordset(LcCriSql) ', dbOpenDynaset)

LcTamanho = MostraCliente.Rows
a = 2
Me.Caption = msg
MostraCliente.Rows = 1
LcAchou = False
Do Until rsCliente.EOF
  LcAchou = True
  If Len(Trim(rsCliente!almox)) > 0 Then
   If Not IsNull(rsCliente!almox) Then
     
     MostraCliente.Rows = a
     MostraCliente.TextMatrix(a - 1, 0) = rsCliente!almox & ""
     MostraCliente.TextMatrix(a - 1, 1) = rsCliente!Estoque & ""
     a = a + 1
     rsCliente.MoveNext
    End If
  End If
  
Loop



rsCliente.Close
Set rsCliente = Nothing
MostraCliente.SetFocus
Exit Function
errorExibeCli:
If err = 5 Then Resume Next
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   'Resume 0
End If


End Function

Private Sub MostraCliente_DblClick()
LcLinha = MostraCliente.Row
FrmSaidaProduto.almox.Text = MostraCliente.TextMatrix(LcLinha, 0)
Unload Me
End Sub

Private Sub MostraCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   LcLinha = MostraCliente.Row
   FrmSaidaProduto.almox.Text = MostraCliente.TextMatrix(LcLinha, 0)
   Unload Me
End If
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
