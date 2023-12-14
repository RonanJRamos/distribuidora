VERSION 5.00
Begin VB.Form atualizapreco 
   BackColor       =   &H00C9DADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualiza Preço"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LimitePreco 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox maximo 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox minimo 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox venda 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limite de Venda"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E2F1F1&
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
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E2F1F1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E2F1F1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E2F1F1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E2F1F1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximo Estoque"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minimo de Venda"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   465
   End
End
Attribute VB_Name = "atualizapreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim Rs As Recordset
On Error GoTo errat
Dim LcPreco As String
Dim LcMinimoVenda As String
Dim LcMaximo As String
Dim LcLimite As String
Dim LcSq As String

If Len(venda.Text) = 0 Then venda.Text = 0
If Len(maximo.Text) = 0 Then maximo.Text = 0
If Len(minimo.Text) = 0 Then minimo.Text = 0
If Len(LimitePreco.Text) = 0 Then LimitePreco.Text = 0

If Not IsNumeric(venda.Text) Then
   MsgBox "o valor de venda deve ser numérico.", 64, "Aviso"
   venda.SetFocus
   Exit Sub
End If
If Not IsNumeric(minimo.Text) Then
   MsgBox "o valor de mímimo de venda deve ser numérico.", 64, "Aviso"
   minimo.SetFocus
   Exit Sub
End If
If Not IsNumeric(maximo.Text) Then
   MsgBox "o valor de máximo deve ser numérico.", 64, "Aviso"
   maximo.SetFocus
   Exit Sub
End If
If Not IsNumeric(LimitePreco.Text) Then
   MsgBox "o valor de Limite de Venda deve ser numérico.", 64, "Aviso"
   LimitePreco.SetFocus
   Exit Sub
End If

LcPreco = Replace(venda.Text, ",", ".")
LcMinimoVenda = Replace(minimo.Text, ",", ".")
LcMaximo = Replace(maximo.Text, ",", ".")
LcLimite = Replace(LimitePreco.Text, ",", ".")

LcSq = "UPDATE Produtos SET preco=" & LcPreco & ",MinimoVenda=" & LcMinimoVenda & ",maximoEstoque=" & LcMaximo & ",LimiteVenda=" & LcLimite
LcSq = LcSq & " where codigo=" & FrmEntradaProduto.Txt(1).Text

conexaoAdo.BeginTrans
ExecutaSql LcSq

conexaoAdo.CommitTrans
Unload Me
FrmEntradaProduto.SetFocus
Exit Sub
errat:
conexaoAdo.RollbackTrans
MsgBox "Ocorreu o seguinte erro Processando os Preços:" & Chr(13) & err.Description, 64, "Aviso"



End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
Dim Cestoque As ControleDb
On Error Resume Next

Set Cestoque = New ControleDb
Cestoque.CodProduto = FrmEntradaProduto.Txt(1).Text

'AbreBase
'Set Rs = Dbbase.OpenRecordset("Select * from alid009 where cod='" & FrmEntradaProduto.txt(1).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'If Not Rs.EOF Then
   LcOldV = FrmEntradaProduto.Valor(0).Text
   If CDbl(Cestoque.PrecoDeCusto) < CDbl(LcOldV) Then
      Label4.Caption = "Custo Anterior       : " & Format(Cestoque.PrecoDeCusto, "Currency")
      Label5.Caption = "Venda                    : " & Format(Cestoque.PrecoVenda, "currency")
      Label6.Caption = "Minimo de Venda : " & Format(Cestoque.PrecoMinimo, "Currency")
      Label7.Caption = "Maximo                  : " & Format(Cestoque.maximoEstoque, "currency")
      Label8.Caption = "Entre com o Novo Valor de Venda."
      'Label4.Caption = lcprompt
      venda.Text = Cestoque.PrecoVenda & ""
      minimo.Text = Cestoque.PrecoMinimo & ""
      maximo.Text = Cestoque.maximoEstoque & ""
      LimitePreco.Text = Cestoque.LimiteVenda & ""
   End If
'End If
'Rs.Close
'Dbbase.Close
Set Cestoque = Nothing
End Sub

Private Sub maximo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub maximo_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub minimo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub minimo_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub venda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub venda_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub
