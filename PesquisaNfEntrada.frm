VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PesquisaNfEntrada 
   Caption         =   "Pesqusa nota de Entrada"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Comfirmar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton CmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid item 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin VB.TextBox Nf 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero da Nota"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "PesquisaNfEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GeraGrid()
On Error Resume Next
Item.TextMatrix(0, 0) = "NF"
Item.TextMatrix(0, 1) = "Entrada"
Item.TextMatrix(0, 2) = "Fornecedor"
Item.TextMatrix(0, 3) = "Valor"
Item.TextMatrix(0, 4) = "CodFornecedor"
Item.ColWidth(0) = 1000
Item.ColWidth(1) = 1000
Item.ColWidth(2) = 5000
Item.ColWidth(3) = 1000
Item.ColWidth(4) = 0
Item.ColWidth(5) = 0
End Sub

Private Sub CmdConfirmar_Click()
'On Error Resume Next
linha = Item.Row

FrmEntradaProduto.BuscaNota Item.TextMatrix(linha, 5), Item.TextMatrix(linha, 4), Item.TextMatrix(linha, 0)
FrmEntradaProduto.EPesquisa.Text = "Pesquisando"
Unload Me
End Sub

Private Sub CmdExibir_Click()
On Error Resume Next
Dim RsNota As ADODB.Recordset
Dim RsFornecedor As Recordset
Dim a As Integer
Dim StrSql As String
Dim db As Database
Dim LcAchou As Boolean
Set db = OpenDatabase(GLBase)
StrSql = "Select * from entradanf where nf like '%" & NF.Text & "%' order by data desc limit 300"
Set RsNota = AbreRecordset(StrSql, True)

Screen.MousePointer = 11
Item.Rows = 1
Do Until RsNota.EOF
   If err.Number <> 0 Then Exit Do
   LcAchou = True
   StrSql = "Select * from alid002 where codigo='" & RsNota!clicred & "'"
   Set RsFornecedor = db.OpenRecordset(StrSql)
   a = Item.Rows
   Item.Rows = a + 1
   Item.TextMatrix(a, 0) = RsNota!NF & ""
   Item.TextMatrix(a, 1) = RsNota!Data & ""
   If Not RsFornecedor.EOF Then
       Item.TextMatrix(a, 2) = RsFornecedor!RazaoSoc & ""
   End If
   Item.TextMatrix(a, 3) = RsNota!valor & ""
   Item.TextMatrix(a, 4) = RsNota!clicred & ""
   Item.TextMatrix(a, 5) = RsNota!Codigo & ""
   RsNota.MoveNext
   Set RsFornecedor = Nothing
Loop
Screen.MousePointer = 0
If Not LcAchou Then
   MsgBox "Não foi encontrada nota com este numero.", 64, "Aviso"
End If
End Sub

Private Sub Form_Load()
GeraGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmEntradaProduto.SetFocus
End Sub

Private Sub Item_Click()
CmdConfirmar.Enabled = True

End Sub
