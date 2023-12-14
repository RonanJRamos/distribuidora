VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ExibecomissaoRepresent 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comissão do Vendedor"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Fechar 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "&Confirma  F2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   " Relação de Comissões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "ExibecomissaoRepresent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MsComissao
        Descricao As String
        valor As Currency
End Type
Private LcMat() As MsComissao
Private LcTam As Long

Private Sub Confirma_Click()
LcLinha = Item.Row
orcamento.Comissao.Text = Item.TextMatrix(LcLinha, 0)
orcamento.ComissaoProduto.Text = Item.TextMatrix(LcLinha, 0)
Unload Me
End Sub

Private Sub Fechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
Inicializacom
'Item.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
GeraGrid
LcTam = 0
'Item.SetFocus
End Sub
Function GeraGrid()
Item.ColAlignment(0) = 7
Item.ColAlignment(1) = 3

Item.ColWidth(0) = 500
Item.ColWidth(1) = 5000


Item.TextMatrix(0, 0) = "Valor"
Item.TextMatrix(0, 1) = "Descrição"


LcTamanhoGrid = 1
End Function

Function Inicializacom()
On Error Resume Next
Dim RsComissao As Recordset
Dim b As Integer
Dim a As Integer
AbreBase
LcSql = "Select * From comissaovendedor where vendedor='" & orcamento.codigoVendedor.Text & "'"
Set RsComissao = Dbbase.OpenRecordset(LcSql)
If RsComissao.EOF Then
   MsgBox "Não Existe Comissão cadastrada para Este Vendedor.", 64, "Aviso"
   FrmVendaOrcam.comisVenda.Text = 0
   Unload Me
   RsComissao.Close
   Set RsComissao = Nothing
   Exit Function
End If
Do Until RsComissao.EOF
  b = 1
  ReDim Preserve LcMat(LcTam)
  LcMat(LcTam).Descricao = RsComissao!Descricao
  LcMat(LcTam).valor = RsComissao!Comissao

  For a = 0 To LcTam
       Item.Rows = b + 1
       Item.TextMatrix(b, 0) = LcMat(a).valor
       Item.TextMatrix(b, 1) = LcMat(a).Descricao
       b = b + 1
  Next
  RsComissao.MoveNext
  LcTam = LcTam + 1
Loop



RsComissao.Close
Dbbase.Close
Set RsComissao = Nothing
Set Dbbase = Nothing
Item.SetFocus

End Function

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub valor_LostFocus()

End Sub

Private Sub Form_Unload(Cancel As Integer)
orcamento.CodigoCliente.SetFocus
LcBuscaVendedor = True
End Sub

Private Sub Item_DblClick()
Dim LcLinha As Long
LcLinha = Item.Row
If LcLinha = 0 Then
   MsgBox "Seleção Inválida, Selecione um Valor na Lista...", 64, "Aviso"
   Exit Sub
End If
orcamento.Comissao.Text = Item.TextMatrix(LcLinha, 0)
orcamento.ComissaoProduto.Text = Item.TextMatrix(LcLinha, 0)
orcamento.CodigoCliente.SetFocus

Unload Me
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    LcLinha = Item.Row
   If LcLinha = 0 Then
      MsgBox "Seleção Inválida, Selecione um Valor na Lista...", 64, "Aviso"
       Exit Sub
   End If
   'If GlComissaoVelha = 0 Then GlComissaoVelha = CCur(FrmVendaOrcam.comisVenda.Text)
   'FrmVendaOrcam.comisVenda.Text = Item.TextMatrix(LcLinha, 0)
    orcamento.Comissao.Text = Item.TextMatrix(LcLinha, 0)
    orcamento.ComissaoProduto.Text = Item.TextMatrix(LcLinha, 0)
    orcamento.CodigoCliente.SetFocus
   Unload Me
End If
If KeyCode = 113 Then
   LcLinha = Item.Row
   If LcLinha = 0 Then
      MsgBox "Seleção Inválida, Selecione um Valor na Lista...", 64, "Aviso"
       Exit Sub
   End If
   orcamento.Comissao.Text = Item.TextMatrix(LcLinha, 0)
   orcamento.ComissaoProduto.Text = Item.TextMatrix(LcLinha, 0)
   CodigoCliente.SetFocus
   Unload Me
End If
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
