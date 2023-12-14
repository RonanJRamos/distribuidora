VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fichadeestoque 
   BackColor       =   &H00B6BEA3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Estoque"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Com 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   14673105
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exibir Ficha"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox nome 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Com"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "fichadeestoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub codigo_Change()
On Error Resume Next
Command1.Enabled = Len(Codigo.Text)
Command1.Default = True
End Sub

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "%+{E}"
'If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Codigo_LostFocus()
On Error Resume Next
Dim Rs As ADODB.Recordset
If Len(Codigo.Text) = 0 Then Exit Sub
Set Rs = AbreRecordset("Select * from produtos where codigo=" & Codigo.Text, True)
If Not Rs.EOF Then
   Nome.Text = Rs!Nome
   Com.Text = Rs!QtdMedida
   Call Command1_Click
Else
   Nome.Text = ""
End If
Set Rs = Nothing

End Sub



Private Sub Command1_Click()
'On Error Resume Next
Dim Rs As ADODB.Recordset
Dim Estoque As ControleDb
Dim LcSaldo As Double
Dim LcSaldoUnit As Double
Dim Cl As New ControleEstoque

'Estoque.CodProduto = codigo.Text
Dim a As Long
Set Rs = AbreRecordset("Select * from historicoproduto where produto='" & Codigo.Text & "' order by codigo", True)
a = Rs.RecordCount
grid.Rows = Rs.RecordCount + 1
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o historico..."
Screen.MousePointer = 11
Do Until Rs.EOF
   If err.Number > 0 Then Exit Do
   Dim LcAnterior As Currency
   Dim LcQuantidade As Currency
   Dim LcQuantUnidade As Currency
   Dim LcCaixa As Currency
   Dim LcUnidade As Currency
   Dim LcCom As Integer
   If IsNumeric(Com.Text) Then LcCom = Com.Text Else LcCom = 1
   If IsNumeric(Rs!Anterior) Then
     
       Transforna_Unidade Rs!Anterior, LcCom, LcAnterior, LcQuantUnidade
   Else
      LcAnterior = 0
   End If

   Transforna_Unidade CCur(Rs!Santa2 + Rs!california), LcCom, LcQuantidade, LcQuantUnidade
   If IsNumeric(Rs!Saldo) Then
      
      Transforna_Unidade Rs!Saldo, LcCom, LcCaixa, LcUnidade
   Else
      LcCaixa = 0
      LcUnidade = 0
   End If
  
   'grid.Rows = a + 1
   grid.TextMatrix(a, 0) = Rs!NF
   grid.TextMatrix(a, 1) = Rs!Tipo
   grid.TextMatrix(a, 2) = Rs!clienteforn
   grid.TextMatrix(a, 3) = Format(Rs!Data, "dd/mm/yy")
   grid.TextMatrix(a, 4) = LcAnterior 'Rs!Anterior  ' Rs!santa + Rs!Santa2 + Rs!California
   
   
   grid.TextMatrix(a, 5) = LcQuantidade 'Rs!Santa2 + Rs!California
   grid.TextMatrix(a, 6) = LcCaixa 'Rs!saldo ' & " Unidades." ' & " e " & LcSaldoUnit & "Unid(s)."
   grid.TextMatrix(a, 7) = LcUnidade
   a = a - 1
   Rs.MoveNext
Loop
Me.Caption = LcCap
Screen.MousePointer = 0
Set Estoque = Nothing
Set Rs = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
GeraGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmPrincipal.SetFocus
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    GlCriterioSql = "select * From produtos where nome like '" & UCase(Nome.Text) & "%'  order by nome"
    FrmPesquisaProdutos.Show , Me
End If
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub nome_LostFocus()
If Len(Codigo.Text) > 0 Then Exit Sub
GlCriterioSql = "select * From produtos where nome like '" & UCase(Nome.Text) & "%'  order by nome"
FrmPesquisaProdutos.Show , Me
End Sub
Function GeraGrid()
On Error Resume Next
grid.ColAlignment(0) = 7
grid.ColAlignment(1) = 1
grid.ColAlignment(2) = 1
grid.ColAlignment(3) = 7
grid.ColAlignment(4) = 7
grid.ColAlignment(6) = 7
grid.ColAlignment(7) = 7
grid.TextMatrix(0, 0) = "NF"
grid.TextMatrix(0, 1) = "Tipo"
grid.TextMatrix(0, 2) = "Ciente/Fornecedor"
grid.TextMatrix(0, 3) = "Emissão"
grid.TextMatrix(0, 4) = "Anterior"
grid.TextMatrix(0, 5) = "Quantidade"
grid.TextMatrix(0, 6) = "Saldo"
grid.TextMatrix(0, 7) = "Quant em Unidade"

grid.ColWidth(0) = 1000
grid.ColWidth(1) = 500
grid.ColWidth(2) = 4800
grid.ColWidth(3) = 1000
grid.ColWidth(4) = 1000
grid.ColWidth(5) = 1000
grid.ColWidth(6) = 1000
grid.ColWidth(7) = 1000
End Function
