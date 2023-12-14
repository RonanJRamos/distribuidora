VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form exibeitem 
   BackColor       =   &H00D8C5B6&
   Caption         =   "Detalhes do Item"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid item 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "exibeitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
Dim LcSql As String
Dim Rs As Recordset
Dim CEST As ControleDb
Dim LcEst As String
Dim LcUni As String
EscreveGrid
Set CEST = New ControleDb
CEST.CodProduto = Me.Tag
'LcEst = Cest.EstoqueTotalFechado
'LcSql = "Select codigo, nome,preco,minimovenda,Fix(QuantEstoque/QtdMedida) AS Est,Round((((QuantEstoque / QtdMedida) - Fix(QuantEstoque / QtdMedida)) * QtdMedida), 0) AS EstU from produto where codigo=" & Me.Tag
'LcSql = "Select codigo, nome,preco,minimovenda from produtos where codigo=" & Me.Tag
AbreBase
Set Rs = Dbbase.OpenRecordset("select * from alid004 where cod='" & CEST.CodigoDaUnidade & "'")
If Not Rs.EOF Then
   LcUni = Rs!Simbolo & ""
End If
Set Rs = Nothing
'Debug.Print LcSql
'Set rs = AbreRecordsetLeitura(LcSql)
'Set grid.DataSource = rs
'grid.Row = 0
Item.TextMatrix(1, 0) = Me.Tag
Item.TextMatrix(1, 1) = CEST.DescricaoProduto & "  " & LcUni & " c/" & CEST.QuantidadeDaUnidade
Item.TextMatrix(1, 2) = CEST.EstoqueTotalFechado
Item.TextMatrix(1, 3) = CEST.EstoqueTotalUnitario
Item.TextMatrix(1, 4) = CEST.PrecoVenda
Item.TextMatrix(1, 5) = CEST.PrecoMinimo
If IsNumeric(CEST.LimiteVenda) Then
   Item.TextMatrix(1, 6) = CEST.LimiteVenda
Else
   Item.TextMatrix(1, 6) = 0
End If

'Data1.DatabaseName = GLBase
'Data1.RecordSource = LcSql
'Data1.Refresh
Set CEST = Nothing
End Sub
Function EscreveGrid()
On Error Resume Next
Item.TextMatrix(0, 0) = "Código"
Item.TextMatrix(0, 1) = "Descrição"
Item.TextMatrix(0, 2) = "Estoque"
Item.TextMatrix(0, 3) = "Est. Unit"
Item.TextMatrix(0, 4) = "Preço Venda"
Item.TextMatrix(0, 5) = "Preço Min"
Item.TextMatrix(0, 6) = "Preço Limite"
Item.ColWidth(0) = 1244
Item.ColWidth(1) = 4734
Item.ColWidth(2) = 1144
Item.ColWidth(3) = 1114
Item.ColWidth(4) = 1095
Item.ColWidth(5) = 1270
Item.ColWidth(6) = 1270
End Function
Private Sub Form_Load()
Me.Top = 800
Me.Left = 250
End Sub

