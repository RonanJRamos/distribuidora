VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BaixaComissaoBel 
   Caption         =   "Fechamento de Comissões"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LucroP 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Lucro 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox PercComissao 
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox TSelecionada 
      Height          =   375
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox tcomissao 
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox tVendas 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Confirma Lançamento  F3"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Marcar &Todos Como Pago  F5"
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   7320
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Marcar Todos como Não Pago  F4"
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exibição"
      Height          =   1095
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option5 
         Caption         =   "Sintético"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Analítico"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1095
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Não Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Comissao 
      Height          =   5175
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Executa F2"
      Height          =   975
      Left            =   7920
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin MSMask.MaskEdBox datai 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label9 
      Caption         =   "Lucro  %"
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
      Left            =   3720
      TabIndex        =   30
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Lucro  R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Perc. Comissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Total de Sel."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   24
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Total Comissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "T. de Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Data Final"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Data Inicial"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "BaixaComissaoBel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Vendedor
        Codigo As String
        Nome As String
End Type
Private LcTamanhoGrid As Long
Private MtVendedor() As Vendedor
Private LcTGrid, a As Long
Private LcLucro As Double
Private LcCucroP As Double
Private RsComissao As Recordset, RsCliente As Recordset, RsSintetico As Recordset

Private Sub Comissao_DblClick()
On Error Resume Next
Dim LcLinha As Integer
LcLinha = Comissao.Row
If Comissao.TextMatrix(LcLinha, 5) = "Sim" Then
   Comissao.TextMatrix(LcLinha, 5) = "Não"
   If Len(TSelecionada.Text) > 0 Then
     TSelecionada.Text = Format(CDbl(TSelecionada.Text) - CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
   End If
Else
  Comissao.TextMatrix(LcLinha, 5) = "Sim"
  
  If Len(TSelecionada.Text) > 0 Then
    TSelecionada.Text = Format(CDbl(TSelecionada.Text) + CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  Else
    TSelecionada.Text = Format(CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  End If
End If
End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
If Comissao.TextMatrix(LcLinha, 5) = "Sim" Then
   Comissao.TextMatrix(LcLinha, 5) = "Não"
   If Len(TSelecionada.Text) > 0 Then
     TSelecionada.Text = Format(CDbl(TSelecionada.Text) - CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
   End If
Else
  Comissao.TextMatrix(LcLinha, 5) = "Sim"
  
  If Len(TSelecionada.Text) > 0 Then
    TSelecionada.Text = Format(CDbl(TSelecionada.Text) + CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  Else
    TSelecionada.Text = Format(CDbl(Comissao.TextMatrix(LcLinha, 4)), "Currency")
  End If
End If
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcCap As String
If Len(Vendedor.Text) = 0 Then
   MsgBox "Escolha o Vendedor Para Listar as Comissões...", 64, "Aviso"
   Vendedor.SetFocus
   Exit Sub
End If

If datai.Text = "  /  /  " Then
   MsgBox "Escolha a Data Inicial do Periodo ...", 64, "Aviso"
   Vendedor.SetFocus
   Exit Sub
End If
If Dataf.Text = "  /  /  " Then
   Dataf.Text = Date
   'Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, Filtrando Registros..."
Comissao.Rows = 1
If Option4 Then montagrid
If Option5 Then GeraSintetico
Me.Caption = LcCap
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
For a = 1 To Comissao.Rows - 1
    Comissao.TextMatrix(a, 5) = "Não"
Next
TSelecionada.Text = ""

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim LcTotal As Double
For a = 1 To Comissao.Rows - 1
    Comissao.TextMatrix(a, 5) = "Sim"
    LcTotal = LcTotal + CDbl(Comissao.TextMatrix(a, 4))
Next
TSelecionada.Text = Format(LcTotal, "Currency")
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command4_Click()
' On Error Resume Next
If Option4 Then BaixaComissaoAnalitico
If Option5 Then Baixasintetico
Comissao.Rows = 1
Vendedor.Text = ""
datai.Text = "  /  /  "
Dataf.Text = "  /  /  "
tVendas.Text = ""
tcomissao.Text = ""
TSelecionada.Text = ""
End Sub
Function BaixaComissaoAnalitico()
On Error Resume Next
Dim RsComissao As Recordset
AbreBase
Set RsComissao = Dbbase.OpenRecordset("alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
For a = 1 To Comissao.Rows - 1
    If Comissao.TextMatrix(a, 5) = "Sim" Then
       LcCri = "Codigo=" & Val(Comissao.TextMatrix(a, 6))
       RsComissao.FindFirst LcCri
       If Not RsComissao.NoMatch Then
          RsComissao.Edit
          RsComissao("pago") = True
          RsComissao.Update
       End If
    Else
       LcCri = "Codigo=" & Val(Comissao.TextMatrix(a, 6))
       RsComissao.FindFirst LcCri
       If Not RsComissao.NoMatch Then
          RsComissao.Edit
          RsComissao("pago") = False
          RsComissao.Update
       End If
    End If
Next
RsComissao.Close

End Function
Function Baixasintetico()
On Error Resume Next
Dim RsComissao As Recordset
AbreBase

For a = 1 To Comissao.Rows - 1
   If Comissao.TextMatrix(a, 5) = "Sim" Then
      LcSql = "select * from alid201 where nf='" & Comissao.TextMatrix(a, 0) & "'"
      Set RsComissao = Dbbase.OpenRecordset(LcSql) '"alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Do Until RsComissao.EOF
         RsComissao.Edit
         RsComissao("pago") = True
         RsComissao.Update
         RsComissao.MoveNext
      Loop
   Else
     LcSql = "select * from alid201 where nf='" & Comissao.TextMatrix(a, 0) & "'"
      Set RsComissao = Dbbase.OpenRecordset(LcSql) '"alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Do Until RsComissao.EOF
         RsComissao.Edit
         RsComissao("pago") = False
         RsComissao.Update
         RsComissao.MoveNext
      Loop
   End If
Next
'RsComissao.Close

End Function

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command5_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Dataf_LostFocus()
On Error Resume Next
If Dataf.Text = "  /  /  " Then Exit Sub
If Not IsDate(Dataf.Text) Then
   MsgBox "Data Inválida", 64, "Aviso"
   Dataf.SetFocus
Else
    If IsDate(datai.Text) Then
        CalculaLucroVenda
    End If
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Datai_LostFocus()
On Error Resume Next
If datai.Text = "  /  /  " Then Exit Sub
If Not IsDate(datai.Text) Then
   MsgBox "Data Inválida", 64, "Aviso"
   datai.SetFocus
Else
    If IsDate(Dataf.Text) Then
        CalculaLucroVenda
    End If
End If
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraVendedor
GeraGrid
End Sub

Function GeraVendedor()
On Error Resume Next
Dim RsVendedor As Recordset
LcTGrid = 0
LcCriSql = "VENDEDOR='"
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("select * from alid200 order by nome") ', dbOpenDynaset)
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Nome) Then
      If err.Number > 0 Then Exit Do
      ReDim Preserve MtVendedor(LcTGrid)
      MtVendedor(LcTGrid).Codigo = RsVendedor!Codigo
      MtVendedor(LcTGrid).Nome = RsVendedor!Nome
      Vendedor.AddItem RsVendedor!Nome
      LcTGrid = LcTGrid + 1
    End If
    RsVendedor.MoveNext
Loop
LcTGrid = LcTGrid - 1
RsVendedor.Close
Set RsVendedor = Nothing

End Function
Function GeraGrid()
On Error Resume Next
Comissao.ColAlignment(0) = 1
Comissao.ColAlignment(1) = 1
Comissao.ColAlignment(2) = 1
Comissao.ColAlignment(3) = 7
Comissao.ColAlignment(4) = 7
Comissao.ColAlignment(5) = 1

Comissao.ColWidth(0) = 950
Comissao.ColWidth(1) = 1100
Comissao.ColWidth(2) = 5000
Comissao.ColWidth(3) = 1000
Comissao.ColWidth(4) = 1000
Comissao.ColWidth(5) = 700
Comissao.ColWidth(6) = 0
Comissao.TextMatrix(0, 0) = "Documento"
Comissao.TextMatrix(0, 1) = "Pag.Comissão"
Comissao.TextMatrix(0, 2) = "Cliente"
Comissao.TextMatrix(0, 3) = "V.Venda"
Comissao.TextMatrix(0, 4) = "Comissão"
Comissao.TextMatrix(0, 5) = "Pago"

LcTamanhoGrid = 1
End Function
Function montagrid()
On Error Resume Next
Dim bb As Database
Dim RsProduto As Recordset
Dim LcCodigoVendedor As String
Dim LcTotalVendas, LcTotalComissao, LcTotalSelec As Double
Dim LcPago As Integer
If Option1 Then LcPago = False
If Option2 Then LcPago = True
'=== Busca Codigo Vendedor
For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).Codigo
       Exit For
    End If
Next
LcCriSql = "select * from alid201 where VENDEDOR='" & LcCodigoVendedor & "' And DATAVENDA >= #" & Format(datai.Text, "mm/dd/yyyy") & "# and datavenda <= #" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
If Not Option3 Then
   LcCriSql = LcCriSql & " And pago=" & LcPago
End If
LcCriSql = LcCriSql & " order by nf"
Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsComissao = bb.OpenRecordset(LcCriSql) ', dbOpenDynaset)
Set RsCliente = bb.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsProduto = bb.OpenRecordset("alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

LcTamanho = Comissao.Rows
a = 2
Me.Caption = msg
Comissao.Rows = 1
LcAchou = False
Comissao.TextMatrix(0, 2) = "Produto"
Do Until RsComissao.EOF
  LcAchou = True
  If Len(Trim(RsComissao!NF)) > 0 Then
   If Not IsNull(RsComissao!NF) Then
       
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsComissao!NF & ""
     Comissao.TextMatrix(a - 1, 1) = RsComissao!DATAVENDA & ""
     LcPesq = "cod='" & RsComissao!Produto & "'"
     RsProduto.FindFirst LcPesq
     Comissao.Rows = a
     If Not RsProduto.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsProduto!Nome & ""
     End If
    ' Comissao.TextMatrix(a - 1, 2) = RsComissao!Cliente & ""
     Comissao.TextMatrix(a - 1, 3) = Format(RsComissao!ValorTotal, "currency")
     Comissao.TextMatrix(a - 1, 4) = Format(RsComissao!Comissao, "currency")
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsComissao!ValorTotal, "currency"))
     LcTotalComissao = LcTotalComissao + CDbl(Format(RsComissao!Comissao, "Currency"))
     
     
     If RsComissao!pago Then
        Comissao.TextMatrix(a - 1, 5) = "Sim"
        LcTotalSelec = LcTotalSelec + CDbl(Format(RsComissao!Comissao, "Currency"))
     Else
        Comissao.TextMatrix(a - 1, 5) = "Não"
     End If
     Comissao.TextMatrix(a - 1, 6) = RsComissao!Codigo
     a = a + 1
    End If
   End If
   RsComissao.MoveNext
Loop
tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
End Function
Function GeraSintetico()
On Error Resume Next
Dim LcComissao, LcTotal As Currency
Dim LcMuda, LcGrava As Integer
Dim LcTotalSelec As Currency
Dim LcPago As Integer
Dim RsCliente As Recordset

For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).Codigo
       Exit For
    End If
Next

AbreBase
LcCriterio1 = "Select * from alid201 where VENDEDOR='" & LcCodigoVendedor & "' and "
LcCriterio1 = LcCriterio1 & " DATAVENDA>=#" & Format(datai.Text, "mm/dd/yyyy") & "# and DATAVENDA <=#" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
'LcCriterio1 = LcCriterio1 & " Order by Nf"
If Option1 Then LcPago = 0
If Option2 Then LcPago = -1
If Not Option3 Then
   LcCriterio1 = LcCriterio1 & " And pago=" & LcPago
End If
LcCriterio1 = LcCriterio1 & " order by nf"

'MsgBox LcCriterio1
Set RsCliente = Dbbase.OpenRecordset("alid001") ', dbOpenDynaset)

Comissao.TextMatrix(0, 2) = "Cliente"

'MsgBox LcCriterio1
Set RsComissao = Dbbase.OpenRecordset(LcCriterio1)
Set RsSintetico = Dbbase.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTotalSelec = 0
Do Until RsSintetico.EOF
   RsSintetico.Delete
   RsSintetico.MoveNext
Loop
LcNota = RsComissao!NF
Do Until RsComissao.EOF
   If LcMuda Then
      LcNota = RsComissao!NF
      LcMuda = False
   End If
   If LcNota = RsComissao!NF Then
      LcComissao = LcComissao + RsComissao!Comissao
      LcTotal = LcTotal + RsComissao!ValorTotal
      LcGrava = True
   Else
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
      LcGrava = False
   End If
   RsComissao.MoveNext
   
Loop
If LcGrava Then
   RsComissao.MovePrevious
   Call GravaSintetico(LcComissao, LcTotal)
   LcGrava = False
End If
RsSintetico.MoveFirst
LcTamanho = Comissao.Rows
a = 2
Do Until RsSintetico.EOF
  LcAchou = True
  If Len(Trim(RsSintetico!NF)) > 0 Then
   If Not IsNull(RsSintetico!NF) Then
     LcPesq = "codigo='" & RsSintetico!Cliente & "'"
    ' RsCliente.FindFirst LcPesq
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsSintetico!NF & ""
     Comissao.TextMatrix(a - 1, 1) = RsSintetico!DATAVENDA & ""
     'If Not RsCliente.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsSintetico!Cliente & ""
     'End If
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsSintetico!ValorTotal, "currency"))
     LcTotalComissao = LcTotalComissao + CDbl(Format(RsSintetico!Comissao, "Currency"))
     Comissao.TextMatrix(a - 1, 3) = Format(RsSintetico!ValorTotal, "currency")
     Comissao.TextMatrix(a - 1, 4) = Format(RsSintetico!Comissao, "currency")

     If RsSintetico!pago Then
        Comissao.TextMatrix(a - 1, 5) = "Sim"
        LcTotalSelec = LcTotalSelec + CCur(RsComissao!Comissao)
     Else
        Comissao.TextMatrix(a - 1, 5) = "Não"
     End If
     a = a + 1
    End If
   End If
   RsSintetico.MoveNext
Loop
tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
Lucro.Text = AcertaNumero(CCur(tVendas.Text) - CCur(LcLucro), 2)
End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
Dim RsCliente As Recordset
On Error Resume Next
LcCriterio22 = "Select * from alid001 where codigo='" & RsComissao!Cliente & "'"
Set RsCliente = Dbbase.OpenRecordset(LcCriterio22)

RsSintetico.AddNew
   RsSintetico!Vendedor = RsComissao!Vendedor
   RsSintetico!NF = RsComissao!NF
   RsSintetico!Comissao = LcComissao
   RsSintetico!ValorTotal = LcTotal
   RsSintetico!ItemBaixo = RsComissao!ItemBaixo
   RsSintetico!DATAVENDA = RsComissao!DATAVENDA
    RsSintetico!pago = RsComissao!pago
   RsSintetico!Cliente = RsCliente!razaosoc
RsSintetico.Update
RsCliente.Close
Set RsCliente = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Option5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub tcomissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TSelecionada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub tVendas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
Sub CalculaLucroVenda()
On Error Resume Next
Dim LcSql As String
Dim Lucro As Double
Dim Rs As adodb.Recordset
Dim a As Long
For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).Codigo
       Exit For
    End If
Next
AbreBase
LcSql = "SELECT ALID050.NUMNF, ALID050.Vendedor, ALID050.DTEMIS, ALID050.status, ALID052.codProd, ALID052.QTDE, ALID052.VALUNIT, ALID052.QTDUM"
LcSql = LcSql & " FROM ALID050 INNER JOIN ALID052 ON ALID050.NUMNF = ALID052.NUMNF where ALID050.Vendedor='" & LcCodigoVendedor & "'"
LcSql = LcSql & " and (ALID050.DTEMIS between #" & Format(CDate(datai.Text), "mm/dd/yy") & "# and #" & Format(CDate(Dataf.Text), "mm/dd/yy") & "#)"
Set Rs = AbreRecordset(LcSql, True)
LcCap = Me.Caption
Me.Caption = "Buscando Venda do Período. Aguarde..."
Do Until Rs.EOF
    LcLucro = LcLucro + VerificaLucratividade(Rs!codProd, Rs!VALUNIT, Rs!QTDE, Rs!QTDUM)
    Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
Me.Caption = LcCap
End Sub
Function VerificaLucratividade(CodProduto As String, LcValor As Double, LcQuant As Double, Com As Double) As Double
On Error Resume Next
Dim RsL     As adodb.Recordset
Dim LcSql   As String
Dim LcCusto As Double
Dim LcCustoBase As Double
Dim LcComBase As Double
Dim LcLucro As Double
Dim a       As Integer
Dim StrSql As String
'======BuscaPrecoCusto
LcCusto = 0
StrSql = "Select * from produtos where codigo=" & CodProduto
Set RsL = AbreRecordset(StrSql, True)
If Not RsL.EOF Then
    DoEvents
    LcComBase = RsL!QtdMedida
    LcCustoBase = CCur(RsL!Custo / LcComBase) * LcQuant * Com
    LcCusto = CCur(LcCusto) + LcCustoBase
End If
VerificaLucratividade = CCur(LcCusto) '
'LcLucro = (LcLucro * 100) / LcCusto
'Lucratividade.Text = AcertaNumero(CCur(LcLucro), 3)
End Function

Private Sub Vendedor_LostFocus()
If IsDate(datai.Text) And IsDate(Dataf.Text) Then
    CalculaLucroVenda
End If
End Sub
