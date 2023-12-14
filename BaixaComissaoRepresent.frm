VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form BaixaComissaoRepresent 
   Caption         =   "Fechamento de Comissões"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TSelecionada 
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox tcomissao 
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox tVendas 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Confirma Lançamento F3"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Marcar &Todos Como Pago F5"
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Marcar Todos como Não Pago F4"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exibição"
      Height          =   975
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton Option5 
         Caption         =   "Sintético"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Analítico"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1095
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   1455
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
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Executa F2"
      Height          =   975
      Left            =   8520
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin MSMask.MaskEdBox datai 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      Caption         =   "Total de Sel.  para Pagar"
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
      Left            =   5040
      TabIndex        =   24
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Total de Comissão"
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
      Left            =   5040
      TabIndex        =   22
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Total de Vendas"
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
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Data Final"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Data Inicial"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
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
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "BaixaComissaoRepresent"
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
'If Len(Vendedor.Text) = 0 Then
'   MsgBox "Escolha o Vendedor Para Listar as Comissões...", 64, "Aviso"
'   Vendedor.SetFocus
'   Exit Sub
'' End If

If Datai.Text = "  /  /  " Then
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
If Option4 Then montagrid
If Option5 Then GeraSintetico
Me.Caption = LcCap
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
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
If KeyCode = 13 Then SendKeys "{TAB}"
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
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Option4 Then BaixaComissaoAnalitico
If Option5 Then Baixasintetico
Comissao.Rows = 1
Vendedor.Text = ""
Datai.Text = "  /  /  "
Dataf.Text = "  /  /  "
tVendas.Text = ""
tcomissao.Text = ""
TSelecionada.Text = ""
End Sub
Function BaixaComissaoAnalitico()
On Error Resume Next
Dim RsComissao As Recordset
AbreBase
Set RsComissao = Dbbase.OpenRecordset("ComissaoRepresentante", dbOpenDynaset, dbSeeChanges, dbOptimistic)
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
      LcSql = "select * from ComissaoRepresentante where nf='" & Comissao.TextMatrix(a, 0) & "'"
      Set RsComissao = Dbbase.OpenRecordset(LcSql) '"alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Do Until RsComissao.EOF
         RsComissao.Edit
         RsComissao("pago") = True
         RsComissao.Update
         RsComissao.MoveNext
      Loop
   Else
      LcSql = "select * from ComissaoRepresentante where nf='" & Comissao.TextMatrix(a, 0) & "'"
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
If KeyCode = 13 Then SendKeys "{TAB}"
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
If KeyCode = 13 Then SendKeys "{TAB}"
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
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"


End Sub

Private Sub Datai_LostFocus()
On Error Resume Next
If Datai.Text = "  /  /  " Then Exit Sub
If Not IsDate(Datai.Text) Then
   MsgBox "Data Inválida", 64, "Aviso"
   Datai.SetFocus
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
Exit Function
Dim RsVendedor As Recordset
LcTGrid = 0
LcCriSql = "VENDEDOR='"
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("select * from alid200 order by nome") ', dbOpenDynaset)
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Nome) Then
      If Err.Number > 0 Then Exit Do
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
LcCriSql = "select * from ComissaoRepresentante where DATAVENDA >= #" & Format(Datai.Text, "mm/dd/yyyy") & "# and datavenda <= #" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
If Not Option3 Then
   LcCriSql = LcCriSql & " And pago=" & LcPago
End If
LcCriSql = LcCriSql & " order by nf"
AbreBase
Set RsComissao = Dbbase.OpenRecordset(LcCriSql) ', dbOpenDynaset)
Set RsCliente = Dbbase.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsProduto = Dbbase.OpenRecordset("alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

LcTamanho = Comissao.Rows
Comissao.TextMatrix(0, 2) = "Produto"
a = 2
Me.Caption = msg
Comissao.Rows = 1
LcAchou = False
Do Until RsComissao.EOF
  LcAchou = True
  If Len(Trim(RsComissao!nf)) > 0 Then
   If Not IsNull(RsComissao!nf) Then
       
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsComissao!nf & ""
     Comissao.TextMatrix(a - 1, 1) = RsComissao!DATAVENDA & ""
     LcPesq = "cod='" & RsComissao!Produto & "'"
     RsProduto.FindFirst LcPesq
     Comissao.Rows = a
     If Not RsProduto.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsProduto!Nome & ""
     End If
    ' Comissao.TextMatrix(a - 1, 2) = RsComissao!Cliente & ""
     Comissao.TextMatrix(a - 1, 3) = Format(RsComissao!VALORTOTAL, "currency")
     Comissao.TextMatrix(a - 1, 4) = Format(RsComissao!Comissao, "currency")
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsComissao!VALORTOTAL, "currency"))
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
If Not LcAchou Then
   Comissao.Rows = 1
End If

tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
End Function
Function GeraSintetico()
On Error Resume Next
Dim LcComissao, LcTotal As Currency
Dim LcMuda, LcGrava As Integer
Dim LcTotalSelec As Currency

Dim RsCliente As Recordset

For a = 0 To LcTGrid
    If MtVendedor(a).Nome = Vendedor.Text Then
       LcCodigoVendedor = MtVendedor(a).Codigo
       Exit For
    End If
Next
Comissao.TextMatrix(0, 2) = "Cliente"
AbreBase
LcCriterio1 = "Select * from ComissaoRepresentante where "
LcCriterio1 = LcCriterio1 & "DATAVENDA>=#" & Format(Datai.Text, "mm/dd/yyyy") & "# and DATAVENDA <=#" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
'LcCriterio1 = LcCriterio1 & " Order by Nf"
If Option1 Then LcPago = 0
If Option2 Then LcPago = -1
If Not Option3 Then
   LcCriterio1 = LcCriterio1 & " And pago=" & LcPago
End If
LcCriterio1 = LcCriterio1 & " order by nf"

'MsgBox LcCriterio1
Set RsCliente = Dbbase.OpenRecordset("alid001") ', dbOpenDynaset)



'MsgBox LcCriterio1
Set RsComissao = Dbbase.OpenRecordset(LcCriterio1)
Set RsSintetico = Dbbase.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTotalSelec = 0
Do Until RsSintetico.EOF
   RsSintetico.Delete
   RsSintetico.MoveNext
Loop
LcNota = RsComissao!nf
Do Until RsComissao.EOF
   If LcMuda Then
      LcNota = RsComissao!nf
      LcMuda = False
   End If
   If LcNota = RsComissao!nf Then
      LcComissao = LcComissao + RsComissao!Comissao
      LcTotal = LcTotal + RsComissao!VALORTOTAL
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
  
  If Len(Trim(RsSintetico!nf)) > 0 Then
   If Not IsNull(RsSintetico!nf) Then
     LcAchou = True
     LcPesq = "codigo='" & RsSintetico!Cliente & "'"
    ' RsCliente.FindFirst LcPesq
     Comissao.Rows = a
     Comissao.TextMatrix(a - 1, 0) = RsSintetico!nf & ""
     Comissao.TextMatrix(a - 1, 1) = RsSintetico!DATAVENDA & ""
     'If Not RsCliente.NoMatch Then
       Comissao.TextMatrix(a - 1, 2) = RsSintetico!Cliente & ""
     'End If
     LcTotalVendas = LcTotalVendas + CDbl(Format(RsSintetico!VALORTOTAL, "currency"))
     LcTotalComissao = LcTotalComissao + CDbl(Format(RsSintetico!Comissao, "Currency"))
     Comissao.TextMatrix(a - 1, 3) = Format(RsSintetico!VALORTOTAL, "currency")
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
If Not LcAchou Then
   Comissao.Rows = 1
End If
tVendas.Text = Format(LcTotalVendas, "Currency")
tcomissao.Text = Format(LcTotalComissao, "currency")
TSelecionada.Text = Format(LcTotalSelec, "currency")
End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
Dim RsCliente As Recordset
On Error Resume Next
LcCriterio22 = "Select * from alid001 where codigo='" & RsComissao!Cliente & "'"
Set RsCliente = Dbbase.OpenRecordset(LcCriterio22)

RsSintetico.AddNew
   RsSintetico!Vendedor = RsComissao!Vendedor
   RsSintetico!nf = RsComissao!nf
   RsSintetico!Comissao = LcComissao
   RsSintetico!VALORTOTAL = LcTotal
   RsSintetico!ITEMBAIXO = RsComissao!ITEMBAIXO
   RsSintetico!DATAVENDA = RsComissao!DATAVENDA
    RsSintetico!pago = RsComissao!pago
   RsSintetico!Cliente = RsCliente!Razaosoc
RsSintetico.Update
RsCliente.Close
Set RsCliente = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Option5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub tcomissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub TSelecionada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub tVendas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{M}"
If KeyCode = 116 Then SendKeys "%+{T}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub
