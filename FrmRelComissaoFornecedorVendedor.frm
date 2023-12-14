VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelComissaoFornecedorVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissao por Fornecedor"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   5160
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox codigoVendedor 
      Height          =   405
      Left            =   4320
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox vendedor 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apos Imprimir"
      Height          =   1335
      Left            =   6240
      TabIndex        =   22
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton marcar 
         Caption         =   "Não Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton marcarpg 
         Caption         =   "Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situção"
      Height          =   1335
      Left            =   4320
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
      Begin VB.OptionButton todos 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton naoPago 
         Caption         =   "Não Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton pago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   1335
      Left            =   2040
      TabIndex        =   17
      Top             =   2640
      Width           =   2175
      Begin VB.OptionButton sintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton analitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   2640
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Comissao 
      Height          =   315
      ItemData        =   "FrmRelComissaoFornecedorVendedor.frx":0000
      Left            =   120
      List            =   "FrmRelComissaoFornecedorVendedor.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   6720
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox copias 
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
      Left            =   4800
      TabIndex        =   8
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1035
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   8640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6120
      Y1              =   -240
      Y2              =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelComissaoFornecedorVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      Codigo As String
      Nome As String
End Type
Private LcTamanho, LcTamMatVend, a As Integer
Private MtVendedor() As TipoVend, MtMatVendedor() As TipoVend
Private RsComissao As Recordset, RsSintetico As Recordset
Function CarregaVendedor()
On Error GoTo errc
Dim RsVendedor As Recordset, RsVend As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVend = Dbbase.OpenRecordset("Alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcTamanho = 0
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Razaosoc) Then
      ReDim Preserve MtVendedor(LcTamanho)
      MtVendedor(LcTamanho).Codigo = RsVendedor!Codigo
      MtVendedor(LcTamanho).Nome = RsVendedor!Razaosoc & ""
      Comissao.AddItem RsVendedor!Razaosoc
      LcTamanho = LcTamanho + 1
   End If
   RsVendedor.MoveNext
  
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
LcTamMatVend = 0
Do Until RsVend.EOF
  If Not IsNull(RsVend!Nome) Then
    ReDim Preserve MtMatVendedor(LcTamMatVend)
    MtMatVendedor(LcTamMatVend).Codigo = RsVend!Codigo
    MtMatVendedor(LcTamMatVend).Nome = RsVend!Nome & ""
    Vendedor.AddItem RsVend!Nome
    LcTamMatVend = LcTamMatVend + 1
  End If
  RsVend.MoveNext
   
Loop
 If LcTamMatVend > 0 Then LcTamMatVend = LcTamMatVend - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
RsVend.Close
Set RsVend = Nothing
Exit Function
errc:

Exit Function

End Function

Private Sub analitico_Click()
'naoPago.Enabled = True
'pago.Enabled = True
'todos.Enabled = True

End Sub

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Comissao_Click()
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Comissao.Text Then
       Codigo.Text = MtVendedor(a).Codigo
       Exit For
    End If
Next
End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsBaixa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, Lccelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
If Not IsDate(Datai.Text) Then
   MsgBox "A Data Inicial Não é Válida", 64, "Aviso"
   Exit Sub
End If
If Not IsDate(Dataf.Text) Then
   MsgBox "A Data Final Não é Válida", 64, "Aviso"
   Exit Sub
End If
AbreBase
LcBaixa = "select * from Alid201 where Fornecedor='" & Codigo.Text & "' and vendedor='" & codigoVendedor.Text & "' and DATAVENDA>=#" & Datai.Text & "# and DATAVENDA <=#" & Dataf.Text & "#"
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsBaixa = Dbbase.OpenRecordset(LcBaixa)

'Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!bairro
   LcFone = RsEmpresa!fone
End If
'If Not RsOpcao.EOF Then
 '  LcVer = RsOpcao!msg
 '  LcVer1 = RsOpcao!Msg1
'End If

    'Abertura do relatório de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    If analitico Then
       lctitulo = "Relatório de Comissões por Fornecedor << ANALÍTICO >>"
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissaoFornecVend.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaoFornecVendsl.rpt"
       End If
    Else
       lctitulo = "Relatório de Comissões por Fornecedor << SINTÉTICO >>"
       GeraSintetico
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissaosinteticoFornecVend.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaosinteticoFornecVendsl.rpt"
       End If
          
    End If
    
    CryRelatorio.CopiesToPrinter = Val(Copias.Text)
    If analitico And Comissao.Text <> "TODOS" Then
       LcFormula = "{ALID201.Fornecedor} = '" & Codigo.Text & "'"
       LcFormula = LcFormula & " and {ALID201.vendedor} = '" & codigoVendedor.Text & "'"
    End If

  '== Inicio Filtro
  strData = CDate(Format(Datai.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav1 = " date(" & LcDataInicio & ")"
         
  strData = CDate(Format(Dataf.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav2 = " date(" & LcDataInicio & ")"
  If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
  
  
  If analitico Then LcFormula = LcFormula & "{ALID201.DATAVENDA} >=" & LcChav1 & " And {ALID201.DATAVENDA} <=" & LcChav2
  
  If analitico Then
     If pago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ALID201.pago}=True"
     End If
     If naoPago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ALID201.pago}=False"
     End If
  Else
     If pago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{Sintetico.pago}=True"
     End If
     If naoPago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{Sintetico.pago}=False"
     End If
  
   End If
'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = lctitulo

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Comissões por Fornecedor e Vendedor'"
 CryRelatorio.Formulas(4) = "Fornecedor='" & Comissao.Text & "'"
 CryRelatorio.Formulas(5) = "Vendedor='" & Vendedor.Text & "'"
 CryRelatorio.Formulas(6) = "Celular='" & Lccelular & "'"
 CryRelatorio.Formulas(7) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(8) = "email='" & Lcemail & "'"

If Impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close

Set RsOpcao = Nothing
Set RsEmpresa = Nothing

If CryRelatorio.LastErrorNumber > 0 Then
   MsgBox CryRelatorio.LastErrorString
Else
   If marcarpg Then
   Err.Number = 0
     Do Until RsBaixa.EOF
         If Err.Number > 0 Then
            ' MsgBox Err.Description
             Exit Do
          End If
         RsBaixa.Edit
         RsBaixa("pago") = True
         RsBaixa.Update
         RsBaixa.MoveNext
      Loop
   End If
End If
RsBaixa.Close
Set RsBaixa = Nothing
Dbbase.Close
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Dataf_LostFocus()
If Not IsDate(Dataf.Text) And Dataf.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Dataf.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(Datai.Text) And Datai.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Datai.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 4545
Me.Width = 8685
CarregaVendedor

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Impressora_Click()
Copias.Visible = True
Label3.Visible = True
End Sub
Function GeraSintetico()
Dim LcComissao, LcTotal As Currency
Dim LcMuda As Integer
AbreBase
LcCriterio1 = "Select * from alid201 where (Fornecedor='" & Codigo.Text & "') and (vendedor='" & codigoVendedor.Text & "') and (DATAVENDA Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#) order by nf"

Set RsComissao = Dbbase.OpenRecordset(LcCriterio1)
Set RsSintetico = Dbbase.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
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
   Else
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
   End If
   RsComissao.MoveNext
Loop
If LcTotal > 0 Then
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
End If
RsComissao.Close

End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
Dim RsCliente As Recordset
'On Error Resume Next
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
   RsSintetico!fornecedor = RsComissao!fornecedor
   RsSintetico!Cliente = RsCliente!Razaosoc
RsSintetico.Update
RsCliente.Close
Set RsCliente = Nothing
End Function
Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub sintetico_Click()
'naoPago.Enabled = False
'pago.Enabled = False
'todos.Enabled = False

End Sub

Private Sub sintetico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Vendedor_Click()

For a = 0 To LcTamMatVend
    If MtMatVendedor(a).Nome = Vendedor.Text Then
       codigoVendedor.Text = MtMatVendedor(a).Codigo
       Exit For
    End If
Next
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Video_Click()
Copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
