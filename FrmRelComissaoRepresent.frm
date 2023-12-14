VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelComissaoRepresent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissão da Representada"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apos Imprimir"
      Height          =   1335
      Left            =   6240
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton marcar 
         Caption         =   "Não Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton marcarpg 
         Caption         =   "Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situção"
      Height          =   1335
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton todos 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton naoPago 
         Caption         =   "Não Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton pago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   1335
      Left            =   2040
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
      Begin VB.OptionButton sintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton analitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   6720
      TabIndex        =   7
      Top             =   0
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
      TabIndex        =   6
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
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
      TabIndex        =   1
      Top             =   720
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
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   8640
      Y1              =   1440
      Y2              =   1440
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
      TabIndex        =   12
      Top             =   360
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
      TabIndex        =   11
      Top             =   360
      Width           =   1080
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6120
      Y1              =   -240
      Y2              =   1440
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
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelComissaoRepresent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      Codigo As String
      Nome As String
End Type
Private LcTamanho, a As Integer
Private MtVendedor() As TipoVend
Private RsComissao As Recordset, RsSintetico As Recordset
Function CarregaVendedor()
On Error GoTo errc

Exit Function
errc:
MsgBox Err.Description & Err.Number
'Resume 0
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
'For a = 0 To LcTamanho
'    If MtVendedor(a).Nome = Comissao.Text Then
'       codigo.Text = MtVendedor(a).codigo
 '      Exit For
 '   End If
'Next
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
LcBaixa = "select * from ComissaoRepresentante where DATAVENDA>=#" & Datai.Text & "# and DATAVENDA <=#" & Dataf.Text & "#"
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
       lctitulo = "Relatório de Comissões << ANALÍTICO >>"
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissaoRepresentada.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaoRepresentadasl.rpt"
       End If
          
       CryRelatorio.SortFields(0) = "+{ComissaoRepresentante.nf}"
    Else
       lctitulo = "Relatório de Comissões << SINTÉTICO >>"
       GeraSintetico
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissaosinteticoRepresentada.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaosinteticoRepresentadasl.rpt"
       End If
       
       CryRelatorio.SortFields(0) = ""
    End If
    'CryRelatorio.SortFields(0) = "+{ALID201.VENDEDORr}"
    
    CryRelatorio.CopiesToPrinter = Val(Copias.Text)
    'If analitico And Comissao.Text <> "TODOS" Then
    '   LcFormula = "{ComissaoRepresentante.VENDEDOR} = '" & codigo.Text & "'"
   ' End If

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
  
  
  If analitico Then LcFormula = LcFormula & "{ComissaoRepresentante.DATAVENDA} >=" & LcChav1 & " And {ComissaoRepresentante.DATAVENDA} <=" & LcChav2
  
  If analitico Then
     If pago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ComissaoRepresentante.pago}=True"
     End If
     If naoPago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ComissaoRepresentante.pago}=False"
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
 CryRelatorio.Formulas(3) = "titulo='Relatório de Comissões da Representada.'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & Lccelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

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
Me.Height = 3675
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
On Error GoTo ErroSint
Dim LcComissao, LcTotal As Currency
Dim LcMuda, LcGrava As Integer
AbreBase
LcCriterio1 = "Select * from ComissaoRepresentante where "
LcCriterio1 = LcCriterio1 & " DATAVENDA>=#" & Format(Datai.Text, "mm/dd/yyyy") & "# and DATAVENDA <=#" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
LcCriterio1 = LcCriterio1 & " Order by Nf"
'MsgBox LcCriterio1

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
Exit Function
ErroSint:
If Err.Number = 3021 Then Resume Next
MsgBox Err.Description & " " & Err.Number, 16, "Erro Encontrado"
Resume Next

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

Private Sub marcar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub marcarpg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub naoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub pago_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub todos_KeyDown(KeyCode As Integer, Shift As Integer)
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
