VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelCaixaPerido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Caixa por Período"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   2880
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   600
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
      Left            =   2280
      TabIndex        =   4
      Text            =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1440
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
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Período"
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Inicial"
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
      TabIndex        =   10
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Final"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Width           =   585
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelCaixaPerido"
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
Private Sub Cbo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
AbreBase
'==> Exclui os anteriores
Dbbase.Execute "Delete from caixa"
PreencheDespesas
PreencheReceitas
AcertaSaldo
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
If Not RsOpcao.EOF Then
   LcVer = RsOpcao!Msg
   LcVer1 = RsOpcao!Msg1
End If

    'Abertura do relatório de vendas
  CryRelatorio.DiscardSavedData = True
    CryRelatorio.DataFiles(0) = GLBase
    'If analitico Then
       'lctitulo = "Relatório de Comissões << ANALÍTICO >>"
      If GlImprimeSemLinha Then
         CryRelatorio.ReportFileName = App.Path & "\CaixaDetalhe.rpt"
      Else
         CryRelatorio.ReportFileName = App.Path & "\CaixaDetalheSl.rpt"
      End If
    'Else
       'lctitulo = "Relatório de Comissões << SINTÉTICO >>"
   ' End If
    'CryRelatorio.SortFields(0) = "+{ALID201.VENDEDORr}"
    
    CryRelatorio.CopiesToPrinter = Val(Txt1.Text)
    'If Comissao.Text <> "TODOS" Then
       'LcFormula = "{ALID201.VENDEDOR} = '" & codigo.Text & "'"
    'End If

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
  LcFormula = LcFormula & "{caixa.Data} >=" & LcChav1 & " And {caixa.Data} <=" & LcChav2


'== fim filtro
'== fim filtro

CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Caixa por Período"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "EnderecoEmpresa='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Receitas por Período'"
 'CryRelatorio.Formulas(4) = "Versiculo='" & LcVer & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & LcCelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub
Sub PreencheDespesas()
Dim RsDespesas As Recordset
Dim RsCaixa As Recordset

StrSql = "SELECT ALID014.DTVENC, Sum(ALID014.VALOR) AS Valores, Sum(ALID014.VALPAGO) AS ValorPago"
StrSql = StrSql & " From ALID014 "
StrSql = StrSql & " GROUP BY ALID014.DTVENC "
StrSql = StrSql & " HAVING (((ALID014.DTVENC) Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#));"

AbreBase
Set RsDespesas = Dbbase.OpenRecordset(StrSql)
Do Until RsDespesas.EOF
  '===> Verifica se ja tem lancamento para ele
    If IsDate(RsDespesas!DTVENC) Then
        StrSql = "Select * from Caixa where Data=#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#"
        Set RsCaixa = Dbbase.OpenRecordset(StrSql)
        If RsCaixa.EOF Then
           StrSql = "Insert into caixa (Data,Recebimentos,Pagamentos,SaldoAnterior,SaldoAtual,fechado,TrocoDia,TrocoProximoDia) values("
           StrSql = StrSql & "#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & Replace(RsDespesas!Valores, ",", ".") & ","
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0)"
        Else
           StrSql = "Update caixa set Pagamentos=" & Replace(RsDespesas!Valores, ",", ".") & " where data=#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#"
        End If
    End If
 
  Dbbase.Execute StrSql
  RsDespesas.MoveNext
Loop

End Sub
Sub PreencheReceitas()
Dim RsDespesas As Recordset
Dim RsCaixa As Recordset

StrSql = "SELECT alid015.DTVENC, Sum(alid015.VALOR) AS SomaDeVALOR, Sum(alid015.VALPAGO) AS SomaDeVALPAGO"
StrSql = StrSql & " From alid015"
StrSql = StrSql & " GROUP BY alid015.DTVENC"
StrSql = StrSql & " HAVING (((alid015.DTVENC) Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#));"

AbreBase
Debug.Print StrSql
Set RsDespesas = Dbbase.OpenRecordset(StrSql)
Do Until RsDespesas.EOF
  '===> Verifica se ja tem lancamento para ele
  If IsDate(RsDespesas!DTVENC) Then
        StrSql = "Select * from Caixa where Data=#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#"
        Set RsCaixa = Dbbase.OpenRecordset(StrSql)
        If RsCaixa.EOF Then
           StrSql = "Insert into caixa (Data,Recebimentos,Pagamentos,SaldoAnterior,SaldoAtual,fechado,TrocoDia,TrocoProximoDia) values("
           StrSql = StrSql & "#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#,"
           StrSql = StrSql & Replace(RsDespesas!SomaDeVALOR, ",", ".") & ","
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0,"
           StrSql = StrSql & "0)"
        Else
           StrSql = "Update caixa set Recebimentos=" & Replace(RsDespesas!SomaDeVALOR, ",", ".") & " where data=#" & Format(RsDespesas!DTVENC, "mm/dd/yy") & "#"
        End If
  End If
 
  Dbbase.Execute StrSql
  RsDespesas.MoveNext
Loop

End Sub
Sub AcertaSaldo()

Dim RsCaixa As Recordset
Dim Saldo As Currency
StrSql = "Select * from Caixa order by Data"
AbreBase
Set RsCaixa = Dbbase.OpenRecordset(StrSql)

Do Until RsCaixa.EOF
   Dim Saldodia As Currency
   Saldodia = RsCaixa!Recebimentos - RsCaixa!Pagamentos
   Saldo = Saldo + Saldodia
   StrSql = "Update caixa set TrocoDia=" & Replace(Saldodia, ",", ".") & ",SaldoAtual=" & Replace(Saldo, ",", ".") & " where data=#" & Format(RsCaixa!Data, "mm/dd/yy") & "#"
   Dbbase.Execute StrSql
   Debug.Print StrSql
   RsCaixa.MoveNext
Loop
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
Me.Height = 3465
Me.Width = 5370


End Sub

Private Sub Impressora_Click()
copias.Visible = True
Label3.Visible = True
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub sintetico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Video_Click()
copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
