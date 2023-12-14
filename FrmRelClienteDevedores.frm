VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClienteDevedores 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Clientes Devedores"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4440
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   3960
      TabIndex        =   4
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
      TabIndex        =   3
      Text            =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Saída"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton impressora 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
   Begin MSMask.MaskEdBox dataf 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
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
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento Inferior a"
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
      TabIndex        =   8
      Top             =   120
      Width           =   2205
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
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelClienteDevedores"
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
Function GeraNota()
On Error GoTo errGera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String
'LcFormula = LcFormula & "{ALID015.DTVENC} <=" & LcChav1 & " And {ALID015.DTVENC} >=" & LcChav2
'LcFormula = LcFormula & " AND {ALID015.VALOR} > {ALID015.VALPAGO}"

LcSql = "Select * from alid015 where DTVENC <= '" & Format(DataI.Text, "yyyy-mm-dd") & "'" ' And '" & Format(Datai.Text, "yyyy-mm-dd") & "'"
LcSql = LcSql & " and valor>valpago"

AbreBase
'abreconexao
Set RsNota = AbreRecordsetRel(LcSql, RsNota)
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from alid015", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsNota.Requery
'===> Apagando Registros antigos
err.Number = 0
'MsgBox LcSql
Do Until RsNotaMdb.EOF
    If err.Number > 0 Then Exit Do
    RsNotaMdb.Delete
    RsNotaMdb.MoveNext
Loop
err.Number = 0
Do Until RsNota.EOF
    'If err.Number > 0 Then Stop: Exit Do
    RsNotaMdb.AddNew
    On Error Resume Next
    For C = 0 To RsNota.Fields.Count - 1
    
        LcNome = RsNota.Fields(C).Name
        RsNotaMdb(LcNome) = RsNota.Fields(C)
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
'MsgBox err.Description & err.Number
'FechaConexao
RsNotaMdb.Close
Exit Function
errGera:
If err.Number = 438 Then
   Resume Next
Else
  MsgBox err.Description & err.Number
End If
End Function

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim RsReceita As ADODB.Recordset

Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
AbreBase
'Set RsReceita = Dbbase.OpenRecordset("alid015", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsReceita = AbreRecordset("select * from alid015 where isnull(valpago)", RsReceita)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
'Dataf.Text = "01/01/88"
'abreconexao
LcSql = "Update alid015 SET Valpago=0 where isnull(valpago)"
LcRegistrosAfetados = ExecutaSql(LcSql)
GeraNota
'Do Until RsReceita.EOF
'   If IsNull(RsReceita.VALPAGO) Then
'      RsReceita.Edit
'      RsReceita!VALPAGO = 0
'      RsReceita.Update
'   End If
'   RsReceita.MoveNext
'Loop
'RsReceita.Close
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
If Not RsOpcao.EOF Then
   LcVer = RsOpcao!msg
   LcVer1 = RsOpcao!Msg1
End If

    'Abertura do relatório de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    'If analitico Then
       'lctitulo = "Relatório de Comissões << ANALÍTICO >>"
      If GlImprimeSemLinha Then
         CryRelatorio.ReportFileName = App.Path & "\Receita.rpt"
      Else
         CryRelatorio.ReportFileName = App.Path & "\Receitasl.rpt"
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
  CryRelatorio.DiscardSavedData = True
  strData = CDate(Format(DataI.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav1 = " date(" & LcDataInicio & ")"
         
  strData = CDate(Format(DataF.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav2 = " date(" & LcDataInicio & ")"
  If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
  LcFormula = LcFormula & "{ALID015.DTVENC} <=" & LcChav1 '& " And {ALID015.DTVENC} >=" & LcChav2
  LcFormula = LcFormula & " AND {ALID015.VALOR} > {ALID015.VALPAGO}"

'== fim filtro
'== fim filtro

CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Receitas por Período"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
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
If Not IsDate(DataF.Text) And DataF.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      DataF.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(DataI.Text) And DataI.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      DataI.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 2835
Me.Width = 5370

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
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
Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function


