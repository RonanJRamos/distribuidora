VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmConferenciaQuitacao 
   BackColor       =   &H00E6E4D2&
   Caption         =   "Confere as Receitas Recebidas"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Sa�da"
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
      Begin VB.OptionButton Video 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton impressora 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
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
      Left            =   2160
      TabIndex        =   2
      Text            =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer tempo 
      Left            =   4440
      Top             =   2280
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   0
      TabIndex        =   6
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
      Left            =   2040
      TabIndex        =   7
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
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
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
      Left            =   2040
      TabIndex        =   10
      Top             =   480
      Width           =   585
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
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   690
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Periodo de Recebimento"
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
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   2670
   End
End
Attribute VB_Name = "FrmConferenciaQuitacao"
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
Private Aguarda As Boolean
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


Function GeraNota()
On Error Resume Next
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String

LcSql = "Select * from alid015 where DTPAGTO Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
Debug.Print LcSql
AbreBase
'abreconexao
Set RsNota = AbreRecordsetRel(LcSql, RsNota)
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from alid015")
RsNota.Requery
'===> Apagando Registros antigos
Dbbase.Execute ("Delete from alid015")
Do Until RsNota.EOF
    RsNotaMdb.AddNew
    For C = 0 To RsNota.Fields.Count - 1
        LcNome = RsNota.Fields(C).Name
        RsNotaMdb(LcNome) = RsNota.Fields(C)
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
'FechaConexao
RsNotaMdb.Close

End Function
Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
AbreBase
GeraNota
Aguarda = True
tempo.Interval = 3600
Do While Aguarda
   DoEvents
Loop
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relat�rio..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
If Not RsOpcao.EOF Then
   LcVer = RsOpcao!Msg
   LcVer1 = RsOpcao!Msg1
End If

    'Abertura do relat�rio de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    'If analitico Then
       'lctitulo = "Relat�rio de Comiss�es << ANAL�TICO >>"
   
        CryRelatorio.ReportFileName = App.Path & "\ReceitaRecebidaSl.rpt"
   
    'Else
       'lctitulo = "Relat�rio de Comiss�es << SINT�TICO >>"
   ' End If
    CryRelatorio.SortFields(0) = "+{ALID015.codigo}"
    
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
  LcFormula = LcFormula & "{ALID015.DTPAGTO} >=" & LcChav1 & " And {ALID015.DTPAGTO} <=" & LcChav2


'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Receitas Recebidas por periodo"

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relat�rio de Receitas por Per�odo'"
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
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub copias_KeyDown(KeyCode As Integer, Shift As Integer)
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
      MsgBox "A data digitada n�o � V�lida...", 48, "Aviso"
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
      MsgBox "A data digitada n�o � V�lida...", 48, "Aviso"
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

Private Sub impressora_LostFocus()
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub tempo_Timer()
Aguarda = False
End Sub

Private Sub Video_Click()
copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_LostFocus()
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
