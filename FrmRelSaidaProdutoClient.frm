VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelSaidaProdutoClien 
   BackColor       =   &H00E6E4D2&
   Caption         =   "Relatório de Saida de Estoque por Cliente"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Cliente 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Nome 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Copias 
      Alignment       =   2  'Center
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
      Left            =   5760
      TabIndex        =   8
      Text            =   "1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Tipo de Pesquisa"
      Height          =   1335
      Left            =   2280
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton Igual 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Iniciado 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Iniciado por"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Saída"
      Height          =   1335
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
      Begin VB.OptionButton Video 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Impressora 
         BackColor       =   &H00E6E4D2&
         Caption         =   "Impressora"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
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
      Caption         =   "Codigo"
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
      Index           =   2
      Left            =   4920
      TabIndex        =   19
      Top             =   720
      Width           =   765
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
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   1440
      Width           =   1080
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
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   1185
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
      Left            =   5760
      TabIndex        =   15
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relatório de Saida de Estoque por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "FrmRelSaidaProdutoClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadoCliente
        Codigo As String
        Nome As String
End Type
Private MtCliente() As DadoCliente
Private LcTam, a As Long
Function carregaCliente()
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim RsEmpresa As Recordset
AbreBase
LcTam = 0
Set RsEmpresa = Dbbase.OpenRecordset("select * from alid001 order by razaosoc", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsEmpresa.EOF
    ReDim Preserve MtCliente(LcTam)
    DoEvents
    If Not IsNull(RsEmpresa!razaosoc) Then
        MtCliente(LcTam).Codigo = RsEmpresa!Codigo
        MtCliente(LcTam).Nome = RsEmpresa!razaosoc
        Cliente.AddItem RsEmpresa!razaosoc
        LcTam = LcTam + 1
    End If
    RsEmpresa.MoveNext
Loop
If LcTam > 0 Then LcTam = LcTam - 1
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set dbbasee = Nothing

End Function

Private Sub CLIENTE_Change()
On Error Resume Next
Nome.Text = MtCliente(Cliente.ListIndex).Codigo

End Sub

Private Sub Cliente_Click()
On Error Resume Next
Nome.Text = MtCliente(Cliente.ListIndex).Codigo
'For a = 0 To LcTam
'    If MtCliente(a).nome = CLIENTE.Text Then
'       nome.Text = MtCliente(a).codigo
'       Exit For
'    End If
'Next

End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Function GeraNota()
On Error GoTo ErrGera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String

LcSql = "Select * from alid050 where cliente = '" & Nome.Text & "'"
'LcSql = "Select * from alid015 where cliente = '" & UCase(nome.Text) & "'"

LcSql = LcSql & " and DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "'"

AbreBase
'abreconexao
Set RsNota = AbreRecordsetRel(LcSql, RsNota)
RsNota.Requery
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from alid050", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'===> Apagando Registros antigos
'MsgBox LcSql
Do Until RsNotaMdb.EOF
    RsNotaMdb.Delete
    RsNotaMdb.MoveNext
Loop
err.Number = 0
Do Until RsNota.EOF
    'If err.Number > 0 Then Exit Do
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
Exit Function
ErrGera:

If err.Number = 438 Then
   Resume Next
Else
 ' MsgBox err.Description & err.Number
End If
Resume Next

End Function

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, rsCliente As Recordset
Dim LcEmpresa, LcEndereco, LcFone, Lccelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
If Len(Cliente.Text) = o Then
   MsgBox "Escolha o Cliente a Exibir...", 64, "Aviso"
   Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."

GeraNota
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsCliente = Dbbase.OpenRecordset("select * from alid001 where RAZAOSOC='" & MtCliente(Cliente.ListIndex).codigo & "'")
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If


'Abertura do relatório de vendas
    
    
    CryRelatorio.DataFiles(0) = GLBase
    If GlImprimeSemLinha Then
       CryRelatorio.ReportFileName = App.Path & "\NotaSaida.rpt"
    Else
       CryRelatorio.ReportFileName = App.Path & "\NotaSaidasl.rpt"
    End If
   LcFormula = "{ALID050.CLIENTE}='" & Nome.Text & "'"
   CryRelatorio.SortFields(0) = "+{ALID050.NUMNF}"
   
    CryRelatorio.CopiesToPrinter = Val(copias.Text)
 '== Inicio Filtro
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
 ' If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
 ' LcFormula = LcFormula & "{ALID050.DTEMIS} >=" & LcChav1 & " And {ALID050.DTEMIS} <=" & LcChav2


'== fim filtro
    
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Saída de Produtos por Cliente"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
'CryRelatorio.Formulas(3) = "Versiculo='" & LcVer & "'"
'CryRelatorio.Formulas(4) = "Versiculo1='" & LcVer1 & "'"
CryRelatorio.Formulas(5) = "titulo='Saída de Produtos por Cliente'"
CryRelatorio.Formulas(3) = "Celular='" & Lccelular & "'"
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
'RsOpcao.Close
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
End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
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
On Error Resume Next
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

Private Sub Form_Click()
For a = 0 To LcTam
    If MtCliente(a).Nome = Cliente.Text Then
       Nome.Text = MtCliente(a).Codigo
       Exit For
    End If
Next
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
'Me.Height = 4215
'Me.Width = 7350
carregaCliente
'abreconexao
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaConexao
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Iniciado_Click()
'Escolha
'BuscaExpressao

End Sub

Private Sub Iniciado_GotFocus()
On Error Resume Next
'Txt(0).Text = ""
'Txt(1).Text = ""
End Sub

Private Sub Iniciado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
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

