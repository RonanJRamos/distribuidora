VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClienteTelem 
   Caption         =   "Relatório de Clientes/ Telemarketing"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox GrupoEconomico 
      Height          =   315
      ItemData        =   "FrmRelClienteTelem.frx":0000
      Left            =   240
      List            =   "FrmRelClienteTelem.frx":0002
      TabIndex        =   16
      Top             =   1560
      Width           =   4335
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4680
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Ordem 
      Height          =   315
      Left            =   4560
      TabIndex        =   14
      Top             =   2040
      Width           =   2415
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
      Left            =   5280
      TabIndex        =   6
      Text            =   "1"
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Pesquisa"
      Height          =   1335
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton Igual 
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton Iniciado 
         Caption         =   "Iniciado por"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton Video 
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Impressora 
         Caption         =   "Impressora"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Economico"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Label Label4 
      Caption         =   "Classificar Por"
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
      Left            =   4560
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
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
      Left            =   5280
      TabIndex        =   13
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "  Relatório de Clientes/ Telemarketing"
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
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   630
   End
End
Attribute VB_Name = "FrmRelClienteTelem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If


'Abertura do relatório de vendas
    
    CryRelatorio.DiscardSavedData = True
    CryRelatorio.DataFiles(0) = GLBase
    If GlImprimeSemLinha Then
       CryRelatorio.ReportFileName = App.Path & "\relclienteTelem.rpt"
    Else
       CryRelatorio.ReportFileName = App.Path & "\relclienteTelemsl.rpt"
    End If
    If Iniciado Then LcFormula = "{ALID001.RAZAOSOC} like '" & UCase(Nome.Text) & "*'"
    If Qualquer Then LcFormula = "{ALID001.RAZAOSOC} like '*" & UCase(Nome.Text) & "*''"
    If Igual Then LcFormula = "{ALID001.RAZAOSOC}='" & UCase(Nome.Text) & "'"
    If Len(GrupoEconomico.Text) > 0 Then
       LcFormula = LcFormula & "  and UpperCase({ALID001.GrupoEconomicoNome})='" & UCase(GrupoEconomico.Text) & "'"
    End If
    Debug.Print LcFormula
    Select Case Ordem.Text
            Case Is = "Razão Social"
                 CryRelatorio.SortFields(0) = "+{ALID001.razaosoc}"
            Case Is = "Fantasia"
                 CryRelatorio.SortFields(0) = "+{ALID001.FANTASIA}"
            Case Is = "Código"
                 CryRelatorio.SortFields(0) = "+{ALID001.CODIGO}"
            Case Is = "Endereço"
                 CryRelatorio.SortFields(0) = "+{ALID001.END}"
            Case Is = "Bairro"
                 CryRelatorio.SortFields(0) = "+{ALID001.BAIRRO}"
            'Case Is = "Cidade"
                 'CryRelatorio.SortFields(0) = "+{ALID001.NOME}"
            Case Is = "Estado"
                 CryRelatorio.SortFields(0) = "+{ALID001.ESTADO}"
            Case Is = "Fone"
                CryRelatorio.SortFields(0) = "+{ALID001.FONE1}"
            Case Is = "CGC"
                CryRelatorio.SortFields(0) = "+{ALID001.CGC}"
            Case Is = "Inscrição Estadual"
                CryRelatorio.SortFields(0) = "+{ALID001.INSCEST}"
           Case Else
               CryRelatorio.SortFields(0) = "+{ALID001.razaosoc}"
    End Select
    
    CryRelatorio.CopiesToPrinter = Val(Copias.Text)

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Clientes/Telemarketing"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
'CryRelatorio.Formulas(3) = "Versiculo='" & LcVer & "'"
'CryRelatorio.Formulas(4) = "Versiculo1='" & LcVer1 & "'"
CryRelatorio.Formulas(5) = "titulo='Clientes por Nome'"
CryRelatorio.Formulas(3) = "Celular='" & LcCelular & "'"
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
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub


Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub
Function CarregaCombo()
Ordem.AddItem "Razão Social"
Ordem.AddItem "Fantasia"
Ordem.AddItem "Código"
Ordem.AddItem "Endereço"
Ordem.AddItem "Bairro"
Ordem.AddItem "Cidade"
Ordem.AddItem "Estado"
Ordem.AddItem "Fone"
Ordem.AddItem "CGC"
Ordem.AddItem "Inscrição Estadual"
Ordem.Text = "Razão Social"
End Function
Function CarregaGrupoEconomico()
On Error GoTo errc
Debug.Print conexaoAdo.ConnectionString
Dim RsVendedor As ADODB.Recordset
Set RsVendedor = AbreRecordset("Select * from GrupoEconomico order by nome", True)
LcTamanhoGr = 0
Do Until RsVendedor.EOF
   GrupoEconomico.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
Loop
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:
MsgBox err.Description & " " & err.Number
Exit Function

End Function

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3495
Me.Width = 7080
CarregaCombo
CarregaGrupoEconomico
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
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
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Ordem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub
