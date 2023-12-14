VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRelEntradaProduto 
   BackColor       =   &H00EAE8DD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Entrada de Produto "
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Tipo de Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
      Begin VB.OptionButton Igual 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton Iniciado 
         BackColor       =   &H00EAE8DD&
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
      BackColor       =   &H00EAE8DD&
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton Video 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Impressora 
         BackColor       =   &H00EAE8DD&
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
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4215
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
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relatório de Entrada de Produto "
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
      Caption         =   "Nota fiscal"
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
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRelEntradaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Function GeraDados()
On Error GoTo errgera
Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As Recordset
Dim LcSql As String
Dim LcNome As String
Dim db As Database
Dim C As Integer
LcSql = "Select * from EntradaNf where nf like '" & Nome.Text & "%'"
'AbreBase
Set db = OpenDatabase(App.Path & "\relatorios.mdb")
'==> Apagando Registros
db.Execute "Delete * from EntradaNf"

Set RsNota = AbreRecordset(LcSql, True)
Set RsNotaMdb = db.OpenRecordset("Select * from EntradaNf", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsNota.Requery
'===> Apagando Registros antigos
'Do Until RsNotaMdb.EOF
'    DoEvents
'    RsNotaMdb.Delete
'    RsNotaMdb.MoveNext
'Loop

Do Until RsNota.EOF
    RsNotaMdb.AddNew
    For C = 0 To RsNota.Fields.Count - 1
        LcNome = RsNota.Fields(C).Name
        If UCase(LcNome) = "CODIGO" Then
            RsNotaMdb(LcNome) = CDbl(RsNota.Fields(C))
        Else
            RsNotaMdb(LcNome) = RsNota.Fields(C)
        End If
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
RsNotaMdb.Close

Exit Function
errgera:
'MsgBox err.Description & err.Number
Resume Next
End Function

Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, Lccelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
GeraDados
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
    
    
    CryRelatorio.DataFiles(0) = App.Path & "\relatorios.mdb"
    If GlImprimeSemLinha Then
       CryRelatorio.ReportFileName = App.Path & "\NotaEntrada.rpt"
    Else
       CryRelatorio.ReportFileName = App.Path & "\NotaEntradasl.rpt"
    End If
    If Iniciado Then LcFormula = "{EntradaNf.NF} like '" & UCase(Nome.Text) & "*'"
    If Qualquer Then LcFormula = "{EntradaNf.NF} like '*" & UCase(Nome.Text) & "*'"
    If Igual Then LcFormula = "{EntradaNf.NF}='" & UCase(Nome.Text) & "'"
    If Len(Nome.Text) > 0 Then
       CryRelatorio.SortFields(0) = "+{EntradaNf.NF}"
      End If
    CryRelatorio.CopiesToPrinter = Val(copias.Text)

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Entrada de Produtos"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
'CryRelatorio.Formulas(3) = "Versiculo='" & LcVer & "'"
'CryRelatorio.Formulas(4) = "Versiculo1='" & LcVer1 & "'"
CryRelatorio.Formulas(5) = "titulo='Entrada de Produtos'"
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
On Error Resume Next
Set GlFormA = Me
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3495
Me.Width = 7080

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
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
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
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
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub
