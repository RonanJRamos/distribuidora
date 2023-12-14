VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Caixa 
   Caption         =   "Caixa"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form3"
   ScaleHeight     =   5580
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TrocoProximo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Atual 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   19
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Pagamentos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Recebimento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox TrocoDia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Anterior 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ver Mais Caixas F4"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin MSMask.MaskEdBox DataUtil 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
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
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Sair  F10"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Confirmar Caixa F3"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Pesquisar Caixa  F11"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      Caption         =   "Para Detalhar Recebimentos e Pagamentos, dê um Duplo Clique em cima do Campo."
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   4920
      TabIndex        =   21
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Próximo Dia Util"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Troco do Dia"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Troco Proximo Dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6735
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   4800
      Y1              =   6000
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label5 
      Caption         =   "Data  do Caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Saldo Atual"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Pagamentos"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Recebimentos"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Saldo Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Anterior_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Atual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Command1_Click()
FrmPesquisaCaixa.Show , Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim RsCaixa As Recordset, Db As Database
Set Db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsCaixa = Db.OpenRecordset("caixa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterio = "data=#" & Format(Data.Text, "mm/dd/yy") & "#"
RsCaixa.FindFirst LcCriterio
If Not RsCaixa.NoMatch Then
   RsCaixa.Edit
Else
   RsCaixa.AddNew
End If
RsCaixa!Data = CDate(Data.Text)
RsCaixa!Recebimentos = CDbl(Recebimento.Text)
RsCaixa!Pagamentos = CDbl(Pagamentos.Text)
RsCaixa!SaldoAnterior = CDbl(Anterior.Text)
RsCaixa!SaldoAtual = CDbl(Atual.Text)
RsCaixa!TrocoDia = CDbl(TrocoDia.Text)
RsCaixa!TrocoProximoDia = CDbl(TrocoProximo.Text)
RsCaixa!fechado = True
RsCaixa.Update

'==== Gera Caixa Proximo Dia
RsCaixa.AddNew
RsCaixa!Data = CDate(DataUtil.Text)
RsCaixa!Recebimentos = 0
RsCaixa!Pagamentos = 0
RsCaixa!SaldoAnterior = CDbl(Atual.Text)
RsCaixa!SaldoAtual = CDbl(Atual.Text)
RsCaixa!TrocoDia = CDbl(TrocoProximo.Text)
RsCaixa!TrocoProximoDia = 0
RsCaixa!fechado = False
RsCaixa.Update


Data.Text = "  /  /  "
Recebimento.Text = ""
Pagamentos.Text = ""
Anterior.Text = ""
Atual.Text = ""
TrocoDia.Text = ""
TrocoProximo.Text = ""
DataUtil.Text = "  /  /  "
Command2.Enabled = False
Status.Caption = "Aguardando Caixa"

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Command4_Click()
DetalhaCaixa.Show , Me
End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Data_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"

End Sub

Private Sub DataUtil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Call PesquisaCaixa(Date, False)
End Sub
Function PesquisaCaixa(LcData As Date, LcAvisa As Boolean)
On Error Resume Next
Dim LcCriterio As String
Dim RsCaixa As Recordset, Db As Database
Set Db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsCaixa = Db.OpenRecordset("caixa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterio = "data=#" & Format(LcData, "mm/dd/yy") & "#"
RsCaixa.FindFirst LcCriterio
If Not RsCaixa.NoMatch Then
   Data.Text = Format(RsCaixa!Data, "dd/mm/yy") & ""
   Recebimento.Text = AcertaNumero(CStr(RsCaixa!Recebimentos), 2) & ""
   Pagamentos.Text = AcertaNumero(CStr(RsCaixa!Pagamentos), 2) & ""
   Anterior.Text = AcertaNumero(CStr(RsCaixa!SaldoAnterior), 2) & ""
   Atual.Text = AcertaNumero(CStr(RsCaixa!SaldoAtual), 2) & ""
   TrocoDia.Text = AcertaNumero(CStr(RsCaixa!TrocoDia), 2)
   TrocoProximo.Text = AcertaNumero(CStr(RsCaixa!TrocoProximoDia), 2)
   
   If RsCaixa!fechado Then
      msg = "Caixa Fechado."
      Command2.Enabled = False
      DataUtil.Text = Format(Date + 1, "dd/mm/yy")
   Else
      Command2.Enabled = True
      msg = "Caixa Aberto."
      DataUtil.Text = Format(CDate(Data.Text) + 1, "dd/mm/yy")
   End If
   Status.Caption = msg
 Else
   If LcAvisa Then
      MsgBox "Não Foi Encontrado o Caixa do Dia " & LcData, 64, Aviso
      DataUtil.Text = Format(Date + 1, "dd/mm/yy")
   End If
   Status.Caption = "Aguardando Caixa"
 End If
RsCaixa.Close
Db.Close
Set RsCaixa = Nothing
Set Db = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
FrmPrincipal.SetFocus
End Sub

Private Sub Pagamentos_DblClick()
GlRec = "D"
Me.Tag = "D"
DetalhaCaixaTm.Show , Me
End Sub

Private Sub Pagamentos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Recebimento_DblClick()
GlRec = "R"
Me.Tag = "R"
DetalhaCaixaTm.Show , Me
End Sub

Private Sub Recebimento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub TrocoDia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub TrocoProximo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{V}"
If KeyCode = 122 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub
