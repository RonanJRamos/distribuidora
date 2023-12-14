VERSION 5.00
Begin VB.Form FrmPesquisaNota 
   Caption         =   "Pesquisa Orçamento"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Não Acrescentar zeros no nº da NF"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Por Cliente F3"
      Height          =   495
      Left            =   1380
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entre Com o Nº do Orçamento"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3105
   End
End
Attribute VB_Name = "FrmPesquisaNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
Select Case GlFormA.Name
      Case Is = "Orcamento"
           Orcamento.BuscaNota (Right("000000" & txt.Text, 6))
     Case Is = "FrmSaidaProduto"
           If Check1.Value = 0 Then
              FrmSaidaProduto.BuscaNota (Right("000000" & txt.Text, 6))
           Else
             FrmSaidaProduto.BuscaNota (txt.Text)
           End If
    Case Is = "FrmSaidaProdutoAlternativo"
           FrmSaidaProdutoAlternativo.BuscaNota (txt.Text)
     Case Is = "FrmProposta"
                FrmProposta.BuscaNota (Right("000000" & txt.Text, 6))
     Case Is = "FrmVales"
           FrmVales.BuscaNota (Right("000000" & txt.Text, 6))
           FrmVales.CmdSalvar.Enabled = False
End Select
Unload Me
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Load FrmBuscaOrcCliente
FrmBuscaOrcCliente.Tag = Me.Tag
FrmBuscaOrcCliente.Show , Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
'If GlFormA.Name <> "Orcamento" And GlFormA.Name <> "FrmProposta" Then
'   Command1.Visible = False
'   Me.Width = 3555
'Else
   Me.Width = 4665
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Select Case GlFormA.Name
      Case Is = "FrmVendaOrcam"
           'FrmVendaOrcam.BuscaNota (Right("000000" & txt.Text, 6))
     Case Is = "FrmSaidaProduto"
           FrmSaidaProduto.SetFocus
End Select
GlFormA.SetFocus
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
