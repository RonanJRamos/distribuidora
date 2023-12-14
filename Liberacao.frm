VERSION 5.00
Begin VB.Form Liberacao 
   Caption         =   "Libera Preço"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a Senha de Liberação"
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
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Liberacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub CmdFechar_Click()
On Error Resume Next
GlLibera = False
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub CmdOk_Click()
If Len(txt.Text) = 0 Then
   MsgBox "É Necessário Digitar a Senha...", 64, "Aviso"
   txt.SetFocus
   Exit Sub
End If
If txt.Text = GlSenhaLiberacao Then
   GlLibera = True
   Unload Me
Else
   MsgBox "Senha Inválida...", 64, "Senha Não Confere"
   txt.SetFocus
End If
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlEscolha = False
GlFormA.SetFocus
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
