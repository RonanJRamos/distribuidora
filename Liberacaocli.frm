VERSION 5.00
Begin VB.Form LiberacaoCli 
   Caption         =   "Libera Preço"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Utilizado 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   75
   End
   Begin VB.Label Credito 
      AutoSize        =   -1  'True
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
      TabIndex        =   6
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Limite Utilizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Limite de Crédito"
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   2985
   End
End
Attribute VB_Name = "LiberacaoCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub CmdFechar_Click()
On Error Resume Next
GlBuscaProdutoAgora = True
'GlLibera = False
GlEscolha = False
Me.Visible = False
If GlFormA.Name = "Orcamento" Then
    orcamento.CodigoCliente.Text = ""
    orcamento.NomeCliente.Text = ""
    orcamento.CodigoCliente.SetFocus
End If

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
If txt.Text = GlSenhaCredito Then
   GlLibera = True
   GlBuscaProdutoAgora = True
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
utilizado.Caption = Format(GlUtilizado, "Currency")
Credito.Caption = Format(GlCredito, "currency")
GlUtilizado = 0
GlCredito = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlEscolha = False
If GlFormA.Name = "Orcamento" Then
   orcamento.SetFocus
End If
GlFormA.SetFocus
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
