VERSION 5.00
Begin VB.Form ConfirmaSenha 
   BackColor       =   &H00CAE0E6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Senha para Exclusão de Vales"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirma"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha de Acesso"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "ConfirmaSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Senha.Text = GlLiberaPedidoVendas Then
    FrmExclusaodeVales.Show , FrmPrincipal
    Unload Me
Else
    MsgBox "Senha Inválida.", 64, "Exclusão Negada"
    Senha.SetFocus
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Senha_Change()
On Error Resume Next
Command1.Enabled = Len(Senha.Text)
End Sub
