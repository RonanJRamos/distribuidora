VERSION 5.00
Begin VB.Form SenhaAteraPreco 
   BackColor       =   &H00E4DED3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha Para Alteração "
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Senha 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite a Senha Para Acesso"
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
      TabIndex        =   2
      Top             =   360
      Width           =   2940
   End
End
Attribute VB_Name = "SenhaAteraPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Senha.Text = GlLiberaPedidoVendas Then
    'FrmProposta.LiberaPedido
    GlLiberaSenhaAlteraPr = True
    Unload Me
    FrmSaidaProduto.SetFocus
    FrmExibeSenha = False
Else
    MsgBox "Senha Inválida.", 64, "Liberação Negada"
    Senha.SetFocus
    FrmExibeSenha = False
    
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
FrmSaidaProduto.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmExibeSenha = False
End Sub

Private Sub Senha_Change()
On Error Resume Next
Command1.Enabled = Len(Senha.Text)
End Sub
