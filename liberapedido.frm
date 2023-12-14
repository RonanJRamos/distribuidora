VERSION 5.00
Begin VB.Form Frmliberapedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libera Pedido"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "liberapedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entre com a Senha de Liberação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2805
   End
End
Attribute VB_Name = "Frmliberapedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txt.Text = GlLiberaPedidoVendas Then
    FrmProposta.LiberaPedido
    Unload Me
Else
    MsgBox "Senha Inválida.", 64, "Liberação Negada"
    txt.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmProposta.SetFocus
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
