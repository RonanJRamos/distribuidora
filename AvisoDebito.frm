VERSION 5.00
Begin VB.Form AvisoDebito 
   Caption         =   "Debitos do Cliente"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fecha F10"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox senha 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Senha Para Liberação"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "AvisoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If GlSenhaDebito = Senha.Text Then
   GlLiberaAtraso = True
   Unload Me
Else
   MsgBox "Senha Inválida.", 64, "Aviso"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim LcCrit As String

LcCrit = "O Cliente " & GlCliDeb & " Tem Valores em Atraso."
Label1.Caption = LcCrit
End Sub

Private Sub Form_Unload(Cancel As Integer)
GlSaiuAtraso = True
End Sub
