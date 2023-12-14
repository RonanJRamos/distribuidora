VERSION 5.00
Begin VB.Form DiretorioDbsase 
   BackColor       =   &H00B3E9FD&
   Caption         =   "Localize o Diretório da Base de Dados"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   5055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleção"
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
      TabIndex        =   5
      Top             =   3600
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diretório"
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
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drive"
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
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "DiretorioDbsase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
If Len(Txt.Text) > 0 Then
   GLBase = Txt.Text
   GlEscolha = False
   Unload Me
Else
   MsgBox "Não foi escolhido nenhuma pasta....", 48, "Aviso"
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Dir1_Change()
Txt.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
