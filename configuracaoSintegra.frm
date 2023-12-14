VERSION 5.00
Begin VB.Form configuracaoSintegra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "configuracaoSintegra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F2"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   3240
      X2              =   3240
      Y1              =   0
      Y2              =   5040
   End
End
Attribute VB_Name = "configuracaoSintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next
Dim LcArq   As Integer
Dim Arquivo As String
Dim a As Long
LcArq = FreeFile
'===> Grava o Arquivo no diretorio do Banco de Dados
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then
       Exit For
    End If
Next
Arquivo = Mid(GLBase, 1, a)
Arquivo = Arquivo & "Configuracaosintegra.txt"
Open Arquivo For Output Access Write As #LcArq
   
Print #LcArq, Drive1
Print #LcArq, Dir1.Path

Close #LcArq
Opcoes.Sintegra.Text = Dir1.Path

LcDrive = Drive1
GlPathPDV = Dir1.Path
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
If Len(LcDrive) > 0 Then
   Drive1 = LcDrive
   If Len(LcPAth) > 0 Then
      Dir1.Path = LcPAth
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Opcoes.SetFocus
End Sub
