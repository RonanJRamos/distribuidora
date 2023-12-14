VERSION 5.00
Begin VB.Form FrmApresentacao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   840
   End
   Begin VB.PictureBox FormShape1 
      Height          =   480
      Left            =   720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   2760
      TabIndex        =   0
      Top             =   4440
      Width           =   15
   End
End
Attribute VB_Name = "FrmApresentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DataRepeater1_Click()

End Sub

Private Sub Form_Activate()
On Error Resume Next
Timer1.Interval = 6000
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
FormShape1.hWnd = FrmApresentacao.hWnd
FormShape1.ShapePicture = FrmApresentacao.Picture
moving = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Unload Me
frmLogin.Show
End Sub
