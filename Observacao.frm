VERSION 5.00
Begin VB.Form Observacao 
   Caption         =   "Observacao"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3615
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox obs 
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Observacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
FrmProposta.obs.Text = obs.Text
Unload Me
FrmProposta.SetFocus
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
FrmProposta.SetFocus
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub obs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
