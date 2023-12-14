VERSION 5.00
Begin VB.Form FrmBuscaProposta 
   BackColor       =   &H00CBB19C&
   Caption         =   "Busca Pedido"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form2"
   ScaleHeight     =   1485
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar F10"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox tx5t 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Pedido"
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
      TabIndex        =   1
      Top             =   600
      Width           =   1410
   End
End
Attribute VB_Name = "FrmBuscaProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
FrmSaidaProduto.BuscaProposta (Right("000000" & tx5t.Text, 6))
Unload Me
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlFormA.SetFocus
End Sub

Private Sub tx5t_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
