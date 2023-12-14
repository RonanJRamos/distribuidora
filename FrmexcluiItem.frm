VERSION 5.00
Begin VB.Form FrmExcluiItem 
   BackColor       =   &H00ADC1FE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item a Excluir"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar F10"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entre Com o Nº do Item"
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
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmExcluiItem"
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
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
Select Case GlFormA.Name
   Case Is = "FrmSaidaProduto"
       FrmSaidaProduto.ExcluiItem (Val(txt.Text))
   Case Is = "FrmVales"
       FrmVales.ExcluiItem (Val(txt.Text))
   Case Is = "FrmPedido"
       FrmPedido.ExcluiItem (Val(txt.Text))
   Case Is = "FrmEntradaProduto"
       FrmEntradaProduto.ExcluiItem (Val(txt.Text))
   Case Is = "Orcamento"
       Orcamento.ExcluiItem (txt.Text)
   Case Is = "comissao"
       comissao.ExcluiItem (CCur(txt.Text))
   Case Is = "FrmProposta"
       FrmProposta.ExcluiItem (Val(txt.Text))
End Select
Unload Me
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
If GlFormA.Name = "comissao" Then
   Label1.Caption = "Entre Com o Valor a Excluir:"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlFormA.SetFocus
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
