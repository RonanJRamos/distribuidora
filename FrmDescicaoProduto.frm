VERSION 5.00
Begin VB.Form FrmDescicaoProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descrição Detalhada do Produto"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancela 
      Caption         =   "&Cancelar F10"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Ok  F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Descricao 
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "FrmDescicaoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub cancela_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cancela_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{TAB}"
If KeyCode = 121 Then SendKeys "%+{TAB}"
End Sub

Private Sub Descricao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "%+{TAB}"

End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Select Case GlFormA.Name
    Case Is = "FrmSaidaProduto"
         Descricao.Text = FrmSaidaProduto.txt(2).Text
    Case Is = "Orcamento"
         Descricao.Text = orcamento.NomeProduto.Text
End Select
End Sub

Private Sub ok_Click()
On Error Resume Next
Select Case GlFormA.Name
    Case Is = "FrmSaidaProduto"
          FrmSaidaProduto.txt(2).Text = Descricao.Text
          
    Case Is = "Orcamento"
          orcamento.NomeProduto.Text = Descricao.Text
End Select
Unload Me
Select Case GlFormA.Name
   Case Is = "FrmSaidaProduto"
          FrmSaidaProduto.SetFocus
    Case Is = "Orcamento"
          orcamento.SetFocus
End Select
End Sub

Private Sub ok_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "%+{TAB}"
End Sub
