VERSION 5.00
Begin VB.Form DescontoRepresentante 
   BackColor       =   &H00C5FEE1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconto Sobre Item "
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Novo 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Digitado 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   552
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F2"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox ValorDesconto 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1416
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Alterado"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1848
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Digitado"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Desconto em %"
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
      Left            =   120
      TabIndex        =   0
      Top             =   984
      Width           =   2655
   End
End
Attribute VB_Name = "DescontoRepresentante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
If Len(ValorDesconto.Text) > 0 Then

   orcamento.Unitario.Text = Novo.Text
   orcamento.preconormal.Text = Novo.Text
   Unload Me
Else
   MsgBox "Não Foi Digitado Nenhum Desconto.", vbInformation, "Avso"
   ValorDesconto.SetFocus
End If

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Digitado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Digitado.Text = orcamento.Unitario.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlFormA.SetFocus
End Sub

Private Sub Novo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub ValorDesconto_Change()
On Error Resume Next
Dim LcValor, LcDesconto As Double
LcDesconto = (CDbl(ValorDesconto.Text) / 100) * CDbl(Digitado.Text)
LcValor = CDbl(Digitado.Text) - LcDesconto
Novo.Text = AcertaNumero(CStr(LcValor), GlDecimais)

End Sub

Private Sub ValorDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub
