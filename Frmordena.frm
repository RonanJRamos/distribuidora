VERSION 5.00
Begin VB.Form FrmOrdena 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordena"
   ClientHeight    =   1875
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5145
   Icon            =   "Frmordena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1107.812
   ScaleMode       =   0  'User
   ScaleWidth      =   4830.876
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmdPesquisa 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame Condicao 
      Caption         =   "Condição Registro"
      Height          =   1455
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Manter Atual"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ir Primeiro"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ir Último"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK F2"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   1140
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel F10"
      Height          =   390
      Left            =   1560
      TabIndex        =   1
      Top             =   1140
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar por"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "FrmOrdena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LcPosAtual, a As Integer
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
Dim LcOr, LcPos, a As Integer
Dim LcCri As String
For a = 0 To CmdPesquisa.ListCount - 1
    If MtPesquisa(a).Campo = CmdPesquisa.Text Then
       LcIndice = MtPesquisa(a).Indice
       LcPos = a
       Exit For
    End If
Next
'LcIndice = MtPesquisa(CmdPesquisa.ListIndex).Indice
'RsAtual.Index = LcIndice
LcOr = InStr(1, GlStringBase, "order")
If LcOr > 0 Then
   GlStringBase = GlordemAnterior
Else
   GlordemAnterior = GlStringBase
End If

GlStringBase = GlStringBase & " order by " & LcIndice
Call AbreBanco(GlFormAtual)
If Option1 Then
   GlChave = GlFormA.txt(LcPos).Text
   AchaReg (1)
   If RsAtual.NoMatch Then
      Exit Sub
   End If
   'MsgBox RsAtual!Cod
End If
If Option2 Then
   RsAtual.MoveFirst
Else
   If Option3 Then
      RsAtual.MoveLast
   End If
End If
Call GlFormA.VinculaDados
   'Call FechaBanco
'Call FechaBanco
SaiOk:
LcRegAtual = False
Unload Me
End Sub

Private Sub CmdPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim a As Integer, LcCampoAt As String
For a = 0 To 31
   If Len(Trim(MtPesquisa(a).Campo)) <> 0 Then
      CmdPesquisa.AddItem MtPesquisa(a).Campo
      If MtPesquisa(a).Indice = LcIndice Then
         LcCampoAt = MtPesquisa(a).Campo
         LcPosAtual = a
      End If
   End If
Next
CmdPesquisa.Text = LcCampoAt
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Sincroniza
LcRegAtual = False
GlFormA.SetFocus
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub
