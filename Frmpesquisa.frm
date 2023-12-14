VERSION 5.00
Begin VB.Form frmPesquisa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa"
   ClientHeight    =   2340
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5445
   Icon            =   "Frmpesquisa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1382.549
   ScaleMode       =   0  'User
   ScaleWidth      =   5112.56
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F3"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1860
      TabIndex        =   11
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton CmdProximo 
      Caption         =   "&Próximo F4"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3000
      TabIndex        =   10
      Top             =   1860
      Width           =   1140
   End
   Begin VB.Frame Condicao 
      Caption         =   "Condição Campo"
      Height          =   1575
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option3 
         Caption         =   "Exatamente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Qualquer Parte"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Inicio "
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtChave 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ComboBox CmdPesquisa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK F2"
      Enabled         =   0   'False
      Height          =   390
      Left            =   720
      TabIndex        =   0
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar F10"
      Height          =   390
      Left            =   4140
      TabIndex        =   1
      Top             =   1860
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo Para Pesquisa"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LcTipoPesq As Integer, LcPosAtual As Integer
Dim LCEntrouPesquisa, a As Integer
Public LoginSucceeded As Boolean
Public Event Pesquisou(Chave As String, Valor As String, Tipo As String, Posicao As Integer)
Private Sub CmdAnterior_Click()
On Error GoTo erroanterior
GLPesquisa = True
If GlFormA.Name = "FrmGrupoEconomico" Then
   RaiseEvent Pesquisou(MtPesquisa(LcPosAtual).Indice, TxtChave.Text, MtPesquisa(LcPosAtual).Tipo, 2)
      Exit Sub
End If

RsAtual.FindPrevious LcCriterio
If Not RsAtual.NoMatch Then
  If GlFormA.Name = "Despesas" Then
       Call GlFormA.pesquisa(LcCriterio, 2)
   Else
      Call GlFormA.VinculaDados
   End If

Else
   MsgBox "Não Existe Mais Registros...", 64, "Não Encontrado"
   TxtChave.SetFocus
   Exit Sub
End If
GLPesquisa = False
CmdAnterior.SetFocus
Exit Sub
erroanterior:
MsgBox "Ocorreu o seguinte erro no sistema:" & Chr(13) & err.Description & " de Nº :" & err.Number, vbInformation, "Erro Encontrado."
Resume Next
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub cmdCancel_Click()
   On Error Resume Next
    LoginSucceeded = False
       
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdOk_Click()
On Error Resume Next
Dim LcMsg, S As String, LcCondicao, z As Integer

If Option1 Then
   LcCondicao = 1
   LcTipoPesq = 2
Else
   If Option2 Then
      LcCondicao = 2
      LcTipoPesq = 2
   Else
      LcCondicao = 3
      LcTipoPesq = 1
   End If
End If
For z = 0 To 32
    If MtPesquisa(z).campo = CmdPesquisa.Text Then
       LcPosAtual = z
       Exit For
    End If
Next
If GlFormA.Name = "FrmGrupoEconomico" Then
   RaiseEvent Pesquisou(MtPesquisa(LcPosAtual).Indice, TxtChave.Text, MtPesquisa(LcPosAtual).Tipo, 0)
End If



Select Case MtPesquisa(LcPosAtual).Tipo
    Case Is = "T"
    
         Select Case LcCondicao
                Case Is = 1
                     LcCriterio = MtPesquisa(LcPosAtual).Indice & " Like '" & TxtChave & "*'"
                Case Is = 2
                     LcCriterio = MtPesquisa(LcPosAtual).Indice & " Like '*" & TxtChave & "*'"
                Case Is = 3
                     LcCriterio = MtPesquisa(LcPosAtual).Indice & " Like '" & TxtChave & "'"
         End Select
         LcMsg = TxtChave & " Não Foi Encontrado..."
    Case Is = "D"
         LcCriterio = MtPesquisa(LcPosAtual).Indice & "=#" & TxtChave & "#"
         LcMsg = "A Data " & TxtChave & " Não Foi Encontrada..."
    Case Is = "N"
         LcCriterio = MtPesquisa(LcPosAtual).Indice & "=" & Val(TxtChave)
         LcMsg = "O Valor " & TxtChave & " Não Foi Encontrado..."
    Case Is = "M"
        LcCriterio = MtPesquisa(LcPosAtual).Indice & "=" & PermiteNumero(TxtChave)
        'LcCriterio = MtPesquisa(LcPosAtual).Indice & "=21.20"
        LcMsg = "O Valor" & CCur(TxtChave) & " Não Foi Encontrado..."
End Select
GLPesquisa = True
Call AbreBanco(GlFormAtual)
RsAtual.Requery
RsAtual.FindFirst LcCriterio
If Not RsAtual.NoMatch Then
   If GlFormA.Name = "Despesas" Then
       Call GlFormA.pesquisa(LcCriterio, 0)
   Else
      Call GlFormA.VinculaDados
   End If
   
Else
   MsgBox LcMsg, 64, "Não Encontrado"
   TxtChave.SetFocus
   Exit Sub
End If
Fim:
'GLPesquisa = False
CmdAnterior.Enabled = True
CmdProximo.Enabled = True
CmdProximo.SetFocus
End Sub
Function PermiteNumero(LcNumero As String) As String
Dim a, Saida As Integer
Dim LCLEtra, LcPrefixo, LcSufixo As String
Saida = False
For a = Len(LcNumero) To 1 Step -1
   LCLEtra = Mid(LcNumero, a, 1)
   If LCLEtra = "," Then
      Saida = True
      Exit For
   End If
Next
If Saida Then
   LcPrefixo = Mid(LcNumero, 1, a - 1)
   LcSufixo = Mid(LcNumero, a + 1)
   PermiteNumero = LcPrefixo & "." & LcSufixo
Else
   PermiteNumero = LcNumero
End If


End Function

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdPesquisa_Click()
On Error Resume Next

If MtPesquisa(CmdPesquisa.ListIndex).Tipo = "T" Then
   Condicao.Enabled = True
   Option1.Enabled = True
   Option2.Enabled = True
   Option3.Enabled = True
   
Else
   Condicao.Enabled = False
   Option1.Enabled = False
   Option2.Enabled = False
   Option3.Enabled = False
End If
LcIndice = MtPesquisa(CmdPesquisa.ListIndex).Indice
LcPosAtual = CmdPesquisa.ListIndex
End Sub
Function limpa()
On Error Resume Next
cmdOK.Enabled = False
CmdProximo.Enabled = False
CmdAnterior.Enabled = False
TxtChave = ""
End Function
Private Sub CmdPesquisa_GotFocus()
limpa

End Sub

Private Sub CmdPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   TxtChave.SetFocus
End If
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdProximo_Click()
On Error GoTo ErroProximo
GLPesquisa = True
If GlFormA.Name = "FrmGrupoEconomico" Then
   RaiseEvent Pesquisou(MtPesquisa(LcPosAtual).Indice, TxtChave.Text, MtPesquisa(LcPosAtual).Tipo, 2)
   Exit Sub
End If

RsAtual.FindNext LcCriterio
If Not RsAtual.NoMatch Then
   If GlFormA.Name = "Despesas" Then
       Call GlFormA.pesquisa(LcCriterio, 1)
   Else
      Call GlFormA.VinculaDados
   End If
Else
    MsgBox "Não Existe Mais Registros...", 64, "Não Encontrado"
   TxtChave.SetFocus
   Exit Sub
End If
CmdProximo.SetFocus
Exit Sub
ErroProximo:
MsgBox "Ocorreu o seguinte erro no sistema:" & Chr(13) & err.Description & " de Nº :" & err.Number, vbInformation, "Erro Encontrado."
Resume Next


End Sub



Private Sub CmdProximo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Not LCEntrouPesquisa Then
   TxtChave.SetFocus
   LCEntrouPesquisa = True
End If
End Sub

Private Sub Form_Load()
'Monta Combo
On Error Resume Next
Dim a As Integer, LcCampoAt As String
CmdPesquisa.Clear
For a = 0 To 31
   If Len(Trim(MtPesquisa(a).campo)) <> 0 Then
      CmdPesquisa.AddItem MtPesquisa(a).campo
      If MtPesquisa(a).Indice = LcIndice Then
         LcCampoAt = MtPesquisa(a).campo
         LcPosAtual = a
      End If
   End If
Next
CmdPesquisa.Text = LcCampoAt
CmdPesquisa.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim LcIndiceAntigo As String
On Error Resume Next
GLPesquisa = False
Sincroniza
LCEntrouPesquisa = False
LcRegAtual = False
GlFormA.SetFocus
End Sub

Private Sub Option1_Click()
limpa
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option2_Click()
limpa
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option3_Click()
limpa
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub TxtChave_Change()
On Error Resume Next
If Len(Trim(TxtChave)) = 0 Then
   cmdOK.Enabled = False
Else
   cmdOK.Enabled = True
End If
End Sub

Private Sub TxtChave_GotFocus()
limpa
End Sub

Private Sub TxtChave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   cmdOK.SetFocus
End If
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub TxtChave_LostFocus()
On Error Resume Next
TxtChave.Text = UCase(TxtChave.Text)
End Sub
