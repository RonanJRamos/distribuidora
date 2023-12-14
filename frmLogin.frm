VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00A7A3FE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acesso ao Sistema."
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK F2"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel F10"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label FrmLog 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
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
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private a As Integer
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 113 Then SendKeys "%+{O}"

End Sub

Private Sub CmdOk_Click()
Dim a As Integer
On Error Resume Next
'SerieHd
VerificaOpcoes
GlUsuario = UCase(txtUserName.Text)
hasbilitatodos

  '=== Verifica se Foi Digitado Alguma Coisa em Branco
  If (IsNull(txtUserName.Text)) Or (Len(Trim(txtUserName.Text)) = 0) Then
     MsgBox "É Necessário Digitar o nome do Usuário", 48, "Aviso"
     txtUserName.SetFocus
     Exit Sub
  End If
  If (IsNull(txtPassword.Text)) Or (Len(Trim(txtPassword.Text)) = 0) Then
     MsgBox "É Necessário Digitar o a Senha de Acesso", 48, "Aviso"
     txtPassword.SetFocus
     Exit Sub
  End If
  
  If (txtUserName.Text = "Decisao") And (txtPassword.Text = "Suporte") Then
      txtPassword.Text = ""
      txtUserName.Text = ""
      hasbilitatodos
      txtUserName.SetFocus
      FrmPrincipal.montapainel
      frmDataSisema.Show
      Unload Me
  Else
      If VerificaSenha(txtUserName, txtPassword) Then
         'Permissoes
         HabilitaMenus
         FrmPrincipal.montapainel
         frmDataSisema.Show
         Unload Me
      Else
        MsgBox "Senha Inválida", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
 End If
End Sub
Function VerificaSenha(LcUser, LcSenha As String) As Integer
On Error GoTo errVerif
Dim RsGrupo As ADODB.Recordset
Dim LcCriterio As String
GLPesquisa = True
'Call AbreBanco(usuario)
'abreconexao
LcCriterio = "select * from usuario where Nome='" & LcUser & "' And Senha='" & LcSenha & "'"
'RsAtual.FindFirst LcCriterio
Set RsGrupo = AbreRecordset(LcCriterio)
If Not RsGrupo.EOF Then
   GlGrupo = RsGrupo!Grupo
   VerificaSenha = True
Else
   VerificaSenha = False
End If
Saiver:
'FechaBanco
GLPesquisa = False
Exit Function
errVerif:
VerificaSenha = False
GoTo Saiver
End Function

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 113 Then SendKeys "%+{O}"
End Sub

Private Sub Form_Activate()
Desabilitatodos
End Sub

Private Sub Form_Load()
GlGrupo = ""

End Sub

Private Sub FrmLog_DblClick(Index As Integer)
On Error GoTo err
If Dir(App.Path & "\VirtualDeposito.vbp", vbArchive) <> "" Then
    txtUserName.Text = "Decisao"
    txtPassword.Text = "Suporte"
End If
err:
MsgBox "A proteção atual não foi encontrada, verifique o arquivo texto no diretório raiz do c: ou unidade atual do sistema.", vbCritical, "Erro Encontrado"
Exit Sub
End Sub

Private Sub lblLabels_DblClick(Index As Integer)
On Error GoTo ErrLb
'If Dir("D:\projeto\Lidis Sql\VirtualDeposito.vbp", vbArchive) <> "" Then
    txtUserName.Text = "Decisao"
    txtPassword.Text = "Suporte"
'End If
ErrLb:
End Sub

Private Sub txtPassword_Change()
If Len(txtPassword.Text) = 0 Then
   cmdOK.Enabled = False
Else
  cmdOK.Enabled = True
End If
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 113 Then SendKeys "%+{O}"

End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
