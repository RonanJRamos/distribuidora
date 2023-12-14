VERSION 5.00
Begin VB.Form FrmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ExcluiReceita 
      Caption         =   "Permitir excluir Receita"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txt 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CmdExcluiUsuario 
      Caption         =   "&Excluir Usuario F3"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdNovo 
      Caption         =   "&Novo Usuario F4"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ListBox ListUsuarios 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ComboBox CboGrupo 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Confirmação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcPermissao As Long
Private LcAlterou, a As Integer
Private LcSenha As String

Private Sub CboGrupo_Change()
LcAlterou = True
End Sub

Private Sub CboGrupo_Click()
LcAlterou = True
End Sub

Private Sub CboGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CboGrupo_LostFocus()
On Error Resume Next
CmdSalvar.SetFocus
End Sub

Private Sub CmdExcluiUsuario_Click()
On Error GoTo erroexcluir
Dim RsUsuario As ADODB.Recordset
Dim a, Item, LcResposta As Long
Dim LcCriterio As String

LcResposta = MsgBox("Confirma a Exclusão deste Usuario ?", 36, "Confirmação")
If LcResposta = 7 Then Exit Sub

LcCriterio = "Select * From GrpSenhas where Grupo='" & txt(0).Text & "'"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsUsuario = AbreRecordset("select * from usuario") ', dbOpenDynaset)
LcCriterio = "Nome='" & txt(0).Text & "'"
RsUsuario.Find LcCriterio
If Not RsUsuario.EOF Then
   RsUsuario.Delete
   For Item = ListUsuarios.ListCount - 1 To 0 Step -1
      If ListUsuarios.List(Item) = txt(0).Text Then
         ListUsuarios.RemoveItem (Item)
         Exit For
      End If
   Next
   
End If
RsUsuario.Close
Set RsUsuario = Nothing
LcAlterou = False
txt(0).Text = ""
txt(1).Text = ""
txt(2).Text = ""
CboGrupo.Text = ""
ExcluiReceita.Value = 0
Exit Sub
erroexcluir:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Sub
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Sub

Private Sub CmdExcluiUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me

End Sub


Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdNovo_Click()
On Error Resume Next
LcAlterou = False
txt(0).Text = ""
txt(1).Text = ""
txt(2).Text = ""
CboGrupo.Text = ""
ExcluiReceita.Value = 0
End Sub

Private Sub CmdNovo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdSalvar_Click()
Salva
LcAlterou = False
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Form_Activate()
 Set GlFormA = Me
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaUsuario
CarregaGrupo
LcAlterou = False
End Sub
Function VerificaLista(LcNome As String) As Integer
Dim Item As Long

For Item = ListUsuarios.ListCount - 1 To 0 Step -1
     If ListUsuarios.List(Item) = LcNome Then
        VerificaLista = True
        Exit For
     Else
        VerificaLista = False
     End If
Next
End Function

Function CarregaUsuario()
On Error GoTo errousuario
Dim RsUsuario As ADODB.Recordset
Dim a As Long

Dim LcCriterio As String
LcCriterio = "Select * From GrpSenhas where Grupo='" & txt(0).Text & "'"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsUsuario = AbreRecordset("select * from usuario", True) ', dbOpenDynaset)

Do Until RsUsuario.EOF
   If Not VerificaLista(RsUsuario!Nome) Then
      ListUsuarios.AddItem RsUsuario!Nome
   End If
   RsUsuario.MoveNext
Loop
RsUsuario.Close
Set RsUsuario = Nothing
Exit Function
errousuario:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function
Function CarregaGrupo()
On Error GoTo errogrupo
Dim RsGrupo As ADODB.Recordset
Dim a As Long

Dim LcCriterio As String
LcCriterio = "Select * From GrpSenhas where Grupo='" & txt(0).Text & "'"
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("select * from GrpSenhas", True) ', dbOpenDynaset)

Do Until RsGrupo.EOF
   If Not VerificaGrupo(RsGrupo!Grupo) Then
      CboGrupo.AddItem RsGrupo!Grupo
   End If
   RsGrupo.MoveNext
Loop
RsGrupo.Close
Set RsGrupo = Nothing
Exit Function
errogrupo:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function
Function VerificaGrupo(LcNome As String) As Integer
Dim Item As Long

For Item = CboGrupo.ListCount - 1 To 0 Step -1
     If CboGrupo.List(Item) = LcNome Then
        VerificaGrupo = True
        Exit For
     Else
        VerificaGrupo = False
     End If
Next
End Function
Function BuscaDados()
On Error GoTo erroBusca

Dim RsGrupo As ADODB.Recordset
Dim a As Long

Dim LcCriterio As String

'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("select * from Usuario", True) ', dbOpenDynaset)
LcCriterio = "Nome='" & txt(0).Text & "'"
RsGrupo.Find LcCriterio
If Not RsGrupo.EOF Then
   txt(1).Text = RsGrupo!Senha
   txt(2).Text = ""
   CboGrupo.Text = RsGrupo!Grupo
   ExcluiReceita.Value = IIf(RsGrupo!ExcluiReceita, 1, 0)
Else
   txt(1).Text = ""
   CboGrupo.Text = ""
   ExcluiReceita.Value = 0
End If
RsGrupo.Close
Set RsGrupo = Nothing
Exit Function
erroBusca:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function

Private Sub ListUsuarios_DblClick()
Dim LcResposta As Integer
If LcAlterou Then
   LcResposta = MsgBox("As Alterações deste Usuário Não foram Salvas." & Chr(13) & "Salva-as Agora ?", 36, "Aviso")
   If LcResposta = 6 Then Salva
   LcAlterou = False
End If

txt(0).Text = ListUsuarios.Text
BuscaDados
LcAlterou = False
txt(0).SetFocus

End Sub

Private Sub ListUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Txt_Change(Index As Integer)
LcAlterou = True
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   Select Case Index
     Case Is = 0
        txt(1).SetFocus
     Case Is = 1
        txt(2).SetFocus
     Case Is = 2
        VerificaSenha
  End Select
End If

If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Select Case Index
   Case Is = 0
         BuscaDados
End Select
End Sub

Function Salva()
On Error GoTo errosalva
Dim RsGrupo As ADODB.Recordset
Dim a As Long
Dim LcNovo As Boolean
Dim LcCriterio As String
Dim afetados As Integer
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("select * from Usuario where Nome='" & txt(0).Text & "'")
If RsGrupo.EOF Then
    LcNovo = True
    ListUsuarios.AddItem txt(0).Text
End If
RsGrupo.Close
Set RsGrupo = Nothing
If LcNovo Then
    LcSql = "Insert into Usuario(Nome,Grupo,Senha,ExcluiReceita)Values("
    LcSql = LcSql & "'" & txt(0).Text & "',"
    LcSql = LcSql & "'" & CboGrupo.Text & "',"
    LcSql = LcSql & "'" & txt(1).Text & "',"
    LcSql = LcSql & ExcluiReceita.Value & ")"
Else
    LcSql = "Update Usuario set "
    LcSql = LcSql & "Nome='" & txt(0).Text & "',"
    LcSql = LcSql & "Grupo='" & CboGrupo.Text & "',"
    LcSql = LcSql & "ExcluiReceita=" & ExcluiReceita.Value & ","
    LcSql = LcSql & "Senha='" & txt(1).Text & "' where Nome='" & txt(0).Text & "'"
End If
Debug.Print LcSql
afetados = ExecutaSql(LcSql)
Debug.Print DEscricaoErro
LcAlterou = False
txt(0).Text = ""
txt(1).Text = ""
txt(2).Text = ""
CboGrupo.Text = ""
ExcluiReceita.Value = 0
Exit Function
errosalva:
MsgBox err.Description & err.Number
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select

End Function
Function VerificaSenha()
If Len(Trim(txt(1).Text)) = 0 Then
   MsgBox "Deve-se Digitar uma Sennha para Acesso...", 48, "Aviso"
   txt(1).SetFocus
   Exit Function
End If
If txt(1).Text <> txt(2).Text Then
   MsgBox "Senha Não Confere...", 48, "Aviso"
   txt(1).Text = ""
   txt(2).Text = ""
   txt(1).SetFocus
   Exit Function
End If
CboGrupo.SetFocus
End Function
