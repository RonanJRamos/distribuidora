VERSION 5.00
Begin VB.Form FrmGrupoEconomico 
   BackColor       =   &H00CAE1A2&
   Caption         =   "Cadastro Grupo Economico"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNovo 
      Caption         =   "&Incluir Registro"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   1980
      Width           =   2385
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   855
      Width           =   1185
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   855
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   1230
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   1605
      Width           =   1185
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   1605
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   1230
      Width           =   1185
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   2355
      Width           =   2385
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Tag             =   "S/N/S/00/N/CODIGO"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   1
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "S/T/S/01/S/DESCRICAO"
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Cadastro Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13515
   End
End
Attribute VB_Name = "FrmGrupoEconomico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado, a As Integer
Private WithEvents FrmPesquisar As frmPesquisa
Attribute FrmPesquisar.VB_VarHelpID = -1
Private RsAtual_Grupo_Ec As ADODB.Recordset

Private Function Desabilitatodos()
Dim a As Integer
txt(0).Enabled = False
txt(1).Enabled = False

End Function
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Custo

Set GlFormA = Me
For a = 0 To 1
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    If txt(a).Visible Then
       Select Case LcNome
           Case Is = "CODIGO"
                MtPesquisa(a).campo = "CODIGO"
                MtPesquisa(a).Indice = "ID"
           Case Is = "DESCRICAO"
                MtPesquisa(a).campo = "Descrição"
                MtPesquisa(a).Indice = "Nome"
           Case Else
                MtPesquisa(a).campo = LcNome
                MtPesquisa(a).Indice = "Nome"
        End Select
     End If
 Next
 LcIndice = "CODIGO"
End Function
Private Function DesabilitaCtr()
CmdPrimeiro.Enabled = False
CmdAnterior.Enabled = False
CmdUltimo.Enabled = False
CmdSeguinte.Enabled = False
MnMovimento.Enabled = False
MnRegistro.Enabled = False
CmdExcluir.Enabled = False
CmdPesquisar.Enabled = False
CmdOrdenar.Enabled = False
End Function
Function VinculaDados(Optional ID As Long = 0, Optional Nome As String = "", Optional Ordem As String = "", Optional PosicaoReg As Integer = -1)
On Error GoTo ErroVinculo
Dim StrWhere As String
If Len(Nome) > 0 Then StrWhere = " Nome='" & Nome & "'"
If ID > 0 Then StrWhere = " ID=" & ID
    If Len(StrWhere) > 0 Or Len(Ordem) > 0 Then
       If Len(StrWhere) > 0 Then StrWhere = " Where " & StrWhere
       Set RsAtual_Grupo_Ec = AbreRecordset("Select * from GrupoEconomico " & StrWhere & Ordem, True)
       If Not RsAtual_Grupo_Ec.EOF Then
        txt(0).Text = RsAtual_Grupo_Ec!ID
        txt(1).Text = RsAtual_Grupo_Ec!Nome
        'txt(1).SetFocus
        CmdSalvar.Enabled = False
        LcRegAtual = False
    End If
Else
    If PosicaoReg > -1 Then
     If Not RsAtual_Grupo_Ec.EOF And Not RsAtual_Grupo_Ec.BOF Then
         If PosicaoReg = 0 Then RsAtual_Grupo_Ec.MoveFirst
            If PosicaoReg = 1 Then RsAtual_Grupo_Ec.MovePrevious
            If PosicaoReg = 2 Then RsAtual_Grupo_Ec.MoveNext
            If PosicaoReg = 3 Then RsAtual_Grupo_Ec.MoveLast
            If Not RsAtual_Grupo_Ec.EOF And Not RsAtual_Grupo_Ec.BOF Then
               txt(0).Text = RsAtual_Grupo_Ec!ID
               txt(1).Text = RsAtual_Grupo_Ec!Nome
            Else
                If RsAtual_Grupo_Ec.EOF Then RsAtual_Grupo_Ec.MoveLast
               If RsAtual_Grupo_Ec.BOF Then RsAtual_Grupo_Ec.MoveFirst
            End If
     End If
      
    End If
 
End If

Exit Function
ErroVinculo:
    MsgBox err.Description & " " & err.Number
    
    Resume Next
    CmdSalvar.Enabled = False
End Function
Function Salv() As Boolean
LcP = GLPesquisa
LcI = LcIndice
Dim LcIncluir As Boolean
Dim StrSql    As String
Dim LcSeguranca As Double
On Error GoTo errSlvar
Dim RsAtualP As ADODB.Recordset
'===> Se for Consulta não Salva
Set RsAtualP = AbreRecordset("select * from GrupoEconomico ", True)
'===> Verifica se é inclusão
If Len(txt(0).Text) = 0 Then
   '==> Verifica se já cadastrou o nome do produto
   RsAtualP.Find "nome='" & txt(1).Text & "'"
   If Not RsAtualP.EOF Then
      MsgBox "O Grupo Economico " & txt(1).Text & " já foi cadastrado com o código:" & RsAtualP!ID, 64, "Aviso"
      Exit Function
   End If
   'RsAtualP.AddNew
   LcIncluir = True
Else
   If GlConfirmaAlteracao Then
      Resposta = MsgBox("Confirma a Alteração deste registro?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Aviso")
   Else
      Resposta = vbYes
   End If
   If Resposta = 7 Then GoTo Saida
   '===> Pesquisa o produto
   LcPes = "ID=" & txt(0).Text
   RsAtualP.Find LcPes
   If RsAtualP.EOF Then
     ' RsAtualP.AddNew
     LcIncluir = True
   End If
End If

If Len(txt(0).Text) = 0 Then
    StrSql = "Insert into GrupoEconomico (Nome) values ('" & txt(1).Text & "')"
    
Else
    StrSql = "Update GrupoEconomico Set " & _
           "Nome ='" & Replace(txt(1).Text, "'", "''") & "'," & _
           ID = " Where Codigo=" & txt(0).Text
End If
'Debug.Print StrSql
afetados = ExecutaSql(StrSql)
If LcTipoDados = 1 Then
   If Mantem.Value = 1 Then
      txt(0).Text = ""
   Else
      NovoReg
   End If
End If
Salv = True
RetornaCorFundo
txt(1).SetFocus
CmdSalvar.Enabled = False
Exit Function
Saida:
GLPesquisa = LcP
LcIndice = LcI
RsAtualP.Close
'Dbbase.Close
Set RsAtualP = Nothing
'Set Dbbase = Nothing

errSlvar:
MsgBox err.Description & " Nº: " & err.Number
'Resume 0
Exit Function
End Function
Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
VinculaDados 0, "", "", 1
GlMov = False
LcRegAtual = False
txt(1).SetFocus
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
GlTab = "DecricaoCusto"
If IsNumeric(txt(0).Text) Then
   If CLng(txt(0).Text) > 0 Then
     GlSq = "delete from GrupoEconomico where ID=" & txt(0).Text
      If ExecutaSql(GlSq) > 0 Then
        VinculaDados 0, "", " order by ID"
        VinculaDados 0, "", "", 3
      End If
   Else
        MsgBox "Registro não selecionado para a exclusâo", 64, "Aviso"
   End If
Else
   MsgBox "Registro não selecionado para a exclusâo", 64, "Aviso"
End If
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload frmPesquisa
Unload Me

End Sub

Private Sub CmdNovo_Click()
txt(0).Text = ""
    txt(1).Text = ""
    txt(1).SetFocus
End Sub

Private Sub CmdPesquisar_Click()
On Error Resume Next

Set FrmPesquisar = frmPesquisa
FrmPesquisar.Show , Me
LcRegAtual = False
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
VinculaDados 0, "", "", 0
GlMov = False
LcRegAtual = False
txt(1).SetFocus
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
If Len(txt(1).Text) = 0 Then
   MsgBox "Informe o nome do grupo economico para salvar!", 64, "Aviso"
Else
    If Salv() Then
        VinculaDados 0, txt(1).Text
        LcRegAtual = False
        NovoReg
    End If
End If

txt(1).SetFocus

End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
VinculaDados 0, "", "", 2
GlMov = False
txt(1).SetFocus
LcRegAtual = False

End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
VinculaDados 0, "", "", 3
txt(1).SetFocus
GlMov = False
LcRegAtual = False
End Sub

Private Sub Form_Activate()
On Error Resume Next
 Set GlFormA = Me
 CarreGamatriz
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
CmdSalvar.Enabled = True
End Sub

Private Sub Form_Load()
VinculaDados 0, "", " order by ID"
LcIndice = "ID"

End Sub

Private Sub FrmPesquisar_Pesquisou(Chave As String, Valor As String, Tipo As String, Posicao As Integer)
'On Error Resume Next
'Dim RsAtual_Grupo_Ec As ADODB.Recordset
Dim StrWhere As String
If Tipo = "N" Then
   StrWhere = Chave & "=" & Valor
Else
  StrWhere = Chave & " Like '%" & Valor & "%'"
End If
'Set RsAtual_Grupo_Ec = AbreRecordset("Select * from GrupoEconomico " & StrWhere, True)
  Dim SAir As Boolean
  SAir = False
  Do While Not RsAtual_Grupo_Ec.EOF And Not RsAtual_Grupo_Ec.BOF
      Select Case Posicao
        Case Is = 0
          RsAtual_Grupo_Ec.MoveFirst
          RsAtual_Grupo_Ec.Find StrWhere, 0, adSearchForward
          SAir = True
        Case Is = 1
          RsAtual_Grupo_Ec.Find StrWhere, RsAtual_Grupo_Ec.RecordCount - 1, adSearchBackward
          If Not RsAtual_Grupo_Ec.BOF Then
               If Tipo = "N" Then
                   If txt(0).Text <> RsAtual_Grupo_Ec!ID Then SAir = True
                Else
                   If txt(1).Text <> RsAtual_Grupo_Ec!Nome Then SAir = True
                End If
          End If
        Case Is = 2
          RsAtual_Grupo_Ec.Find StrWhere, 1, adSearchForward
          If Not RsAtual_Grupo_Ec.EOF Then
               If Tipo = "N" Then
                   If txt(0).Text <> RsAtual_Grupo_Ec!ID Then SAir = True
                Else
                   If txt(1).Text <> RsAtual_Grupo_Ec!Nome Then SAir = True
                End If
          End If
        End Select
        If SAir Then Exit Do
  Loop
If Not RsAtual_Grupo_Ec.EOF And Not RsAtual_Grupo_Ec.BOF Then
    txt(0).Text = RsAtual_Grupo_Ec!ID
    txt(1).Text = RsAtual_Grupo_Ec!Nome
    
    txt(1).SetFocus
    CmdSalvar.Enabled = False
    'MnSalvar.Enabled = False
    LcRegAtual = False
ElseIf RsAtual_Grupo_Ec.EOF Then RsAtual_Grupo_Ec.MoveFirst
ElseIf RsAtual_Grupo_Ec.BOF Then RsAtual_Grupo_Ec.MoveLast
    
End If
Exit Sub
ErroVinculo:
    Resume Next
    Dim a As Long
    On Error Resume Next
   
    txt(0).Text = ""
    txt(1).Text = ""
    CmdSalvar.Enabled = False
End Sub

