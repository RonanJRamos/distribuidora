VERSION 5.00
Begin VB.Form FrmUnidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Unidades"
   ClientHeight    =   2505
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   6840
      TabIndex        =   2
      Tag             =   "S/N/S/01/S/QUANTIDADE"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   9480
      TabIndex        =   39
      Top             =   2100
      Width           =   2385
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   375
      Left            =   10680
      TabIndex        =   38
      Top             =   1350
      Width           =   1185
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   375
      Left            =   9480
      TabIndex        =   37
      Top             =   1725
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   375
      Left            =   10680
      TabIndex        =   36
      Top             =   1725
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   375
      Left            =   9480
      TabIndex        =   35
      Top             =   1350
      Width           =   1185
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   375
      Left            =   9480
      TabIndex        =   34
      Top             =   975
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   375
      Left            =   10680
      TabIndex        =   33
      Top             =   975
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   6480
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   7800
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   6600
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   5880
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   5400
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   5280
      MaxLength       =   1
      TabIndex        =   6
      Tag             =   "S/T/N/30/N/MOVCAIXA"
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   19
      Left            =   5280
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "S/D/N/19/N/VP"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   5520
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   5760
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   6000
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "S/T/N/08/N/SIMBOLO"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "S/T/N/02/N/COMPRA"
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "S/T/S/01/S/NOME"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Tag             =   "S/T/S/00/N/COD"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   360
   End
   Begin VB.TextBox DataS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox HoraS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   375
      Left            =   10680
      TabIndex        =   22
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   600
      Width           =   1185
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
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
      Left            =   5520
      TabIndex        =   43
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Movi. Caixa (S/N)"
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
      Left            =   3240
      TabIndex        =   42
      Top             =   2040
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Vista / Prazo (V/P)"
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
      Left            =   3240
      TabIndex        =   41
      Top             =   1680
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Na Compra (S/N)"
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
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9360
      X2              =   9360
      Y1              =   480
      Y2              =   2520
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs.:"
      Enabled         =   0   'False
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
      Left            =   7080
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
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
      Left            =   3720
      TabIndex        =   24
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simbolo"
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
      TabIndex        =   30
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
      Enabled         =   0   'False
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
      Left            =   4680
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   210
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
      TabIndex        =   29
      Top             =   1080
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9360
      Y1              =   960
      Y2              =   960
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
      TabIndex        =   28
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Unidades"
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
      Left            =   360
      TabIndex        =   27
      Top             =   120
      Width           =   11835
   End
   Begin VB.Menu MnArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu MnSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu MnRegistro 
      Caption         =   "&Registro"
      Begin VB.Menu MnSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnPesquisar 
         Caption         =   "&Pesquisar"
      End
      Begin VB.Menu MnOrdenar 
         Caption         =   "&Ordenar"
      End
      Begin VB.Menu MnExcluir 
         Caption         =   "&Excluir"
      End
   End
   Begin VB.Menu MnMovimento 
      Caption         =   "&Movimentar"
      Begin VB.Menu MnPrimeiro 
         Caption         =   "&Primeiro"
      End
      Begin VB.Menu MnAnterior 
         Caption         =   "&Anterior"
      End
      Begin VB.Menu MSeguinte 
         Caption         =   "&Seguinte"
      End
      Begin VB.Menu MnUltimo 
         Caption         =   "&Último"
      End
   End
   Begin VB.Menu MnPop 
      Caption         =   "&Pop"
      Visible         =   0   'False
      Begin VB.Menu PopPesquisar 
         Caption         =   "&Pesquisar"
      End
      Begin VB.Menu PopOrdenar 
         Caption         =   "&Ordenar"
      End
   End
End
Attribute VB_Name = "FrmUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado, a As Integer
Private LcDesunidade As String
Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    txt(a).Enabled = False
Next
End Function

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
txt(1).SetFocus
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
GlTab = "alid004"
GlSq = "Select * from alid004 where cod='" & txt(0).Text & "'"

If Exclui(Unidade) = 1 Then
      VinculaDados
End If
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload frmPesquisa
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdOrdenar_Click()
On Error Resume Next
FrmOrdena.Show , Me
End Sub

Private Sub CmdOrdenar_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPesquisar_Click()
MnPesquisar_Click
End Sub

Private Sub CmdPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
txt(1).SetFocus
End Sub

Private Sub CmdPrimeiro_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Call SalvaRegistro(Unidade)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
txt(1).SetFocus
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Unidade) Then VinculaDados
GlMov = False
txt(1).SetFocus
LcRegAtual = False
End Sub

Private Sub CmdSeguinte_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enultimo, Unidade) Then VinculaDados
txt(1).SetFocus
GlMov = False
LcRegAtual = False
End Sub



Private Sub CmdUltimo_KeyDown(KeyCode As Integer, Shift As Integer)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
 Set GlFormA = Me
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        LcCap = "   <<Modo Inclusão>>"
        DesabilitaCtr
   Case Is = 2
        LcCap = "   <<Modo Alteração>>"
      Call AbreBanco(Unidade)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Unidade)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
txt(1).SetFocus

End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Unidade

Set GlFormA = Me
For a = 0 To 30
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).tipo = LcTipo
    If txt(a).Visible Then
       Select Case LcNome
           Case Is = "COD"
                MtPesquisa(a).Campo = "CODIGO"
           Case Is = "NOME"
                MtPesquisa(a).Campo = "Descrição"
           Case Is = "RD"
                MtPesquisa(a).Campo = "TIPO"
           Case Is = "CPF3"
                MtPesquisa(a).Campo = "CEP 3 DEP."
           Case Is = "QudLocacao"
                MtPesquisa(a).Campo = "QUT LOCAÇÃO"
           Case Is = "UltimaLocacao"
                MtPesquisa(a).Campo = "ÚLTIMA LOCAÇÃO"
           Case Is = "ValorDevido"
                MtPesquisa(a).Campo = "VALOR DEVIDO"
           Case Is = "UltimoProduto"
                MtPesquisa(a).Campo = "ÚLTIMO PRODUTO"
           Case Is = "CodigoConvenio"
                MtPesquisa(a).Campo = "CÓDIGO CONVENIO"
           Case Else
                MtPesquisa(a).Campo = LcNome
        End Select
     End If
 Next
 
End Function

Private Sub Form_Load()
On Error Resume Next
Me.Height = 3165
Me.Width = 12000
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
 
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FechaBanco

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (Unidade)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Unidade)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False

End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(Unidade) = 1 Then
      VinculaDados
End If
LcRegAtual = False
End Sub

Private Sub MnOrdenar_Click()
On Error Resume Next
FrmOrdena.Show , Me
End Sub

Private Sub MnPesquisar_Click()
On Error Resume Next
frmPesquisa.Show , Me
LcRegAtual = False
End Sub

Private Sub MnPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(Unidade)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Unidade) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
HoraS.Text = Format(Time, "hh:mm:ss")
End Sub
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
Function VinculaDados()
On Error Resume Next
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Unidade)


txt(0).Text = GlCampo0
txt(1).Text = GlCampo1
txt(8).Text = GlCampo8
txt(2).Text = GlCampo2
txt(19).Text = GlCampo19
txt(30).Text = GlCampo30
txt(4).Text = GlCampo4
txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub txt_Change(Index As Integer)
Call Alterado

End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub
Function limpa()
Dim a As Long
On Error Resume Next
For a = 0 To 36
  txt(a).Text = ""
Next
txt(1).SetFocus
CmdSalvar.Enabled = False
End Function


Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 4 Then
   If Len(txt(Index).Text) = 0 Then Exit Sub
   If Not IsNumeric(txt(Index).Text) Then
      MsgBox "Digite Um Valor Numérico.", vbInformation, "Aviso"
      txt(Index).Text = ""
      txt(Index).SetFocus
      Exit Sub
   End If
End If
End Sub
