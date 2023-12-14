VERSION 5.00
Begin VB.Form FrmTransportadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Fonecedores"
   ClientHeight    =   4890
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   9480
      TabIndex        =   51
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   10680
      TabIndex        =   50
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   9480
      TabIndex        =   49
      Top             =   2280
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   10680
      TabIndex        =   48
      Top             =   3000
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   8760
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   7560
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   7560
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   3600
      TabIndex        =   12
      Tag             =   "S/T/N/14/N/Celular"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "S/T/N/13/N/Fax"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   960
      MaxLength       =   20
      TabIndex        =   14
      Tag             =   "S/T/N/30/S/CGC"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   5520
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   15
      Tag             =   "S/T/N/12/N/INSCEST"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   10
      Tag             =   "S/T/N/11/N/FONE2"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   960
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "S/T/N/10/N/FONE1"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "S/T/N/09/N/CEP"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   840
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "S/T/N/08/N/ESTADO"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   840
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "S/T/N/04/N/CIDADE"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   840
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "S/T/N/03/N/END"
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   840
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "S/T/S/01/N/RAZAOSOC"
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   0
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "S/T/S/00/S/CODIGO"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   120
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
      Left            =   7800
      TabIndex        =   30
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
      Left            =   9840
      TabIndex        =   29
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   9480
      TabIndex        =   24
      Top             =   3000
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   10680
      TabIndex        =   23
      Top             =   2280
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   10680
      TabIndex        =   22
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9480
      TabIndex        =   20
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   9480
      TabIndex        =   21
      Top             =   3840
      Width           =   2385
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
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite F5 Para Escolher a Cidade"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5160
      TabIndex        =   52
      Top             =   2520
      Width           =   2340
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9360
      X2              =   9360
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Label cidade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   47
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Celular"
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
      Index           =   2
      Left            =   2760
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Index           =   1
      Left            =   6600
      TabIndex        =   45
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone Opção"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   43
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
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
      TabIndex        =   31
      Top             =   3120
      Width           =   480
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
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fant."
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
      TabIndex        =   28
      Top             =   4680
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
      TabIndex        =   27
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uf"
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
      TabIndex        =   37
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.E.P.:"
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
      Left            =   2400
      TabIndex        =   39
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
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
      TabIndex        =   38
      Top             =   3600
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InscrIção"
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
      Left            =   4560
      TabIndex        =   25
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End."
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
      TabIndex        =   35
      Top             =   1440
      Width           =   420
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
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl"
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
      Left            =   5160
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
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
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   42
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razão"
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
      TabIndex        =   34
      Top             =   1080
      Width           =   600
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
      TabIndex        =   33
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Transportadoras"
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
      TabIndex        =   32
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
Attribute VB_Name = "FrmTransportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado, a As Integer
Private LcDesCidade As String
Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    txt(a).Enabled = False
Next
End Function

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Transportadora) Then VinculaDados
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
GlTab = "transportadora"
GlSq = "Select * from transportadora where codigo='" & txt(0).Text & "'"

If Exclui(Transportadora) = 1 Then
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
Call Teclas(KeyCode)
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
If MovImentacao(enPrimeiro, Transportadora) Then VinculaDados
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
Call SalvaRegistro(Transportadora)
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
If MovImentacao(enSeguinte, Transportadora) Then VinculaDados
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
If MovImentacao(enultimo, Transportadora) Then VinculaDados
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
      Call AbreBanco(Transportadora)
      VinculaDados
   Case Is = 3
      LcCap = "   <<Modo Consultar>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Transportadora)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True

txt(0).Enabled = False


End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Transportadora
For a = 0 To 30
   MtPesquisa(a).Campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).tipo = ""
Next

For a = 0 To 30
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).tipo = LcTipo
    If Err = 0 Then
       Select Case LcNome
           Case Is = "RAZAOSOC"
                MtPesquisa(a).Campo = "NOME"
           Case Is = "FONEOPC"
                MtPesquisa(a).Campo = "FONE OPCIONAL"
           Case Is = "CPF1"
                MtPesquisa(a).Campo = "CEP 1 DEP."
           Case Is = "CPF2"
                MtPesquisa(a).Campo = "CEP 2 DEP."
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
     Err = 0
 Next
 
End Function

Private Sub Form_Load()
On Error Resume Next
Me.Height = 5205
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
   SalvaRegistro (Transportadora)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Transportadora)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False

End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Transportadora) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(Transportadora) = 1 Then
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
If MovImentacao(enPrimeiro, Transportadora) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(Transportadora)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Transportadora) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Transportadora) Then VinculaDados
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
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Transportadora)


txt(0).Text = GlCampo0
txt(1).Text = GlCampo1
txt(2).Text = GlCampo2
txt(3).Text = GlCampo3
txt(4).Text = GlCampo4
'=== Exibe o nome da cidade
txt(5).Text = GlCampo5
txt(6).Text = GlCampo6
txt(7).Text = GlCampo7
BuscaCidade
txt(8).Text = GlCampo8
txt(9).Text = GlCampo9
txt(10).Text = GlCampo10
txt(11).Text = GlCampo11
txt(12).Text = GlCampo12
txt(13).Text = GlCampo13
txt(14).Text = GlCampo14
txt(15).Text = GlCampo15
txt(16).Text = GlCampo16
txt(18).Text = GlCampo18
txt(19).Text = GlCampo19
txt(25).Text = GlCampo25
txt(30).Text = GlCampo30

txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub txt_Change(Index As Integer)
On Error Resume Next

If Index = 0 Then
   GlAlteraCodigo = True
   GlCodigoAnterior = GlCampo0
End If
Call Alterado

End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
     BuscaCidade
  Else
    Call Teclas(KeyCode)
  End If
End If
If KeyCode < 112 Or KeyCode > 123 Then
      If Index = 0 Then
         If Not GlAlteraCodigo Then
            GlAlteraCodigo = True
            GlCodigoAnterior = GlCampo0
         End If
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
Function BuscaCidade()

Dim RsCidade As Recordset
AbreBase
Set RsCidade = Dbbase.OpenRecordset("select * from alid005")
txt(7).Text = Right("0000" & txt(7).Text, 4)
LcCriterioCi = "cod='" & txt(7).Text & "'"
RsCidade.FindFirst LcCriterioCi
If Not RsCidade.NoMatch Then
   cidade.Caption = RsCidade!Nome
   LcDesCidade = RsCidade!Nome
Else
   cidade.Caption = ""
   'MsgBox "O código da cidade não foi encontrado...,", 64, "Aviso"
End If
RsCidade.Close
Set RsCidade = Nothing



End Function

Private Sub Txt_LostFocus(Index As Integer)
If Index = 7 Then BuscaCidade
'If Index = 0 Then
'   If Not GLCalculacodigotransportadora Then txt(0).Text = Trim(txt(0).Text)
'End If
'If Not GLCalculacodigotransportadora Then If VerificaDuplicado(Index) Then txt(Index).SetFocus

End Sub
