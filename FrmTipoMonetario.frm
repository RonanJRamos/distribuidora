VERSION 5.00
Begin VB.Form FrmTipoMonetario 
   BackColor       =   &H00CBB19C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Tipo Monetário"
   ClientHeight    =   3180
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Codigo 
      Height          =   285
      Left            =   6720
      TabIndex        =   45
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox DescricaoNFE 
      Height          =   315
      ItemData        =   "FrmTipoMonetario.frx":0000
      Left            =   1680
      List            =   "FrmTipoMonetario.frx":0002
      TabIndex        =   43
      Top             =   1440
      Width           =   3735
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
      TabIndex        =   18
      Tag             =   "S/N/N/18/N/VALORULTCOMPRA"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   7800
      TabIndex        =   17
      Tag             =   "S/D/N/16/N/DATAULTIMACOMPRA"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   6600
      TabIndex        =   15
      Tag             =   "S/T/N/15/N/EMAIL"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   5880
      TabIndex        =   14
      Tag             =   "S/T/N/14/N/Celular"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   5400
      TabIndex        =   13
      Tag             =   "S/T/N/13/N/Fax"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "S/T/N/30/N/MOVCAIXA"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Tag             =   "S/T/N/25/N/OBSERVACAO"
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   19
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "S/D/N/19/N/VP"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   5520
      TabIndex        =   16
      Tag             =   "S/T/N/12/N/incest"
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   12
      Tag             =   "S/T/N/11/N/FONE2"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   5760
      TabIndex        =   11
      Tag             =   "S/T/N/10/N/FONE1"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   6000
      TabIndex        =   10
      Tag             =   "S/T/N/09/N/CEP"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "S/T/N/08/N/VENDA"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   9
      Tag             =   "S/T/N/04/N/CIDADE"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   8
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   7
      Tag             =   "S/T/N/05/N/COMPLEMENTO"
      Top             =   -120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   6
      Tag             =   "S/T/N/03/N/END"
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "S/T/N/02/N/COMPRA"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Tag             =   "S/T/S/01/S/XTPMONET"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Tag             =   "S/T/S/00/N/TPMONET"
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
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
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
      Left            =   7800
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   375
      Left            =   10680
      TabIndex        =   21
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      TabIndex        =   20
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
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição NFe"
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
      TabIndex        =   44
      Top             =   1440
      Width           =   1365
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
      Left            =   3360
      TabIndex        =   42
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   41
      Top             =   2280
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
      Left            =   240
      TabIndex        =   40
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9360
      X2              =   9360
      Y1              =   480
      Y2              =   3240
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
      TabIndex        =   23
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Na Venda (S/N)"
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
      Left            =   240
      TabIndex        =   30
      Top             =   2280
      Width           =   1440
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
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   615
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
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Tipos Monetários"
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
      TabIndex        =   26
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
Attribute VB_Name = "FrmTipoMonetario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado, a As Integer
Private LcDesmonetario As String
Private Rs As Recordset

Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    Txt(a).Enabled = False
Next
DescricaoNFE.Enabled = False
End Function
Private Sub IniciaRecordset()
Dim StSql As String
StSql = "Select * from alid008 order by codigo"
AbreBase
Set Rs = Dbbase.OpenRecordset(StSql)
End Sub

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enAnterior, monetario) Then VinculaDados
Rs.MovePrevious
If Rs.BOF Then
   MsgBox "Este é o primeiro Registro.", 64, "Aviso"
   Rs.MoveFirst
End If
VinculaDados
GlMov = False
LcRegAtual = False
Txt(1).SetFocus
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
GlTab = "alid008"
GlSq = "Select * from alid008 where TPMONET='" & Txt(0).Text & "'"

If Exclui(monetario) = 1 Then
   VinculaDados
End If
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
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
 Txt(1).SetFocus
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
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPesquisar_Click()
MnPesquisar_Click
End Sub

Private Sub CmdPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enPrimeiro, monetario) Then VinculaDados
Rs.MoveFirst
VinculaDados
GlMov = False
LcRegAtual = False
Txt(1).SetFocus
End Sub

Private Sub CmdPrimeiro_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
'Call SalvaRegistro(monetario)
Salva
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
Txt(1).SetFocus
End Sub
Sub Salva()
Dim Inclusao As Boolean
Dim StrSql As String

Inclusao = True

If IsNumeric(Codigo.Text) Then
    If CLng(Codigo.Text) > 0 Then
       Inclusao = False
    End If
 End If
 If Inclusao Then
    Dim LcCodigo As String
    StrSql = "Insert into alid008 (TPMONET,XTPMONET,venda,COMPRA,VP,MOVCAIXA,DescricaoNFE) values("
    StrSql = StrSql & "'" & LcCodigo & "',"
    StrSql = StrSql & "'" & Replace(Txt(1).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(Txt(8).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(Txt(2).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(Txt(19).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(Txt(30).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(DescricaoNFE, "'", "''") & "')"
 Else
    StrSql = "Update alid008 Set "
    StrSql = StrSql & "XTPMONET='" & Txt(1).Text & "',"
    StrSql = StrSql & "venda='" & Replace(Txt(8).Text, "'", "''") & "',"
    StrSql = StrSql & "COMPRA='" & Replace(Txt(2).Text, "'", "''") & "',"
    StrSql = StrSql & "VP='" & Replace(Txt(19).Text, "'", "''") & "',"
    StrSql = StrSql & "MOVCAIXA='" & Replace(Txt(30).Text, "'", "''") & "',"
    StrSql = StrSql & "DescricaoNFE='" & Replace(DescricaoNFE, "'", "''") & "'"
    StrSql = StrSql & " where codigo=" & Codigo.Text
 End If
 'Debug.Print StrSql
 Dbbase.Execute StrSql
End Sub
Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enSeguinte, monetario) Then VinculaDados
Rs.MoveNext
If Rs.EOF Then
   MsgBox "Este é o ultimo Registro", 64, "Aviso"
   Rs.MoveLast
End If
VinculaDados
GlMov = False
Txt(1).SetFocus
LcRegAtual = False
End Sub

Private Sub CmdSeguinte_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enultimo, monetario) Then VinculaDados
Rs.MoveLast
VinculaDados
Txt(1).SetFocus
GlMov = False
LcRegAtual = False
End Sub



Private Sub CmdUltimo_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub DescricaoNFE_Click()
If LcTipoDados <> 3 Then CmdSalvar.Enabled = True
End Sub

Private Sub DescricaoNFE_KeyDown(KeyCode As Integer, Shift As Integer)
If LcTipoDados <> 3 Then CmdSalvar.Enabled = True
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
      'Call AbreBanco(monetario)
      IniciaRecordset
      VinculaDados
   Case Is = 3
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
     ' Call AbreBanco(monetario)
      IniciaRecordset
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
Txt(1).SetFocus

End Sub
Sub carregaTipoNFe()
        DescricaoNFE.AddItem ("")
        DescricaoNFE.AddItem ("01=Dinheiro")
        DescricaoNFE.AddItem ("02=Cheque")
        DescricaoNFE.AddItem ("03=Cartão de Crédito")
        DescricaoNFE.AddItem ("04=Cartão de Débito")
        DescricaoNFE.AddItem ("05=Crédito Loja")
        DescricaoNFE.AddItem ("10=Vale Alimentação")
        DescricaoNFE.AddItem ("11=Vale Refeição")
        DescricaoNFE.AddItem ("12=Vale Presente")
        DescricaoNFE.AddItem ("13=Vale Combustível")
        'DescricaoNFE.AddItem ("14=Duplicata Mercantil")
        DescricaoNFE.AddItem ("15=Boleto Bancário")
        DescricaoNFE.AddItem ("90= Sem pagamento")
        DescricaoNFE.AddItem ("99=Outros")
End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.monetario
For a = 0 To 30
   MtPesquisa(a).campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).Tipo = ""
Next
Set GlFormA = Me
For a = 0 To 30
    LcNome = Mid$(Txt(a).Tag, 12)
    LcTipo = Mid$(Txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    If Txt(a).Visible Then
       Select Case LcNome
           Case Is = "TPMONET"
                MtPesquisa(a).campo = "CODIGO"
           Case Is = "XTPMONET"
                MtPesquisa(a).campo = "Descrição"
           Case Is = "CPF2"
                MtPesquisa(a).campo = "CEP 2 DEP."
           Case Is = "CPF3"
                MtPesquisa(a).campo = "CEP 3 DEP."
           Case Is = "QudLocacao"
                MtPesquisa(a).campo = "QUT LOCAÇÃO"
           Case Is = "UltimaLocacao"
                MtPesquisa(a).campo = "ÚLTIMA LOCAÇÃO"
           Case Is = "ValorDevido"
                MtPesquisa(a).campo = "VALOR DEVIDO"
           Case Is = "UltimoProduto"
                MtPesquisa(a).campo = "ÚLTIMO PRODUTO"
           Case Is = "CodigoConvenio"
                MtPesquisa(a).campo = "CÓDIGO CONVENIO"
           Case Else
                MtPesquisa(a).campo = LcNome
        End Select
     End If
 Next
 
End Function

Private Sub Form_Load()
On Error Resume Next
'Me.Height = 3165
'Me.Width = 12000
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
'Top = 800
'Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
carregaTipoNFe
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FechaBanco

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (monetario)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (monetario)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False

End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, monetario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(monetario) = 1 Then
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
If MovImentacao(enPrimeiro, monetario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(monetario)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, monetario) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, monetario) Then VinculaDados
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
If LcTipoDados = 1 Then NovoReg 'Else Call RegistroAtual(monetario)
'Dim Rs As New ADODB.Recordset
'Set Rs = AbreRecordset("Select * from ALID008 where codigo=" & Txt(0).Text, True)

Txt(0).Text = Rs!TPMONET & ""
Txt(1).Text = Rs!XTPMONET & ""
Txt(8).Text = Rs!venda & ""
Txt(2).Text = Rs!COMPRA & ""
Txt(19).Text = Rs!vp & ""
Txt(30).Text = Rs!MOVCAIXA & ""
DescricaoNFE.Text = Rs!DescricaoNFE & ""
Codigo.Text = Rs!Codigo & ""
Txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
Call Alterado
GlCampo3 = DescricaoNFE.Text
End Sub


Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
  Txt(a).Text = ""
Next
Txt(1).SetFocus
CmdSalvar.Enabled = False
End Function


Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 8 Or Index = 2 Or Index = 30 Then
   If Txt(Index).Text <> "S" And Txt(Index).Text <> "s" _
   And Txt(Index).Text <> "N" And Txt(Index).Text <> "n" Then
      MsgBox "Escolha <S> para SIM  ou <N> para Não.", 64, "Aviso"
      Txt(Index).Text = ""
      'txt(Index).SetFocus
   End If
End If
If Index = 19 Then
    If Txt(Index).Text <> "P" And Txt(Index).Text <> "p" _
   And Txt(Index).Text <> "V" And Txt(Index).Text <> "v" Then
      MsgBox "Escolha <V> para A VISTA  ou <P> para A PRAZO.", 64, "Aviso"
      Txt(Index).Text = ""
      Txt(Index).SetFocus
   End If
End If
End Sub
