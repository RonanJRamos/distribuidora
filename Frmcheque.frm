VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmcheques 
   BackColor       =   &H00CBB19C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de cheques"
   ClientHeight    =   4065
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10995
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   2
      Tag             =   "S/T/S/05/N/CHEQUE"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Devolvido 
      BackColor       =   &H00CBB19C&
      Caption         =   "Devolvido"
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
      Left            =   4440
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   6840
      TabIndex        =   9
      Tag             =   "S/T/N/4/N/TelEmitente"
      Top             =   2400
      Width           =   1455
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   6840
      TabIndex        =   42
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Compensado 
      BackColor       =   &H00CBB19C&
      Caption         =   "Compensado"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin MSMask.MaskEdBox emissao 
      Height          =   405
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   8520
      TabIndex        =   18
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   8520
      TabIndex        =   20
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   9720
      TabIndex        =   23
      Top             =   2760
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   8
      Tag             =   "S/T/N/10/N/Emitente"
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1560
      TabIndex        =   10
      Tag             =   "S/T/N/08/N/PASSADOPARA"
      Top             =   2880
      Width           =   6735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   3480
      TabIndex        =   3
      Tag             =   "S/T/N/06/N/BANCO"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   6480
      TabIndex        =   4
      Tag             =   "S/T/N/03/N/AGENCIA"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   0
      Tag             =   "S/T/N/02/N/PEDIDO"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Tag             =   "S/T/S/01/S/CLIENTE"
      Top             =   600
      Width           =   5175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   15
      Tag             =   "S/T/N/00/N/codigo"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
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
      Left            =   4080
      Top             =   720
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
      Left            =   6720
      TabIndex        =   29
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
      Left            =   8880
      TabIndex        =   28
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   8520
      TabIndex        =   22
      Top             =   2760
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   9720
      TabIndex        =   21
      Top             =   2040
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   9720
      TabIndex        =   17
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8520
      TabIndex        =   16
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   8520
      TabIndex        =   25
      Top             =   3480
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
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin MSMask.MaskEdBox dataentrada 
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox deposito 
      Height          =   405
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   714
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Devolucao 
      Height          =   405
      Index           =   0
      Left            =   7320
      TabIndex        =   14
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Dev."
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
      Left            =   6360
      TabIndex        =   44
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fone Emitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   43
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Efetivação"
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
      TabIndex        =   41
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N. Cheque"
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
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programado para"
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
      TabIndex        =   39
      Top             =   1680
      Width           =   1680
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recebido em "
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
      Left            =   2640
      TabIndex        =   38
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   8400
      X2              =   8400
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   -120
      X2              =   9360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
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
      TabIndex        =   27
      Top             =   600
      Width           =   675
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
      TabIndex        =   26
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passado Para"
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
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agencia"
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
      Left            =   5640
      TabIndex        =   32
      Top             =   1200
      Width           =   780
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
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
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
      Left            =   2760
      TabIndex        =   36
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emitente"
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
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   2280
      TabIndex        =   31
      Top             =   600
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Cheques Recebidos"
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
      TabIndex        =   30
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
Attribute VB_Name = "Frmcheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado As Integer
Private LcDesCidade As String

Private Type TipoVend
      Codigo As String
      Nome As String
End Type
Private LcNavegando, LcBuscaCliente As Integer
Private LcTamanho, a As Integer
Private MtVendedor() As TipoVend

Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    Txt(a).Enabled = False
Next
End Function




Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Cheques) Then VinculaDados
GlMov = False
LcRegAtual = False
Txt(2).SetFocus
LcBuscaCliente = False
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
LcNavegando = True
GlTab = "cheques"
GlSq = "Select * from cheques where cheque='" & Txt(5).Text & "' and agencia='" & Txt(3).Text & "' and banco='" & Txt(6).Text & "'"

If Exclui(Cheques) = 1 Then
      VinculaDados
End If

LcBuscaCliente = False

End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
  Txt(2).SetFocus
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
  Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdOrdenar_Click()
On Error Resume Next
LcNavegando = True
FrmOrdena.Show , Me
LcBuscaCliente = False
End Sub

Private Sub CmdOrdenar_KeyDown(KeyCode As Integer, Shift As Integer)
 
LcNavegando = True
Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If


End Sub

Private Sub CmdPesquisar_Click()

LcNavegando = True
MnPesquisar_Click
End Sub

Private Sub CmdPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
 LcNavegando = True
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  LcBuscaCliente = False
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
LcNavegando = True
GlMov = True
If MovImentacao(enPrimeiro, Cheques) Then VinculaDados
GlMov = False
LcRegAtual = False
Txt(2).SetFocus
LcBuscaCliente = False
End Sub

Private Sub CmdPrimeiro_KeyDown(KeyCode As Integer, Shift As Integer)
 LcNavegando = True
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Dim RsTipo      As Recordset
Dim LcPesquisa  As String
Dim LcTipoM     As String
Dim LcCodigo    As String
Dim LcValor     As Double

LcNavegando = True
Call SalvaRegistro(Cheques)
If GlInclusaoCheque Then
   AbreBase
   Set RsTipo = Dbbase.OpenRecordset("select * from Alid008 order by TPMONET", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcPesquisa = "XTPMONET='CHEQUE'"
   RsTipo.FindFirst LcPesquisa
   If Not RsTipo.NoMatch Then
      LcTipoM = RsTipo!TPMONET
   Else
      If Not RsTipo.EOF Then
         RsTipo.MoveLast
         LcCodigo = Right("00" & CStr(CInt(RsTipo!TPMONET) + 1), 2)
      Else
         LcCodigo = "01"
      End If
      RsTipo.AddNew
      RsTipo("TPMONET") = LcCodigo
      RsTipo("XTPMONET") = "CHEQUE"
      RsTipo("VENDA") = "S"
      RsTipo("COMPRA") = "S"
      RsTipo("VP") = "V"
      RsTipo("MOVCAIXA") = "S"
      RsTipo.Update
      LcTipoM = LcCodigo
   End If
    LcValor = CDbl(valor.Text)
    Call lancacaixa("Receita", Txt(5).Text, LcTipoM, LcValor)
End If

VinculaDados
NovoReg
If LcTipoDados = 1 Then limpa
Txt(2).SetFocus
LcRegAtual = False
LcBuscaCliente = False
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 LcNavegando = True
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If

End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
LcNavegando = True
GlMov = True
If MovImentacao(enSeguinte, Cheques) Then VinculaDados
GlMov = False
Txt(2).SetFocus
LcRegAtual = False
LcBuscaCliente = False
End Sub

Private Sub CmdSeguinte_KeyDown(KeyCode As Integer, Shift As Integer)
  LcNavegando = True
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  cBuscaCliente = False
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
LcNavegando = True
GlMov = True
If MovImentacao(enultimo, Cheques) Then VinculaDados
Txt(2).SetFocus
GlMov = False
LcRegAtual = False
LcBuscaCliente = False
End Sub



Private Sub CmdUltimo_KeyDown(KeyCode As Integer, Shift As Integer)
   LcNavegando = True
 Txt(2).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub
Private Sub Compensado_Click()
LcCor = Compensado.BackColor
Call Alterado
GlCampo25 = Compensado
Compensado.BackColor = LcCor
End Sub

Private Sub Compensado_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub dataentrada_Change()
If LcRegAtual Then Exit Sub

GlCampo23 = emissao(0).Text
GlCampo22 = dataentrada.Text
GlCampo24 = deposito(1).Text
Call Alterado
End Sub

Private Sub dataentrada_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub dataentrada_LostFocus()
If dataentrada.Text = "  /  /  " Then Exit Sub
If Not IsDate(dataentrada.Text) Then
   MsgBox "A Data digitada não é válida...", 64, "Aviso"
   dataentrada.SetFocus
End If
End Sub

Private Sub deposito_Change(Index As Integer)
If LcRegAtual Then Exit Sub

GlCampo23 = emissao(0).Text
GlCampo22 = dataentrada.Text
GlCampo24 = deposito(1).Text
GlCampo26 = Devolucao(0).Text

Call Alterado
End Sub


Private Sub deposito_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub deposito_LostFocus(Index As Integer)

If deposito(1).Text = "  /  /  " Then Exit Sub
If Not IsDate(deposito(1).Text) Then
   MsgBox "A Data digitada não é válida...", 64, "Aviso"
   deposito(1).SetFocus
End If
End Sub

Private Sub Devolucao_Change(Index As Integer)
If LcRegAtual Then Exit Sub

GlCampo23 = emissao(0).Text
GlCampo22 = dataentrada.Text
GlCampo24 = deposito(1).Text
GlCampo26 = Devolucao(0).Text

Call Alterado
End Sub

Private Sub Devolucao_LostFocus(Index As Integer)
If Devolucao(0).Text = "  /  /  " Then Exit Sub
If Not IsDate(Devolucao(0).Text) Then
   MsgBox "A Data digitada não é válida...", 64, "Aviso"
   Devolucao(0).SetFocus
End If
End Sub

Private Sub Devolvido_Click()
LcCor = Devolvido.BackColor
Call Alterado
GlCampo27 = Devolvido
Devolvido.BackColor = LcCor
End Sub

Private Sub emissao_Change(Index As Integer)
If LcRegAtual Then Exit Sub

GlCampo23 = emissao(0).Text
GlCampo22 = dataentrada.Text
GlCampo24 = deposito(1).Text
Call Alterado
End Sub

Private Sub emissao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub emissao_LostFocus(Index As Integer)
If emissao(0).Text = "  /  /  " Then Exit Sub
If Not IsDate(emissao(0).Text) Then
   MsgBox "A Data digitada não é válida...", 64, "Aviso"
   emissao(0).SetFocus
End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        DesabilitaCtr
        LcCap = "   <<Modo Inclusão>>"
   Case Is = 2
        LcCap = "   <<Modo Alteração>>"
      Call AbreBanco(Cheques)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Cheques)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
If Not GLCalculacodigoCliente Then
   Txt(2).SetFocus
Else
  'txt(0).Enabled = False
End If
Label1.Caption = Label1.Caption & LcCap

End Sub
Function CarreGamatriz()
On Error Resume Next
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Cheques

For a = 0 To 30
    LcNome = Mid$(Txt(a).Tag, 12)
    LcTipo = Mid$(Txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    If err = 0 Then
       Select Case LcNome
           Case Is = "FONEOPC"
                MtPesquisa(a).campo = "FONE OPCIONAL"
           Case Is = "CPF1"
                MtPesquisa(a).campo = "CEP 1 DEP."
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
    err = 0
 Next
 
End Function

Private Sub Form_Load()
On Error Resume Next
Me.Height = 4725
Me.Width = 11085
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
   SalvaRegistro (Cheques)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Cheques)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False
FrmPrincipal.SetFocus
End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Cheques) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(Cheques) = 1 Then
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
If MovImentacao(enPrimeiro, Cheques) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(Cheques)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Cheques) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Cheques) Then VinculaDados
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
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Cheques)


Txt(0).Text = GlCampo0
Txt(1).Text = GlCampo1
Txt(2).Text = GlCampo2
Txt(3).Text = GlCampo3
Txt(4).Text = GlCampo4
'=== Exibe o nome da cidade
Txt(5).Text = GlCampo5
Txt(6).Text = GlCampo6
Txt(7).Text = GlCampo7
'BuscaCidade
Txt(8).Text = GlCampo8
valor.Text = GlCampo9
Txt(10).Text = GlCampo10
Txt(11).Text = GlCampo11
Txt(12).Text = GlCampo12
Txt(13).Text = GlCampo13
Txt(14).Text = GlCampo14
Txt(15).Text = GlCampo15
Txt(16).Text = GlCampo16
Txt(18).Text = GlCampo18
Txt(17).Text = GlCampo17
Txt(19).Text = GlCampo19
Txt(20).Text = GlCampo20
Txt(25).Text = GlCampo25
Txt(30).Text = GlCampo30
Txt(21).Text = GlCampo21
Compensado = CInt(GlCampo25)
Devolvido = CInt(GlCampo27)
If GlCampo23 = "" Then
   emissao(0).Text = "  /  /  "
Else
   emissao(0).Text = Format(GlCampo23, "dd/mm/yy")
End If

If GlCampo26 = "" Then
   Devolucao(0).Text = "  /  /  "
Else
   Devolucao(0).Text = Format(GlCampo26, "dd/mm/yy")
End If

If GlCampo22 = "" Then
   dataentrada.Text = "  /  /  "
Else
   dataentrada.Text = Format(GlCampo22, "dd/mm/yy")
End If

If GlCampo24 = "" Then
   deposito(1).Text = "  /  /  "
Else
   deposito(1).Text = Format(GlCampo24, "dd/mm/yy")
End If

Txt(2).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
Call Alterado

End Sub
Function BuscaCliente()
Dim rsCliente As Recordset
If LcBuscaCliente Then Exit Function
If LcNavegando Then Exit Function
If Len(Txt(1).Text) = 0 Then Exit Function
AbreBase
Set rsCliente = Dbbase.OpenRecordset("select * from alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If GLCalculacodigoCliente Then
   If IsNumeric(Txt(1).Text) Then
      Txt(1).Text = Right("00000" & Txt(1).Text, 5)
   End If
End If
LcCriterioche = "codigo='" & Txt(1).Text & "'"
rsCliente.FindFirst LcCriterioche
If Not rsCliente.NoMatch Then
   Txt(1).Text = rsCliente!RAZAOSOC
   Codigo.Text = rsCliente!Codigo
Else
   FrmPesquisaCliente.Txt.Text = Txt(1).Text
   GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt(1).Text) & "*'  order by RAZAOSOC"
   FrmPesquisaCliente.Show , Me
   LcAlteradoCliente = False
   LcBuscaCliente = True
End If
rsCliente.Close
Set rsCliente = Nothing
End Function
Function BuscaPedido()
Dim RsPedido As Recordset, rsCliente As Recordset
If Len(Txt(2).Text) = 0 Then Exit Function
AbreBase
Txt(2).Text = Right("000000" & Txt(2).Text, 6)
Set RsPedido = Dbbase.OpenRecordset("select * from Orcamento where doc='" & Txt(2).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set rsCliente = Dbbase.OpenRecordset("select * from alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsPedido.EOF Then
   'MsgBox "O pedido " & txt(2).Text & " Não foi encontrado...", 64, "Aviso"
   'txt(2).SetFocus
   Exit Function
End If
LcCriterio = "CODIGO='" & RsPedido!Cliente & "'"
rsCliente.FindFirst LcCriterio
If Not rsCliente.NoMatch Then
   Codigo.Text = rsCliente!Codigo
   Txt(1).Text = rsCliente!RAZAOSOC
   LcBuscaCliente = True
End If
rsCliente.Close
RsPedido.Close

Set rsCliente = Nothing
Set RsPedido = Nothing


End Function

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
LcNavegando = False
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If Index = 1 Then LcBuscaCliente = False
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
  emissao(a).Text = "  /  /  "
  
Next
deposito(1).Text = "  /  /  "
Txt(2).SetFocus
Codigo.Text = ""
CmdSalvar.Enabled = False
Compensado = 0
End Function

Private Sub Txt_LostFocus(Index As Integer)
If Index = 2 Then BuscaPedido
If Index = 1 Then BuscaCliente
End Sub

Private Sub valor_Change()
GlCampo9 = valor.Text
Call Alterado
End Sub

Private Sub valor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub valor_LostFocus()
If Len(valor.Text) = 0 Then Exit Sub
If Not IsNumeric(valor.Text) Then
   MsgBox "O Valor Digitado Não é Válido...", 64, "Aviso"
   valor.Text = ""
   valor.SetFocus
End If

End Sub
