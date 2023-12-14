VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCliente 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de CLientes"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox GrupoEconomico 
      Height          =   315
      ItemData        =   "Frmcliente.frx":0000
      Left            =   9000
      List            =   "Frmcliente.frx":0002
      TabIndex        =   89
      Top             =   503
      Width           =   2775
   End
   Begin VB.CheckBox Bloqueado 
      BackColor       =   &H00CAE1A2&
      Caption         =   "Bloqueado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2280
      TabIndex        =   84
      Top             =   495
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   83
      Tag             =   "S/T/S/00/S/CODIGO"
      Top             =   473
      Width           =   975
   End
   Begin VB.TextBox EmailFinanceiro 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Tag             =   "S/T/N/32/N/email"
      ToolTipText     =   "para mais de um email coloque separado por ;"
      Top             =   4080
      Width           =   7455
   End
   Begin VB.ComboBox TipoContr 
      Height          =   315
      ItemData        =   "Frmcliente.frx":0004
      Left            =   4440
      List            =   "Frmcliente.frx":0011
      TabIndex        =   17
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton CmdObs 
      Caption         =   "Observações"
      Height          =   495
      Left            =   9480
      TabIndex        =   78
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Numero 
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   2160
      TabIndex        =   26
      Top             =   6720
      Width           =   7095
   End
   Begin VB.CheckBox Comodato 
      BackColor       =   &H00CAE1A2&
      Caption         =   "Comodato"
      Height          =   375
      Left            =   9600
      TabIndex        =   76
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "S/T/N/11/N/FONE2"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   32
      Left            =   840
      TabIndex        =   13
      Tag             =   "S/T/N/32/N/email"
      ToolTipText     =   "para mais de um email coloque separado por ;"
      Top             =   3720
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proposta de Compra"
      Height          =   495
      Left            =   9480
      TabIndex        =   71
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   23
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   14
      Tag             =   "S/T/N/23/N/CondicaoEspecial"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      ItemData        =   "Frmcliente.frx":0068
      Left            =   2160
      List            =   "Frmcliente.frx":006A
      TabIndex        =   23
      Top             =   6000
      Width           =   5175
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   7440
      TabIndex        =   68
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   20
      Left            =   5640
      TabIndex        =   25
      Tag             =   "S/M/N/20/N/CreditoUtilizado"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   17
      Left            =   2160
      TabIndex        =   24
      Tag             =   "S/M/N/17/N/LimiteCredito"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   9480
      TabIndex        =   63
      Top             =   1920
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   10680
      TabIndex        =   62
      Top             =   1920
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   9480
      TabIndex        =   61
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   10680
      TabIndex        =   60
      Top             =   3600
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   8760
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   7560
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   15
      Left            =   840
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "S/T/N/15/N/CONTATO"
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   3600
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   7200
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "S/T/N/13/N/Fax"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   7560
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   840
      MaxLength       =   20
      TabIndex        =   10
      Tag             =   "S/T/N/10/N/FONE1"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "S/T/N/08/N/ESTADO"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   840
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "S/T/N/04/N/CIDADE"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   840
      MaxLength       =   20
      TabIndex        =   6
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   6360
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "S/T/N/03/N/END"
      Top             =   2160
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   840
      MaxLength       =   60
      TabIndex        =   1
      Tag             =   "S/T/N/02/N/FANTASIA"
      Top             =   1440
      Width           =   6255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   840
      MaxLength       =   60
      TabIndex        =   0
      Tag             =   "S/T/S/01/N/RAZAOSOC"
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
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
      Left            =   7680
      TabIndex        =   43
      Top             =   0
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
      Left            =   9720
      TabIndex        =   42
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   9480
      TabIndex        =   37
      Top             =   3600
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   10680
      TabIndex        =   36
      Top             =   2880
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   10680
      TabIndex        =   35
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9480
      TabIndex        =   33
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   9480
      TabIndex        =   34
      Tag             =   "f"
      Top             =   5520
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
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   5
      Mask            =   "99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   18
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99.999-999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   19
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "999.999.999-99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   20
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   17
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox InscMunic 
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Top             =   5400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Suframa 
      Height          =   285
      Left            =   4800
      TabIndex        =   22
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   17
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataCadastro 
      Height          =   375
      Left            =   5640
      TabIndex        =   85
      Top             =   473
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Economico"
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
      TabIndex        =   88
      Top             =   600
      Width           =   1695
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
      TabIndex        =   87
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastrado em"
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
      Left            =   4200
      TabIndex        =   86
      Top             =   540
      Width           =   1440
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Financeiro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suframa"
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
      Left            =   3840
      TabIndex        =   81
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscr. Munic"
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
      TabIndex        =   80
      Top             =   5400
      Width           =   1140
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Contribuinte"
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
      TabIndex        =   79
      Top             =   4560
      Width           =   1650
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Dados Adicionais (NF)"
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RG"
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
      Left            =   3840
      TabIndex        =   75
      Top             =   5040
      Width           =   285
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
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
      TabIndex        =   74
      Top             =   4560
      Width           =   390
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Aniversário"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   72
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Cond. Especial"
      Height          =   255
      Left            =   6600
      TabIndex        =   70
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contato"
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
      TabIndex        =   69
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito Utilizado"
      Height          =   255
      Left            =   4080
      TabIndex        =   67
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Limite de Crédito"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Telemarketing que Atende"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite F5 Para Escolher a Cidade"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6960
      TabIndex        =   64
      Top             =   2880
      Width           =   2340
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   9360
      X2              =   9360
      Y1              =   960
      Y2              =   7320
   End
   Begin VB.Label cidade 
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   59
      Top             =   2880
      Width           =   2175
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
      TabIndex        =   58
      Top             =   2520
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
      TabIndex        =   57
      Top             =   3360
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
      TabIndex        =   55
      Top             =   3360
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
      TabIndex        =   44
      Top             =   3360
      Width           =   480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   -120
      X2              =   9360
      Y1              =   4920
      Y2              =   4920
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
      TabIndex        =   56
      Top             =   5400
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
      TabIndex        =   41
      Top             =   1440
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
      TabIndex        =   40
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   3240
      Y2              =   3240
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
      Left            =   3960
      TabIndex        =   49
      Top             =   2880
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
      Left            =   4680
      TabIndex        =   51
      Top             =   2880
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
      Left            =   240
      TabIndex        =   50
      Top             =   5040
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscrição"
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
      Left            =   6720
      TabIndex        =   38
      Top             =   4560
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
      TabIndex        =   47
      Top             =   2160
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
      Left            =   6720
      TabIndex        =   53
      Top             =   5400
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
      TabIndex        =   48
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
      TabIndex        =   52
      Top             =   2520
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
      TabIndex        =   54
      Top             =   2880
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
      TabIndex        =   46
      Top             =   1080
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Clientes"
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
      TabIndex        =   45
      Top             =   0
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
      Begin VB.Menu MnRecuperaDataCadastro 
         Caption         =   "&Recuperar Data cadastro"
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
Attribute VB_Name = "FrmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado As Integer
Private LcDesCidade As String

Private Type TipoVend
      codigo As String
      Nome As String
End Type

Private LcTamanho, a As Integer
Private LcTamanhoGr As Integer
Private MtVendedor() As TipoVend
Private MtGrupoEc() As TipoVend
Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    txt(a).Enabled = False
Next
End Function


Private Sub Bloqueado_Click()
CmdSalvar.Enabled = True

End Sub

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Cliente) Then VinculaDados
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
GlTab = "Alid001"
GlSq = "Select * from alid001 where codigo='" & txt(0).Text & "'"
If Exclui(Cliente) = 1 Then
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

Private Sub CmdObs_Click()
Load FrmObsCliente
FrmObsCliente.codigo.Caption = txt(0).Text
FrmObsCliente.Show , Me
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
LcIndice = "RAZAOSOC"
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
If MovImentacao(enPrimeiro, Cliente) Then VinculaDados
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
Dim RsClient    As Recordset
Dim CNPJ        As String
Dim Inscricao   As String
Dim Estado      As String

CNPJ = Mask(1).Text
CNPJ = Replace(CNPJ, ",", "")
CNPJ = Replace(CNPJ, ".", "")
CNPJ = Replace(CNPJ, "-", "")
CNPJ = Replace(CNPJ, "/", "")
CNPJ = Replace(CNPJ, "\", "")
CNPJ = Replace(CNPJ, " ", "")

If Len(CNPJ) = 0 Then
    CNPJ = Mask(4).Text
    CNPJ = Replace(CNPJ, ",", "")
    CNPJ = Replace(CNPJ, ".", "")
    CNPJ = Replace(CNPJ, "-", "")
    CNPJ = Replace(CNPJ, "/", "")
    CNPJ = Replace(CNPJ, "\", "")
    CNPJ = Replace(CNPJ, " ", "")
End If
If TipoContr.Text = "" Then
        MsgBox "Informe o Tipo de Contribuição do Cliente.", 64, "Aviso"
        Mask(2).SetFocus
        Exit Sub
End If
If Len(CNPJ) = 0 Then
    MsgBox "É nescessario cadastrar o CNPJ ou CPF do cliente.", 64, "Aviso"
    Mask(1).SetFocus
    SendKeys "+{home}+{end}"
    Exit Sub
 End If
Estado = txt(8).Text
Estado = Trim(Estado)
If Len(Estado) = 0 Then
   MsgBox "É nescessario cadastrar o estado do cliente.", 64, "Aviso"
   txt(8).SetFocus
   SendKeys "+{home}+{end}"
   Exit Sub
End If
If Len(CNPJ) > 11 Then
   If Not Calc_CNPJ(CNPJ) Then
      MsgBox "O CNPJ do cliente é invalido.", 64, "Aviso"
      Exit Sub
   End If
Else
   If Not Calc_CPF(CNPJ) Then
      MsgBox "O CPF do cliente é invalido.", 64, "Aviso"
      Exit Sub
   End If
End If
Inscricao = Mask(2).Text
Inscricao = Replace(Inscricao, ",", "")
Inscricao = Replace(Inscricao, ".", "")
Inscricao = Replace(Inscricao, "-", "")
Inscricao = Replace(Inscricao, "/", "")
Inscricao = Replace(Inscricao, "\", "")
Inscricao = Replace(Inscricao, " ", "")
If TipoContr.Text = "1 - Contribuinte ICMS" Then
    If Len(Inscricao) = 0 Then
        MsgBox "Informe a inscrição estadual do Cliente.", 64, "Aviso"
        Mask(2).SetFocus
        Exit Sub
    Else
        If Consiste(Inscricao, Estado) <> 0 Then
           MsgBox "A Inscrição Estadual do cliente é invalida.", 64, "Aviso"
           'ValidaEntradaSintegra = False
           Exit Sub
        End If
    End If
End If
If TipoContr.Text = "2 - Contribuinte isento de Inscrição" Then
    
End If



AbreBase


Set RsClient = Dbbase.OpenRecordset("ALID001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If LcTipoDados = 1 Then
    If Mask(1).Text <> "  .   .   /    -  " Then
       LcPes = "cgc='" & Mask(1).Text & "'"
    Else
       LcPes = "cpf='" & Mask(4).Text & "'"
    End If
       
    RsClient.FindFirst LcPes
    If Not RsClient.NoMatch Then
       If Mask(1).Text <> "  .   .   /    -  " Then
          LcResp = MsgBox("O CNPJ " & Mask(1).Text & "  Já Foi Cadastrado para o Cliente" & Chr(10) & RsClient!razaosoc, 65, "Aviso")
          Exit Sub
       Else
          LcResp = MsgBox("O CPF " & Mask(4).Text & "  Já Foi Cadastrado para o Cliente" & Chr(10) & RsClient!razaosoc, 65, "Aviso")
          Exit Sub
       End If
       
       If LcResp = 2 Then Exit Sub
    End If
Else
    If Mask(1).Text <> "  .   .   /    -  " Then
       LcPes = "cgc='" & Mask(1).Text & "' and codigo<>'" & txt(0).Text & "'"
    Else
       LcPes = "cpf='" & Mask(4).Text & "' and codigo<>'" & txt(0).Text & "'"
    End If
    RsClient.FindFirst LcPes
    If Not RsClient.NoMatch Then
     '  If Mask(1).Text <> "  .   .   /    -  " Then
     '     LcResp = MsgBox("O CNPJ " & Mask(1).Text & "  Já Foi Cadastrado para o Cliente" & Chr(10) & RsClient!razaosoc, 65, "Aviso")
     '  Else
     '     LcResp = MsgBox("O CPF " & Mask(1).Text & "  Já Foi Cadastrado para o Cliente" & Chr(10) & RsClient!razaosoc, 65, "Aviso")
     '  End If
       
     '  If LcResp = 2 Then Exit Sub
    End If
End If

RsClient.Close

Call SalvaRegistro(Cliente)
VinculaDados
LcRegAtual = False
'NovoReg
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
If MovImentacao(enSeguinte, Cliente) Then VinculaDados
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
If MovImentacao(enultimo, Cliente) Then VinculaDados
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
Function CarregaGrupoEconomico()
On Error GoTo errc
'Dim RsAtual_Grupo_Ec As ADODB.Recordset
'AbreBase
'Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Debug.Print conexaoAdo.ConnectionString
Dim RsVendedor As ADODB.Recordset
Set RsVendedor = AbreRecordset("Select * from GrupoEconomico order by nome", True)
LcTamanhoGr = 0
Do Until RsVendedor.EOF
   If err.Number <> 0 Then Exit Do
   ReDim Preserve MtGrupoEc(LcTamanhoGr)
   MtGrupoEc(LcTamanhoGr).codigo = RsVendedor!ID
   MtGrupoEc(LcTamanhoGr).Nome = RsVendedor!Nome
   GrupoEconomico.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
   LcTamanhoGr = LcTamanhoGr + 1
   
Loop
If LcTamanhoGr > 0 Then LcTamanhoGr = LcTamanhoGr - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:
MsgBox err.Description & " " & err.Number
Resume 0
Exit Function

End Function
Function CarregaTelemarketing()
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsVendedor.EOF
   ReDim Preserve MtVendedor(LcTamanho)
   MtVendedor(LcTamanho).codigo = RsVendedor!codigo
   MtVendedor(LcTamanho).Nome = RsVendedor!Nome
   vendedor.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
   LcTamanho = LcTamanho + 1
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:

Exit Function

End Function

Private Sub Command1_Click()
FrmPropostaCompra.Show
End Sub

Private Sub Command2_Click()
FrmPropostaCompra.Show
End Sub

Private Sub Comodato_Click()
If LcTipoDados <> 3 Then
   CmdSalvar.Enabled = True
End If
End Sub

Private Sub DataCadastro_Change()
If LcTipoDados <> 3 Then
    CmdSalvar.Enabled = True
End If
End Sub

Private Sub EmailFinanceiro_Change()
CmdSalvar.Enabled = True
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        DesabilitaCtr
        DataCadastro.Text = Format(Date, "dd/mm/yyyy")
        LcCap = "   <<Modo Inclusão>>"
   Case Is = 2
        LcCap = "   <<Modo Alteração>>"
      Call AbreBanco(Cliente)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Cliente)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
If Not GLCalculacodigoCliente Then
   txt(0).SetFocus
Else
  txt(0).Enabled = False
End If
Label1.Caption = Label1.Caption & LcCap
CarregaTelemarketing
End Sub
Function CarreGamatriz()
On Error Resume Next
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Cliente
For a = 0 To 31
   MtPesquisa(a).campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).Tipo = ""
Next
For a = 0 To 30
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    If err = 0 Then
       Select Case LcNome
           Case Is = "CGC"
                MtPesquisa(a).campo = "CNPJ"
           Case Is = "inscest"
                MtPesquisa(a).campo = "INSC.EST."
           Case Is = "RAZAOSOC"
                MtPesquisa(a).campo = "RAZÃO SOCIAL"
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

 MtPesquisa(a).Indice = "cgc"
 MtPesquisa(a).Tipo = "T"
 MtPesquisa(a).campo = "CNPJ"
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If LcTipoDados = 3 Then
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
CarregaGrupoEconomico
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FechaBanco

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   If Len(txt(0).Text) > 0 And Len(txt(1).Text) > 0 Then Call CmdSalvar_Click
   
End If
If (LcTipoDados = 2) And LcAlterado Then
   Call CmdSalvar_Click
End If
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False
GlAlteraCodigo = False
FrmPrincipal.SetFocus
End Sub

Private Sub GrupoEconomico_Click()
If LcTipoDados <> 3 Then
   CmdSalvar.Enabled = True
End If
End Sub

Private Sub Mask_Change(Index As Integer)

If LcRegAtual Then Exit Sub
GlCampo9 = Mask(3).Text
GlCampo12 = Mask(2).Text
GlCampo30 = Mask(1).Text
GlCampo31 = Mask(0).Text
GlCampo26 = Mask(4).Text
GlCampo27 = Mask(5).Text
CmdSalvar.Enabled = True
End Sub

Private Sub Mask_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub Mask_LostFocus(Index As Integer)
If Mask(1).Text = "  .   .   /    -  " And Mask(4).Text = "   .   .   -  " Then Exit Sub
If Index = 1 Then
  If Mask(1).Text <> "  .   .   /    -  " Then
   If Not Calc_CNPJ(Mask(1).Text) Then
      MsgBox "CNPJ Inválido...", 64, "Aviso"
      'Mask(1).Text = "  .   .   /    -  "
      Mask(1).SetFocus
   End If
 End If
End If
If Index = 4 Then
 If Mask(4).Text <> "   .   .   -  " Then
  If Not Calc_CPF(Mask(4).Text) Then
      MsgBox "CPF Inválido...", 64, "Aviso"
      Mask(4).SetFocus
   End If
 End If
End If
End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(Cliente) = 1 Then
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
If MovImentacao(enPrimeiro, Cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnRecuperaDataCadastro_Click()
On Error GoTo Erro
Dim LcCap As String
LcCap = Me.Caption
If MsgBox("Confirma a alteração da data de cadastro de todos os CLientes", vbYesNo, "Confirmação") = vbYes Then
        Screen.MousePointer = 11
        Dim rsCliente As Recordset, RsProduto As ADODB.Recordset
        AbreBase
        Set rsCliente = Dbbase.OpenRecordset("Select Codigo from alid001 order by codigo")
        Dim total As Integer
        If Not rsCliente.EOF Then
           rsCliente.MoveLast
           total = rsCliente.RecordCount
           rsCliente.MoveFirst
        End If
        
        Dim a As Integer
        Do Until rsCliente.EOF
           a = a + 1
           Me.Caption = "Acertando cliente codigo " & a & " de " & total
           DoEvents
           Dim RsNota As ADODB.Recordset
           Dim StrSql As String
           StrSql = "Select DTEMIS from alid050 where CLIENTE='" & rsCliente!codigo & "' order by dtemis limit 1"
           Set RsNota = AbreRecordset(StrSql, True)
           Debug.Print StrSql
           If Not RsNota.EOF Then
              StrSql = "Update alid001 set datacadastro=#" & Format(RsNota!DtEmis, "mm/dd/yyyy") & "# where codigo='" & rsCliente!codigo & "'"
              Dbbase.Execute StrSql
           End If
           rsCliente.MoveNext
        Loop
        Me.Caption = LcCap
        Screen.MousePointer = vbDefault
        MsgBox ("Acabei")

Else
   Me.Caption = LcCap
   Screen.MousePointer = vbDefault
   MsgBox "Operação cancelada pelo Usuario"

End If
Exit Sub
Erro:
Screen.MousePointer = vbDefault
Me.Caption = LcCap
MsgBox err.Description
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(Cliente)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub Numero_Change()
CmdSalvar.Enabled = True
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
Dim LcGrID As Long
LcGrID = 0
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Cliente)
TipoContr.Text = RsAtual!TipoContribuinte & ""
Suframa.Text = RsAtual!InscricaoSuframa & ""
InscMunic.Text = RsAtual!InscricaoMunicipal & ""
If IsNumeric(RsAtual!GrupoEconomicoID) Then LcGrID = RsAtual!GrupoEconomicoID
GrupoEconomico.Text = Get_Nome_Gr_Economico(LcGrID)
Numero.Text = RsAtual!Numero & ""
EmailFinanceiro.Text = RsAtual!EmailFinanceiro & ""
If IsDate(RsAtual!DataCadastro) Then
    DataCadastro.Text = Format(RsAtual!DataCadastro, "dd/mm/yyyy")
Else
    DataCadastro.Text = "  /  /    "
End If
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
Mask(3).Text = GlCampo9
txt(10).Text = GlCampo10
txt(11).Text = GlCampo11
Mask(2).Text = GlCampo12
txt(13).Text = GlCampo13
txt(14).Text = GlCampo14
txt(15).Text = GlCampo15
txt(16).Text = GlCampo16
txt(18).Text = GlCampo18
txt(17).Text = GlCampo17
txt(19).Text = GlCampo19
txt(20).Text = GlCampo20
txt(21).Text = GlCampo21
vendedor = GlCampo22
txt(23).Text = GlCampo23
txt(25).Text = GlCampo25
Mask(4).Text = ExibeCpf(GlCampo26)
Mask(5).Text = GlCampo27
Mask(1).Text = ExibeCgc(GlCampo30)
Mask(0).Text = GlCampo31
txt(32).Text = GlCampo32
txt(12).Text = GlCampo35

If RsAtual!Comodato Then
   Comodato.Value = 1
Else
   Comodato.Value = 0
End If
If RsAtual!Bloqueado Then
   Bloqueado.Value = 1
Else
   Bloqueado.Value = 0
End If
txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False

Exit Function
ErroVinculo:
Resume Next
End Function
Private Function Get_Nome_Gr_Economico(GrEconID As Long) As String
  Dim Resposta As String
  Resposta = ""
  For a = 0 To LcTamanhoGr
        If MtGrupoEc(a).codigo = GrEconID Then
           Resposta = MtGrupoEc(a).Nome
           Exit For
        End If
    Next
    Get_Nome_Gr_Economico = Resposta
End Function
Public Function Get_ID_Gr_Economico(Gr_Nome As String) As Long
  Dim Resposta As Long
  Resposta = 0
  For a = 0 To LcTamanhoGr
        If MtGrupoEc(a).Nome = Gr_Nome Then
           Resposta = MtGrupoEc(a).codigo
           Exit For
        End If
    Next
    Get_ID_Gr_Economico = Resposta
End Function
Private Sub TipoContr_Click()
If LcTipoDados <> 3 Then
    If TipoContr.Text = "1 - Contribuinte ICMS" Then
       Mask(2).Enabled = True
    End If
    If TipoContr.Text = "2 - Contribuinte isento de Inscrição" Then
       Mask(2).Enabled = False
    End If
    If TipoContr.Text = "9 - Não Contribuinte" Then
       Mask(2).Enabled = True
    End If
    'Call Alterado
    
    CmdSalvar.Enabled = True
End If
End Sub


Private Sub Txt_Change(Index As Integer)

Call Alterado

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
'If (KeyCode >= 47 And KeyCode <= 111) Or KeyCode = 32 Then
      If Index = 0 Then
         If Not GlAlteraCodigo Then
            GlAlteraCodigo = True
            GlCodigoAnterior = GlCampo0
         End If
      End If
 

'End If
End Sub
Function limpa()
Dim a As Long
On Error Resume Next
For a = 0 To 36
  txt(a).Text = ""
Next
txt(1).SetFocus
CmdSalvar.Enabled = False
vendedor.Text = " "
Mask(1).Text = "  .   .   /    -  "
Mask(2).Text = ""
Mask(4).Text = "   .   .   -  "
Mask(5).Text = ""
Numero.Text = ""
EmailFinanceiro.Text = ""
If LcTipoDados = 1 Then
   DataCadastro.Text = Format(Date, "dd/mm/yyyy")
Else
    DataCadastro.Text = "  /  /    "
End If
End Function
Function BuscaCidade()

Dim RsCidade As Recordset
AbreBase
Set RsCidade = Dbbase.OpenRecordset("select * from alid005")
txt(7).Text = Right("0000" & txt(7).Text, 4)
LcCriterioCi = "cod='" & txt(7).Text & "'"
RsCidade.FindFirst LcCriterioCi
If Not RsCidade.NoMatch Then
   Cidade.Caption = RsCidade!Nome
   LcDesCidade = RsCidade!Nome
Else
  Cidade.Caption = ""
  ' MsgBox "O código da cidade não foi encontrado...,", 64, "Aviso"
End If
RsCidade.Close
Set RsCidade = Nothing



End Function

Private Sub Txt_LostFocus(Index As Integer)
If Index = 17 Or Index = 20 Then
   If Not IsNumeric(txt(Index).Text) Then
      If Len(txt(Index).Text) = 0 Then Exit Sub
      MsgBox "Digite Um Valor Numérico.", vbInformation, "Aviso"
      txt(Index).Text = ""
      txt(Index).SetFocus
      Exit Sub
   End If
End If
If Index = 7 Then BuscaCidade
If Index = 0 Then
   If Not GLCalculacodigoCliente Then txt(0).Text = Trim(txt(0).Text)
End If

If Not GLCalculacodigoCliente Then If VerificaDuplicado(Index) Then txt(Index).SetFocus
   
End Sub

Private Sub Vendedor_Change()
GlCampo22 = vendedor.Text

End Sub

Private Sub Vendedor_Click()
On Error Resume Next
If LcTipoDados <> 3 Then
    For a = 0 To LcTamanho
        If MtVendedor(a).Nome = vendedor.Text Then
           codigo.Text = MtVendedor(a).codigo
           Exit For
        End If
    Next
    GlCampo22 = vendedor.Text
    CmdSalvar.Enabled = True
End If
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If LcTipoDados <> 3 Then
    If KeyCode = 13 Then
       SendKeys "{TAB}"
    Else
      If KeyCode = 116 And Index <> 7 Then
      Else
        Call Teclas(KeyCode)
      End If
    End If
End If
End Sub
