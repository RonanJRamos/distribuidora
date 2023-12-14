VERSION 5.00
Begin VB.Form FrmFornecedor 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Fonecedores"
   ClientHeight    =   6345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Email 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   5055
   End
   Begin VB.TextBox Numero 
      Height          =   285
      Left            =   6240
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   840
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "S/T/N/03/N/END"
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox Contato 
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   4560
      Width           =   8175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   17
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   18
      Tag             =   "S/T/N/30/S/CGC"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   20
      Tag             =   "S/N/N/4/N/COMISSAOREPRESENTANTE"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   9480
      TabIndex        =   56
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   10680
      TabIndex        =   55
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   9480
      TabIndex        =   54
      Top             =   2280
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   10680
      TabIndex        =   53
      Top             =   3000
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   18
      Left            =   9000
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   7560
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   5040
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   3600
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   13
      Tag             =   "S/T/N/13/N/Fax"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   840
      MaxLength       =   20
      TabIndex        =   17
      Tag             =   "S/T/N/30/S/CGC"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   5400
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "S/T/N/12/N/INSCEST"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "S/T/N/11/N/FONE2"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   840
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "S/T/N/10/N/FONE1"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "S/T/N/09/N/CEP"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   840
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "S/T/N/08/N/ESTADO"
      Top             =   3240
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
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   840
      MaxLength       =   20
      TabIndex        =   6
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   5055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "S/T/N/02/N/FANTASIA"
      Top             =   1440
      Width           =   6615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   840
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "S/T/S/01/N/RAZAOSOC"
      Top             =   1080
      Width           =   6615
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   9480
      TabIndex        =   29
      Top             =   3000
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   10680
      TabIndex        =   28
      Top             =   2280
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   10680
      TabIndex        =   27
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9480
      TabIndex        =   25
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   9480
      TabIndex        =   26
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
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   120
      TabIndex        =   64
      Top             =   3720
      Width           =   600
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
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
      Left            =   6000
      TabIndex        =   63
      Top             =   1815
      Width           =   210
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl."
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
      TabIndex        =   62
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label17 
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
      Left            =   3000
      TabIndex        =   60
      Top             =   5040
      Width           =   390
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   3960
      TabIndex        =   59
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite F5 Para Escolher a Cidade"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5160
      TabIndex        =   57
      Top             =   3240
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
      Left            =   2280
      TabIndex        =   52
      Top             =   2880
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
      TabIndex        =   51
      Top             =   2880
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
      Left            =   6480
      TabIndex        =   50
      Top             =   4080
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
      Left            =   3000
      TabIndex        =   48
      Top             =   4080
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
      TabIndex        =   36
      Top             =   4080
      Width           =   480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   -240
      X2              =   9240
      Y1              =   5880
      Y2              =   5880
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
      TabIndex        =   49
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   8640
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   -120
      X2              =   9240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9360
      Y1              =   3600
      Y2              =   3600
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
      TabIndex        =   42
      Top             =   3240
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
      TabIndex        =   44
      Top             =   3240
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
      TabIndex        =   43
      Top             =   5040
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
      Left            =   6000
      TabIndex        =   30
      Top             =   5040
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
      TabIndex        =   40
      Top             =   1800
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
      TabIndex        =   46
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
      TabIndex        =   41
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
      TabIndex        =   45
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
      TabIndex        =   47
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
      TabIndex        =   39
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
      TabIndex        =   38
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Fornecedores"
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
      TabIndex        =   37
      Top             =   0
      Width           =   11835
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
      TabIndex        =   61
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão Representante"
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
      Left            =   120
      TabIndex        =   58
      Top             =   5640
      Width           =   2775
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
Attribute VB_Name = "FrmFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCarregado, a As Integer
Private LcDesCidade As String
Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    Txt(a).Enabled = False
Next
End Function

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, fornecedor) Then VinculaDados
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
GlTab = "Alid002"
GlSq = "Select * from alid002 where codigo='" & Txt(0).Text & "'"

If Exclui(fornecedor) = 1 Then
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
Call Teclas(KeyCode)
End Sub

Private Sub CmdPesquisar_Click()
LcIndice = "RAZAOSOC"
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
If MovImentacao(enPrimeiro, fornecedor) Then VinculaDados
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
Dim Inscricao   As String
Dim Estado      As String
Dim CNPJ        As String

CNPJ = Txt(30).Text
CNPJ = Replace(CNPJ, ",", "")
CNPJ = Replace(CNPJ, ".", "")
CNPJ = Replace(CNPJ, "-", "")
CNPJ = Replace(CNPJ, "/", "")
CNPJ = Replace(CNPJ, "\", "")
CNPJ = Replace(CNPJ, " ", "")

If Len(CNPJ) = 0 Then
    CNPJ = Txt(17).Text
    CNPJ = Replace(CNPJ, ",", "")
    CNPJ = Replace(CNPJ, ".", "")
    CNPJ = Replace(CNPJ, "-", "")
    CNPJ = Replace(CNPJ, "/", "")
    CNPJ = Replace(CNPJ, "\", "")
    CNPJ = Replace(CNPJ, " ", "")
End If

If Len(CNPJ) = 0 Then
   MsgBox "É nescessario cadastrar o CNPJ / CPF do fornecedor.", 64, "Aviso"
   Txt(30).SetFocus
   SendKeys "+{home}+{end}"
   Exit Sub
End If
Inscricao = Txt(12).Text
Inscricao = Replace(Inscricao, ",", "")
Inscricao = Replace(Inscricao, ".", "")
Inscricao = Replace(Inscricao, "-", "")
Inscricao = Replace(Inscricao, "/", "")
Inscricao = Replace(Inscricao, "\", "")
Inscricao = Replace(Inscricao, " ", "")
If Len(Inscricao) = 0 Then
    Txt(12).Text = "ISENTO"
    Inscricao = "ISENTO"
End If
Estado = Txt(8).Text
Estado = Trim(Estado)
If Len(Estado) = 0 Then
   MsgBox "É nescessario cadastrar o estado do fornecedor.", 64, "Aviso"
   Txt(8).SetFocus
   SendKeys "+{home}+{end}"
   Exit Sub
End If
If Len(CNPJ) > 11 Then
    If Not Calc_CNPJ(CNPJ) Then
       MsgBox "O CNPJ do fornecedor é invalido.", 64, "Aviso"
       Exit Sub
    End If
Else
    If Not Calc_CPF(CNPJ) Then
       MsgBox "O CPF do fornecedor é invalido.", 64, "Aviso"
       Exit Sub
    End If

End If
If Consiste(Inscricao, Estado) <> 0 Then
   MsgBox "A Inscrição Estadual do fornecedor é invalida.", 64, "Aviso"
   'ValidaEntradaSintegra = False
   Exit Sub
End If
Call SalvaRegistro(fornecedor)
VinculaDados
LcRegAtual = False
'NovoReg
If LcTipoDados = 1 Then limpa
Txt(1).SetFocus
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
If MovImentacao(enSeguinte, fornecedor) Then VinculaDados
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
If MovImentacao(enultimo, fornecedor) Then VinculaDados
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

Private Sub Contato_Change()
GlAlteraCodigo = True
GlCodigoAnterior = GlCampo0
Call Alterado

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
      Call AbreBanco(fornecedor)
      VinculaDados
   Case Is = 3
      LcCap = "   <<Modo Consultar>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(fornecedor)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
If Not GLCalculacodigoFornecedor Then
   Txt(0).SetFocus
Else
  Txt(0).Enabled = False
End If

End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.fornecedor
For a = 0 To 30
   MtPesquisa(a).campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).Tipo = ""
Next

For a = 0 To 30
    LcNome = Mid$(Txt(a).Tag, 12)
    LcTipo = Mid$(Txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    If err = 0 Then
       Select Case LcNome
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
 
End Function

Private Sub Form_Load()
On Error Resume Next
'Me.Height = 5745
'Me.Width = 12000
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Txt(4).Visible = GlRepresentante
Label14.Visible = GlRepresentante
Label15.Visible = GlRepresentante
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FechaBanco

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (fornecedor)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (fornecedor)
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
If MovImentacao(enAnterior, fornecedor) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(fornecedor) = 1 Then
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
If MovImentacao(enPrimeiro, fornecedor) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(fornecedor)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, fornecedor) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, fornecedor) Then VinculaDados
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
Dim CNPJ As String
Dim CPF As String

If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(fornecedor)


Txt(0).Text = GlCampo0
Txt(1).Text = GlCampo1
Txt(2).Text = GlCampo2
Txt(3).Text = GlCampo3
Txt(4).Text = GlCampo4
'=== Exibe o nome da cidade
Txt(5).Text = GlCampo5
Txt(6).Text = GlCampo6
Txt(7).Text = GlCampo7
BuscaCidade
Txt(8).Text = GlCampo8
Txt(9).Text = GlCampo9
Txt(10).Text = GlCampo10
Txt(11).Text = GlCampo11
Txt(12).Text = GlCampo12
Txt(13).Text = GlCampo13
Txt(14).Text = GlCampo14
Txt(15).Text = GlCampo15
Txt(16).Text = GlCampo16
Txt(18).Text = GlCampo18
Txt(19).Text = GlCampo19
Txt(25).Text = GlCampo25
If Len(GlCampo30) >= 14 Then
   CNPJ = Mid(GlCampo30, 1, 2) & "." & Mid(GlCampo30, 3, 3) & "." & Mid(GlCampo30, 6, 3) & "/" & Mid(GlCampo30, 9, 4) & "-" & Mid(GlCampo30, 13, 2)
Else
   CNPJ = GlCampo30
End If
If Len(RsAtual!CPF) >= 11 Then
   CPF = Mid(RsAtual!CPF, 1, 3) & "." & Mid(RsAtual!CPF, 4, 3) & "." & Mid(RsAtual!CPF, 7, 3) & "-" & Mid(RsAtual!CPF, 10, 2)
Else
  CPF = RsAtual!CPF & ""
End If
Txt(30).Text = CNPJ
Txt(17).Text = CPF
Contato.Text = RsAtual!Contato & ""
Email.Text = RsAtual!Email & ""
Numero.Text = RsAtual!Numero & ""
Txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
If Index = 0 Then
   GlAlteraCodigo = True
   GlCodigoAnterior = GlCampo0
End If
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
  Txt(a).Text = ""
Next
Txt(1).SetFocus
CmdSalvar.Enabled = False
End Function
Function BuscaCidade()

Dim RsCidade As Recordset
AbreBase
Set RsCidade = Dbbase.OpenRecordset("select * from alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Txt(7).Text = Right("0000" & Txt(7).Text, 4)
LcCriterioCi = "cod='" & Txt(7).Text & "'"
RsCidade.FindFirst LcCriterioCi
If Not RsCidade.NoMatch Then
   Cidade.Caption = RsCidade!Nome
   LcDesCidade = RsCidade!Nome
Else
   Cidade.Caption = ""
   'MsgBox "O código da cidade não foi encontrado...,", 64, "Aviso"
End If
RsCidade.Close
Set RsCidade = Nothing



End Function

Private Sub Txt_LostFocus(Index As Integer)
 
If Index = 4 Then
  If Not Calc_CPF(Txt(17).Text) Then
      MsgBox "CPF Inválido...", 64, "Aviso"
      Txt(17).SetFocus
   End If
End If
If Index = 7 Then BuscaCidade
If Index = 0 Then
   If Not GLCalculacodigoFornecedor Then Txt(0).Text = Trim(Txt(0).Text)
End If
If Not GLCalculacodigoFornecedor Then If VerificaDuplicado(Index) Then Txt(Index).SetFocus

End Sub
