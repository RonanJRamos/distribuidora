VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmProposta 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saída de Estoque"
   ClientHeight    =   9735
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11850
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox InformacoesComplementares 
      Height          =   855
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Top             =   7440
      Width           =   6375
   End
   Begin VB.TextBox VendedorImprimir 
      Height          =   285
      Left            =   240
      TabIndex        =   30
      Top             =   7440
      Width           =   4935
   End
   Begin VB.TextBox ValidadeCotacao 
      Height          =   285
      Left            =   2520
      TabIndex        =   33
      Text            =   "3 DIAS"
      Top             =   8640
      Width           =   2655
   End
   Begin VB.TextBox PrazoEntrega 
      Height          =   285
      Left            =   240
      TabIndex        =   32
      Text            =   "Imediato"
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox CondPag 
      Height          =   285
      Left            =   240
      TabIndex        =   31
      Top             =   8040
      Width           =   4935
   End
   Begin VB.CommandButton CmdDuplicarPedido 
      Caption         =   "Duplicar Pedido "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   110
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Imprimir c/&SubItem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   108
      Top             =   1935
      Width           =   1575
   End
   Begin VB.TextBox Lucratividade 
      Height          =   285
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox comodato 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Tem Comodato"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   105
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox usuario 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Li&star Pe. Pen.      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   102
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Liberar Pedido      "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   101
      Top             =   2295
      Width           =   1575
   End
   Begin VB.CheckBox Pendente 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Pendente"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   100
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox tipoBlQ 
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   10440
      TabIndex        =   99
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox tipoblP 
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   10800
      TabIndex        =   92
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox HoraLiberacao 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8160
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox MaquinaLiberacao 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8160
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox JaBloqueado 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8160
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox DataLiberacao 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8280
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox bloqueado 
      BackColor       =   &H00FF8080&
      Height          =   375
      Left            =   8160
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox obs 
      Height          =   285
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   86
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "E&xcluir             F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   85
      Top             =   825
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Imprimir           F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   84
      Top             =   1575
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Fechar Pedido F3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5880
      TabIndex        =   83
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Pesquisar        F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   82
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Gravar             F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   81
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Observação    F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Validade 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Ordem 
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Condicoes 
      Height          =   285
      Left            =   11160
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   7440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox faturado 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Faturado"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   75
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox LiberaFaturamento 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Liberado Para Faturamento"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2040
      Width           =   2655
   End
   Begin MSMask.MaskEdBox Previsao 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   17
      Left            =   240
      TabIndex        =   38
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   6
      Left            =   8040
      TabIndex        =   68
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox santamaria1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   66
      Text            =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox santamaria 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   65
      Text            =   "0"
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox california 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   64
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox almox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   63
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox CFOP 
      Height          =   315
      ItemData        =   "FrmProposta.frx":0000
      Left            =   8520
      List            =   "FrmProposta.frx":0016
      TabIndex        =   62
      Text            =   "512"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox utilizado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   60
      Text            =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox limite 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   59
      Text            =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   0
      Left            =   7200
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox tam 
      Height          =   375
      Left            =   7440
      TabIndex        =   57
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Comissao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   56
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   3480
      TabIndex        =   53
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   15
      Left            =   1800
      TabIndex        =   52
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   14
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   8640
      Width           =   6375
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox cst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   46
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox icms 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   45
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox minimo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   44
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   9720
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   16
      Top             =   3120
      Width           =   810
   End
   Begin VB.ComboBox Unidade 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3090
      Width           =   1095
   End
   Begin VB.ComboBox Natureza 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmProposta.frx":003C
      Left            =   5880
      List            =   "FrmProposta.frx":003E
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Custo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   39
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   12
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   37
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   36
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   3015
      Left            =   120
      TabIndex        =   29
      Top             =   4080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   16
      BackColor       =   -2147483624
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Item     F4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   27
      Top             =   1185
      Width           =   1575
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Cancelar        F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   8655
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8040
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   17
      Top             =   3120
      Width           =   810
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   10320
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Só sai na NFe)"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   23
      Left            =   8400
      TabIndex        =   117
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Complementares para NFe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   22
      Left            =   5400
      TabIndex        =   116
      Top             =   7200
      Width           =   2790
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Vendedor a Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   240
      TabIndex        =   115
      Top             =   7200
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Validade Cotação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   2520
      TabIndex        =   114
      Top             =   8400
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   113
      Top             =   8400
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Condições de Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   112
      Top             =   7800
      Width           =   1830
   End
   Begin VB.Label LabelBloqueado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Bloqueado"
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
      Left            =   3000
      TabIndex        =   111
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F11 -> Exibe as obs do cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3120
      TabIndex        =   109
      Top             =   2520
      Width           =   2490
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lrucro %"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8880
      TabIndex        =   106
      Top             =   2160
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Validade Proposta"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   104
      Top             =   2400
      Width           =   1560
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Est. e Pr. Insf."
      Height          =   255
      Left            =   3240
      TabIndex        =   98
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   97
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Baixo"
      Height          =   255
      Left            =   1680
      TabIndex        =   96
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      Height          =   255
      Left            =   2760
      TabIndex        =   95
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Est. Insufic."
      Height          =   255
      Left            =   120
      TabIndex        =   94
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   1200
      TabIndex        =   93
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   7440
      TabIndex        =   79
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Validade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   7560
      TabIndex        =   78
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordem Compra"
      Height          =   255
      Left            =   8160
      TabIndex        =   77
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cond. de Pag. / Prazo"
      Height          =   255
      Left            =   6000
      TabIndex        =   76
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido de Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Width           =   2610
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Prevista"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   73
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Ver Últimas Compras do Cliente Pressione F12 "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   72
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto F11"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7920
      TabIndex        =   71
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   70
      Top             =   9000
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      Height          =   255
      Left            =   7200
      TabIndex        =   69
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS"
      Height          =   255
      Left            =   5040
      TabIndex        =   67
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   61
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Nota"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   3480
      TabIndex        =   55
      Top             =   9000
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produtos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   1800
      TabIndex        =   54
      Top             =   9000
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   5400
      TabIndex        =   51
      Top             =   8400
      Width           =   330
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Calc. ICMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   8160
      TabIndex        =   50
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   9600
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Produto Digite Seu Código,  Nome ou pressione F5 Para Detalhar Produto Pressione F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   3480
      Width           =   7605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Cliente pressione F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   41
      Top             =   1320
      Width           =   2850
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Natureza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   40
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   9960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   0
      X2              =   9960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   9960
      X2              =   9960
      Y1              =   0
      Y2              =   4440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data da Emissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V. Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   58
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V. Unit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   23
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unid. / Com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   22
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
End
Attribute VB_Name = "FrmProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoUnid
      codigo As String
      Descricao As String
      Simbolo As String
      quantidade As Long
End Type
Private LcItem As Long, LcTam, LcQUn, LcQuantiImpressao, LcQuantiImpressaoBoleto As Long
Private FnunNota, FnunBoleto
Private LcNota, LcBoleto, LcEspC As String
Private LcFocus, LcCalculado, LcSalto  As Integer
Private LcPrecoVelho As Currency
Private ComNormal, ComAlterado, LcQuantNesc, LcQtSta1, LcQtSta, LcQtCal As Long
Private LcLinha As String
Private RsOpcoes As Recordset, RsClientes As Recordset
Private RsCidade As Recordset
Private LcValor1, LcValor2, LcValor3, LcUltimo As Currency
Private LcAlteradoCliente, LcAlteradoProduto, LcAlteradoFuncionario As Integer
Private LcMat() As DadosEntrada, LcLimpa As Integer
Private Liberado, LcBuscaCliente, LcBuscaNota As Integer
Private MtUnidade() As TipoUnid, MtImpressao(), MtBoleto() As String
Private LcImpressoes, LcProximo, LcLimpaValor, LcPesquisaCli As Integer
Private LcSaldoCaixa, LcSaldoUnit As Double
Private TotalCaixa, TotalUnitario As Double
Private LcFechaitem, a As Integer
Private LcCorPadrao     As Variant
Private Rel As New CrysPropostaVenda
Private RelLidis As New CrysPropostaVendaLidis


Private Sub CFOP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub CmdDuplicarPedido_Click()
FrmDuplicaPedido.Show , Me
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
FrmExcluiItem.Show
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
ReDim LcMat(0)
LcTam = 0
LcItem = 0
End Sub



Function CalculaValores()

Dim LcTotal As Currency, LcQuant As Double, LcUnit As Currency
On Error Resume Next

If LcCalculado Then Exit Function
LcCalculado = True
'=== Converte os Valores
If Len(Trim(Txt(3).Text)) = 0 Then Exit Function
valor(1).Text = 0
If Len(Trim(Txt(3).Text)) > 0 Then
   LcQuant = CDbl(Txt(3).Text)
Else
   LcQuant = 1
End If
If CCur(Txt(3).Text) > 0 Then
   LcQuant = CDbl(Txt(3).Text)
Else
   LcQuant = 1
End If
'MsgBox Txt(3).Text

LcUnit = CDbl(valor(0).Text)

LcTotal = CDbl(LcQuant) * CDbl(LcUnit)
valor(1).Text = LcTotal

LcCalculado = False
End Function
Function GeraGrid()
On Error Resume Next
Item.ColAlignment(0) = 7
Item.ColAlignment(1) = 3
Item.ColAlignment(2) = 1
Item.ColAlignment(3) = 3
Item.ColAlignment(4) = 1
Item.ColAlignment(5) = 3
Item.ColAlignment(6) = 3
Item.ColAlignment(7) = 3
Item.ColAlignment(8) = 3
Item.ColAlignment(9) = 3
Item.ColAlignment(10) = 3
Item.ColAlignment(11) = 3
Item.ColAlignment(12) = 3
Item.ColAlignment(13) = 3
Item.ColAlignment(14) = 3
Item.ColAlignment(15) = 3

Item.ColWidth(0) = 500
Item.ColWidth(1) = 700
Item.ColWidth(2) = 4600
Item.ColWidth(3) = 500
Item.ColWidth(4) = 1000
Item.ColWidth(5) = 900
Item.ColWidth(6) = 1200
Item.ColWidth(7) = 1200
Item.ColWidth(8) = 600
Item.ColWidth(9) = 0
Item.ColWidth(10) = 0
Item.ColWidth(11) = 0
Item.ColWidth(12) = 0
Item.ColWidth(13) = 0
Item.ColWidth(14) = 0
Item.ColWidth(15) = 0

Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Código"
Item.TextMatrix(0, 2) = "Descrição"
Item.TextMatrix(0, 3) = "CST"
Item.TextMatrix(0, 4) = "Unidade"
Item.TextMatrix(0, 5) = "Quant"
Item.TextMatrix(0, 6) = "Unitário"
Item.TextMatrix(0, 7) = "Total"
Item.TextMatrix(0, 8) = "ICMS"
Item.TextMatrix(0, 10) = "blo"
Item.TextMatrix(0, 11) = "jaBlo"
Item.TextMatrix(0, 12) = "tipoblo"
Item.TextMatrix(0, 13) = "MaquinaLi"
Item.TextMatrix(0, 14) = "DataLib"
Item.TextMatrix(0, 15) = "HoraLib"

LcTamanhoGrid = 1
End Function

Public Function MondaGridAutomatico(LcCodigo As Long, quantidade As Long, ValorUnirario As Currency, Com As Long, Lc_Unidade As String, LcItem As Long)
Dim Dados As DadosEntrada
'MsgBox LcTipo
'On Error GoTo errBuscaFor
Dim RsProduto As ADODB.Recordset
Dim Rs        As ADODB.Recordset
Dim LcValorDigitado
'If Not LcAlteradoProduto Then Exit Function
AbreBase
Set RsProduto = AbreRecordset("select * from Produtos where codigo=" & LcCodigo, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
GlCriterioSql = ""
LcCalculado = True
         If Not RsProduto.EOF Then
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            Dados.CodPro = RsProduto!codigo
            Dados.produto = RsProduto!Nome
            Dados.Und = LcUnidade
            Dados.Com = RsProduto!QtdMedida
            Dados.VUnit = RsProduto!Preco
            Dados.cst = RsProduto!cst
            LcPrecoVelho = RsProduto!Preco
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = RsProduto!Preco / RsProduto!QtdMedida Else PrecoVendaNormal = 0
            ComNormal = RsProduto!QtdMedida
            'minimo.Text = RsProduto!MinimoVenda & ""
            If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MinimoVenda / RsProduto!QtdMedida Else PrecoMimimodeVendaAlterado = 0
            If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
               If Val(Dados.cst) = 60 Or Val(Dados.cst) = 160 Or Val(Dados.cst) = 260 Then
                  Dados.icms = "00"
               Else
                  Dados.icms = "18"
                End If
            Else
               icms = RsProduto!icms
            End If
            Dados.Venda1 = RsProduto!CustoTotal
            LcAchou = True
           ' SendKeys "{TAB}"
         Else
           ' Txt(2).Text = ""
         End If


RsProduto.Close
Set RsProduto = Nothing


LcItem = Item.Rows
ReDim Preserve LcMat(LcTam)
LcMat(LcTam).Item = Right("00" & LcItem + 1, 2)
LcMat(LcTam).CodPro = Dados.CodPro
LcMat(LcTam).produto = Dados.produto
LcMat(LcTam).Qut = quantidade
If Len(Lc_Unidade) > 0 Then
   LcMat(LcTam).Und = Lc_Unidade
Else
   LcMat(LcTam).Und = LcUnidade
End If

If Com > 0 Then
   LcMat(LcTam).Com = Com
Else
   LcMat(LcTam).Com = Dados.Com
End If

If ValorUnirario > 0 Then
   LcMat(LcTam).VUnit = ValorUnirario
Else
   LcMat(LcTam).VUnit = Dados.VUnit
End If
LcMat(LcTam).Vtotal = LcMat(LcTam).Qut * LcMat(LcTam).VUnit 'CCur(valor(0).Text) * CCur(Txt(3).Text)
LcMat(LcTam).Venda1 = Dados.Venda1 ' CCur(Custo.Text)
LcMat(LcTam).cst = Dados.cst ' cst.Text
LcMat(LcTam).icms = Dados.icms ' icms.Text
LcMat(LcTam).almox = Dados.almox ' almox.Text
LcMat(LcTam).california = Dados.california ' CLng(california.Text)
LcMat(LcTam).santamaria = Dados.santamaria ' CLng(santamaria.Text)
LcMat(LcTam).santamaria1 = Dados.santamaria1 ' CLng(santamaria1.Text)
LcMat(LcTam).usuario = ""
'If (CInt(tipoblP.Text) > 0) Or (CInt(tipoBlQ.Text) > 0) Then
   LcMat(LcTam).bloqueado = Dados.bloqueado
'Else
 '  LcMat(LcTam).bloqueado = False
'End If

'If Not LcMat(LcTam).jaEsteveBloqueado Then
   LcMat(LcTam).jaEsteveBloqueado = Dados.jaEsteveBloqueado
'End If
LcMat(LcTam).tipoliberacao = Dados.tipoliberacao 'CInt(tipoBlQ.Text) + CInt(tipoblP.Text)
Command9.Enabled = LcMat(LcTam).bloqueado

LcTam = LcTam + 1
EscreveGrid
LcLimpaValor = True

End Function
Function montagrid()
Dim LcAchou, a As Integer
On Error Resume Next
If Not LcFechaitem Then Exit Function

'==== Verifica se Foi digitados todos os campos
If Len(Trim(Txt(1).Text)) = 0 Then
   MsgBox "Necessário Informar o Produto.", 48, "Aviso"
   Txt(1).SetFocus
   Exit Function
End If
If Len(Trim(Txt(3).Text)) = 0 Or (Txt(3).Text = "0") Then
   MsgBox "Necessário Informar a Quantidade de Saída.", 48, "Aviso"
   Txt(3).SetFocus
   Exit Function
End If
If Len(Trim(valor(0).Text)) = 0 Or valor(0).Text = "0" Then
   MsgBox "Necessário Informar o Valor Unitario do Item.", 48, "Aviso"
   valor(0).SetFocus
   Exit Function
End If

'VerificaEstoque (CLng(Txt(4).Text) * ccur(Txt(3).text))

LcItem = LcItem + 1
ReDim Preserve LcMat(LcTam)
LcMat(LcTam).Item = Right("00" & LcItem, 2)
LcMat(LcTam).CodPro = Txt(1).Text
LcMat(LcTam).produto = Txt(2).Text
LcMat(LcTam).Qut = CCur(Txt(3).Text)
LcMat(LcTam).Und = Unidade.Text
LcMat(LcTam).Com = Txt(4).Text
LcMat(LcTam).VUnit = CCur(valor(0).Text)
LcMat(LcTam).Vtotal = CCur(valor(0).Text) * CCur(Txt(3).Text)
LcMat(LcTam).Venda1 = CCur(Custo.Text)
LcMat(LcTam).cst = cst.Text
LcMat(LcTam).icms = icms.Text
LcMat(LcTam).almox = almox.Text
LcMat(LcTam).california = CLng(california.Text)
LcMat(LcTam).santamaria = CLng(santamaria.Text)
LcMat(LcTam).santamaria1 = CLng(santamaria1.Text)
LcMat(LcTam).usuario = ""
If (CInt(tipoblP.Text) > 0) Or (CInt(tipoBlQ.Text) > 0) Then
   LcMat(LcTam).bloqueado = True
Else
   LcMat(LcTam).bloqueado = False
End If

If Not LcMat(LcTam).jaEsteveBloqueado Then
   LcMat(LcTam).jaEsteveBloqueado = LcMat(LcTam).bloqueado
End If
LcMat(LcTam).tipoliberacao = CInt(tipoBlQ.Text) + CInt(tipoblP.Text)
Command9.Enabled = LcMat(LcTam).bloqueado

LcTam = LcTam + 1
EscreveGrid
LcLimpaValor = True
For a = 1 To 6
   If a <> 5 Then
      Txt(a).Text = ""
   End If
   valor(a).Text = ""
Next
'===> seta  a Cor da Celula

LcLimpaValor = False
california.Text = ""
santamaria.Text = ""
santamaria1.Text = ""
Txt(3).Text = " "
valor(0).Text = " "
valor(0).Text = " "
Custo.Text = "0"
icms.Text = "0"
cst.Text = "0"
minimo.Text = "0"
almox.Text = ""
tipoblP.Text = 0
tipoBlQ.Text = 0

Txt(2).SetFocus
End Function
Function limpanota()
On Error Resume Next
Dim a As Integer
Liberado = False
LcTam = 0
LcItem = 0
ReDim LcMat(0)
Item.Rows = 1
For a = 0 To 15
   Txt(a).Text = ""
   valor(a).Text = ""
Next
Txt(17).Text = ""
Txt(16).Text = ""
limite.Text = 0
utilizado.Text = 0
Previsao.Text = Format(Date, "dd/mm/yy")
'CalculaNumeroNota
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
almox.Text = ""
faturado = 0
Liberado = 0
bloqueado.Text = ""
JaBloqueado.Text = ""
MaquinaLiberacao.Text = ""
HoraLiberacao.Text = ""
DataLiberacao.Text = ""
Command9.Enabled = False
Txt(6).Text = "EM LANCAMENTO"
Unload DadosTransp
Txt(0).Locked = True
Lucratividade.Text = ""

Txt(12).SetFocus
End Function
Function EscreveGrid(Optional GridAlterado As Boolean)
On Error Resume Next
Dim b, a As Integer
Dim primeiro As Integer
Dim Segundo As Integer
Dim Terceiro As Integer
Dim LcBloqueio As Boolean

primeiro = &HC0C000
Segundo = &HC000&
Terceiro = &H8080FF

b = 1
Item.Rows = 1
LcBloqueio = False

For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item.Rows = b + 1
        LcMat(a).Item = Right("00" & a + 1, 2)
       Item.TextMatrix(b, 0) = Right("00" & a + 1, 2)
       Item.TextMatrix(b, 1) = LcMat(a).CodPro
       Item.TextMatrix(b, 2) = LcMat(a).produto
       Item.TextMatrix(b, 3) = LcMat(a).cst
       Item.TextMatrix(b, 4) = LcMat(a).Und & " C/" & LcMat(a).Com
       Item.TextMatrix(b, 5) = LcMat(a).Qut
       Item.TextMatrix(b, 6) = Format(LcMat(a).VUnit, "Currency")
       Item.TextMatrix(b, 7) = Format(LcMat(a).Vtotal, "Currency")
       Item.TextMatrix(b, 8) = LcMat(a).icms
       Item.TextMatrix(b, 10) = LcMat(a).bloqueado
       Item.TextMatrix(b, 11) = LcMat(a).jaEsteveBloqueado
       Item.TextMatrix(b, 12) = LcMat(a).tipoliberacao
       Item.TextMatrix(b, 13) = LcMat(a).MaquinaLiberacao
       Item.TextMatrix(b, 14) = LcMat(a).DataLiberacao
       Item.TextMatrix(b, 15) = LcMat(a).HoraLiberacao
       If LcCorPadrao = 0 Then
          Item.Row = 1
          Item.Col = 1
          LcCorPadrao = Item.CellBackColor
       End If
       If GridAlterado Then
            BloqueioQuant = 0
            BloqueioValor = 0
            BloqueioQuant = VerificaDisponivelGrid(LcMat(a).CodPro, LcMat(a).Qut, CDbl(LcMat(a).Com))
            BloqueioValor = ConferePrecoGrid(LcMat(a).CodPro, CDbl(LcMat(a).VUnit), CDbl(LcMat(a).Com))

            If BloqueioQuant > 0 Or BloqueioValor > 0 Then
                LcMat(a).bloqueado = True
                LcMat(a).tipoliberacao = BloqueioQuant + BloqueioValor
            Else
                LcMat(a).bloqueado = False
            End If
       End If
       If LcMat(a).bloqueado Then
          Select Case LcMat(a).tipoliberacao
              Case Is = 1
                  Cor = &HC0C000
              Case Is = 2
                  Cor = &HC000&
              Case Is = 3
                  Cor = &H8080FF
              End Select
             Item.Row = b
            For x = 0 To 15
                Item.Col = x
                Item.CellBackColor = Cor
            Next
            
       End If
       If Not LcBloqueio Then
          If LcMat(a).bloqueado Then LcBloqueio = True
       End If
       b = b + 1
    End If
Next
If LcBloqueio Then
   Pendente.Value = 1
   bloqueado.Text = True
   JaBloqueado.Text = True
Else
    Pendente.Value = 0
   bloqueado.Text = False
   JaBloqueado.Text = False
End If
CalculaIcms
VerificaLucratividade
If faturado = 0 Then
   Command3.Enabled = True
   CmdSalvar.Enabled = True
   CmdExcluir.Enabled = True
End If

End Function
Function CalculaIcms()
On Error Resume Next
Dim LcBaseCalculo, LcIcms, LcPRodutos, LcNota As Currency
Dim LcItem As String, LcComp As String
Dim LcQuantItemSubs, a As Integer
'LcItem = 0
LcQuantItemSubs = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      If LcMat(a).icms > 0 Then
         LcBaseCalculo = LcBaseCalculo + LcMat(a).Vtotal
         LcIcms = LcIcms + ((LcMat(a).icms / 100) * LcMat(a).Vtotal)
      Else
         LcQuantItemSubs = LcQuantItemSubs + 1
         If LcQuantItemSubs > 1 Then
            LcItem = LcItem & ", "
         End If
         LcItem = LcItem & Right("00" & CStr(LcMat(a).Item), 2)
      End If
      LcPRodutos = LcPRodutos + LcMat(a).Vtotal
      LcNota = LcNota + LcMat(a).Vtotal
   End If
Next
If LcQuantItemSubs > 0 Then
   If LcQuantItemSubs > 1 Then
      LcComp = "Itens " & LcItem & " ICMS cobrado por subst. Tributária."
   Else
      LcComp = "Item " & LcItem & " ICMS cobrado por subst. Tributária."
   End If
   If Natureza.Text <> "TRANSFERENCIA" Then Txt(14).Text = LcComp
End If
LcPercDivicao = LcIcms / LcBaseCalculo

If Len(Trim(Txt(17).Text)) = 0 Or Txt(17).Text = "0" Then
   Txt(13).Text = Format(LcBaseCalculo, "Currency")
   Txt(11).Text = Format(LcIcms, "Currency")
   Txt(15).Text = Format(LcPRodutos, "Currency")
   Txt(16).Text = Format(LcNota, "Currency")
Else
   If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
   If Not IsEmpty(LcBaseCalculo) Then LcBaseCalculo = LcBaseCalculo - CDbl(Txt(17).Text)
   Txt(13).Text = Format(LcBaseCalculo, "Currency")
   Txt(11).Text = Format(LcBaseCalculo * LcPercDivicao, "Currency")
   Txt(15).Text = Format(LcPRodutos, "Currency")
   Txt(16).Text = Format((LcNota - CCur(Txt(17).Text)), "Currency")
End If
End Function
Function RemontaIndice()
On Error Resume Next
Dim a As Integer
LcItem = 0

For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      LcItem = LcItem + 1
      LcMat(a).Item = Right("00" & LcItem, 2)
   End If
Next


End Function
Function CarregaComboNatureza()
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim LcNome As String
Dim primeiro As Boolean
primeiro = True
StrSql = "Select * from naturezaoperacao where origem='Saida' and NaoCarregaPedido=0 order by codigo"
Set Rs = AbreRecordset(StrSql, True)
Do Until Rs.EOF
  Natureza.AddItem Rs!Nome & ""
  If primeiro Then
     LcNome = Rs!Nome & ""
     primeiro = False
  End If
  Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
Natureza.Text = LcNome
BuscaCFOPPadrao
End Function
Function CarregaCboUnidade()
On Error Resume Next
LcQUn = 0
Dim LcAchou As Integer
Dim RsUnidade As Recordset
Dim LcPrimeiro As String
AbreBase
Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 order By SIMBOLO", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsUnidade.EOF
   ReDim Preserve MtUnidade(LcQUn)
   MtUnidade(LcQUn).codigo = RsUnidade!cod
   MtUnidade(LcQUn).Descricao = RsUnidade!Nome
   MtUnidade(LcQUn).Simbolo = RsUnidade!Simbolo
   MtUnidade(LcQUn).quantidade = RsUnidade!quantidade
   Unidade.AddItem RsUnidade!Simbolo
   RsUnidade.MoveNext
   LcQUn = LcQUn + 1
Loop
If LcQUn > 0 Then LcQUn = LcQUn - 1
RsUnidade.Close
Dbbase.Close
Set RsUnidade = Nothing
Set Dbbase = Nothing


End Function
Function calculaunitario()
On Error Resume Next

valor(0).Text = CDbl(Txt(4).Text) * PrecoVendaNormal

minimo.Text = CLng(Txt(4).Text) * CCur(AcertaNumero(CStr(PrecoMimimodeVendaAlterado), GlDecimais))
End Function
Function BuscaProduto(LcTipo As Integer)
On Error Resume Next
'MsgBox LcTipo
On Error GoTo errBuscaFor
Dim RsProduto As ADODB.Recordset
Dim Rs        As ADODB.Recordset
Dim LcValorDigitado
Dim LcCodigo As String
If Not LcAlteradoProduto Then Exit Function
'AbreBase
Set RsProduto = AbreRecordset("select * from Produtos where Desativado=0", True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
GlCriterioSql = ""
LcCalculado = True
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "Codigo=" & Txt(1).Text
         RsProduto.MoveFirst
         RsProduto.Find LcCriterioCli
         If Not RsProduto.EOF Then
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            Txt(1).Text = RsProduto!codigo
            Txt(2).Text = RsProduto!Nome
            Unidade.Text = LcUnidade
            Txt(4).Text = RsProduto!QtdMedida
            valor(0).Text = RsProduto!Preco
            cst.Text = RsProduto!cst
            LcPrecoVelho = RsProduto!Preco
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = RsProduto!Preco / RsProduto!QtdMedida Else PrecoVendaNormal = 0
            ComNormal = RsProduto!QtdMedida
            minimo.Text = RsProduto!MinimoVenda & ""
            If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MinimoVenda / RsProduto!QtdMedida Else PrecoMimimodeVendaAlterado = 0
            If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
               If Val(cst.Text) = 60 Or Val(cst.Text) = 160 Or Val(cst.Text) = 260 Then
                  icms.Text = "00"
               Else
                 If Len(Txt(5).Text) = 0 Then
                       If IsNull(RsProduto!icms) Then
                          icms.Text = "18"
                       Else
                         If RsProduto!icms = 0 Then
                           icms.Text = "18"
                         Else
                           icms.Text = RsProduto!icms
                         End If
                       End If
                    Else
                       icms.Text = Txt(5).Text
                    End If
                End If
            Else
               icms = RsProduto!icms
            End If
            Custo.Text = RsProduto!CustoTotal
            LcAchou = True
            SendKeys "{TAB}"
         Else
            Txt(2).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(2).Text
        If Len(Txt(2).Text) = 0 Then Exit Function
        
        lcchave = Right("00000" & Txt(2).Text, 5)
        If IsNumeric(lcchave) Then
           LcCriterioCli = "Codigo=" & lcchave
        Else
          LcCriterioCli = "nome='" & lcchave & "'"
        End If
        RsProduto.MoveFirst
        RsProduto.Find LcCriterioCli
        If Not RsProduto.EOF Then
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            Txt(1).Text = RsProduto!codigo
            Txt(2).Text = RsProduto!Nome
            Unidade.Text = LcUnidade
            Txt(4).Text = RsProduto!QtdMedida
            LcAchou = 0
            If GlContrato Then
                LcSql = "SELECT ContratoDados.Valor,ContratoDados.CodProduto,ContratoFornecimento.Datai,ContratoFornecimento.DataF"
                LcSql = LcSql & " FROM ContratoDados INNER JOIN ContratoFornecimento ON ContratoDados.CodContrato = ContratoFornecimento.Codigo"
                LcSql = LcSql & " WHERE ContratoFornecimento.Cliente='" & Txt(9).Text & "' and Codproduto='" & Txt(1).Text & "'"
                LcSql = LcSql & " and ContratoFornecimento.DataI<#" & Format(Date, "mm/dd/yy") & "# and ContratoFornecimento.DataF>#" & Format(Date, "mm/dd/yy") & "#"
                Set Rs = AbreRecordset(LcSql, True)
                Do Until Rs.EOF
                    If Rs!CodProduto = Txt(1).Text Then
                        LcAchou = 1
                        valor(0).Text = AcertaNumero(CDbl(Rs!valor), 2)
                        PrecoVendaNormal = CDbl(Rs!valor)
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close
                Set Rs = Nothing
                If LcAchou = 0 Then
                    valor(0).Text = RsProduto!Preco
                End If
            Else
                valor(0).Text = RsProduto!Preco
            End If
            cst.Text = RsProduto!cst
            LcPrecoVelho = RsProduto!Preco
            If LcAchou = 0 Then
                If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = RsProduto!Preco / RsProduto!QtdMedida Else PrecoVendaNormal = 0
            End If
            ComNormal = RsProduto!QtdMedida
            minimo.Text = RsProduto!MinimoVenda & ""
            If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MinimoVenda / RsProduto!QtdMedida Else PrecoMimimodeVendaAlterado = 0
            If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
               If Val(cst.Text) = 60 Or Val(cst.Text) = 160 Or Val(cst.Text) = 260 Then
                  icms.Text = "00"
               Else
                 If Len(Txt(5).Text) = 0 Then
                       If IsNull(RsProduto!icms) Then
                          icms.Text = "18"
                       Else
                         If RsProduto!icms = 0 Then
                           icms.Text = "18"
                         Else
                           icms.Text = RsProduto!icms
                         End If
                       End If
                    Else
                       icms.Text = Txt(5).Text
                    End If
                End If
            Else
               icms = RsProduto!icms
            End If
            Custo.Text = RsProduto!CustoTotal
            If Not IsNumeric(Custo.Text) Then Custo.Text = 0
            'SendKeys "{TAB}"
        Else
            Txt(2).Text = LcValorDigitado
            GlCriterioSql = "select * From Produtos where nome like '" & UCase(Txt(2).Text) & "%'  order by nome"
            FrmPesquisaProdutos.Txt.Text = Txt(2).Text
            If LcAlteradoProduto Then
               FrmPesquisaProdutos.Show , Me
               LcAlteradoProduto = False
            End If
            'Data(1).SetFocus
        End If

        LcAlteradoProduto = False
End Select

RsProduto.Close
Set RsProduto = Nothing
Exit Function

errBuscaFor:
If err = 383 Then MsgBox "A Unidade deste Produto Não Foi Cadastrada.", 64, "Aviso": Resume Next
If err = 11 Then Resume Next
If err = 3420 Then

   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   Resume Next
End If




End Function
Function ImprimeGalpao()
On Error Resume Next
Dim LcImprimiu, a As Integer
FnunNota = FreeFile

LcSalto = Val(GLSaltoLinhaNota)
LcGalpao = GlEstoqueDisponivel


If IsNull(LcGalpao) Then LcGalpao = "LPT1"
LcImprimiu = False
LcImpressoes = 0
Open LcGalpao For Output Access Write As #FnunNota 'Abre Porta Nf
'Salta linhas no inicio da nota
LcLinha = ""
For a = 0 To LcTam - 1
    If LcMat(a).california > 0 Then
       If Not LcImprimiu Then
          LcLinha = "NOTA FISCAL N :" & Txt(0).Text
          Print #FnunNota, LcLinha + Chr(13)
          LcLinha = "CLIENTE       :" & Txt(9).Text
          Print #FnunNota, Chr(13)
          Print #FnunNota, LcLinha + Chr(13) + Chr(18)
          LcLinha = Left("Produto" & "                                      ", 40) & Left("Galpao" & "                  ", 20) & "Quantidade"
          Print #FnunNota, LcLinha + Chr(13)
          For C = 1 To 80
              LcEsp = LcEsp & "-"
          Next C
          Print #FnunNota, LcEsp + Chr(13)
          LcImprimiu = True
       End If
       LcLinha = Left(LcMat(a).produto & "                                      ", 40)
       LcLinha = LcLinha & Left("CALIFORNIA" & "                                      ", 20)
       LcLinha = LcLinha & LcMat(a).california
       Print #FnunNota, LcLinha + Chr(13)
    End If
    
        If LcMat(a).santamaria1 > 0 Then
       If Not LcImprimiu Then
          LcLinha = "NOTA FISCAL N :" & Txt(0).Text
          Print #FnunNota, LcLinha + Chr(13)
          LcLinha = "CLIENTE       :" & Txt(9).Text
          Print #FnunNota, Chr(13)
          Print #FnunNota, LcLinha + Chr(13) + Chr(18)
          LcLinha = Left("Produto" & "                                      ", 40) & Left("Galpao" & "                  ", 20) & "Quantidade"
          Print #FnunNota, LcLinha + Chr(13)
          For C = 1 To 80
              LcEsp = LcEsp & "-"
          Next C
          Print #FnunNota, LcEsp + Chr(13)
          LcImprimiu = True
       End If
       LcLinha = Left(LcMat(a).produto & "                                      ", 40)
       LcLinha = LcLinha & Left("SANTA MARIA 2" & "                                      ", 20)
       LcLinha = LcLinha & LcMat(a).santamaria1
       Print #FnunNota, LcLinha + Chr(13)
    End If

Next
Print #FnunNota, Chr(15) + Chr(13)
Close #FnunNota
End Function
Function BuscaVendendor(LcTipo As Integer)
'On Error Resume Next
On Error GoTo errBuscaFor
Dim RsVendedor As Recordset
Dim LcValorDigitado
Dim LcCodigo As String
If Not LcAlteradoFuncionario Then Exit Function
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("select * from alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & Txt(10).Text & "'"
         RsVendedor.FindFirst LcCriterioCli
         If Not RsVendedor.NoMatch Then
            Txt(7).Text = RsVendedor!Nome
            If CDbl(Comissao.Text) <> 1 Then
                Comissao.Text = RsVendedor!Comissao
            End If
            SendKeys "{TAB}"
         Else
            Txt(7).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(7).Text
        If Len(Txt(7).Text) = 0 Then Exit Function
        
        lcchave = Txt(7).Text
        LcCriterioCli = "NOME='" & lcchave & "'"
        RsVendedor.FindFirst LcCriterioCli
        If Not RsVendedor.NoMatch Then
            Txt(7).Text = RsVendedor!Nome
            Txt(10).Text = RsVendedor!codigo
            
            If Len(Comissao.Text) > 0 Then
               If CDbl(Comissao.Text) <> 1 Then
                  Comissao.Text = RsVendedor!Comissao & ""
               End If
            Else
              Comissao.Text = RsVendedor!Comissao & ""
            End If
            'SendKeys "{TAB}"
        Else
            Txt(7).Text = LcValorDigitado
            '=== Verifica se foi por nome
            lcchave = Txt(7).Text
            LcCriterioCli = "nome='" & lcchave & "'"
            RsVendedor.FindFirst LcCriterioCli
            If Not RsVendedor.NoMatch Then
               Txt(7).Text = RsVendedor!Nome
               Txt(10).Text = RsVendedor!codigo
            
               If Len(Comissao.Text) > 0 Then
                  If CDbl(Comissao.Text) <> 1 Then
                     Comissao.Text = RsVendedor!Comissao & ""
                  End If
               Else
                  Comissao.Text = RsVendedor!Comissao & ""
               End If
            'SendKeys "{TAB}"
            Else
               FrmPesquisaFuncionarios.Txt.Text = Txt(7).Text
               GlCriterioSql = "select * From alid200 where nome like '" & UCase(Txt(7).Text) & "*'  order by nome"
               If LcAlteradoFuncionario Then
                  FrmPesquisaFuncionarios.Show , Me
                  LcAlteradoFuncionario = False
               End If
            'Data(1).SetFocus
            End If
         End If
  
End Select

RsVendedor.Close
Set RsVendedor = Nothing
Exit Function

errBuscaFor:
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   'Resume 0
End If


End Function
Function BuscaCliente(LcTipo As Integer)
On Error GoTo errBuscaFor
Dim rsCliente As Recordset
Dim LcValorDigitado As String
Dim LcCodigo As String
Dim LcCredito, LcUtilizado As Currency

If LcAlteradoCliente Then Exit Function
AbreBase

GlLibera = False
LcAlteradoCliente = True
Set rsCliente = Dbbase.OpenRecordset("select * from alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & Txt(8).Text & "'"
         rsCliente.FindFirst LcCriterioCli
         If Not rsCliente.NoMatch Then
            Txt(9).Text = rsCliente!razaosoc
            LcDesCidade = rsCliente!razaosoc
            If rsCliente!comodato Then
               comodato.Value = 1
           Else
               comodato.Value = 0
           End If
            Txt(7).Text = rsCliente!TelemarketingAtende
            BuscaVendendor (2)
            
            If Not IsEmpty(rsCliente!LimiteCredito) And (Not IsNull(rsCliente!LimiteCredito)) Then LcCredito = rsCliente!LimiteCredito Else LcCredito = 0
            If Not IsEmpty(rsCliente!CreditoUtilizado) And (Not IsNull(rsCliente!CreditoUtilizado)) Then LcUtilizado = rsCliente!CreditoUtilizado Else LcUtilizado = 0
            limite.Text = LcCredito
            utilizado.Text = LcUtilizado
            
            SendKeys "{TAB}"
         Else
            Txt(9).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(9).Text
        'If Len(txt(9).Text) = 0 Then Exit Function
        If GLCalculacodigoCliente Then
           lcchave = Right("00000" & Txt(9).Text, 5)
        Else
           lcchave = Txt(9).Text
        End If
        
        LcCriterioCli = "CODIGO='" & lcchave & "'"
        rsCliente.FindFirst LcCriterioCli
        If Not rsCliente.NoMatch Then
            Txt(9).Text = rsCliente!razaosoc & ""
            Txt(8).Text = rsCliente!codigo & ""
            LcDesCidade = rsCliente!razaosoc & ""
            Txt(7).Text = rsCliente!TelemarketingAtende & ""
            If rsCliente!comodato Then
               comodato.Value = 1
           Else
               comodato.Value = 0
           End If
            BuscaVendendor (2)
            If Not IsEmpty(rsCliente!LimiteCredito) And (Not IsNull(rsCliente!LimiteCredito)) Then LcCredito = rsCliente!LimiteCredito Else LcCredito = 0
            If Not IsEmpty(rsCliente!CreditoUtilizado) And (Not IsNull(rsCliente!CreditoUtilizado)) Then LcUtilizado = rsCliente!CreditoUtilizado Else LcUtilizado = 0
            limite.Text = LcCredito
            utilizado.Text = LcUtilizado

            'SendKeys "{TAB}"
        Else
            Txt(9).Text = LcValorDigitado
            FrmBuscaCliente.Txt.Text = Txt(9).Text
            GlCriterioSql = "  where RAZAOSOC like '" & UCase(Txt(9).Text) & "*'  order by RAZAOSOC"
            If Not LcBuscaNota Then
               FrmBuscaCliente.Show , Me
               LcAlteradoCliente = True
            End If
            'Data(1).SetFocus
        End If
  
End Select
HabilitaClienteBloqueado Not rsCliente!bloqueado

'AbreOBS
LcPesquisaCli = True
rsCliente.Close
Set rsCliente = Nothing
Exit Function

errBuscaFor:
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   Resume 0
End If

End Function
Sub HabilitaClienteBloqueado(Bloqueia As Boolean)
Command3.Enabled = Bloqueia
Command5.Enabled = Bloqueia
CmdExcluir.Enabled = Bloqueia
Command7.Enabled = Bloqueia
Command11.Enabled = Bloqueia
Command9.Enabled = Bloqueia
CmdDuplicarPedido.Enabled = Bloqueia
Txt(2).Enabled = Bloqueia
Unidade.Enabled = Bloqueia
Txt(4).Enabled = Bloqueia
Txt(3).Enabled = Bloqueia
valor(0).Enabled = Bloqueia
valor(1).Enabled = Bloqueia

LabelBloqueado.Visible = Not Bloqueia


End Sub
Sub AbreOBS()
On Error Resume Next
If Len(Txt(8).Text) > 0 Then
    Load FrmObsCliente
    FrmObsCliente.Financeira.Locked = True
    FrmObsCliente.codigo.Caption = Txt(8).Text
    FrmObsCliente.Show , Me
End If
End Sub
Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command1_Click()
On Error Resume Next
FrmPesquisaCliente.Show , Me
End Sub

Private Sub Command10_Click()
FrmListaPedidoPendente.Show , Me
End Sub

Private Sub Command11_Click()
If SalvaNota Then
   ImprimeSubItem
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
FrmPesquisaProdutos.Show , Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
tam.Text = LcTam
DadosTransp.Show , Me
End Sub

Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Dim conexao As New ADODB.Connection
Dim strConnect As String
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
strConnect = "driver={Microsoft Access Driver (*.mdb)};DBQ=" & GLBase & ";UID=Admin;PWD=;"
Set conexaoo = New ADODB.Connection
conexao.CursorLocation = adUseClient
'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor
LcComentario = "- Função 'abreconexao - Abrindo a Conexão com o DB."
'MsgBox strConnect
conexao.Open strConnect

RsAtual.ActiveConnection = conexao

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function
Sub ImprimeCrys11()
Dim StrSql As String

Dim Rs As ADODB.Recordset
Dim StrSqlObs As String
Dim LcOBS As String

'If Len(GlMsg) > 0 And Len(GlMsg1) > 0 And Len(GlMsg2) > 0 Then
'   LcOBS = GlMsg & Chr(10) & GlMsg1 & Chr(10) & GlMsg2
'ElseIf Len(GlMsg) > 0 And Len(GlMsg1) > 0 Then
'   LcOBS = GlMsg & Chr(10) & GlMsg1
'ElseIf Len(GlMsg) > 0 Then
'   LcOBS = GlMsg
'End If
'If Len(LcOBS) > 0 Then
'   StrSqlObs = "Update proposta set obs= [obs] + '" & LcOBS & "' WHERE NUMNF='" & UCase(Txt(0).Text) & "'"
'   AbreBase
'   Dbbase.Execute StrSqlObs
'End If

StrSql = "SELECT proposta.NUMNF, proposta.DTEMIS, proposta.NATUREZA, proposta.CLIENTE, proposta.TRANSP, proposta.TIPOTRANS, proposta.PLACATRANS, proposta.UFTRANS, proposta.CGCCPFTRAN,"
StrSql = StrSql & " proposta.ENDTRANS , subproposta.Item, subproposta.Galpao, subproposta.QTDE, subproposta.VALUNIT, subproposta.CONTR,subproposta.NCM, subproposta.UNIMED, subproposta.QTDUM,"
StrSql = StrSql & " subproposta.QTDE01, subproposta.QTDE02, subproposta.QTDE03, subproposta.codigo, subproposta.Descricao, subproposta.codProd, subproposta.bloqueado,"
StrSql = StrSql & " subproposta.tipoliberacao, subproposta.jaEsteveBloqueado, subproposta.MaquinaLiberacao, subproposta.DataLiberacao, subproposta.HoraLiberacao,"
StrSql = StrSql & " subproposta.usuario, subproposta.faturado, subproposta.SubItem, subproposta.NCM, subproposta.Compra, subproposta.cst, proposta.MUNICTRANS,"
StrSql = StrSql & " proposta.UFMUNIC, proposta.INSCEST, proposta.OBS02, proposta.OBS03, proposta.OBS04, proposta.CONTR, proposta.codigo, proposta.ValorProduto,"
StrSql = StrSql & " proposta.ValorNota, proposta.Vendedor, proposta.FoneTransp, proposta.Cidade, proposta.Cep, proposta.formapag, proposta.Dias, proposta.vencimento1, proposta.vencimento2, proposta.vencimento3, proposta.vencimento4, proposta.vencimento5, proposta.CondPag, proposta.status, proposta.DESCONTO,"
StrSql = StrSql & " proposta.Previsao, proposta.Liberado, proposta.faturado, proposta.Validade, proposta.OrdemCompra, proposta.obs, proposta.ICMS,"
StrSql = StrSql & " proposta.Bloqueado, proposta.dataliberacao, proposta.JaEsteveBloqueado, proposta.MaquinaLiberacao, proposta.HoraLiberacao, proposta.pendente,"
StrSql = StrSql & " proposta.Usuario, proposta.dias1, proposta.dias2, proposta.dias3, proposta.Romaneio, proposta.EnderecoEntrega, proposta.Oc, ALID001.FANTASIA,"
StrSql = StrSql & " ALID001.RAZAOSOC, [alid001].[end] & ', ' & [numero] AS END, ALID001.BAIRRO, ALID001.CIDADE, ALID001.ESTADO, ALID001.CEP, ALID001.FONE1,"
StrSql = StrSql & " ALID001.FONE2 , ALID001.Contato, ALID001.CGC, ALID001.INSCEST, ALID001.Email, ALID001.cpf, ALID001.rg, proposta.formapag, proposta.Dias, proposta.Validade, "
StrSql = StrSql & " proposta.valorproduto, proposta.ValorNota, proposta.DESCONTO, ALID200.NOME, ALID001.Numero, ALID005.NOME AS NomeCidade"
StrSql = StrSql & " FROM (((proposta INNER JOIN subproposta ON proposta.NUMNF = subproposta.NUMNF) INNER JOIN ALID001 ON proposta.CLIENTE = ALID001.CODIGO) LEFT JOIN ALID200 ON proposta.Vendedor = ALID200.CODIGO) INNER JOIN ALID005 ON ALID001.CIDADE = ALID005.COD"
StrSql = StrSql & " WHERE (((proposta.NUMNF)='" & UCase(Txt(0).Text) & "'))"
StrSql = StrSql & " ORDER BY subproposta.ITEM;"
Debug.Print StrSql

Set Rs = AbreRecordsetRel(StrSql, Rs)
Load Relatorios
If InStr(UCase(GLBase), "LIDIS") <> 0 Then
    With Relatorios
         RelLidis.DiscardSavedData
         RelLidis.Database.SetDataSource Rs
         .CRViewer1.ReportSource = RelLidis
    End With
Else
    With Relatorios
         Rel.DiscardSavedData
         Rel.Database.SetDataSource Rs
         .CRViewer1.ReportSource = Rel
    End With
End If
 setaformula
Relatorios.CRViewer1.ViewReport
    Relatorios.Show

Screen.MousePointer = vbDefault

'Me.Caption = LcCap

End Sub
Sub setaformula()
Dim a As Integer
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset
Dim RsOpcao As Recordset
Dim LcValor As Double
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
Dim lctitulo As String
Dim StrSql As String
Dim bb     As Database

Set db = OpenDatabase(GLBase)
Set RsEmpresa = db.OpenRecordset("Select * from EMPRESA")

If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & "" & IIf(Not IsNull(RsEmpresa!Fax), " Fax:" & RsEmpresa!Fax, "")
   LcInscricao = RsEmpresa!inscricaoestadual & ""
   LcCNPJ = RsEmpresa!CGC & ""
   Celular = "Insc. Estadual: " & LcInscricao '.608783.0021'"
   Lcemail = "CNPJ: " & LcCNPJ '.682.162/0001-88'"
End If
Set RsEmpresa = Nothing
If InStr(UCase(GLBase), "LIDIS") <> 0 Then
    With RelLidis
'Exit Sub
        For a = 1 To .FormulaFields.Count
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("VendedorNome") Then
               If Len(VendedorImprimir.Text) = 0 Then
                   .FormulaFields(a).Text = "totext('" & Txt(7).Text & "')"
               Else
                    .FormulaFields(a).Text = "totext('" & VendedorImprimir.Text & "')"
               End If
              
            End If
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMAIL") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Celular") Then .FormulaFields(a).Text = "totext('" & Celular & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg1") Then .FormulaFields(a).Text = "totext('" & GlMsg & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg2") Then .FormulaFields(a).Text = "totext('" & GlMsg1 & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg3") Then .FormulaFields(a).Text = "totext('" & GlMsg2 & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
               .FormulaFields(a).Text = "totext('" & lctitulo & "')"
            End If
        Next
    End With
Else
    With Rel
'Exit Sub
        For a = 1 To .FormulaFields.Count
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("VendedorNome") Then
               If Len(VendedorImprimir.Text) = 0 Then
                   .FormulaFields(a).Text = "totext('" & Txt(7).Text & "')"
               Else
                    .FormulaFields(a).Text = "totext('" & VendedorImprimir.Text & "')"
               End If
              
            End If
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMAIL") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Celular") Then .FormulaFields(a).Text = "totext('" & Celular & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg1") Then .FormulaFields(a).Text = "totext('" & GlMsg & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg2") Then .FormulaFields(a).Text = "totext('" & GlMsg1 & "')"
            'If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg3") Then .FormulaFields(a).Text = "totext('" & GlMsg2 & "')"
            If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
               .FormulaFields(a).Text = "totext('" & lctitulo & "')"
            End If
        Next
    End With


End If


End Sub
Function Imprime()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim LcCNPJ As String
Dim LcInscricao As String
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & IIf(Not IsNull(RsEmpresa!Fax), " Fax:" & RsEmpresa!Fax, "")
   LcInscricao = RsEmpresa!inscricaoestadual & ""
   LcCNPJ = RsEmpresa!CGC & ""
End If


'Abertura do relatório de vendas
    
    
    CryRelatorio.DataFiles(0) = GLBase
    CryRelatorio.ReportFileName = App.Path & "\PropostasVendas.rpt"
    LcFormula = "{proposta.numnf}='" & UCase(Txt(0).Text) & "'"
    CryRelatorio.CopiesToPrinter = 1

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Pedido de Venda"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'" '(31)3388-1015 - Fax :3388-2520'"
CryRelatorio.Formulas(3) = "Celular='Insc. Estadual: " & LcInscricao & "'" '.608783.0021'"
CryRelatorio.Formulas(4) = "email='CNPJ: " & LcCNPJ & "'" '.682.162/0001-88'"
'CryRelatorio.Formulas(5) = "titulo='Produtos'"
 
LcTipoSaida = 0
Me.Caption = LcCap
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport

'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Function
Function ImprimeSubItem()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim LcCNPJ As String
Dim LcInscricao As String
Dim RsProduto As ADODB.Recordset
Dim Rs        As ADODB.Recordset
Dim LcValorDigitado
Dim LcCodigo As String
'AbreBase

AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & IIf(Not IsNull(RsEmpresa!Fax), " Fax:" & RsEmpresa!Fax, "")
   LcInscricao = RsEmpresa!inscricaoestadual & ""
   LcCNPJ = RsEmpresa!CGC & ""
End If

'==> Acrescenta o sub item na tabela de proposta
For a = 1 To Item.Rows
     Dim LcCodProduto As String
     LcCodProduto = Item.TextMatrix(a, 1)
     If Len(Trim(LcCodProduto)) > 0 Then
        Set RsProduto = AbreRecordset("select * from Produtos where codigo=" & LcCodProduto, True)  ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
        If Not RsProduto.EOF Then
            LcSq = "update subproposta set "
            LcSq = LcSq & "SubItem='" & RsProduto!subitem & "',"
            LcSq = LcSq & "NCM='" & RsProduto!ClassificacaoFiscal & "',"
            LcSq = LcSq & "CST='" & RsProduto!cst & "',"
            LcSq = LcSq & "Compra=" & Replace(RsProduto!Custo, ",", ".") & ""
                        LcSq = LcSq & " where NUMNF='" & Txt(0).Text & "' and codProd='" & LcCodProduto & "'"
            Dbbase.Execute LcSq
            Debug.Print LcSq
            LcAfetados = Dbbase.RecordsAffected
        End If
     End If
Next
'Abertura do relatório de vendas
    CryRelatorio.DataFiles(0) = GLBase
    CryRelatorio.ReportFileName = App.Path & "\PropostasVendasSubItem.rpt"
    LcFormula = "{proposta.numnf}='" & UCase(Txt(0).Text) & "'"
    CryRelatorio.CopiesToPrinter = 1

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Pedido de Venda"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'" '(31)3388-1015 - Fax :3388-2520'"
CryRelatorio.Formulas(3) = "Celular='Insc. Estadual: " & LcInscricao & "'" '.608783.0021'"
CryRelatorio.Formulas(4) = "email='CNPJ: " & LcCNPJ & "'" '.682.162/0001-88'"
'CryRelatorio.Formulas(5) = "titulo='Produtos'"
 
LcTipoSaida = 0
Me.Caption = LcCap
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport

'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Function

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Command4.Caption = "&Pesquisar        F9" Then
   Load FrmPesquisaNota
   FrmPesquisaNota.Tag = "proposta"
   FrmPesquisaNota.Show , Me
   Command4.Caption = "&Pesquisar        F9"
   LcPesquisa = True
   Txt(0).Locked = False
Else
   Command4.Caption = "&Pesquisar        F9"
   limpanota
   LcPesquisa = False
End If

End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command5_Click()
SalvaNota
limpanota

End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command6_Click()
Observacao.Show , Me

End Sub

Private Sub Command6_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 119 Then SendKeys "%+{I}"
If KeyCode = 117 Then SendKeys "%+{O}"
If KeyCode = 120 Then Call Command8_Click
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Command7_Click()
If SalvaNota Then
  'Imprime
  ImprimeCrys11
  limpanota
End If
End Sub

Private Sub Command7_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command8_Click()
Dim RsOrc As Recordset, RsItem As Recordset
Dim LcSql1, LcSql2, LcSql3, LcSql4, LcSql5 As String
Dim LcResp As Integer
Dim LcExcluidos As Integer
Dim LcExcluidos1 As Integer
If Len(Txt(0).Text) = 0 Then
   MsgBox "É nescessario a escolha de um pedido para a exclusão.", 64, "Aviso"
   Exit Sub
End If
LcResp = MsgBox("Excluir Pedido?", vbCritical + vbYesNo, "Aviso")
'Select Case LcResp
If LcResp = vbNo Then Exit Sub
LcPesquisa = True
LcSql1 = "delete from proposta where NUMNF='" & Txt(0).Text & "'"
LcSql2 = "delete from subproposta where NUMNF='" & Txt(0).Text & "'"
LcBuscaNota = False
AbreBase
Dbbase.Execute LcSql1, LcExcluidos
Dbbase.Execute LcSql2, LcExcluidos1

'Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsItem = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'If Not RsOrc.EOF Then
'   RsOrc.Delete
'   LcBuscaNota = True
'End If
'Do Until RsItem.EOF
'   RsItem.Delete
 '  RsItem.MoveNext
 '  LcBuscaNota = True
'Loop
limpanota

   MsgBox "Pedido Excluido com Sucesso.", 64, "Aviso"

'RsOrc.Close
'RsItem.Close
End Sub

Private Sub Command8_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command9_Click()
On Error Resume Next
Frmliberapedido.Show , Me
End Sub
Function LiberaPedido()
On Error Resume Next
Dim LcPesq  As String
Dim LcSql   As String
Dim Rs      As ADODB.Recordset
Dim Ru      As Recordset
Dim LcAchou As Boolean
Dim LcLista As String
Dim LcCm    As String
Dim LcUn()  As String

AbreBase
Set Rs = AbreRecordset("Select * from Produtos", True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Ru = Dbbase.OpenRecordset("Select * from alid004", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcLista = "Os Seguintes Item(s) Continuam com o Estoque Insuficiente." & Chr(13)
'==> vamos Checar se Tem Os Creditos
For a = 1 To Item.Rows - 1
    If CInt(Item.TextMatrix(a, 12)) = 2 Or CInt(Item.TextMatrix(a, 12)) = 3 Then
       LcPesq = "SIMBOLO='" & Left(Item.TextMatrix(a, 4), 2) & "'"
       Ru.FindFirst LcPesq
       LcCod = Ru!cod
       '===> Verifica a diferença de Unidades
       LcPes = "codigo=" & Item.TextMatrix(a, 1)
       Rs.Find LcPes
       'If LcCod = Rs!UnidMedida Then
       '   If CDbl(Item.TextMatrix(a, 5)) > Rs!QuantEstoque Then
       '      LcLista = LcLista & Chr(13) & Item.TextMatrix(a, 2)
       '      LcAchou = True
       '   End If
       'Else
       LcUn = Split(Item.TextMatrix(a, 4), "/")
                 
          
          '===> Procura / para saber a quantidade digitada
        '  For X = Len(Item.TextMatrix(a, 4)) To 1 Step -1
        '        If Mid(Item.TextMatrix(a, 4), X, 1) = "/" Then
        '           Exit For
        '        Else
        '           LcCm = Mid(Item.TextMatrix(a, 4), X, 1) & LcCm
        '        End If
        '  Next
       '   LcPes = "cod='" & Item.TextMatrix(a, 1) & "'"
       '   Rs.FindFirst LcPes
       If (CDbl(LcUn(1)) * CDbl(Item.TextMatrix(a, 5))) > (Rs!QuantEstoque) Then
          LcLista = LcLista & Chr(13) & Item.TextMatrix(a, 2)
          LcAchou = True
      End If
   End If
Next
'===> O Estoque não Foi Corrigido, Então Avisa
If LcAchou Then
    LcLista = LcLista & Chr(13) & Chr(13) & "Libera o Pedido Mesmo Assim ?"
    LcResp = MsgBox(LcLista, vbExclamation + vbYesNo, "Estoque Insuficiente")
    If LcResp = 7 Then GoTo Saida
End If

For a = 0 To LcTam - 1
    LcMat(a).DataLiberacao = Date
    LcMat(a).bloqueado = False
    LcMat(a).HoraLiberacao = Time
    LcMat(a).MaquinaLiberacao = GlNomeMaquina
    LcMat(a).usuario = GlUsuario
    LcMat(a).tipoliberacao = 0
Next
'==> Libera as Cores
For a = 1 To Item.Rows - 1
   For x = 0 To 15
       Item.Row = a
       Item.Col = x
       Item.CellBackColor = LcCorPadrao
   Next

Next

DataLiberacao.Text = Date
HoraLiberacao.Text = Time
MaquinaLiberacao.Text = GlNomeMaquina

usuario.Text = GlUsuario
Pendente.Value = 0
bloqueado.Text = "False"
Command9.Enabled = False

Saida:
Rs.Close
Ru.Close
Dbbase.Close
Set Rs = Nothing
Set Ru = Nothing
Set Dbbase = Nothing

End Function

Private Sub faturado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If GlExibirLucratividade Then
    Lucratividade.Visible = True
    Label19.Visible = True
Else
    Lucratividade.Visible = False
    Label19.Visible = False
End If
If Not GlCarregado Then
   Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
   GlCarregado = True
End If
LcCorPadrao = 0
End Sub
Function BuscaNota(LcNumeroOrc As String)
On Error GoTo ErroBuscaNota
Dim RsOrc As Recordset, RsItem As Recordset
Dim RsProduto As ADODB.Recordset, rsCliente As Recordset
Dim RsVendedor As Recordset
Dim LcSql1, LcSql2, LcSql3, LcSql4, LcSql5 As String
Dim LcSql6 As String

LcPesquisa = True
LcSql1 = "Select * from proposta where NUMNF='" & LcNumeroOrc & "'"
LcSql2 = "Select * from subproposta where NUMNF='" & LcNumeroOrc & "' order by item"
LcSql3 = "Select * from ALid001"
LcSql5 = "Select * from ALid200"
LcSql6 = "Select * from Produtos"

LcBuscaNota = True
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsItem = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set rsCliente = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVendedor = Dbbase.OpenRecordset(LcSql5, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = AbreRecordset(LcSql6, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)

'==== Preenchendo a Nota

If RsOrc.EOF Then
   MsgBox "A Proposta Nº: " & LcNumeroOrc & " Não foi encontrado..."
   Command4.Caption = "&Pesquisar        F9"
   Txt(10).SetFocus
   Exit Function
End If
Txt(0).Text = RsOrc!numnf
Txt(12).Text = Format(RsOrc!DtEmis, "dd/mm/yy")
Txt(6).Text = RsOrc!Status
Txt(17).Text = RsOrc!Desconto
Select Case RsOrc("NATUREZA")
    Case Is = "VV"
         Natureza.Text = "VENDAS A VISTA"
    Case Is = "VP"
         Natureza.Text = "VENDAS A PRAZO"
    Case Is = "EM"
        Natureza.Text = "EMPENHO"
    Case Is = "TR"
        Natureza.Text = "TRANSFERENCIA"
    Case Is = "DE"
      Dim LcAchou As Boolean
      LcAchou = False
      For a = 0 To Natureza.ListCount - 1
          If UCase(Natureza.List(a)) = "DEVOLUCAO" Then
          LcAchou = True
          Exit For
          End If
      Next
      If LcAchou Then
         Natureza.Text = "DEVOLUCAO"
      Else
        Natureza.Text = "VENDAS A VISTA"
      End If
        
End Select
If Len(RsOrc!Comissao) > 0 Then
   Comissao.Text = RsOrc!Comissao
Else
   Comissao.Text = 1.5
End If
If Len(RsOrc!CFOP) > 0 Then
   CFOP.Text = RsOrc!CFOP
Else
   CFOP.Text = "512"
End If
Txt(10).Text = RsOrc!vendedor & ""
LcCriterio = "Codigo='" & RsOrc!vendedor & "'"
RsVendedor.FindFirst LcCriterio
If Not RsVendedor.NoMatch Then
   Txt(7).Text = RsVendedor!Nome
Else
  Txt(7).Text = ""
End If
Txt(8).Text = RsOrc!Cliente
LcCriterio = "Codigo='" & RsOrc!Cliente & "'"
Previsao.Text = RsOrc!Previsao
LiberaFaturamento = RsOrc!Liberado
If RsOrc!faturado Then
   faturado = 1
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
Else
   faturado = 0
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
End If

rsCliente.FindFirst LcCriterio
If Not rsCliente.NoMatch Then
   Txt(9).Text = rsCliente!razaosoc
   Txt(8).Text = rsCliente!codigo
End If
obs.Text = RsOrc!obs
Condicoes.Text = RsOrc!FormaPag
Ordem.Text = RsOrc!ordemcompra
Validade.Text = RsOrc!Validade
Txt(15).Text = RsOrc!ValorProduto
Txt(16).Text = RsOrc!ValorNota
Txt(5).Text = RsOrc!icms & ""
Txt(17).Text = RsOrc!Desconto & ""
Txt(14).Text = RsOrc!obs & ""
If RsOrc!bloqueado Then
   bloqueado.Text = True
Else
   bloqueado.Text = False
End If
If RsOrc!jaEsteveBloqueado Then
   JaBloqueado.Text = True
Else
   JaBloqueado.Text = False
End If
DataLiberacao.Text = RsOrc!DataLiberacao & ""
MaquinaLiberacao.Text = RsOrc!MaquinaLiberacao & ""
HoraLiberacao.Text = RsOrc!HoraLiberacao & ""
If RsOrc!Pendente Then
   Pendente.Value = 1
Else
   Pendente.Value = 0
End If
CondPag.Text = RsOrc("CondPag") & ""
PrazoEntrega.Text = RsOrc("Prazo") & ""
ValidadeCotacao.Text = RsOrc("Validade") & ""
VendedorImprimir.Text = RsOrc("VendedorImprimir") & ""
InformacoesComplementares.Text = RsOrc("InfComplementar") & ""


Load DadosTransp
DadosTransp.Txt(0).Text = RsOrc("TRANSP") & ""
DadosTransp.Tipo.Text = RsOrc("TIPOTRANS") & ""
DadosTransp.Placa.Text = RsOrc("PLACATRANS") & ""
DadosTransp.Txt(1).Text = RsOrc("UFTRANS") & ""
DadosTransp.Txt(2).Text = RsOrc("CGCCPFTRAN") & ""
DadosTransp.Txt(3).Text = RsOrc("ENDTRANS") & ""
DadosTransp.Txt(4).Text = RsOrc("MUNICTRANS") & ""
DadosTransp.Txt(5).Text = RsOrc("UFMUNIC") & ""
DadosTransp.Txt(6).Text = RsOrc("INSCEST") & ""
DadosTransp.Txt(7).Text = RsOrc("OBS02") & ""
DadosTransp.Txt(8).Text = RsOrc("OBS03") & ""
DadosTransp.Txt(9).Text = RsOrc("OBS04") & ""
DadosTransp.TipoMonetario.Text = RsOrc("formapag") & ""
DadosTransp.Dias(0).Text = RsOrc("dias1") & ""
DadosTransp.Dias(1).Text = RsOrc("dias2") & ""
DadosTransp.Dias(2).Text = RsOrc("dias3") & ""
DadosTransp.Txt(10).Text = RsOrc("ENDERECOENTREGA") & ""
DadosTransp.Txt(11).Text = RsOrc("OC") & ""
DadosTransp.Hide
If IsDate(RsOrc("vencimento1")) Then DadosTransp.Vencimento(0).Text = Format(RsOrc("vencimento1"), "dd/mm/yy")
If IsDate(RsOrc("vencimento2")) Then DadosTransp.Vencimento(1).Text = Format(RsOrc("vencimento2"), "dd/mm/yy")
If IsDate(RsOrc("vencimento3")) Then DadosTransp.Vencimento(2).Text = Format(RsOrc("vencimento3"), "dd/mm/yy")


Command9.Enabled = RsOrc!Pendente
'If Len(RsOrc!desconto) > 0 Then Txt(13).Text = RsOrc!desconto Else Txt(13).Text = ""
'If Len(RsOrc!TotalDesconto) > 0 Then desconto.Text = RsOrc!TotalDesconto Else desconto.Text = ""
'===== Escreve dados Grid
LcItem = 0
LcTam = 0
'ReDim LcMat(LcTam)
Do Until RsItem.EOF
    LcItem = LcItem + 1
    ReDim Preserve LcMat(LcTam)
    If Len(RsItem!Item) > 0 Then LcMat(LcTam).Item = RsItem!Item
      LcCriterio = "Codigo=" & RsItem("codProd")
      RsProduto.MoveFirst
      RsProduto.Find LcCriterio
      If Not RsProduto.EOF Then
            cst.Text = RsProduto!cst
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = RsProduto!Preco / RsProduto!QtdMedida Else PrecoVendaNormal = 0
            ComNormal = RsProduto!QtdMedida
            minimo.Text = RsProduto!MinimoVenda & ""
            If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MinimoVenda / RsProduto!QtdMedida Else PrecoMimimodeVendaAlterado = 0
            If Val(cst.Text) = 60 Or Val(cst.Text) = 160 Or Val(cst.Text) = 260 Then
                icms.Text = "00"
            Else
               If Len(Txt(5).Text) = 0 Then
                       If IsNull(RsProduto!icms) Then
                          icms.Text = "18"
                       Else
                         If RsProduto!icms = 0 Then
                           icms.Text = "18"
                         Else
                           icms.Text = RsProduto!icms
                         End If
                       End If
                    Else
                       icms.Text = Txt(5).Text
                    End If
            End If
        
      End If
      LcMat(LcTam).Venda1 = RsProduto!CustoTotal
      LcMat(LcTam).cst = cst.Text & ""
      LcMat(LcTam).icms = icms.Text & ""
      LcMat(LcTam).CodPro = RsItem("codProd") & ""
      LcMat(LcTam).Qut = RsItem("QTDE")
      LcMat(LcTam).VUnit = RsItem("VALUNIT")
      LcMat(LcTam).Und = RsItem("UNIMED")
      LcMat(LcTam).Com = RsItem("QTDUM")
      LcMat(LcTam).produto = RsItem("descricao") & ""
      LcMat(LcTam).Vtotal = LcMat(LcTam).Qut * LcMat(LcTam).VUnit
      
      LcMat(LcTam).bloqueado = RsItem("Bloqueado")
      LcMat(LcTam).tipoliberacao = RsItem("TipoLiberacao")
      LcMat(LcTam).jaEsteveBloqueado = RsItem("jaEsteveBloqueado")
      LcMat(LcTam).MaquinaLiberacao = RsItem("MaquinaLiberacao") & ""
      If IsDate(RsItem("DataLiberacao")) Then LcMat(LcTam).DataLiberacao = RsItem("DataLiberacao")
      LcMat(LcTam).HoraLiberacao = RsItem("HoraLiberacao") & ""
      LcTam = LcTam + 1
      
      RsItem.MoveNext
    LcAchou = True
Loop
 EscreveGrid

 If LcAchou Then
    FrmSaidaProduto.SetFocus
    Txt(2).SetFocus
 Else
    Txt(10).SetFocus
    CmdSalvar.Enabled = True
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    
 End If
 Command8.Enabled = True
 RsOrc.Close
 RsItem.Close
 rsCliente.Close
 RsVendedor.Close
 LcBuscaNota = False
 VerificaLucratividade
 Exit Function
 
ErroBuscaNota:
 'MsgBox Err.Description & Err.Number
 Resume Next
End Function
Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
'Me.Height = 8400
'Me.Width = 11970
GlEscolhe = 1
CarregaCboUnidade
CarregaComboNatureza
Txt(6).Text = "EM LANCAMENTO"
Txt(0).Locked = True
Previsao.Text = Format(Date, "DD/MM/YY")
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
LabelBloqueado.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
GlCarregado = False
End Sub

Private Sub Item_DblClick()
Dim LcRow As Integer
Dim lccol As Integer
On Error Resume Next
If Item.Rows = 1 Then Exit Sub
LcColuna = Item.Col
linha = Item.Row
'If UCase(Txt(6).Text) = "EMITIDA" Then Exit Sub
If LcColuna = 5 Then
    If faturado.Value = 1 Then MsgBox "Pedido já faturado.", 64, "Aviso": Exit Sub
    LcValor = InputBox("Entre com a Nova Quantidade o Produto", "Produto: " & Item.TextMatrix(linha, 2))
    If Len(LcValor) > 0 Then
        If CDbl(LcValor) > 0 Then
            LcValor = Replace(LcValor, ".", ",")
            LcItemc = Item.TextMatrix(linha, 0)
            For z = 0 To UBound(LcMat)
               If Len(LcMat(z).CodPro) > 0 Then
                    If LcMat(z).Item = LcItemc Then
                        LcMat(z).Qut = CCur(LcValor)
                        LcMat(z).Vtotal = CCur(LcMat(z).VUnit) * CCur(LcValor)
                        EscreveGrid True
                       ' Call EscreveGrid(VerificaDisponivelGrid(LcMat(z).CodPro, CCur(LcValor), CCur(LcMat(z).com)), ConferePrecoGrid(LcMat(z).CodPro, CCur(LcMat(z).VUnit), CCur(LcMat(z).com)), True)
                        Exit For
                    End If
               End If
            Next
        End If
    End If
    If linha > 10 Then
         Item.TopRow = linha
        Item.Row = linha
        Item.Col = 5
        Item.ColSel = linha.Cols - 1
    End If
   
   ' Item.Row = linha
   ' Item.Col = 5
    Exit Sub
End If
If Item.Col = 6 Then
    If faturado.Value = 1 Then MsgBox "Pedido já faturado.", 64, "Aviso": Exit Sub

    LcValor = InputBox("Entre com o novo Valor Unitario para o Produto", "Produto: " & Item.TextMatrix(linha, 2))
    If Len(LcValor) > 0 Then
        LcValor = Replace(LcValor, ".", ",")
        LcItemc = Item.TextMatrix(linha, 0)
        For z = 0 To UBound(LcMat)
          If Len(LcMat(z).CodPro) > 0 Then
                If LcMat(z).Item = LcItemc Then
                    LcMat(z).VUnit = CCur(LcValor)
                    LcMat(z).Vtotal = CCur(LcValor) * CCur(LcMat(z).Qut)
                    EscreveGrid True
                    'Call EscreveGrid(VerificaDisponivelGrid(LcMat(z).CodPro, CCur(LcMat(z).Qut), CCur(LcMat(z).com)), ConferePrecoGrid(LcMat(z).CodPro, CCur(LcValor), CCur(LcMat(z).com)), True)
                    Exit For
                End If
          End If
        Next
        If linha > 10 Then
            Item.TopRow = linha
            Item.Row = linha
            Item.Col = 6
            Item.ColSel = linha.Cols - 1
        End If
    End If
    Exit Sub
End If

LcRow = Item.Row
Load exibeitem
exibeitem.Tag = Item.TextMatrix(LcRow, 1)
exibeitem.Show , Me
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub LiberaFaturamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 13 Then SendKeys "{TAB}"
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Natureza_Click()
If Natureza = "TRANSFERENCIA" Then
   Txt(5).Text = 0
Else
   Txt(5).Text = ""
End If
If Natureza.Text = "ORÇAMENTO" Then
   LiberaFaturamento.Enabled = False
Else
   LiberaFaturamento.Enabled = True
End If
BuscaCFOPPadrao
End Sub
Sub BuscaCFOPPadrao()
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim LcNome As String
Dim primeiro As Boolean
primeiro = True
StrSql = "Select * from naturezaoperacao where nome='" & Natureza.Text & "'"
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
   If Not IsNull(Rs!cfoppadrao) Then
      If Len(Rs!cfoppadrao) > 0 Then
         CFOP.Text = Rs!cfoppadrao
      End If
   End If
    
End If
Rs.Close
Set Rs = Nothing
End Sub
Private Sub Natureza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 13 Then SendKeys "{TAB}"
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub



Private Sub Previsao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 13 Then SendKeys "{TAB}"
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Previsao_LostFocus()
On Error Resume Next
If Previsao.Text = "  /  /  " Then Exit Sub
If Not IsDate(Previsao.Text) Then
   MsgBox "Data Inválida.", vbInformation, "Aviso"
   Previsao.Text = "  /  /  "
   Previsao.SetFocus
End If
End Sub

Private Sub Txt_Change(Index As Integer)
On Error Resume Next
LcCalculado = False
If Index = 3 Or Index = 5 Then CalculaValores
If Index = 9 Then LcAlteradoCliente = True
If Index = 2 Then LcAlteradoProduto = True
If Index = 7 Then LcAlteradoFuncionario = True
If Index = 8 Then
   If Len(Txt(8).Text) > 0 Then
      Txt(8).Text = Right("00000" & Txt(8).Text, 5)
      If Len(Trim(Txt(8).Text)) > 0 Then BuscaCliente (2)
   End If
End If
End Sub
Function CalculaDesconto()
On Error Resume Next
If Len(Trim(Txt(17).Text)) = 0 Then Txt(17).Text = 0

If Not IsNumeric(Txt(17).Text) Then
   MsgBox "Digite o Desconto Em Valor Numérico...", 64, "Aviso"
   Txt(17).SetFocus
   Exit Function
End If
CalculaIcms
VerificaLucratividade
End Function
Private Sub txt_GotFocus(Index As Integer)
On Error Resume Next
Dim a As Integer
LcLimpa = True
If Index = 3 Then
   For a = 0 To LcTam - 1
     If LcMat(a).CodPro = Txt(1).Text Then
        MsgBox "O Produto " & Txt(2).Text & " Já está selecionado...", vbInformation, "Item Duplicado."
        Txt(2).Text = ""
        Txt(1).Text = ""
        Txt(2).SetFocus
     End If
   Next
End If
If Index = 7 Then
   If Len(Trim(Txt(9).Text)) = 0 Then
      'LcPesquisaCli = False
      MsgBox "É Necessário Escolher o Cliente para  a Venda.", 64, "Aviso"
      Txt(9).SetFocus
   Else
      ValidaEntradaSintegra
   End If
End If
If Index = 5 Then
   If Len(Trim(Txt(7).Text)) = 0 Then
      LcPesquisaCli = False
      MsgBox "É Necessário Escolher o Vendedor Responsável.", 64, "Aviso"
      Txt(7).SetFocus
     End If
Else
  LcPesquisaCli = True
End If
If Index = 1 Then
   If Len(Trim(Txt(8).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Cliente para a Nota Fiscal.", 64, "Aviso"
      
   End If
End If

If Index = 9 Then LcAlteradoCliente = False
If Index = 2 Then
   LcAlteradoProduto = False
   ValidaEntradaSintegra
End If
If Index = 7 Then
  LcAlteradoFuncionario = False
  
End If

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Index = 9 Then LcAlteradoCliente = False
If KeyCode = 13 Then
   If Index <> 14 Then
      SendKeys "{TAB}"
   End If
End If
If KeyCode = 122 Then
   If Index <> 9 Then
      If Index <> 8 Then
        AbreOBS
        'Txt(17).SetFocus
        Exit Sub
      End If
   End If
   
End If
If Index <> 9 And Index <> 12 And Index <> 0 And _
Index <> 7 And Index <> 5 And Index <> 6 Then
     If KeyCode = 123 Then UltimasComprasCliente.Show , Me
End If
If KeyCode = 38 Then
   VoltaCampo (KeyCode)
End If
If KeyCode = 117 Then FrmDescicaoProduto.Show , Me


If KeyCode = 116 Then
   If Index = 8 Or Index = 9 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(9).Text)) > 0 Then
            FrmPesquisaCliente.Txt.Text = Txt(9).Text
            GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt(9).Text) & "*'  order by RAZAOSOC"
            Txt(2).SetFocus
         Else
            GlCriterioSql = ""
         End If
      Teclas (KeyCode)
   Else
      If Index = 1 Or Index = 2 Then 'Exibe Produtos
         GlEscolhe = 2
         If Len(Trim(Txt(2).Text)) > 0 Then
             GlCriterioSql = "select * From Produtos where nome like '" & UCase(Txt(2).Text) & "%' and Desativado=0 order by nome"
            FrmPesquisaProdutos.Txt.Text = Txt(2).Text
         Else
            GlCriterioSql = ""
         End If
         FrmPesquisaProdutos.Show , Me
      End If
    End If
Else
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then
    If Index = 8 Or Index = 9 Then
       AbreOBS
    Else
       Call Command4_Click
    End If
  End If
  If KeyCode = 121 Then SendKeys "%+{C}"
  Teclas (KeyCode)
End If

   
End Sub
Function VoltaCampo(LcIndex As Integer)
On Error Resume Next
Select Case LcIndex
   Case Is = 12
       Txt(0).SetFocus
   Case Is = 8
       Natureza.SetFocus
   Case Is = 9
       Txt(8).SetFocus
   Case Is = 1
       Txt(8).SetFocus
   Case Is = 2
      Txt(1).SetFocus
   Case Is = 4
     Txt(2).SetFocus
   Case Is = 3
     Txt(4).SetFocus
   Case Is = 5
     Txt(3).SetFocus
   Case Is = 6
     Txt(5).SetFocus
End Select

End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   If Index <> 12 And Index <> 7 And Index <> 14 Then Txt(Index).Text = ""
   LcLimpa = False
End If
If Index = 17 Then
   If KeyAscii = 46 Then KeyAscii = 44
End If
End Sub


Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 6 And GlLibera Then montagrid
If Index = 1 Then
   If Len(Trim(Txt(1).Text)) > 0 Then
      Txt(1).Text = Right("00000" & Txt(1).Text, 5)
      BuscaProduto (2)
   End If
End If

If Index = 2 Then BuscaProduto (2)

If Index = 4 Then calculaunitario

If Index = 3 Then VerificaDisponivel

If Index = 5 Then
   ConferePreco
End If
If Index = 17 Then CalculaDesconto
If Index = 9 Then
   GlPodeAbrirOBS = True
   If LcPesquisaCli And Len(Txt(9).Text) > 0 Then BuscaCliente (2)
End If
If Index = 7 Then BuscaVendendor (2)
If Index = 2 Then BuscaProduto (2)
If Index = 10 And Len(Trim(Txt(Index).Text)) <> 0 Then BuscaVendendor (2)

End Sub
Function VerificaDisponivel()

On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim LcCom As Long
Dim RsNota As ADODB.Recordset
'dim rsnidade As Recordset
'LcSqlUn = "Select * from alid004 where simbolo='" & Unidade.Text & "'"
LcSql = "Select * from Produtos where codigo=" & Txt(1).Text

If Len(Trim(Txt(3).Text)) = 0 Then Exit Function
If Len(Txt(4).Text) > 0 Then LcCom = CLng(Txt(4).Text) Else LcCom = 1
AbreBase
tipoBlQ.Text = 0
Set RsNota = AbreRecordset(LcSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set rsnidade = Dbbase.OpenRecordset(LcSqlUn, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'If rsnidade!cod = RsNota!UnidMedida Then
'   If Not RsNota.EOF Then
'      If RsNota!QuantEstoque < (ccur(Txt(3).text) * CDbl(ccur(Txt(3).text))) Then
         'MsgBox "Não Exite Quantidade Disponivel em Estoque." & Chr(13) & "A quantidade Atual é :" & RsNota!QuantEstoque & " " & Unidade.Text, 64, "Aviso"
         'Txt(3).Text = ""
         'Txt(3).SetFocus
'         tipoBlQ.Text = 2
'      End If
'   Else
      'MsgBox "Não Exite Quantidade Disponivel em Estoque." & Chr(13) & "A quantidade Atual é :0", 64, "Aviso"
      'Txt(3).SetFocus
'      tipoBlQ.Text = 2
'   End If
'Else
If Not IsNull(RsNota("QuantEstoque")) Then LcQuantEstoque = RsNota("QuantEstoque") Else LcQuantEstoque = 0
If LcQuantEstoque < (CDbl(Txt(3).Text) * CDbl(Txt(4).Text)) Then
   tipoBlQ.Text = 2
End If
'End If
'txt(0).Text = LcNumeroNota

End Function


Function ConferePreco()
On Error Resume Next
Dim Rs As ADODB.Recordset
tipoblP.Text = 0
Dim LcPreconovo, LcPRecoAntigo As Currency
Dim LcLcqM As Double
'GlLibera = False
If Len(Txt(1).Text) = o Then Exit Function
If Len(minimo.Text) = 0 Then minimo.Text = 0
'AbreBase
Set Rs = AbreRecordset("select * from Produtos where codigo=" & Txt(1).Text, True) ' & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not Rs.EOF Then
   If IsNull(Rs!QtdMedida) Then
      LcLcqM = 1
   Else
      If Rs!QtdMedida = 0 Then
          LcLcqM = 1
      Else
         LcLcqM = Rs!QtdMedida
      End If
   End If
   LcPRecoAntigo = CDbl(Rs!MinimoVenda) / LcLcqM
Else
  LcPRecoAntigo = 0
End If
If Len(Trim(valor(0).Text)) = 0 Then
   valor(0).Text = 0
End If
If Len(Txt(4).Text) = 0 Then Txt(4).Text = 1
If Txt(4).Text = "0" Then Txt(4).Text = 1
LcPreconovo = CCur(valor(0).Text) / CDbl(Txt(4).Text)
GlEscolha = True

If LcPreconovo < LcPRecoAntigo Then
    'Liberacao.Show
    'GlLibera = False
    ' GlEscolha = True
    ' Do Until Not GlEscolha
    '    DoEvents
    ' Loop
    ' If GlLibera Then
        Comissao.Text = 1
    ' Else
    '    valor(0) = LcPrecoVelho
    '    valor(0).SetFocus
    ' End If
    tipoblP.Text = 1
Else
  GlLibera = True
  If Len(Comissao.Text) = 0 Then Comissao.Text = 0
  If CLng(Comissao.Text) <> 1 Then
     Comissao.Text = 1.5
  End If
End If

End Function
Function ExcluiItem(LcNItem As Integer)
On Error Resume Next
Dim a, b As Integer

For a = 0 To LcTam - 1
    If Val(LcMat(a).Item) = LcNItem Then
       LcMat(a).CodPro = ""
       LcAchou = True
       'Exit For
       
    End If
Next
If Not LcAchou Then
   MsgBox "Item Não encontrado...", 48, "Aviso"
Else
  RemontaIndice
  EscreveGrid
End If
End Function
Function CalculaNumeroNota()
On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim RsNota As Recordset
If Len(Txt(0).Text) = 0 Then
   LcSql = "Select * from proposta order by NUMNF"
   AbreBase
   Set RsNota = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
   If RsNota.EOF Then
      LcNumeroNota = "000001"
   Else
      RsNota.MoveLast
      LcNumeroNota = Right("000000" & CStr(Val(RsNota("NUMNF")) + 1), 6)
   End If
   Txt(0).Text = LcNumeroNota

   RsNota.Close
   Dbbase.Close
   Set RsNota = Nothing
   Set Dbbase = Nothing
Else
   Txt(0).Text = Right("000000" & Txt(0).Text, 6)
End If
End Function

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  SendKeys "{TAB}"
Else
  If KeyCode = 122 Then
   AbreOBS
   'Txt(17).SetFocus
   Exit Sub
End If
  If KeyCode = 117 Then FrmDescicaoProduto.Show , Me
  If KeyCode = 123 Then UltimasComprasCliente.Show , Me
  
End If
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Function GeraComissao()
On Error Resume Next
Dim a As Integer
Dim RsComissao As Recordset
LcSql = "Select * from Alid201"
AbreBase
LcSql = "Select * from Alid201 where nf='" & Txt(0).Text & "'"
Set RsComissao = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)

Do Until RsComissao.EOF
   RsComissao.Delete
   RsComissao.MoveNext
Loop

RsComissao.Close
LcSql = "Select * from Alid201"
Set RsComissao = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Len(Txt(17).Text) > 0 Then
   LcPercDesc = CDbl(Txt(17).Text) / CDbl(Txt(16).Text)
Else
   LcPercDesc = 0
End If
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         RsComissao.AddNew
         RsComissao("Vendedor") = Txt(10).Text
         RsComissao("NF") = Txt(0).Text
         RsComissao("Produto") = LcMat(a).CodPro
         RsComissao("QUANTIDADE") = LcMat(a).Qut
         RsComissao("VALORUNIT") = LcMat(a).VUnit
         RsComissao("VALORTOTAL") = LcMat(a).Vtotal
         If Comissao.Text = "1" Then Ibaixo = True Else Ibaixo = False
         RsComissao("ITEMBAIXO") = Ibaixo
         
         If Len(Comissao.Text) = 0 Then Comissao.Text = "0"
         If Ibaixo Then
            RsComissao("COMISSAO") = (1 / 100) * (LcMat(a).Vtotal - (LcPercDesc * LcMat(a).Vtotal))
         Else
            RsComissao("COMISSAO") = (1.5 / 100) * (LcMat(a).Vtotal - (LcPercDesc * LcMat(a).Vtotal))
         End If
         RsComissao("DATAVENDA") = CDate(Txt(12).Text)
         RsComissao("CLIENTE") = Txt(8).Text
         RsComissao.Update
     End If
Next
RsComissao.Close
Dbbase.Close
Set RsComissao = Nothing
Set Dbbase = Nothing
         
End Function
Function SalvaNota() As Boolean
On Error GoTo ErrSalva
Dim RsNotaFiscal As Recordset, RsItens As Recordset
Dim rsCliente As Recordset, RsProduto As ADODB.Recordset
Dim RsEstoque As Recordset, RsG As Recordset
Dim Estoque As ControleDb
Dim LcCom, a As Long
Dim LcSaldoUnit As Double
Dim LcDesconto As String
Dim wrkDefault As Workspace
Dim Nome_Maquina As String
Dim LcInclusao As Boolean

'===> Verifica se o cliente tem informações financeiras
If GlMostraMsgClientePedido Then
   MostraDetalhesCliente
End If
'NomeMaquina
If Len(GlNomeMaquina) = 0 Then
  NomeMaquina
End If
Nome_Maquina = GlNomeMaquina
Set Estoque = New ControleDb
'CalculaNumeroNota
Set wrkDefault = DBEngine.Workspaces(0)
'==== Verifica pendencia
If Pendente.Value = 0 And LiberaFaturamento.Value = 1 Then LiberaPedido
'==== Grava Os dados da Nota Fiscal
If Len(Txt(0).Text) = 0 Then
   CalculaNumeroNota
   LcInclusao = True
Else
   LcInclusao = False
End If
'===> Exclui a Nota, se existir

AbreBase
wrkDefault.BeginTrans
'LcSq = "Delete from proposta where numnf='" & Txt(0).Text & "'"
'Dbbase.Execute LcSq
LcSq = "Delete from subproposta where numnf='" & Txt(0).Text & "'"
Dbbase.Execute LcSq

'End If
LcNatureza = Natureza.Text
If IsNumeric(Txt(15).Text) Then Txt(15).Text = CCur(Txt(15).Text)
Txt(15).Text = Replace(Txt(15).Text, ",", ".")

If IsNumeric(Txt(16).Text) Then Txt(16).Text = CCur(Txt(16).Text)

Txt(16).Text = Replace(Txt(16).Text, ",", ".")
If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
Txt(17).Text = Replace(Txt(17).Text, ",", ".")

If Len(GlMsg) > 0 And Len(GlMsg1) > 0 And Len(GlMsg2) > 0 Then
   LcOBS = GlMsg & Chr(10) & GlMsg1 & Chr(10) & GlMsg2
ElseIf Len(GlMsg) > 0 And Len(GlMsg1) > 0 Then
   LcOBS = GlMsg & Chr(10) & GlMsg1
ElseIf Len(GlMsg) > 0 Then
   LcOBS = GlMsg
End If
If InStr(Txt(14).Text, LcOBS) = 0 Then
    If Len(LcOBS) > 0 Then
        Txt(14).Text = Txt(14).Text & LcOBS
    End If
End If
LcDesconto = Txt(17).Text
If LcInclusao Then
        LcSq = "INSERT INTO proposta (numnf,dtemis,natureza,status,cliente,transp,tipotrans,placatrans,"
        LcSq = LcSq & "uftrans,CGCCPFTRAN,endtrans,munictrans,ufmunic,INSCEST,obs02,obs03,obs04,valorproduto,"
        LcSq = LcSq & "valornota,icms,formapag,ordemcompra,previsao,liberado,dias1,dias2,dias3"
        If IsDate(DadosTransp.Vencimento(0).Text) Then
           LcSq = LcSq & ",vencimento1"
        End If
        If IsDate(DadosTransp.Vencimento(1).Text) Then
            LcSq = LcSq & ",vencimento2"
        End If
        If IsDate(DadosTransp.Vencimento(2).Text) Then
            LcSq = LcSq & ",vencimento3"
        End If
        LcSq = LcSq & ",Vendedor,Desconto,bloqueado"
        
        If IsDate(DataLiberacao.Text) Then
           LcSq = LcSq & ",dataliberacao"
        End If
        LcSq = LcSq & ",jaEsteveBloqueado,MaquinaLiberacao,HoraLiberacao,pendente,Usuario,EnderecoEntrega,Oc"
        LcSq = LcSq & ",Maquina, obs,CondPag,prazo,Validade,VendedorImprimir,InfComplementar)values('"
        LcSq = LcSq & Txt(0).Text & "','" & Format(Txt(12).Text, "yyyy-mm-dd") & "','" & LcNatureza & "','EMITIDA','"
        LcSq = LcSq & Txt(8).Text & "','" & DadosTransp.Txt(0).Text & "','" & Mid(DadosTransp.Tipo.Text, 1, 1) & "','"
        LcSq = LcSq & DadosTransp.Placa.Text & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(1).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(2).Text) & "','"
        LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.Txt(3).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(4).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(5).Text) & "','"
        LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.Txt(6).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(7).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(8).Text) & "','"
        LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.Txt(9).Text) & "'," & Txt(15).Text & "," & Txt(16).Text & ",'" & Txt(5).Text & "','"
        LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.TipoMonetario.Text) & "','" & Estoque.RetiraCaracter(Ordem.Text) & "','"
        LcSq = LcSq & Format(Previsao.Text, "YYYY-MM-DD") & "'," & CInt(LiberaFaturamento.Value) & ",'"
        LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.Dias(0).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Dias(1).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Dias(2).Text) & "',"
        
        If IsDate(DadosTransp.Vencimento(0).Text) Then
            LcSq = LcSq & "'" & Format(DadosTransp.Vencimento(0).Text, "yyyy-mm-dd") & "',"
        End If
        If IsDate(DadosTransp.Vencimento(1).Text) Then
          LcSq = LcSq & "'" & Format(DadosTransp.Vencimento(1).Text, "yyyy-mm-dd") & "',"
        End If
        If IsDate(DadosTransp.Vencimento(2).Text) Then
           LcSq = LcSq & "'" & Format(DadosTransp.Vencimento(2).Text, "yyyy-mm-dd") & "',"
        End If
        LcSq = LcSq & "'" & Right("00000" & Txt(10).Text, 5) & "'," & LcDesconto & ","
        LcSq = LcSq & bloqueado.Text & ","
        If IsDate(DataLiberacao.Text) Then
            LcSq = LcSq & "'" & Format(DataLiberacao.Text, "yyyy-mm-dd") & "',"
        End If
        LcSq = LcSq & JaBloqueado.Text & ",'" & MaquinaLiberacao.Text & "','" & HoraLiberacao.Text & "',"
        LcSq = LcSq & IIf(GlNaoVerificaEstoque, 0, Pendente.Value) & ",'" & usuario.Text & "'"
        LcSq = LcSq & ",'" & DadosTransp.Txt(10).Text & "','" & DadosTransp.Txt(11).Text & "'"
        LcSq = LcSq & ",'" & Nome_Maquina & "'"
        LcSq = LcSq & ",'" & Replace(Txt(14).Text, "'", "''") & "'"
        LcSq = LcSq & ",'" & Replace(CondPag.Text, "'", "''") & "'"
        LcSq = LcSq & ",'" & Replace(PrazoEntrega.Text, "'", "''") & "'"
        LcSq = LcSq & ",'" & Replace(ValidadeCotacao.Text, "'", "''") & "'"
        LcSq = LcSq & ",'" & Replace(VendedorImprimir.Text, "'", "''") & "'"
        LcSq = LcSq & ",'" & Replace(InformacoesComplementares.Text, "'", "''") & "')"

Else
        LcSq = "Update proposta Set "
        LcSq = LcSq & "natureza='" & LcNatureza & "',"
        LcSq = LcSq & "status='EMITIDA',"
        LcSq = LcSq & "cliente='" & Txt(8).Text & "',"
        LcSq = LcSq & "transp='" & DadosTransp.Txt(0).Text & "',"
        LcSq = LcSq & "tipotrans='" & Mid(DadosTransp.Tipo.Text, 1, 1) & "',"
        LcSq = LcSq & "placatrans='" & DadosTransp.Placa.Text & "',"
        LcSq = LcSq & "uftrans='" & Estoque.RetiraCaracter(DadosTransp.Txt(1).Text) & "',"
        LcSq = LcSq & "CGCCPFTRAN='" & Estoque.RetiraCaracter(DadosTransp.Txt(2).Text) & "',"
        LcSq = LcSq & "endtrans='" & Estoque.RetiraCaracter(DadosTransp.Txt(3).Text) & "',"
        LcSq = LcSq & "munictrans='" & Estoque.RetiraCaracter(DadosTransp.Txt(4).Text) & "',"
        LcSq = LcSq & "ufmunic='" & Estoque.RetiraCaracter(DadosTransp.Txt(5).Text) & "',"
        LcSq = LcSq & "INSCEST='" & Estoque.RetiraCaracter(DadosTransp.Txt(6).Text) & "',"
        LcSq = LcSq & "obs02='" & Estoque.RetiraCaracter(DadosTransp.Txt(7).Text) & "',"
        LcSq = LcSq & "obs03='" & Estoque.RetiraCaracter(DadosTransp.Txt(8).Text) & "',"
        LcSq = LcSq & "obs04='" & Estoque.RetiraCaracter(DadosTransp.Txt(9).Text) & "',"
        LcSq = LcSq & "valorproduto=" & Txt(15).Text & ","
        LcSq = LcSq & "valornota=" & Txt(16).Text & ","
        LcSq = LcSq & "icms='" & Txt(5).Text & "',"
        LcSq = LcSq & "formapag='" & Estoque.RetiraCaracter(DadosTransp.TipoMonetario.Text) & "',"
        LcSq = LcSq & "ordemcompra='" & Estoque.RetiraCaracter(Ordem.Text) & "',"
        LcSq = LcSq & "previsao='" & Format(Previsao.Text, "YYYY-MM-DD") & "',"
        LcSq = LcSq & "liberado=" & CInt(LiberaFaturamento.Value) & ","
        LcSq = LcSq & "dias1='" & Estoque.RetiraCaracter(DadosTransp.Dias(0).Text) & "',"
        LcSq = LcSq & "dias2='" & Estoque.RetiraCaracter(DadosTransp.Dias(1).Text) & "',"
        LcSq = LcSq & "dias3='" & Estoque.RetiraCaracter(DadosTransp.Dias(2).Text) & "',"
             
        If IsNumeric(Txt(16).Text) Then
        
        Else
        
        End If
        If IsDate(DadosTransp.Vencimento(0).Text) Then
            LcSq = LcSq & "vencimento1='" & Format(DadosTransp.Vencimento(0).Text, "yyyy-mm-dd") & "',"
        End If
        If IsDate(DadosTransp.Vencimento(1).Text) Then
          LcSq = LcSq & "vencimento2='" & Format(DadosTransp.Vencimento(1).Text, "yyyy-mm-dd") & "',"
        End If
        If IsDate(DadosTransp.Vencimento(2).Text) Then
           LcSq = LcSq & "vencimento3='" & Format(DadosTransp.Vencimento(2).Text, "yyyy-mm-dd") & "',"
        End If
        LcSq = LcSq & "Vendedor='" & Right("00000" & Txt(10).Text, 5) & "',"
        LcSq = LcSq & "Desconto=" & LcDesconto & ","
        LcSq = LcSq & "bloqueado=" & bloqueado.Text & ","
        If IsDate(DataLiberacao.Text) Then
            LcSq = LcSq & "dataliberacao='" & Format(DataLiberacao.Text, "yyyy-mm-dd") & "',"
        End If
        LcSq = LcSq & "jaEsteveBloqueado=" & JaBloqueado.Text & ","
        LcSq = LcSq & "MaquinaLiberacao='" & MaquinaLiberacao.Text & "',"
        LcSq = LcSq & "HoraLiberacao='" & HoraLiberacao.Text & "',"
        LcSq = LcSq & "pendente=" & IIf(GlNaoVerificaEstoque, 0, Pendente.Value) & ","
        LcSq = LcSq & "Usuario='" & usuario.Text & "',"
        LcSq = LcSq & "EnderecoEntrega='" & DadosTransp.Txt(10).Text & "',"
        LcSq = LcSq & "Oc='" & DadosTransp.Txt(11).Text & "',"
        LcSq = LcSq & "obs='" & Replace(Txt(14).Text, "'", "''") & "',"
        LcSq = LcSq & "CondPag='" & Replace(CondPag.Text, "'", "''") & "',"
        LcSq = LcSq & "Prazo='" & Replace(PrazoEntrega.Text, "'", "''") & "',"
        LcSq = LcSq & "Validade='" & Replace(ValidadeCotacao.Text, "'", "''") & "',"
        LcSq = LcSq & "VendedorImprimir='" & Replace(VendedorImprimir.Text, "'", "''") & "',"
        LcSq = LcSq & "InfComplementar='" & Replace(InformacoesComplementares.Text, "'", "''") & "',"
        LcSq = LcSq & "Maquina='" & Nome_Maquina & "'"
        'LcSq = LcSq & " where codigo=" & Txt(0).Text
        LcSq = LcSq & " where NUMNF='" & Txt(0).Text & "'"
End If
'==> Inclui os dados da proposta

Debug.Print LcSq
Dbbase.Execute LcSq, Processados
If LcInclusao Then
    Dim RsProposta As DAO.Recordset
    Set RsProposta = Dbbase.OpenRecordset("Select top 1 codigo From proposta where Maquina='" & Nome_Maquina & "' order by codigo desc")
    If Not RsProposta.EOF Then
       'Dim NumeroNFe As String
       'NumeroNFe = RsProposta!codigo
       'Txt(0).Text = NumeroNFe
       'LcSq = "Update proposta set NUMNF='" & NumeroNFe & "' where codigo=" & RsProposta!codigo
       'Dbbase.Execute LcSq, Processados
    End If
End If

LcSql2 = "Select * from subproposta where NUMNF = '" & Txt(0).Text & "'"
err.Number = 0
For a = 0 To LcTam - 1
    If Len(LcMat(a).Com) > 0 Then
       LcCom = LcMat(a).Com
    Else
      LcCom = 1
    End If
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         '==> Recupera o NCM
         Dim Str_Sql As String
         Dim LcNCM As String
         Str_Sql = " Select classificacaofiscal from Produtos where codigo=" & LcMat(a).CodPro
         Set RsProduto = AbreRecordset(Str_Sql, True)
         If Not RsProduto.EOF Then
            LcNCM = RsProduto!ClassificacaoFiscal
         End If
         LcQut = Replace(CStr(LcMat(a).Qut), ",", ".")
         LcVUnit = Replace(CStr(LcMat(a).VUnit), ",", ".")
         LcCom = Replace(CStr(LcMat(a).Com), ",", ".")
         LcMat(a).produto = Replace(LcMat(a).produto, "'", "")
         LcMat(a).produto = Replace(LcMat(a).produto, Chr(43), "")
        
        LcTipoLi = LcMat(a).tipoliberacao
        If LcMat(a).jaEsteveBloqueado Then LcJaBlo = -1 Else LcJaBlo = 0
        If LcMat(a).bloqueado Then LcBlo = -1 Else LcBlo = 0
         'Call GeraHistorico(LcMat(a).CodPro, LcMat(a).produto, Txt(0).Text, "E", CDate(Txt(12).Text), LcMat(a).santamaria, LcMat(a).santamaria1, CLng(LcMat(a).california), 0, 0, 0)
         LcSq = "INSERT INTO subproposta (numnf,item,codprod,qtde,valunit,unimed,QTDUM,"
         LcSq = LcSq & "descricao,bloqueado,tipoliberacao,jaestevebloqueado,maquinaliberacao,"
         If IsDate(LcMat(a).DataLiberacao) Then
             'RsItens("DataLiberacao") = LcMat(a).DataLiberacao
             LcSq = LcSq & "DataLiberacao,"
         End If
         LcSq = LcSq & "HoraLiberacao,usuario,NCM) values ('"
         LcSq = LcSq & Txt(0).Text & "','" & Right("00" & LcMat(a).Item, 2) & "','"
         LcSq = LcSq & LcMat(a).CodPro & "'," & LcQut & "," & LcVUnit & ",'"
         LcSq = LcSq & LcMat(a).Und & "'," & LcCom & ",'" & LcMat(a).produto & "',"
         LcSq = LcSq & LcBlo & "," & LcTipoLi & "," & LcJaBlo & ",'"
         LcSq = LcSq & LcMat(a).MaquinaLiberacao & "'"
         If IsDate(LcMat(a).DataLiberacao) Then
            LcSq = LcSq & ",'" & Format(LcMat(a).DataLiberacao, "yyyy-mm-dd") & "'"
         End If
         LcSq = LcSq & ",'" & LcMat(a).HoraLiberacao & "','" & LcMat(a).usuario & "'"
         LcSq = LcSq & ",'" & LcNCM & "')"
         Dbbase.Execute LcSq
         
         '==> Aqui baixa o Estoque, So Baixa se tiver liberado para baixar
         If GlBaixarEstoquenoPedido Then
             If LiberaFaturamento.Value = 1 Then
                If Pendente.Value = 0 Then
                   If faturado.Value = 0 Then
                       Estoque.CodProduto = LcMat(a).CodPro
                       Estoque.CodClien_forn = Txt(8).Text
                       Estoque.NF = Txt(0).Text
                      '==> Efetua a Baixa no Estoque
                       LcComentario = "Atualiza o saldo em Estoque "
                        'Call BaixaEstoque(LcMat(a).CodPro, CDbl(LcMat(a).Qut), CDbl(LcMat(a).com), LcMat(a).Und)
                        LcQSanta = (Estoque.Santa1Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa1Unitario
                        LcQSanta1 = (Estoque.Santa2Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa2Unitario
                        LcqCanifornia = (Estoque.QuantidadeCaliforniaFechado * Estoque.QuantidadeDaUnidade) + Estoque.QuantidadeCaliforniaUnitario
                        
                        'Call BaixaPorNota(LcMat(a).CodPro, CDbl(LcMat(a).QuanTidadeBaixa), CDbl(LcMat(a).Com), LcMat(a).Und, CStr(LcMat(a).Com))
                        If LcMat(a).Qut > 0 Then
                           If Not Estoque.BaixaEstoque(CDbl(LcMat(a).Qut), CDbl(LcMat(a).VUnit), LcMat(a).Und, CDbl(LcMat(a).Com)) Then
                              err.Raise vbObjectError + 513, "Nâo foi efetuada a Atualização.", "Atualização de Estoque do item " & LcMat(a).CodPro & "Não foi Realizada."
                              GoTo ErrSalva
                           End If
                           LcQSanta = LcQSanta - ((Estoque.Santa1Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa1Unitario)
                           LcQSanta1 = LcQSanta1 - ((Estoque.Santa2Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa2Unitario)
                           LcqCanifornia = LcqCanifornia - ((Estoque.QuantidadeCaliforniaFechado * Estoque.QuantidadeDaUnidade) + Estoque.QuantidadeCaliforniaUnitario)
            
                           LcComentario = "Gravando Historico"
                           LcSq = "insert into HistoricoProduto (produto,descricao,santa,santa2,california,nf,data,tipo,unidade,ClienteForn,CodUnid) values ('"
                           LcSq = LcSq & Estoque.CodProduto & "','" & Estoque.RetiraCaracter(Estoque.DescricaoProduto) & "'," & LcQSanta & "," & LcQSanta1 & "," & LcqCanifornia
                           LcSq = LcSq & ",'" & Estoque.NF & "','" & Format(Txt(12).Text, "yyyy-mm-dd") & "','S','" & LcMat(a).Und & "','" & Estoque.RetiraCaracter(Txt(9).Text) & "','" & LcCodLancamento & "')"
                        ' MsgBox LcSq
                         
                           total = ExecutaSql(LcSq)
                           If Len(LcMat(a).NumeroVale) > 0 Then
                              LcSql = "Update HistoricoProduto Set Tipo='S',CodUnid='" & LcCodLancamento & "',nf='" & Estoque.NF & "' where Tipo='V' and nf='" & LcMat(a).NumeroVale & "' And Produto='" & LcMat(a).CodPro & "'"
                              'MsgBox LcSql
                              total = ExecutaSql(LcSql)
                           End If
                        Else
                           '===> Temos o Vale, então vamos mudar o tipo no historico.
                           '===> Alteramos o historico para representar a nota no historico
                           LcSql = "Update HistoricoProduto Set tipo='S',CodUnid='" & LcCodLancamento & "',nf='" & Txt(0).Text & "' where Tipo='V' and nf='" & LcMat(a).NumeroVale & "' And Produto='" & LcMat(a).CodPro & "'"
                           'MsgBox LcSql
                           total = ExecutaSql(LcSql)
                        
                        End If
                   End If
                End If
             End If
         End If
     End If
Next
wrkDefault.CommitTrans
SalvaNota = True
ConfirmaOrcamento.Show , Me
'Me.ControlBox = True
'=== Fecha as Bases
Dbbase.Close
Set Dbbase = Nothing
Exit Function
ErrSalva:
'MsgBox err.Description '& err.Number
'Resume 0
wrkDefault.Rollback
SalvaNota = False
MsgBox "Foram Encontrados erros durante o processamento do Pedido.", 64, err.Description & err.Number
'Resume 0
End Function

Private Sub Unidade_LostFocus()
On Error Resume Next
Dim a As Integer
For a = 0 To LcQUn
    If MtUnidade(a).Simbolo = Unidade.Text Then
       If MtUnidade(a).quantidade <> 0 Then
          Txt(4).Text = MtUnidade(a).quantidade
       End If
       Exit For
    End If
Next
End Sub

Private Sub Validade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = 13 Then SendKeys "{TAB}"
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"


End Sub

Private Sub valor_Change(Index As Integer)
On Error Resume Next
If Not LcLimpaValor Then CalculaValores
'CalculaValores
End Sub

Private Sub valor_GotFocus(Index As Integer)
On Error Resume Next
If Index = 0 Then VerificaDisponivel
LcLimpa = True
If Index = 1 Then
   LcFechaitem = True
End If
End Sub

Private Sub valor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 122 Then
   AbreOBS
   'Txt(17).SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
    SendKeys "{TAB}"
Else
    Teclas (KeyCode)
    LcCalculado = False
    If KeyCode = 123 Then UltimasComprasCliente.Show , Me
End If
On Error Resume Next
  If KeyCode = 27 Then Me.Visible = False
  If KeyCode = 113 Then SendKeys "%+{G}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 119 Then SendKeys "%+{I}"
  If KeyCode = 118 Then SendKeys "%+{X}"
  If KeyCode = 120 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub valor_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   valor(Index).Text = ""
   LcLimpa = False
End If
If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub valor_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Index = 1 Then
   If KeyCode = 38 Then
       LcFechaitem = False
       valor(0).SetFocus
   End If
End If

End Sub

Private Sub valor_LostFocus(Index As Integer)
On Error Resume Next
GlLibera = True
If Index = 0 Then ConferePreco
If Index = 1 And GlLibera Then
   montagrid
End If
End Sub
Sub MostraDetalhesCliente()
Dim Rs As Recordset
Dim db As Database
Dim RsReceita As ADODB.Recordset
Dim LcValorVendido As Currency
Dim ValorDisponivel As Currency
Dim LcMsg As String
Set db = OpenDatabase(GLBase)
Debug.Print Txt(8).Text
If Len(Txt(8).Text) = 0 Then Exit Sub
'Set Rs = db.OpenRecordset("Select * from alid001 where codigo='" & Txt(8).Text & "'")
'If Not Rs.EOF Then
'   If IsNumeric(Txt(16).Text) Then
'       LcValorVendido = CCur(Txt(16).Text)
'        ValorDisponivel = CCur(Rs!LimiteCredito) - Rs!CreditoUtilizado
'        If (ValorDisponivel - LcValorVendido) < 0 Then
'            LcMsg = "Limite de Crédito do cliente é insuficiente!" & Chr(13) & "Limite disponivel:" & FormatNumber(ValorDisponivel, 2) & Chr(13) & "ValorCompra:" & FormatNumber(LcValorVendido, 2) & Chr(13) & Chr(13) & "Venda Somente a Vista"
'        End If
'   End If
'End If

Set RsReceita = AbreRecordset("Select * from alid015 where cliente='" & Right("00000" & Txt(8).Text, 5) & "' and DTVENC<#" & Format(Date, "mm/dd/yy") & "# and VALPAGO=0;")
Dim LcValorAtraso As Currency
Do Until RsReceita.EOF
   LcValorAtraso = LcValorAtraso + RsReceita!valor
   RsReceita.MoveNext
Loop
If LcValorAtraso > 0 Then
   If Len(LcMsg) > 0 Then LcMsg = LcMsg & Chr(13) & Chr(13)
   LcMsg = LcMsg & "O cliente possui  " & FormatNumber(LcValorAtraso, 2) & " em atraso!"
End If
If Len(LcMsg) > 0 Then
   Load FrmMostraMsgAtrado
   FrmMostraMsgAtrado.Msg.Caption = LcMsg
   FrmMostraMsgAtrado.Show (1)
End If
End Sub
Function ValidaEntradaSintegra() As Boolean
On Error GoTo errorVali
Dim Rs As Recordset
Dim db As Database
Dim Estado As String
Dim CNPJ As String
Dim Inscricao As String

Set db = OpenDatabase(GLBase)
Debug.Print Txt(8).Text
If Len(Txt(8).Text) = 0 Then Exit Function
Set Rs = db.OpenRecordset("Select * from alid001 where codigo='" & Txt(8).Text & "'")

If Rs.EOF Then
   ValidaEntradaSintegra = False
   MsgBox "Cliente não encontrado.", 64, "Aviso"
Else
   '==> Verifica o Estado do Fornecedor
   If IsNull(Rs!Estado) Then
      Estado = ""
   Else
      Estado = UCase(Rs!Estado)
   End If
   If Len(Estado) = 0 Then
      ValidaEntradaSintegra = False
      MsgBox "O Estado do Cliente não foi cadastrado." & Chr(13) & "cadastre-o antes de emitir a nota fiscal.", 64, "Aviso"
      GoTo Saida
   Else
     '==> Verifica se o Cfop é Valido
     If Estado = "MG" Then
        If Mid(CFOP.Text, 1, 1) <> "5" Then
           MsgBox "O CFOP é invalido para clientes do estado de MG.", 64, "Aviso"
           CFOP.SetFocus
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     Else
        If Mid(CFOP.Text, 1, 1) <> "6" Then
           MsgBox "O CFOP é invalido para clientes do fora do estado de MG.", 64, "Aviso"
           CFOP.SetFocus
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     End If
     '==> Valida o cnpj
     CNPJ = Rs!CGC & ""
     CNPJ = Replace(CNPJ, ",", "")
     CNPJ = Replace(CNPJ, ".", "")
     CNPJ = Replace(CNPJ, "-", "")
     CNPJ = Replace(CNPJ, "/", "")
     CNPJ = Replace(CNPJ, "\", "")
     CNPJ = Replace(CNPJ, " ", "")
     CNPJ = Trim(CNPJ)
     If Len(CNPJ) = 0 Then
        CNPJ = Rs!cpf & ""
        CNPJ = Replace(CNPJ, ",", "")
        CNPJ = Replace(CNPJ, ".", "")
        CNPJ = Replace(CNPJ, "-", "")
        CNPJ = Replace(CNPJ, "/", "")
        CNPJ = Replace(CNPJ, "\", "")
        CNPJ = Replace(CNPJ, " ", "")
        CNPJ = Trim(CNPJ)
     End If
     Inscricao = Rs!INSCEST & ""
     Inscricao = Replace(Inscricao, ",", "")
     Inscricao = Replace(Inscricao, ".", "")
     Inscricao = Replace(Inscricao, "-", "")
     Inscricao = Replace(Inscricao, "/", "")
     Inscricao = Replace(Inscricao, "\", "")
     Inscricao = Replace(Inscricao, " ", "")
     Inscricao = Trim(Inscricao)
     If Len(CNPJ) = 0 Then
        MsgBox "O CNPJ / CPF do cliente não foi cadastrado.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
     If Len(CNPJ) > 11 Then
        If Not Calc_CNPJ(CNPJ) Then
           MsgBox "O CNPJ do cliente é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     Else
        If Not Calc_CPF(CNPJ) Then
           MsgBox "O CPF do cliente é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     End If
     '==> Verifica a Inscricao estadual
     'If Len(Inscricao) = 0 Then
    '    MsgBox "A inscrição Estadual do cliene não foi cadastrada." & Chr(13) & "Caso ele não possua inscrição estadual ou seje pessoa física, casatre como ISENTO.", 64, "Aviso"
     '   ValidaEntradaSintegra = False
    '    GoTo Saida
    ' End If
    ' If Consiste(Inscricao, Estado) <> 0 Then
     '   MsgBox "A Inscrição Estadual do cliente é invalida.", 64, "Aviso"
     '   ValidaEntradaSintegra = False
     '   GoTo Saida
     'End If
   End If
End If
ValidaEntradaSintegra = True
Saida:
Set Rs = Nothing

Exit Function
errorVali:
ValidaEntradaSintegra = False
GoTo Saida
End Function
Sub VerificaLucratividade()
On Error Resume Next
Dim RsL     As ADODB.Recordset
Dim LcSql   As String
Dim LcCusto As Double
Dim LcCustoBase As Double
Dim LcComBase As Double
Dim LcLucro As Double
Dim a       As Integer
Dim StrSql As String
'======BuscaPrecoCusto
LcCusto = 0
For a = 0 To UBound(LcMat)
   If Len(LcMat(a).CodPro) > 0 Then
     StrSql = "Select * from produtos where codigo=" & LcMat(a).CodPro
     Set RsL = AbreRecordset(StrSql, True)
     If Not RsL.EOF Then
        LcComBase = RsL!QtdMedida
     Else
        LcComBase = 1
     End If
     LcCustoBase = (CCur(LcMat(a).Venda1) / LcComBase) * LcMat(a).Com * LcMat(a).Qut
     LcCusto = CCur(LcCusto) + LcCustoBase
   End If
Next
LcLucro = CCur(Txt(16).Text) - CCur(LcCusto)
LcLucro = (LcLucro * 100) / CCur(Txt(16).Text) 'LcCusto
Lucratividade.Text = AcertaNumero(CCur(LcLucro), 3)
End Sub

Function VerificaDisponivelGrid(LcCodProduto As String, LcQuantidade As Double, LcComG As Double) As Double
On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim LcCom As Long
Dim RsNota As ADODB.Recordset
LcSql = "Select * from Produtos where codigo=" & LcCodProduto
AbreBase
VerificaDisponivelGrid = 0
Set RsNota = AbreRecordset(LcSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not IsNull(RsNota("QuantEstoque")) Then LcQuantEstoque = RsNota("QuantEstoque") Else LcQuantEstoque = 0
If LcQuantEstoque < (CDbl(LcQuantidade) * CDbl(LcComG)) Then
   VerificaDisponivelGrid = 2
End If
End Function

Function ConferePrecoGrid(LcCodProduto As String, LcValor As Double, LcComG As Double) As Long
On Error Resume Next
Dim Rs As ADODB.Recordset
ConferePrecoGrid = 0
Dim LcPreconovo, LcPRecoAntigo As Currency
Dim LcLcqM As Double
Set Rs = AbreRecordset("select * from Produtos where codigo=" & LcCodProduto, True)
If Not Rs.EOF Then
   If IsNull(Rs!QtdMedida) Then
      LcLcqM = 1
   Else
      If Rs!QtdMedida = 0 Then
          LcLcqM = 1
      Else
         LcLcqM = Rs!QtdMedida
      End If
   End If
   LcPRecoAntigo = CDbl(Rs!MinimoVenda) / LcLcqM
Else
  LcPRecoAntigo = 0
End If
LcPreconovo = CCur(LcValor) / CDbl(LcComG)
GlEscolha = True

If LcPreconovo < LcPRecoAntigo Then
 '   Comissao.Text = 1
    ConferePrecoGrid = 1
Else
  GlLibera = True
  If CLng(Comissao.Text) <> 1 Then
'     Comissao.Text = 1.5
  End If
End If
End Function
