VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Orcamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orçamento e Vendas"
   ClientHeight    =   7740
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11880
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
   ScaleHeight     =   7740
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   10080
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox DescontoGerado 
      Height          =   285
      Left            =   10920
      TabIndex        =   72
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox DEscontoItem 
      Height          =   285
      Left            =   8640
      TabIndex        =   39
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox ComissaoFabricante 
      Height          =   285
      Left            =   2400
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Representada 
      Height          =   285
      Left            =   8040
      TabIndex        =   67
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Industria 
      Height          =   285
      Left            =   7560
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox DadosComplementares 
      Height          =   285
      Left            =   2880
      TabIndex        =   65
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox FoneTransp 
      Height          =   285
      Left            =   4320
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Transportadora 
      Height          =   285
      Left            =   5400
      TabIndex        =   63
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox ipi 
      Height          =   285
      Left            =   360
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox codigounidade 
      Height          =   285
      Left            =   1800
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox acrescimo 
      Height          =   285
      Left            =   8400
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox preconormal 
      Height          =   285
      Left            =   7560
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox total 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Unitario 
      Height          =   285
      Left            =   7320
      TabIndex        =   38
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Quantidade 
      Height          =   285
      Left            =   6240
      TabIndex        =   37
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox ComissaoProduto 
      Height          =   285
      Left            =   6600
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox LimiteUtilizado 
      Height          =   285
      Left            =   4680
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox LimiteCredito 
      Height          =   285
      Left            =   3120
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSMask.MaskEdBox ValorAcrescimo 
      Height          =   495
      Left            =   7980
      TabIndex        =   48
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DescricaoAcrescimo 
      Height          =   495
      Left            =   6525
      TabIndex        =   47
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ValorDesconto 
      Height          =   495
      Left            =   5070
      TabIndex        =   46
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin VB.TextBox DescricaoDesconto 
      Height          =   495
      Left            =   3135
      TabIndex        =   45
      Top             =   7080
      Width           =   1935
   End
   Begin MSMask.MaskEdBox TotalIpi 
      Height          =   495
      Left            =   1680
      TabIndex        =   44
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   16776960
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TotalOrcamento 
      Height          =   495
      Left            =   9480
      TabIndex        =   49
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TotalProduto 
      Height          =   495
      Left            =   120
      TabIndex        =   43
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   16777215
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   3375
      Left            =   120
      TabIndex        =   42
      Top             =   3360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   16
      BackColor       =   -2147483624
   End
   Begin VB.TextBox Comissao 
      Height          =   285
      Left            =   5640
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Unidade 
      Height          =   315
      Left            =   5640
      TabIndex        =   35
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox NomeProduto 
      Height          =   285
      Left            =   1560
      TabIndex        =   34
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox CodigoProduto 
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   33
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox NomeCliente 
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   1320
      Width           =   7455
   End
   Begin VB.TextBox CodigoCliente 
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox NomeVendedor 
      Height          =   285
      Left            =   2400
      TabIndex        =   29
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox CodigoVendedor 
      Height          =   285
      Left            =   1320
      TabIndex        =   28
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Status 
      Height          =   375
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Natureza 
      Height          =   315
      ItemData        =   "orcamento.frx":0000
      Left            =   6000
      List            =   "orcamento.frx":000D
      TabIndex        =   26
      Text            =   "Venda"
      Top             =   420
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Emissao 
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Documento 
      Height          =   375
      Left            =   600
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pes&quisa F7"
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
      TabIndex        =   53
      Top             =   1116
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar &Pedido F3"
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
      TabIndex        =   50
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir Item F4"
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
      TabIndex        =   52
      Top             =   744
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Gravar  F2"
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
      TabIndex        =   51
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Sair F10"
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
      TabIndex        =   55
      Top             =   1485
      Width           =   1575
   End
   Begin VB.TextBox Acomodacao 
      Height          =   285
      Left            =   5520
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
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
      Height          =   195
      Index           =   18
      Left            =   8520
      TabIndex        =   71
      Top             =   1920
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Ver as Últimas Compras Do Cliente, Pressione F12."
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
      Left            =   120
      TabIndex        =   69
      Top             =   3000
      Width           =   3960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código Representada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6480
      TabIndex        =   68
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pedido de Cliente"
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
      Left            =   600
      TabIndex        =   54
      Top             =   0
      Width           =   2520
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   7560
      TabIndex        =   23
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Alterar o Valor da Comissão no Item que está Sendo Digitado, Pressione F11"
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
      TabIndex        =   22
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para dar Acréscimo pressione F9"
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
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   2325
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acréscimo R$"
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
      Index           =   16
      Left            =   8040
      TabIndex        =   20
      Top             =   6840
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total IPI"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   15
      Left            =   1680
      TabIndex        =   19
      Top             =   6840
      Width           =   750
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Detalhar o Produto Pressione F8"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   17
      Top             =   2520
      Width           =   3225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para dar Desconto Geral pressione F6"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   2520
      Width           =   2700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Orçamento"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   9480
      TabIndex        =   14
      Top             =   6840
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Produtos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acrecimo %"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   6840
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto %"
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
      Index           =   10
      Left            =   3120
      TabIndex        =   11
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto R$"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   6840
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Produto pressione F5"
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
      TabIndex        =   8
      Top             =   2520
      Width           =   2925
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Vendedor ou  Cliente  pressione F5"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   3900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   645
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   0
      X2              =   9960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   9960
      X2              =   9960
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      TabIndex        =   5
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      TabIndex        =   4
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   5
      Left            =   10320
      TabIndex        =   3
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   4
      Left            =   7200
      TabIndex        =   2
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unid."
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
      Left            =   5520
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   2
      Left            =   6480
      TabIndex        =   0
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AC"
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
      Left            =   5640
      TabIndex        =   32
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Orcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===Seta Vaiaveis para Controlar o Fluxo das Informações

Private LcBuscaVendedor, LcBuscaCliente As Integer
Private LcCalculatotal, LcCalculaIpi, LcCalculaDesconto, LcLiberaCalculo As Integer
Private LcCalculaAcrescimo, LcGeraComissao, LcMontaComissao As Integer
Private LcItem, LcTamanhoPedido, LcLinhaAtual As Long
Private LcPerguntaPRoduto, LcLimpa, LcFechaitem As Integer
Private MtPedido(1000)
Private RsClientes As Recordset, RsI As Recordset
Private LcMargem As String
Private Lcarq   As Integer
Private az, a, b As Long
Private LcLinha As Long
Private LcTotalIten As Long
Private LcItensImpressos As Long

Function cabecalhoMeia()
On Error Resume Next
Dim RsClientes As Recordset
Dim RsEmpresa As Recordset
Dim RsCidade    As Recordset

Dim LcCpf As String
Dim LcCgc As String
Dim LcRua As String
Dim LcBairro As String
Dim LcCidade As String
Dim LcEstado As String
Dim LcCep As String

LigaTitulo = Chr(27) & Chr(87) & Chr(1) & Chr(27) & Chr(71) & Chr(27) & Chr(69)
DesligaTitulo = Chr(27) & Chr(87) & Chr(0) & Chr(27) & Chr(72) & Chr(27) & Chr(70)
LigaNegrito = Chr(27) & Chr(71)
DesligaNegrito = Chr(27) & Chr(72)
LigaDraft = Chr(27) & "x" & Chr(0)

LcSql = "Select * from alid001 where codigo='" & CodigoCliente.Text & "'"
AbreBase

Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'LcProcInd = "codigo='" & Industria.Text & "'"
'If Not Rsforn.EOF Then
'    Rsforn.FindFirst LcProcInd
'    If Not Rsforn.NoMatch Then
'       LcIndustria = Rsforn!Fantasia
'    Else
'       LcIndustria = ""
'    End If
'Else
  LcIndustria = ""
'End If

'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LCCidadeEm = RsEmpresa!cidade
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!bairro
   LcFone = RsEmpresa!fone
   If Len(RsEmpresa!Fax) > 0 Then
      LcFax = " Fax: " + RsEmpresa!Fax & ""
   Else
      LcFax = ""
   End If
   If Len(RsEmpresa!celular) > 0 Then
      Lccelular = "Celular: " & RsEmpresa!celular & ""
   Else
      Lccelular = " "
   End If
   If Len(RsEmpresa!celular1) > 0 Then
      Lccelular1 = " Celular: " & RsEmpresa!celular1 & ""
   Else
      Lccelular1 = " "
   End If
   If Len(RsEmpresa!email) > 0 Then
      Lcemail = "E-Mail: " & RsEmpresa!email & ""
   Else
      Lcemail = " "
   End If
   LcUf = RsEmpresa!estado
   LcCep = RsEmpresa!Cep
End If
RsEmpresa.Close
'If GlEscolheCliente Then
'   LcCriterio = "cod='" & RsClientes!cidade & "'"
'   RsCidade.FindFirst LcCriterio
'   If Not RsCidade.NoMatch Then
'      LcCidade = RsCidade!Nome
'   End If
'End If
LcCidade = LcCidade
LcUf = LcUf + ""
LcCep = LcCep & ""
 
LcEndereco = LcEndereco + " " + LCCidadeEm + LcUf + " " + LcCep
Set RsEmpresa = Nothing
'=== Imprime Cabecalho Nota
'Print #FnunNota,
'== Liga Modo Draft
If GlCabecalhoMeiaFolha Then
    Print #Lcarq, LigaDraft + Chr(18)
    Print #Lcarq, LigaTitulo + LcMargem & LcEmpresa + DesligaTitulo + Chr(13)
    Print #Lcarq, LcMargem & LcEndereco + Chr(13)
    Print #Lcarq, LcMargem & "Fone:  " + LcFone + LcFax + " " + Lccelular & Lccelular1 + Chr(13)
    LcTamanhoPedido = LcTamanhoPedido + 1
    Print #Lcarq, LcMargem & Lcemail '+ Chr(13)
    
    LcString = String(81 - Len(GlMargem), "-")
    Print #Lcarq, LcMar & LcString
    'LcTamanhoPedido = LcTamanhoPedido + 1
    'MtPedido(LcTamanhoPedido) = Chr(13)
End If

Set RsClientes = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCidade = Dbbase.OpenRecordset("Select * from alid005 where cod='" & RsClientes!cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Not RsClientes.EOF Then
   LcCpf = RsClientes!cpf & ""
   LcCgc = RsClientes!cgc & ""
   LcRua = RsClientes!End & ""
   LcBairro = RsClientes!bairro & ""
   LcEstado = RsClientes!estado & ""
   LcCep = RsClientes!Cep & ""
   LcFone = RsClientes!fone1 & ""
End If
RsClientes.Close
If Not RsCidade.EOF Then
   LcCidade = RsCidade!nome & ""
Else
   LcCidade = "Nao Cadastrada"
End If
RsCidade.Close
Dbbase.Close
Set RsClientes = Nothing
Set RsCidade = Nothing
Set Dbbase = Nothing

If Natureza.Text = "Orçamento" Then
    LcString = "ORCAMENTO/PEDIDO               NUMERO :" & Documento.Text & "  EMISSAO :" & emissao.Text
Else
    LcString = "ORCAMENTO/PEDIDO               NUMERO :" & Documento.Text & "  EMISSAO :" & emissao.Text
End If
Print #Lcarq, LcMar & LcString
Print #Lcarq, DesligaTitulo

LcString = Left("Vendedor" & "          ", 8) & ":" & codigoVendedor.Text & "   " & NomeVendedor.Text
Print #Lcarq, LcMar & LcString
'===> Monta dados do Cliente
LcString = ""
LcString = Left("Cliente" & "          ", 8) & ":" & NomeCliente.Text & ""
Print #Lcarq, LcMar & LcString
LcString = ""
If Len(LcCgc) = 0 Or LcCgc = "  .   .   /    -  " Then
   LcString = Left("C.P.F." & "          ", 8) & ":" & LcCpf & ""
Else
   LcString = Left("C.N.P.J." & "          ", 8) & ":" & LcCgc & ""
End If
Print #Lcarq, LcMar & LcString
LcString = ""
LcString = Left("Endereco" & "          ", 8) & ":" & LcRua & ""
Print #Lcarq, LcMar & LcString
LcString = ""
LcString = Left("Bairro" & "                            ", 16) & ":" & Left(LcBairro & "                  ", 15) & ""
LcString = LcString & Left("Cidade" & "                 ", 16) & ":" & LcCidade & ""
Print #Lcarq, LcMar & LcString
LcString = ""
LcString = Left("Estado" & "                           ", 16) & ":" & Left(LcEstado & "                  ", 15) & ""
LcString = LcString & Left("C.E.P" & "          ", 8) & ":" & LcCep & " Fone:" & LcFone
Print #Lcarq, LcMar & LcString

LcString = ""
LcString = Right("                    " & "Codigo", 17)
LcString = LcString & Left("    Descricao" & "                                                  ", 32)
LcString = LcString & Left("    Quant" & "       ", 9)
LcString = LcString & Right("           " & "V.Unit", 11)

LcString = LcString & Right("           " & "V.Total", 11)
Print #Lcarq, LcMar & LcString
LcString = ""
LcString = String(81 - Len(GlMargem), "-")
Print #Lcarq, LcMar & LcString
LcLinha = 13

End Function
Function ImprimeMeiaFolha()
Dim a As Integer
Dim LcFechado As Boolean
Dim RsImpressora As Recordset
Dim Db1 As Database

Set Db1 = OpenDatabase(GLBase, False, False)
Set RsImpressora = Db1.OpenRecordset("Select * from impressoras where impressora='" & GlPortaOrcamento & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   
If Not RsImpressora.EOF Then
   If RsImpressora!Maquina = GlNomeMaquina Then
      glimpressoraor = RsImpressora!EnderecoLocal
   Else
      glimpressoraor = RsImpressora!EnderecoRede
   End If
Else
  glimpressoraor = "LPT1"
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, Imprimindo..."

LeConfiguracaoMeiaFolha
LcItensImpressos = 0
Lcarq = FreeFile
LcTotalIten = CInt(GlItensMeiaFolha)
LcLinha = 1
LcMar = ""
For a = 1 To CInt(GlMargemMeiaFolha)
    LcMar = LcMar & " "
Next
LcFechado = False
'If IsNull(GlPortaOrcamento) Then GlPortaOrcamento = "LPT1"
'GlPortaOrcamento = "LPT1"
Open glimpressoraor For Output Access Write As #Lcarq
'===> Imprime Cabecalho
For a = 1 To Item.Rows - 1
    If LcLinha = 1 Then cabecalhoMeia
    Call imprimeitemMeia(a)
    LcFechado = False
    LcLinha = LcLinha + 1
    If LcItensImpressos >= LcTotalIten - 3 Then
       Call FechaImpressaoMeia
       LcFechado = True
       LcLinha = 1
    End If
    
Next
If Not LcFechado Then
  Call FechaImpressaoMeia
End If

Close #Lcarq
Me.Caption = LcCap

End Function
Function imprimeitemMeia(LcPos As Integer)
On Error Resume Next
lcsp = String(81, " ")
LcString = ""
LcString = Right(lcsp & Item.TextMatrix(LcPos, 1), 17) '& "/"
'LcString = LcString & Right(lcsp & Item.TextMatrix(LcPos, 2), 10) & "/"
'LcString = LcString & Right(lcsp & Item.TextMatrix(LcPos, 3), 3) & "  "
LcString = Left(LcString & lcsp, 21) & Left(Item.TextMatrix(LcPos, 2) & lcsp, 30) & "  "
LcString = LcString & Right(lcsp & Item.TextMatrix(LcPos, 5), 5)
LcString = LcString & Right(lcsp & AcertaNumero(Item.TextMatrix(LcPos, 6), 2), 11)
LcString = LcString & Right(lcsp & AcertaNumero(Item.TextMatrix(LcPos, 8), 2), 11)
Print #Lcarq, LcMar & LcString
LcItensImpressos = LcItensImpressos + 1
End Function
Function FechaImpressaoMeia()
On Error Resume Next
LcString = ""

For a = LcItensImpressos To LcTotalIten
    Print #Lcarq, Chr(13)
Next

LcString = String(80 - Len(GlMargemor), "-")
LcSt = String(80, " ")
Print #Lcarq, LcString

LcString = "Total Produto " & Right("                                                                                                                                          " & AcertaNumero(TotalProduto.Text, 2), 66)
Print #Lcarq, LcMar & LcString
LcString = "Desconto " & Right("                                                                                                                                          " & AcertaNumero(ValorDesconto.Text, 2), 71)
Print #Lcarq, LcMar & LcString
LcString = "Total " & Right("                                                                                                                                          " & AcertaNumero(TotalOrcamento.Text, 2), 74)
Print #Lcarq, LcMar & LcString
LcString = String(80 - Len(GlMargemor), "-")
Print #Lcarq, LcString
Print #Lcarq, "Cond de Pag." & DadosOrcamento.TipoPag.Text & "  Forma de Pag.: " & DadosOrcamento.TipoMonetario.Text
LcVencimentos = ""
b = 0
For a = 0 To CInt(DadosOrcamento.Quantidade.Text) - 1
       If Len(LcVencimentos) > 0 Then LcVencimentos = LcVencimentos & "  "
       If a = 0 Then LcVencimentos = "Vencimentos:"
       
       If b = 3 Then
            Print #Lcarq, LcVencimentos
            LcVencimentos = "            "
            b = 0
       Else
            b = b + 1
       End If
       LcVencimentos = LcVencimentos & Format(DadosOrcamento.Vencimento(a).Text, "dd/mm/yy") & "    " & Format(DadosOrcamento.valor.Text, "Currency")
Next
Print #Lcarq, LcVencimentos
For a = 1 To CInt(GlSaltoFinalMeiaFolha)
   Print #Lcarq, Chr(13)
Next

End Function
Function Imprimeorcamento()
Dim a As Integer
On Error GoTo Errimpr
LcMargem = ""
Dim bb As Database
Dim Item, Descricao, cst, icms, Unidade As String
Dim quant, Unitario, total, LcITemsImpressos As Long
Dim LcImpressoes, LcSegunda As Integer

Dim LcValor1, LcValor2, LcValor3, LcValor4, LcValor5 As Double
For a = 0 To 1000
   MtPedido(a) = ""
Next
'CalculaNumeroNota
LcSegunda = False
For a = 1 To GlMargem
    LcMargem = LcMargem & " "
Next
Set bb = OpenDatabase(GLBase, False, False)
Set RsClientes = bb.OpenRecordset("select * from alid001 where codigo='" & CodigoCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = bb.OpenRecordset("select * from empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsI = bb.OpenRecordset("select * from DadosOrcamento where doc='" & Documento.Text & "' order by item", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1

If IsNull(GlPortaOrcamento) Then GlPortaOrcamento = "LPT1"
LcImpressoes = 0
Call cabecalho

LcITemsImpressos = 0
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(15)

Do Until RsI.EOF
     'Call ImprimeItem(az)
     Call imprimeitem
     az = az + 1
     LcITemsImpressos = LcITemsImpressos + 1
     If LcITemsImpressos = 20 Then
        LcSegunda = True
        FechaImpressao
        cabecalho
        LcITemsImpressos = 0
        LcTamanhoPedido = LcTamanhoPedido + 1
        MtPedido(LcTamanhoPedido) = Chr(15)
        
     End If
     RsI.MoveNext
     
Loop

FechaImpressao
'Close #FnunNota
GeraSpool
RsClientes.Close
RsI.Close
bb.Close
Set RsClientes = Nothing
Set RsI = Nothing
Set bb = Nothing
Exit Function
Errimpr:
If err = 76 Then
   
   MsgBox "A Porta de Impressão " & GlPortaOrcamento & " Não Foi encontrada," & Chr(13) & "Verifique se a impressora está em linha e o cabo Conectado, ou a conexão da Rede.", 64, "Aviso"
   Exit Function
Else
   MsgBox err.Description & err.Number
   Resume Next
   
   
End If
Resume Next
End Function
Function cabecalho()
Dim BbC As Database
Dim RsCidade As Recordset, RsForn As Recordset
Dim LcSepara, LcSeparaItem As String
Dim LcFax, Lccelular, Lccelular1, Lcemail As String
Dim LCCidadeEm, LcTipo, LcIndustria As String
Dim a As Integer
LcSepara = ""
For a = 0 To 79 - CLng(GlMargem)
    LcSepara = LcSepara + "="
Next
For a = 0 To 79 - CLng(GlMargem)
    LcSeparaItem = LcSeparaItem + Chr(95)
Next
Set BbC = OpenDatabase(GLBase, False, False)
Set RsCidade = BbC.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = BbC.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsForn = BbC.OpenRecordset("alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcProcInd = "codigo='" & Industria.Text & "'"
If Not RsForn.EOF Then
    RsForn.FindFirst LcProcInd
    If Not RsForn.NoMatch Then
       LcIndustria = RsForn!Fantasia
    Else
       LcIndustria = ""
    End If
Else
  LcIndustria = ""
End If

'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LCCidadeEm = RsEmpresa!cidade
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!bairro
   LcFone = RsEmpresa!fone
   If Len(RsEmpresa!Fax) > 0 Then
      LcFax = " Fax: " + RsEmpresa!Fax & ""
   Else
      LcFax = ""
   End If
   If Len(RsEmpresa!celular) > 0 Then
      Lccelular = "Celular: " & RsEmpresa!celular & ""
   Else
      Lccelular = " "
   End If
   If Len(RsEmpresa!celular1) > 0 Then
      Lccelular1 = " Celular: " & RsEmpresa!celular1 & ""
   Else
      Lccelular1 = " "
   End If
   If Len(RsEmpresa!email) > 0 Then
      Lcemail = "E-Mail: " & RsEmpresa!email & ""
   Else
      Lcemail = " "
   End If
   LcUf = RsEmpresa!estado
   LcCep = RsEmpresa!Cep
End If
RsEmpresa.Close
If GlEscolheCliente Then
   LcCriterio = "cod='" & RsClientes!cidade & "'"
   RsCidade.FindFirst LcCriterio
   If Not RsCidade.NoMatch Then
      LcCidade = RsCidade!nome
   End If
End If
LcCidade = LcCidade
LcUf = LcUf + ""
LcCep = LcCep & ""
 
LcEndereco = LcEndereco + " " + LCCidadeEm + LcUf + " " + LcCep
Set RsEmpresa = Nothing
'=== Imprime Cabecalho Nota
'Print #FnunNota,
'== Liga Modo Draft
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LigaDraft + Chr(18)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LigaTitulo + LcMargem & LcEmpresa + DesligaTitulo + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1

MtPedido(LcTamanhoPedido) = LcMargem & LcEndereco + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & "Fone:  " + LcFone + LcFax + " " + Lccelular & Lccelular1 + Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & Lcemail '+ Chr(13)

'LcTamanhoPedido = LcTamanhoPedido + 1
'MtPedido(LcTamanhoPedido) = Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcSepara ' + Chr(13)

''LcTamanhoPedido = LcTamanhoPedido + 1
'MtPedido(LcTamanhoPedido) = Chr(13)
If Natureza.Text = "Orçamento" Then
   LcTipo = Left(" Cotacao:" & Documento.Text, 15)
Else
   LcTipo = Left(" Pedido:" & Documento.Text, 15)
End If
LcTamanhoPedido = LcTamanhoPedido + 1
If GlRepresentante Then
    Lcco = Left("Ind.: " + LcIndustria + "                                                                                ", 35) & Left("Emissao:" & emissao.Text, 18) & LigaNegrito & LcTipo & DesligaNegrito
    'MsgBox Lcco
    MtPedido(LcTamanhoPedido) = Lcco
Else
   Lcco = LigaNegrito & Left(LcTipo & "                                                                                    ", 35) & DesligaNegrito & Left("Emissao:" & emissao.Text, 18)
   MtPedido(LcTamanhoPedido) = Lcco
End If
'MsgBox MtPedido(LcTamanhoPedido)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(15) + Chr(13)

If GlEscolheCliente Then
   LcLinha = "Cliente   : " & NomeCliente.Text
Else
  LcLinha = "Cliente   : Consumidor Final"
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

If GlEscolheCliente Then
   LcLinha = Left("Endereco  : " & RsClientes!End & "                                                                                     ", 60)
   LcLinha = LcLinha & Left("Bairro    : " & RsClientes!bairro, 40)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   
   LcLinha = Left("Cidade    : " & LcCidade & "                                                                            ", 40)
   LcLinha = LcLinha & Left(" CEP:" & RsClientes!Cep & "                              ", 15)
   LcLinha = LcLinha & Left("     UF: " & RsClientes!estado, 16)
      
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
    
   LcLinha = Left("CNPJ      : " & RsClientes!cgc & "                                                                             ", 60)
   LcLinha = LcLinha & Left("Insc.Est. : " & RsClientes!INSCEST & "                                                                            ", 30)
   LcLinha = LcLinha & Left("Fone:" & RsClientes!fone1 & "              ", 25)
   LcLinha = LcLinha & Left("Fax:" & RsClientes!Fax & "           ", 24)

   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   
End If
If GlRepresentante Then
   If Not IsNull(Representada.Text) Then
      LcRepre = Left("Represent.:" & Representada.Text & "                                                                                     ", 60)
   Else
      LcRepre = Left("Represent.:" & "" & "                                                                                                        ", 60)
   End If
Else
   LcRepre = ""
End If
If Not GlEsclheVendedor Then
   LcLinha = LcRepre & Left("Vendedor  : " & LcEmpresa & "                                             ", 40)
Else
   If GlFuncCodigo Then
      LcLinha = LcRepre & Left("Vendedor  : " & codigoVendedor.Text & "                                                  ", 40)
   Else
      If GlFuncNome Then
         LcLinha = LcRepre & Left("Vendedor  : " & NomeVendedor.Text & "                                                  ", 40)
      Else
         LcLinha = LcRepre & Left("Vendedor  : " & LcEmpresa & "                                                  ", 40)
      End If
   End If
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)

LcLinha = " ITEM"
For a = 1 To 1
    LcLinha = LcLinha & " "
Next
'LcLinha = "Quantidade"
For a = 1 To 1
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "CODIGO"
For a = 1 To 8
    LcLinha = LcLinha & " "
Next

LcLinha = LcLinha & "QUANT."
For a = 1 To 4
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "UN"
For a = 1 To 3
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "DESCRICAO DO MATERIAL"
For a = 1 To 37
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & " AC"

For a = 1 To 2
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "  P.UNIT"

For a = 1 To 4
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "DES."

For a = 1 To 5
    LcLinha = LcLinha & " "
Next

LcLinha = LcLinha & " P.TOTAL"

For a = 1 To 5
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "IPI"

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha '& Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(18) + LcMargem & LcSeparaItem '+ Chr(13)
az = 18


End Function
Function FechaImpressao()
Dim RsOrc As Recordset
Dim BbFc As Database
Dim LcFormaAt As String
Dim a As Integer
Set BbFc = OpenDatabase(GLBase, False, False)
Set RsOrc = BbFc.OpenRecordset("orcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)
For a = az To 36
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = Chr(13)
Next
LcCri = "doc='" & Documento.Text & "'"
RsOrc.FindFirst LcCri
Dim lcLinhasSalto As Integer
For q = 1 To 100
    LcEspaco = LcEspaco & " "
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(15) + Chr(13)

For a = 0 To 78 - CLng(GlMargem)
    LcSepara = LcSepara + Chr(95)
    Next
LcLinha = LcSepara

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(18) + LcMargem & LcLinha + Chr(15) '& Chr(13)

If Len(DadosOrcamento.TipoPag) > 0 Then
   LcCondicoes = DadosOrcamento.TipoPag
Else
   LcCondicoes = ""
End If
LcLinha = ""
If Len(LcCondicoes) > 0 Then
  LcLinha = LigaNegrito & LcMargem & "Condicoes de Pag.: " & LcCondicoes & DesligaNegrito
End If
LcFormaAt = DadosOrcamento.TipoMonetario.Text
If Len(DadosOrcamento.TipoMonetario.Text) = 0 Then
   DadosOrcamento.TipoMonetario.Text = "Vencimento"
End If

LcLinha = LcLinha & Left(LcEspaco, 103 - Len(LcLinha)) & "   Total Produtos: " & Right("            " & AcertaNumero(CStr(RsOrc!TotalProduto), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
If Not GlImprimeDetalhaDesconto Then
   LcLinha = Left(LcEspaco, 99) & "  Desconto      : " & Right("            " & AcertaNumero(CStr(0), 2), 10)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
Else
   LcLinha = Left(LcEspaco, 99) & "  Desconto      : " & Right("            " & AcertaNumero(CStr(RsOrc!TotalDesconto), 2), 10)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
End If
LcLinha = ""
'==== Coloca a Primeira Duplicata
If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
   LcLinha = DadosOrcamento.TipoMonetario.Text & ":" & DadosOrcamento.Vencimento(0).Text
   LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7) & "  "
End If
If DadosOrcamento.Vencimento(1).Text <> "  /  /  " Then
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario.Text & ":" & DadosOrcamento.Vencimento(1).Text
   LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7) & "  "
End If
If Len(LcLinha) > 0 Then
   LcLinha = LcLinha & Left(LcEspaco, 99 - Len(LcLinha)) & Chr(179) & "  Acrescimo     : " & Right("           " & AcertaNumero(CStr(RsOrc!Acrecimo), 2), 10) & "       " & Chr(179)
Else
   LcLinha = LcLinha & Left(LcEspaco, 99) & Chr(179) & "  Acrecimo      : " & Right("           " & AcertaNumero(CStr(RsOrc!Acrecimo), 2), 10) & "       " & Chr(179)
End If

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcLinha = ""
'==== Coloca a Segunda Paarte dos Vencimentos
If DadosOrcamento.Vencimento(2).Text <> "  /  /  " Then
   LcLinha = DadosOrcamento.TipoMonetario.Text & ":" & DadosOrcamento.Vencimento(2).Text
   LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7) & "  "
End If
If DadosOrcamento.Vencimento(3).Text <> "  /  /  " Then
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario.Text & ":" & DadosOrcamento.Vencimento(3).Text
   LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7) & "  "
End If

If Not GlIpi Then
   TotalIpi.Text = 0
End If
If Len(LcLinha) > 0 Then
    LcLinha = LcLinha & Left(LcEspaco, 99 - Len(LcLinha)) & Chr(179) & "  IPI           : " & Right("            " & AcertaNumero(CStr(TotalIpi.Text), 2), 10) & "       " & Chr(179)

Else
  LcLinha = LcLinha & Left(LcEspaco, 99) & Chr(179) & "  IPI           : " & Right("            " & AcertaNumero(CStr(TotalIpi.Text), 2), 10) & "       " & Chr(179)
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
'==== Coloca a Terceira Parte dos Vencimentos
LcLinha = ""

If DadosOrcamento.Vencimento(4).Text <> "  /  /  " Then
   LcLinha = DadosOrcamento.TipoMonetario.Text & ":" & DadosOrcamento.Vencimento(4).Text
   LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
End If
If Len(LcLinha) > 0 Then
    LcLinha = LcLinha & Left(LcEspaco, 99 - Len(LcLinha)) & Chr(179) & "  Total Pagar   : " & Right("           " & AcertaNumero(CStr(RsOrc!TotalGeral), 2), 10) & "       " & Chr(179)
Else
   LcLinha = LcLinha & Left(LcEspaco, 99) & Chr(179) & "  Total Pagar   : " & Right("           " & AcertaNumero(CStr(RsOrc!TotalGeral), 2), 10) & "       " & Chr(179)
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1

MtPedido(LcTamanhoPedido) = Chr(18) + LcMargem & LcSepara + Chr(15) + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
DadosOrcamento.TipoMonetario.Text = LcFormaAt
'If Len(DadosOrcamento.TipoPag) > 0 Then
'   LcCondicoes = DadosOrcamento.TipoPag
'' Else
'   LcCondicoes = ""
'End If

'MtPedido(LcTamanhoPedido) = LcMargem & "  Condicoes de Pag.: " & DadosOrcamento.TipoPag & "     Forma de Pag.:" & DadosOrcamento.TipoMonetario & Chr(13) & "  "
'If Len(LcCondicoes) > 0 Then
'   MtPedido(LcTamanhoPedido) = LcMargem & "  Condicoes de Pag.: " & LcCondicoes
'End If
If Len(DadosOrcamento.TipoMonetario.Text) > 0 Then
  If DadosOrcamento.TipoMonetario <> "Nenhum" Then
     MtPedido(LcTamanhoPedido) = MtPedido(LcTamanhoPedido) & Chr(18) & "Forma de Pag.:" & DadosOrcamento.TipoMonetario.Text
  End If
End If
'LcTamanhoPedido = LcTamanhoPedido + 1
'MtPedido(LcTamanhoPedido) = Chr(13)
If GlImprimeDetalhaDesconto Then
   If Len(DescricaoDesconto.Text) > 0 Then
      LcLinha = "Descricao do Desconto: " & DescricaoDesconto.Text & " << Em Percentual >>"
      LcTamanhoPedido = LcTamanhoPedido + 1
      MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
      LcTamanhoPedido = LcTamanhoPedido + 1
      'MtPedido(LcTamanhoPedido) = Chr(18) + LcMargem & LcSepara & Chr(15) + Chr(13)
      'LcTamanhoPedido = LcTamanhoPedido + 1
     ' MtPedido(LcTamanhoPedido) = Chr(13)
    End If
End If

'If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
'  Select Case DadosOrcamento.quantidade.Text
 '      Case Is = "1"
 '          LcLinha = "Vencimento:" & DadosOrcamento.Vencimento(0).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '      Case Is = 2
 '          LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '          LcTamanhoPedido = LcTamanhoPedido + 1
  '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 ''           LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '      Case Is = 3
 '          LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '           LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '          LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '          LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '       Case Is = 4
 '          LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '          LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
 '           LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '          LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '          LcTamanhoPedido = LcTamanhoPedido + 1
 '          MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
 '          LcLinha = "4 Vencimento:" & DadosOrcamento.Vencimento(3).Text
 '          LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '          LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  '      Case Is = 5
  '         LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
  '         LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '         LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  '         LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
  '         LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '         LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  '         LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
  '         LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '         LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  '         LcLinha = "4 Vencimento:" & DadosOrcamento.Vencimento(3).Text
  '         LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '         LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  '         LcLinha = "5 Vencimento:" & DadosOrcamento.Vencimento(4).Text
  '         LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
  '          LcTamanhoPedido = LcTamanhoPedido + 1
  '         MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)'

           
 ' End Select
'End If
If Len(DadosOrcamento.Txt(3).Text) > 0 Then

   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13) & Chr(13)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)

   LcLinha = "DADOS PARA ENTREGA:"
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)
   LcLinha = "ENDERECO:" & DadosOrcamento.Txt(3).Text & "   BAIRRO:" & DadosOrcamento.Txt(4).Text
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
   LcLinha = "CIDADE:" & DadosOrcamento.Txt(6).Text & "    UF:" & DadosOrcamento.Txt(5).Text & "   C.E.P.:" & DadosOrcamento.Txt(12).Text
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcSepara '& Chr(13)
   
End If

LcCri = "doc='" & Documento.Text & "'"
RsOrc.FindFirst LcCri
If Not RsOrc.NoMatch Then
  If Len(RsOrc!Transp) > 0 Then
    ' LcTamanhoPedido = LcTamanhoPedido + 1
    'MtPedido(LcTamanhoPedido) = Chr(13)
    LcLinha = Chr(18) & LigaNegrito & "Transportadora: " & RsOrc!Transp & "    Fone: " & RsOrc!FoneTransp & DesligaNegrito & ""
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = Chr(13)
      
  End If
  'If Len(RsOrc!FoneTransp) > 0 Then
  ' LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = Chr(13)
  ' LcLinha = "Fone:" & RsOrc!FoneTransp
  ' LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  'End If

  If Len(RsOrc!Obs) > 0 Then
   'LcTamanhoPedido = LcTamanhoPedido + 1
   'MtPedido(LcTamanhoPedido) = Chr(13)

   LcLinha = Chr(15) & "Dados Complementares:"
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   'LcTamanhoPedido = LcTamanhoPedido + 1
   'MtPedido(LcTamanhoPedido) = Chr(13)
   LcLinha = RsOrc!Obs
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha '& Chr(13)
   
  ' LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = Chr(13)
  End If
End If
'LcTamanhoPedido = LcTamanhoPedido + 1
'MtPedido(LcTamanhoPedido) = Chr(13)
lctammsg = Len(GlMsg)
lcspa = ""
For a = 1 To (((80 - CLng(GlMargem)) / 2) - (lctammsg / 2))
    lcspa = lcspa & " "
Next
LcSepara = ""
For a = 0 To 79 - CLng(GlMargem)
    LcSepara = LcSepara + Chr(95)
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = lcspa & GlMsg & Chr(13)
'LcTamanhoPedido = LcTamanhoPedido + 1
' MtPedido(LcTamanhoPedido) = Chr(18) + LcMargem & LcSepara & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = lcspa & GlMsg1 & Chr(13)
If Gl40colunas Then
   For v = 1 To GlSaltoLinhasOrcamento
       LcTamanhoPedido = LcTamanhoPedido + 1
       MtPedido(LcTamanhoPedido) = Chr(13)
   Next
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(18) + Chr(13)

Else
  LcTamanhoPedido = LcTamanhoPedido + 1
  MtPedido(LcTamanhoPedido) = Chr(12)
  LcTamanhoPedido = LcTamanhoPedido + 1
 ' MtPedido(LcTamanhoPedido) = " "
  ' For v = 1 To GlSaltoLinhasOrcamento
  '     LcTamanhoPedido = LcTamanhoPedido + 1
  '     MtPedido(LcTamanhoPedido) = Chr(13)
  ' Next
   'LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = Chr(18) + Chr(13)
End If
'LcTamanhoPedido = LcTamanhoPedido + 1
'MtPedido(LcTamanhoPedido) = Chr(13)
orcamento.Caption = "Orçamento e Vendas"

End Function
Function imprimeitem()
Dim a, b As Integer
If Len(RsI!codigoproduto) = 0 Then Exit Function

LcLinha = " " & Left(RsI!Item & "             ", 6)
LcLinha = LcLinha & Left(RsI!codigoproduto & "                      ", 13)
LcLinha = LcLinha & "  " & Right("              " & RsI!quant, 7)
LcLinha = LcLinha & " " & Right("   " & RsI!unid, 5)
LcLinha = LcLinha & " " & Left(RsI!Descricao & "                                                                                                                   ", 55)
LcLinha = LcLinha & " " & Right("     " & RsI!Acomodacao & "", 4)
LcLinha = LcLinha & "   " & Right("               " & AcertaNumero(CStr(RsI!Unit), GlDecimais), 8)
If Not GlDescUnit Then
   LcLinha = LcLinha & Right("              " & AcertaNumero(CStr(RsI!Desconto), 2), 8)
Else
   LcLinha = LcLinha & Right("              " & AcertaNumero(CStr(0), 2), 8)
End If
For b = 1 To 2
   LcLinha = LcLinha & " "
Next

LcLinha = LcLinha & " " & Right("              " & AcertaNumero(CStr(RsI!total), GlDecimais), 10)

If GlIpi Then
   For b = 1 To 2
      LcLinha = LcLinha & " "
   Next
   LcLinha = LcLinha & " " & Right("            " & AcertaNumero(CStr(RsI!ipi), 0), 5) & " "
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha + Chr(13)

LcLinha = ""


End Function
Function BuscaNota(LcNumeroOrc As String)
On Error Resume Next
Dim bb As Database
Dim RsOrc As Recordset, RsItem As Recordset
Dim RsProduto As Recordset, RsCliente As Recordset
Dim RsVendedor As Recordset
Dim LcSql1, LcSql2, LcSql3, LcSql4, LcSql5 As String
Dim lcbloqueia As Integer
lcbloqueia = True
LcPesquisa = True
LcLiberaCalculo = False
LcSql1 = "Select * from orcamento where doc='" & LcNumeroOrc & "'"
LcSql2 = "Select * from DadosOrcamento where doc='" & LcNumeroOrc & "' order by item"
LcSql3 = "Select * from ALid001"
LcSql5 = "Select * from ALid200"

Set bb = OpenDatabase(GLBase, False, False)

Set RsOrc = bb.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsItem = bb.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCliente = bb.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVendedor = bb.OpenRecordset(LcSql5, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'==== Preenchendo a Nota

If RsOrc.EOF Then
   MsgBox "O Orçamento Nº: " & LcNumeroOrc & " Não foi encontrado..."
   Command4.Caption = "Pes&quisa F7"
   codigoVendedor.SetFocus
   Exit Function
End If
   
Documento.Text = RsOrc!doc
emissao.Text = Format(RsOrc!DTEMIS, "dd/mm/yy")
If RsOrc!Natureza = "Or" Then
   Natureza.Text = "Orçamento"
Else
   Natureza.Text = "Venda"
   
End If
codigoVendedor.Text = RsOrc!Vendedor & ""
LcCriterio = "Codigo='" & RsOrc!Vendedor & "'"
RsVendedor.FindFirst LcCriterio
If Not RsVendedor.NoMatch Then
   NomeVendedor.Text = RsVendedor!nome
Else
  NomeVendedor.Text = ""
End If
CodigoCliente.Text = RsOrc!CLIENTE
LcCriterio = "Codigo='" & RsOrc!CLIENTE & "'"
RsCliente.FindFirst LcCriterio
If Not RsCliente.NoMatch Then
   NomeCliente.Text = RsCliente!razaosoc
   LimiteCredito.Text = RsCliente!LimiteCredito
   LimiteUtilizado.Text = RsCliente!CreditoUtilizado
End If
Comissao.Text = RsOrc!Comissao
'===> Retirado 15/11/01  ComissaoProduto.Text = RsOrc!orcamento
TotalProduto.Text = AcertaNumero(CStr(RsOrc!TotalProduto), 2)
TotalOrcamento.Text = AcertaNumero(CStr(RsOrc!TotalGeral), 2)
DescricaoDesconto.Text = RsOrc!DetalhaDesconto & ""
DescricaoAcrescimo.Text = RsOrc!DetahaAcrecimo & ""
Status.Text = RsOrc!Status
ValorAcrescimo.Text = RsOrc!Acrecimo
'If Len(RsOrc!desconto) > 0 Then valordesconto.Text = RsOrc!desconto Else valordesconto.Text = ""
If Len(RsOrc!TotalDesconto) > 0 Then ValorDesconto.Text = RsOrc!TotalDesconto Else ValorDesconto.Text = ""
Transportadora.Text = RsOrc!Transp & ""
FoneTransp.Text = RsOrc!FoneTransp & ""
DadosComplementares.Text = RsOrc!Obs & ""
Industria.Text = RsOrc!Industria & ""
Representada.Text = RsOrc!Representada & ""

TotalIpi.Text = AcertaNumero(CStr(RsOrc!ipi), 2)
'===== Escreve dados Grid
LcItem = 0
err.Number = 0
Do Until RsItem.EOF
   If err.Number > 0 Then Exit Do
   LcItem = LcItem + 1
   Item.Rows = LcItem + 1
   LcRow = LcItem
   Item.TextMatrix(LcRow, 0) = Right("000" & RsItem!Item, 3) & ""
   Item.TextMatrix(LcRow, 1) = RsItem!codigoproduto & ""
   Item.TextMatrix(LcRow, 2) = RsItem!Descricao & ""
   Item.TextMatrix(LcRow, 3) = RsItem!unid & ""
   Item.TextMatrix(LcRow, 4) = RsItem!Acomodacao & ""
   Item.TextMatrix(LcRow, 5) = RsItem!quant & ""
   Item.TextMatrix(LcRow, 6) = RsItem!Unit & ""
   Item.TextMatrix(LcRow, 7) = RsItem!Desconto & ""
   Item.TextMatrix(LcRow, 8) = RsItem!total & ""
   Item.TextMatrix(LcRow, 10) = RsItem!Comissao & ""
   Item.TextMatrix(LcRow, 11) = RsItem!Unit & ""
   Item.TextMatrix(LcRow, 12) = RsItem!acrescimo & ""
   Item.TextMatrix(LcRow, 13) = RsItem!codigounidade & ""
   Item.TextMatrix(LcRow, 14) = RsItem!ComissaoFabricante & ""
   Item.TextMatrix(LcRow, 15) = RsItem!DescricaoDesconto & ""
   If RsItem!Desconto > 0 Then lcbloqueia = False
   If GlIpi Then
     Item.TextMatrix(LcRow, 9) = RsItem!ipi & ""
   End If
   RsItem.MoveNext
   LcAchou = True
Loop
DescricaoDesconto.Enabled = lcbloqueia
ValorDesconto.Enabled = lcbloqueia
If LcAchou Then
   codigoproduto.SetFocus
Else
   codigoVendedor.SetFocus
    'CmdSalvar.Visible = True
    'Command3.Visible = False
End If
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True
RsOrc.Close
RsItem.Close
RsCliente.Close
RsVendedor.Close
bb.Close

Set RsOrc = Nothing
Set RsItem = Nothing
Set RsCliente = Nothing
Set RsVendedor = Nothing
Set bb = Nothing

codigoproduto.SetFocus
LcLiberaCalculo = True
End Function

Function limpanota()
On Error Resume Next
Documento.Text = ""
emissao.Text = Date
Status.Text = "Em Lançamento"
codigoVendedor.Text = ""
NomeVendedor.Text = ""
CodigoCliente.Text = ""
NomeCliente.Text = ""
codigoproduto.Text = ""
NomeProduto.Text = ""
Unidade.Text = ""
Acomodacao.Text = ""
Quantidade.Text = ""
Unitario.Text = ""
total.Text = ""
DescricaoDesconto.Text = ""
ValorDesconto = ""
DescricaoAcrescimo.Text = ""
ValorAcrescimo.Text = ""
TotalIpi.Text = ""
TotalProduto.Text = ""
TotalOrcamento.Text = ""
Transportadora.Text = ""
FoneTransp.Text = ""
Comissao.Text = ""
ComissaoProduto.Text = ""
ipi.Text = ""
DadosComplementares.Text = ""
codigounidade.Text = ""
LimiteCredito.Text = ""
acrescimo.Text = ""
LimiteUtilizado.Text = ""
preconormal.Text = ""
LcItem = 0
Item.Rows = 1
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
DescricaoDesconto.Enabled = True
ValorDesconto.Enabled = True
Industria.Text = ""
End Function

Private Sub Acomodacao_GotFocus()
LcLimpa = True
End Sub

Private Sub Acomodacao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 123 Then UltimasComprasCliente.Show , Me
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub Acomodacao_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   Acomodacao.Text = ""
   LcLimpa = False
End If
End Sub


Private Sub Acomodacao_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   Unidade.SetFocus
End If

End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
FrmExcluiItem.Show
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub CmdSalvar_Click()
SalvaOrcamento
ConfirmaOrcamento.Show , Me
limpanota
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub CodigoCliente_Change()
'GlBuscaProdutoAgora = False
End Sub

Private Sub CodigoCliente_GotFocus()
On Error Resume Next
If LcMontaComissao Then
   If GlVariasComissao Then
      Exibecomissao.Show , Me
   End If
   LcMontaComissao = False
End If
End Sub

Private Sub CodigoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   GlCriterioSql = " where RAZAOSOC like '" & NomeCliente.Text & "*' order by RAZAOSOC"
   LcPerguntaCliente = False
   FrmBuscaCliente.Show , Me
End If

If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub


Private Sub CodigoCliente_KeyPress(KeyAscii As Integer)
GlBuscaProdutoAgora = False
End Sub

Private Sub CodigoCliente_LostFocus()
On Error Resume Next

Dim bb As Database, RsCliente As Recordset
'===> Verifica se deve Buscar Produto
If Not LcBuscaCliente Then Exit Sub
If Len(Trim(CodigoCliente.Text)) = 0 Then Exit Sub
   
'===> Verifica se o Valor digitado é Númerico
If GLCalculacodigoCliente Then
   If Not IsNumeric(CodigoCliente.Text) And Len(CodigoCliente.Text) > 0 Then
      MsgBox "O Código do Cliente deve ser Numérico...", vbExclamation, "Aviso"
      CodigoCliente.SetFocus
      Exit Sub
   End If
   CodigoCliente.Text = Right("00000" & CodigoCliente.Text, 5)
End If

Set bb = OpenDatabase(GLBase, False, False)
Set RsCliente = bb.OpenRecordset("select * From alid001 where CODIGO='" & CodigoCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
If Not RsCliente.EOF Then
   NomeCliente.Text = RsCliente!razaosoc
   LimiteCredito.Text = RsCliente!LimiteCredito
   LimiteUtilizado.Text = RsCliente!CreditoUtilizado
  codigoproduto.SetFocus
Else
'   LcPerguntaCliente = False
 '  FrmPesquisaCliente.Show , Me
    CodigoCliente.Text = ""
End If
RsCliente.Close
bb.Close
End Sub
Function InprimeNotaOLinto()

On Error GoTo Errimpr
LcMargem = ""
Dim bb As Database
Dim Item, Descricao, cst, icms, Unidade As String
Dim quant, Unitario, total, LcITemsImpressos As Long
Dim LcImpressoes, LcSegunda, a As Integer
Dim az As Long
Dim LcValor1, LcValor2, LcValor3, LcValor4, LcValor5 As Double
For a = 0 To 1000
   MtPedido(a) = ""
Next
'CalculaNumeroNota
LcSegunda = False

Set bb = OpenDatabase(GLBase, False, False)
Set RsClientes = bb.OpenRecordset("select * from alid001 where codigo='" & CodigoCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = bb.OpenRecordset("select * from empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsI = bb.OpenRecordset("select * from DadosOrcamento where doc='" & Documento.Text & "' order by descricao", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1

If IsNull(GlPortaOrcamento) Then GlPortaOrcamento = "LPT1"
LcImpressoes = 0
Call cabecalhoOlinto
'=== Determina a Quantidade de Itens disponiveis
'For az = 1 To Item.Rows - 1
'    If Len(LcMat(az).CodPro) > 0 Then
'       LcTotalitem = LcTotalitem + 1
'    End If
'Next
az = 1
LcITemsImpressos = 0
Do Until RsI.EOF
     Call imprimeimtemOlinto(az)
     LcITemsImpressos = LcITemsImpressos + 1
     If LcITemsImpressos = 25 Then
        LcSegunda = True
        FechaImpressaoolinto
        cabecalhoOlinto
        LcITemsImpressos = 0
     End If
     RsI.MoveNext
     az = az + 1
Loop
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
If LcSegunda Then
   LcLinhaAtual = LcLinhaAtual + 2
Else
   LcLinhaAtual = LcLinhaAtual + 1
End If

LcUltimo = CDbl(TotalOrcamento.Text) - (CDbl(AcertaNumero(DadosOrcamento.valor.Text, 2)) * CDbl(DadosOrcamento.Quantidade.Text))

LcValor1 = CCur(AcertaNumero(DadosOrcamento.valor.Text, 2))
LcValor2 = CCur(AcertaNumero(DadosOrcamento.valor.Text, 2))
LcValor3 = CCur(AcertaNumero(DadosOrcamento.valor.Text, 2))
LcValor4 = CCur(AcertaNumero(DadosOrcamento.valor.Text, 2))
LcValor5 = CCur(AcertaNumero(DadosOrcamento.valor.Text, 2))

Select Case Val(DadosOrcamento.Quantidade.Text)
 Case Is = 1
      LcValor1 = LcValor1 + LcUltimo
 Case Is = 2
      LcValor2 = LcValor2 + LcUltimo
 Case Is = 3
      LcValor3 = LcValor3 + LcUltimo
 Case Is = 4
      LcValor4 = LcValor4 + LcUltimo
 Case Is = 5
      LcValor5 = LcValor5 + LcUltimo
      
End Select


If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(0).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(CStr(LcValor1), 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(1).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(1).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(CStr(LcValor2), 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(2).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(2).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(CStr(LcValor3), 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(3).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(3).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(CStr(LcValor4), 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(4).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(4).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(CStr(LcValor5), 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
FechaImpressaoolinto
'Close #FnunNota
GeraSpool
RsClientes.Close
RsI.Close
bb.Close
Set RsClientes = Nothing
Set RsI = Nothing
Set bb = Nothing
Exit Function
Errimpr:
If err = 76 Then
   
   MsgBox "A Porta de Impressão " & GlPortaOrcamento & " Não Foi encontrada," & Chr(13) & "Verifique se a impressora está em linha e o cabo Conectado, ou a conexão da Rede.", 64, "Aviso"
   Exit Function
Else
   MsgBox err.Description & err.Number
   Resume 0
   
End If
Resume 0
End Function
Function imprimeimtemOlinto(az As Long)
Dim LcTamanhodes, LcImp As Long
Dim a As Integer
LcMargem = " "
For a = 1 To 2
    LcMargem = LcMargem & " "
Next

Dim LcDes As String
For we = 1 To 22
    lces = lces & " "
Next
If Len(RsI!codigoproduto) = 0 Then Exit Function

LcItensImpressos = LcItensImpressos + 1

LcLinha = Left(RsI!codigoproduto & "            ", 5)
LcLinha = LcLinha & " "
LcLinha = LcLinha & Right("     " & RsI!quant, 5)
LcLinha = LcLinha & " "
LcLinha = LcLinha & Right("  " & RsI!unid, 2)
LcLinha = LcLinha & " "
LcTamanhodes = Len(RsI!Descricao)

If LcTamanhodes <= 42 Then
    LcLinha = LcLinha & Left(RsI!Descricao & "                                                  ", 42)
    LcLinha = LcLinha & " "
    LcLinha = LcLinha & Right("             " & AcertaNumero(CStr(RsI!Unit), GlDecimais), 8)
    LcLinha = LcLinha & " "
    LcLinha = LcLinha & Right("              " & AcertaNumero(CStr(RsI!total), 2), 9)
    
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = LcMargem & LcLinha
    
        LcLinha = ""
Else
    LcImp = LcTamanhodes / 42
    If (LcTamanhodes - LcImp) > 0 Then
       LcImp = LcImp + 1
    End If
    For r = 0 To LcImp - 1
        LcDes = Mid(RsI!Descricao, (r * 42) + 1, 42)
        If r = 0 Then
           LcLinha = LcLinha & Left(LcDes & "                                           ", 42)
           LcLinha = LcLinha & " "
           LcLinha = LcLinha & Right("             " & AcertaNumero(CStr(RsI!Unit), GlDecimais), 8)
           LcLinha = LcLinha & " "
           LcLinha = LcLinha & Right("              " & AcertaNumero(CStr(RsI!total), 2), 9)
       Else
           LcLinha = lces & Left(LcDes & "                            ", 42)

       End If
      
       'MsgBox LcLinha
       LcTamanhoPedido = LcTamanhoPedido + 1
       MtPedido(LcTamanhoPedido) = LcMargem & LcLinha
       LcLinha = ""
    Next
    LcLinhaAtual = LcLinhaAtual + LcImp
End If
LcLinhaAtual = LcLinhaAtual + 1
End Function

Function GeraSpool()
On Error Resume Next
Dim a As Integer
Dim bb As Database
Dim RsLogNota As Recordset, RsImpressora As Recordset
Set bb = OpenDatabase(GLBase, False, False)
Set RsLogNota = bb.OpenRecordset("select * from LogImpressaoOrcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsImpressora = Dbbase.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'")
Set RsImpressora = bb.OpenRecordset("select * from impressoras where Impressora='" & GlPortaOrcamento & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

For a = 0 To LcTamanhoPedido
   RsLogNota.AddNew
   RsLogNota!impressora = GlPortaOrcamento
   If UCase(RsImpressora!Maquina) <> "LOCAL" Then
      If Len(Trim(GlNomeMaquina)) > 0 Then
         RsLogNota!Maquina = RsImpressora!Maquina
      Else
         RsLogNota!Maquina = "LOCAL"
      End If
   Else
      RsLogNota!Maquina = "LOCAL"
   End If
   RsLogNota!NF = Documento.Text
   RsLogNota!dados = MtPedido(a)
   RsLogNota.Update
Next
LcTamanhoPedido = 0
RsLogNota.Close
bb.Close
Set RsLogNota = Nothing
Set bb = Nothing
End Function
Function cabecalhoOlinto()
Dim RsCidade As Recordset
Dim LcSepara As String
Dim LcSalto As Long
Dim a As Integer
Dim bb As Database
LcSalto = Val(GLSaltoLinhaNota)
For a = 1 To 12
    LcMargem = LcMargem & " "
Next
'=== Salta o Espaço para Logotipo
LcEspa = "                                                     "
'LcTamanhoPedido = LcTamanhoPedido - 1
For a = 1 To LcSalto
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)
Next

Set bb = OpenDatabase(GLBase, False, False)
Set RsCidade = bb.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
'Me.Caption = "Aguarde, Gerando o Relatório..."
LcCriterio = "cod='" & RsClientes!cidade & "'"
RsCidade.FindFirst LcCriterio
If Not RsCidade.NoMatch Then
   LcCidade = RsCidade!nome
End If
Set RsEmpresa = Nothing
'=== Imprime Cabecalho Nota
For a = 1 To Len(NomeVendedor.Text)
    LCLEtra = Mid$(NomeVendedor.Text, a, 1)
    If LCLEtra = " " Then Exit For
    LcVend = LcVend & LCLEtra
Next


LcLinha = Left(NomeCliente.Text & LcEspa, 47) & _
Left(LcVend & LcEspa, 12) & Documento.Text '=== O Nome do Cliente e o Contato
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)


LcLinha = Left(RsClientes!End & LcEspa, 44) & _
RsClientes!bairro '=== O endereço e o Bairro
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)

LcLinha = Left(RsClientes!Cep & LcEspa, 26) & _
Left(LcCidade & LcEspa, 20) & RsClientes!estado '== Dados da Cidade
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)

LcLinha = Left(RsClientes!cgc & LcEspa, 35) & _
RsClientes!INSCEST  '=== Cgc e Inscricao
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
LcLinha = Left(RsClientes!fone1 & LcEspa, 22) _
& Left(DadosOrcamento.TipoPag.Text & LcEspa, 13)  '==== Cod pag e data emissao

'=== Separa para imprimir vencimentos
LcTamanho = Len(DadosOrcamento.Txt(11).Text)
lcvezes = 1
For a = 1 To LcTamanho
   LCLEtra = Mid(DadosOrcamento.Txt(11).Text, a, 1)
   If IsNumeric(LCLEtra) Then
      LcPrazo = LcPrazo & LCLEtra
   Else
     LcLinha = LcLinha & " " & Right("    " & LcPrazo, 3)
     LcPrazo = ""
     lcvezes = lcvezes + 1
   End If
Next

If Len(LcPrazo) > 0 Then
     LcLinha = LcLinha & " " & Right("    " & LcPrazo, 3)
     LcPrazo = ""
     lcvezes = lcvezes + 1
End If
'==== Acerta Quantidade de vezes para gerar espacos dos quadros
For X = 1 To 5 - lcvezes
    LcLinha = LcLinha & "    "
Next

LcLinha = LcLinha & Right("                " & emissao.Text, 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
LcLinhaAtual = 26
End Function
Function FechaImpressaoolinto()
Dim a As Integer
Dim lcLinhasSalto As Integer
'=== Posisiona na linha Certa
For wq = LcLinhaAtual To 57
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = Chr(13)
Next
LcLinha = " "
'=== Posiona na Coluna certa
For wq = 2 To 66
    LcEsp = LcEsp & " "
Next
LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(TotalProduto.Text), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(ValorDesconto.Text), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(TotalOrcamento.Text), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
For wq = 1 To 3
  LcTamanhoPedido = LcTamanhoPedido + 1
  MtPedido(LcTamanhoPedido) = Chr(13)
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & GlMsg & Chr(13)

For a = 1 To Val(GlPuloFim)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)
Next

End Function

Function SalvaOrcamento()
On Error Resume Next
Dim a As Integer
Dim bb As Database, Rsorcamento As Recordset
Dim RsDados As Recordset, RsCliente As Recordset
Dim RsProduto As Recordset, RsContas As Recordset
Dim RsCaixa As Recordset, RsComissao As Recordset
Dim RsTipoMonetario As Recordset, RsComissaoRepr As Recordset

Dim LcCriterioBusca, LcCriterioProduto, LcNatureza, LcCliente As String
Dim LcNumero As String, LcTipoMOne As String
Dim LcTotal, LcPercAcres, LcDifAcr As Double
Dim LcDifDes, LcPercDesc, LcValorComissao, LcValor As Double
'===> Abre As Bases de Dados

Set bb = OpenDatabase(GLBase, False, False)
Set RsCliente = bb.OpenRecordset("select * From alid001 where CODIGO='" & CodigoCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
Set Rsorcamento = bb.OpenRecordset("Orcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsDados = bb.OpenRecordset("DadosOrcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = bb.OpenRecordset("alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsContas = bb.OpenRecordset("Alid015", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCaixa = bb.OpenRecordset("Alid016", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsComissao = bb.OpenRecordset("Alid201", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsTipoMonetario = bb.OpenRecordset("Alid008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsComissaoRepr = bb.OpenRecordset("ComissaoRepresentante", dbOpenDynaset, dbSeeChanges, dbOptimistic)

'==== Verifica se é Alteracao ou inclusao
If Len(Documento.Text) > 0 Then 'é inclusao
  '== É Alteracao, Então Limpa os dados anteriores e relança
  '== Limpa orcamento
  LcCriterioBusca = "doc='" & Documento.Text & "'"
  Rsorcamento.FindFirst LcCriterioBusca
  If Not Rsorcamento.NoMatch Then
     LcTotal = Rsorcamento!TotalGeral
     LcNatureza = Rsorcamento!condpag
     LcCliente = Rsorcamento!CLIENTE
     Rsorcamento.Delete
     
  End If
  RsDados.FindFirst LcCriterioBusca
  Do Until RsDados.NoMatch
     
     
     LcCriterioProduto = "cod='" & RsDados!codigoproduto & "'"
     RsProduto.FindFirst LcCriterioProduto
     If Not RsProduto.NoMatch Then
        RsProduto.Edit
        RsProduto!QuantEstoque = RsProduto!QuantEstoque + CDbl(RsDados!quant)
        RsProduto.Update
     End If
     RsDados.Delete
     RsDados.FindNext LcCriterioBusca
     If err.Number = 3021 Then Exit Do
  Loop
  err.Number = 0
  If LcNatureza = "A Prazo" Then
     LcCriterioBusca = "Codigo='" & LcCliente & "'"
     RsCliente.FindFirst LcCriterioBusca
     If Not RsCliente.NoMatch Then
       If RsCliente!CreditoUtilizado > 0 Then
          RsCliente.Edit
          RsCliente!ULTCOMPRA = CDate(emissao.Text)
          RsCliente!CreditoUtilizado = CreditoUtilizado - LcTotal
          RsCliente.Update
       End If
     End If
  End If
  err.Number = 0
  LcCriterioBusca = "nf like '" & Documento.Text & "*'"
  RsContas.FindFirst LcCriterioBusca
  Do Until RsContas.NoMatch
     RsContas.Delete
     RsContas.FindNext LcCriterioBusca
     If err.Number = 3021 Then Exit Do
  Loop
  err.Number = 0
  LcCriterioBusca = "nf='" & Documento.Text & "' and CLICRED='R'"
  RsCaixa.FindFirst LcCriterioBusca
  Do Until RsCaixa.NoMatch
     RsCaixa.Delete
     RsCaixa.FindNext LcCriterioBusca
     If err.Number = 3021 Then Exit Do
  Loop
  err.Number = 0
  LcCriterioBusca = "nf='" & Documento.Text & "'"
  RsComissao.FindFirst LcCriterioBusca
  Do Until RsComissao.NoMatch
     RsComissao.Delete
     RsComissao.FindNext LcCriterioBusca
     If err.Number = 3021 Then Exit Do
  Loop
  err.Number = 0
  LcCriterioBusca = "nf='" & Documento.Text & "'"
  RsComissaoRepr.FindFirst LcCriterioBusca
  Do Until RsComissaoRepr.NoMatch
     RsComissaoRepr.Delete
     RsComissaoRepr.FindNext LcCriterioBusca
     If err.Number > 0 Then Exit Do
  Loop
  err.Number = 0
End If


'=== Gera O numero da Nota
If Len(Documento.Text) = 0 Then
   LcNumero = numeronota()
   Documento.Text = LcNumero
Else
   LcNumero = Documento.Text
End If
'===Grava Primeiro o Orcamento

Rsorcamento.AddNew
Rsorcamento!doc = LcNumero
Rsorcamento("DTEMIS") = CDate(emissao.Text)
Rsorcamento("NATUREZA") = Left(Natureza.Text, 2)
Rsorcamento("CLiente") = CodigoCliente.Text
Rsorcamento("TIPOTRANS") = Mid(DadosOrcamento.tipo.Text, 1, 1)
Rsorcamento("OrcVenda") = Natureza.Text
Rsorcamento("Vendedor") = codigoVendedor.Text
Rsorcamento("CondPag") = DadosOrcamento.TipoPag.Text
LcNatureza = DadosOrcamento.TipoPag.Text
If Len(DescricaoDesconto.Text) > 0 Then Rsorcamento!Desconto = DescricaoDesconto.Text
If Len(ValorDesconto.Text) > 0 Then Rsorcamento!TotalDesconto = CDbl(ValorDesconto.Text)
If Len(TotalOrcamento.Text) > 0 Then Rsorcamento!TotalGeral = CDbl(TotalOrcamento.Text)
If Len(TotalProduto.Text) > 0 Then Rsorcamento!TotalProduto = CDbl(TotalProduto.Text)
Rsorcamento!formapag = DadosOrcamento.TipoMonetario.Text
Rsorcamento!Dias = DadosOrcamento.Txt(11).Text
Rsorcamento!Acrecimo = CDbl(ValorAcrescimo.Text)
Rsorcamento!DetalhaDesconto = DescricaoDesconto.Text & ""
Rsorcamento!DetahaAcrecimo = DescricaoAcrescimo.Text & ""
Rsorcamento!Status = "Confirmado"
Rsorcamento!ipi = CDbl(TotalIpi.Text)
Rsorcamento!PrevComissao = CDate(DadosOrcamento.PrevCommisao.Text)
If Len(Comissao.Text) > 0 Then
   Rsorcamento!Comissao = CDbl(Comissao.Text)
End If
If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
   Rsorcamento!Vencimento1 = DadosOrcamento.Vencimento(0).Text
   Rsorcamento!vencimento2 = DadosOrcamento.Vencimento(1).Text
   Rsorcamento!vencimento3 = DadosOrcamento.Vencimento(2).Text
   Rsorcamento!vencimento4 = DadosOrcamento.Vencimento(3).Text
   Rsorcamento!vencimento5 = DadosOrcamento.Vencimento(4).Text
End If
Rsorcamento!FoneTransp = FoneTransp.Text & ""
Rsorcamento!Transp = Transportadora.Text & ""
Rsorcamento!Obs = DadosComplementares.Text & ""
Rsorcamento!Industria = Industria.Text & ""
Rsorcamento!Representada = Representada.Text & ""

Rsorcamento.Update
'==== Salva os Itens do pedido
For a = 1 To Item.Rows - 1
    RsDados.AddNew
    Call Ficha(LcNumero, Item.TextMatrix(a, 1), Item.TextMatrix(a, 2), CDbl(Item.TextMatrix(a, 5)), CDbl(Item.TextMatrix(a, 6)), CDbl(Item.TextMatrix(a, 8)), "S", NomeCliente.Text, Item.TextMatrix(a, 3), " ")
    RsDados("Doc") = LcNumero
    RsDados("CodigoProduto") = Item.TextMatrix(a, 1)
    RsDados("Descricao") = Item.TextMatrix(a, 2)
    RsDados("Quant") = CDbl(Item.TextMatrix(a, 5))
    RsDados("Unit") = CDbl(Item.TextMatrix(a, 6))
    RsDados("Total") = CDbl(Item.TextMatrix(a, 8))
    RsDados("unid") = Item.TextMatrix(a, 3)
    RsDados("item") = Item.TextMatrix(a, 0)
    RsDados("Acomodacao") = Item.TextMatrix(a, 4)
    RsDados("ComissaoFabricante") = Item.TextMatrix(a, 14)
    RsDados("DescricaoDesconto") = Item.TextMatrix(a, 15)
    If GlIpi Then
       RsDados("ipi") = Item.TextMatrix(a, 9) & ""
    Else
       RsDados("ipi") = "0"
    End If
    RsDados("valorUnitarioReal") = CDbl(Item.TextMatrix(a, 11))
    If Len(Item.TextMatrix(a, 7)) > 0 Then RsDados("desconto") = CDbl(Item.TextMatrix(a, 7))
    If Len(Item.TextMatrix(a, 12)) > 0 Then RsDados("acrescimo") = CDbl(Item.TextMatrix(a, 12))
    RsDados("codigounidade") = Item.TextMatrix(a, 13) & ""
    RsDados("comissao") = Item.TextMatrix(a, 10) & ""
    RsDados.Update
Next

'===> se for orçamento nao baixa estoque e etc
If Natureza.Text = "Orçamento" Then GoTo SaidaSistema
'=== Baixa o Estoque
For a = 1 To Item.Rows - 1
      LcCriterioFornec = "cod='" & Item.TextMatrix(a, 1) & "'"
      RsProduto.FindFirst LcCriterioFornec
      If Not RsProduto.NoMatch Then
         RsProduto.Edit
         RsProduto!QuantEstoque = RsProduto!QuantEstoque - CDbl(Item.TextMatrix(a, 5))
         RsProduto.Update
      End If
Next
'===> se for Emprestimo nao Gera Financeiro
If Natureza.Text = "Emprestimo" Then GoTo SaidaSistema

'===Gera a Comissao do Pedido
If Len(ValorDesconto.Text) > 0 Then
   LcPercDesc = CDbl(ValorDesconto.Text) / CDbl(TotalProduto.Text)
Else
   LcPercDesc = 0
End If
If Len(ValorAcrescimo.Text) > 0 Then
   LcPercAcres = CDbl(ValorAcrescimo.Text) / CDbl(TotalProduto.Text)
Else
   LcPercDesc = 0
End If
'=== Lança a comissão do Vendnedor
For a = 1 To Item.Rows - 1
      RsComissao.AddNew
      LcValorComissao = AcertaNumero(CStr((CDbl(Comissao.Text) / 100) * CDbl(Item.TextMatrix(a, 8))), GlDecimais)
      RsComissao("Vendedor") = codigoVendedor.Text
      RsComissao("NF") = LcNumero
      RsComissao("Produto") = Item.TextMatrix(a, 1)
      RsComissao("QUANTIDADE") = CDbl(Item.TextMatrix(a, 5))
      RsComissao("VALORUNIT") = CDbl(Item.TextMatrix(a, 6))
      RsComissao("VALORTOTAL") = CDbl(Item.TextMatrix(a, 8))
    ' If Comissao.Text = "1" Then Ibaixo = True Else Ibaixo = False
    ' RsComissao("ITEMBAIXO") = Ibaixo
      LcDifAcr = LcPercAcres * LcValorComissao
      LcDifDes = LcPercDesc * LcValorComissao
      RsComissao("COMISSAO") = LcValorComissao + LcDifAcr - LcDifDes
      RsComissao("DATAVENDA") = CDate(DadosOrcamento.PrevCommisao.Text)
      RsComissao("CLIENTE") = CodigoCliente.Text & ""
      RsComissao("percentual") = Comissao.Text & ""
      LcCriterioFornec = "cod='" & Item.TextMatrix(a, 1) & "'"
      RsProduto.FindFirst LcCriterioFornec
      If Not RsProduto.NoMatch Then
         RsComissao("Fornecedor") = RsProduto!fornecedor
      End If
      RsComissao.Update
Next
'==== Atualiza ComissaoFornecedor
For a = 1 To Item.Rows - 1
      RsComissaoRepr.AddNew
      LcValorComissao = AcertaNumero(CStr((CDbl(Item.TextMatrix(a, 14)) / 100) * CDbl(Item.TextMatrix(a, 8))), GlDecimais)
      RsComissaoRepr("Vendedor") = codigoVendedor.Text
      RsComissaoRepr("NF") = LcNumero
      RsComissaoRepr("Produto") = Item.TextMatrix(a, 1)
      RsComissaoRepr("QUANTIDADE") = CDbl(Item.TextMatrix(a, 5))
      RsComissaoRepr("VALORUNIT") = CDbl(Item.TextMatrix(a, 6))
      RsComissaoRepr("VALORTOTAL") = CDbl(Item.TextMatrix(a, 8))
    ' If Comissao.Text = "1" Then Ibaixo = True Else Ibaixo = False
    ' RsComissao("ITEMBAIXO") = Ibaixo
      LcDifAcr = LcPercAcres * LcValorComissao
      LcDifDes = LcPercDesc * LcValorComissao
      RsComissaoRepr("COMISSAO") = LcValorComissao + LcDifAcr - LcDifDes
      RsComissaoRepr("DATAVENDA") = CDate(DadosOrcamento.PrevCommisao.Text)
      RsComissaoRepr("CLIENTE") = CodigoCliente.Text
      RsComissaoRepr("percentual") = Item.TextMatrix(a, 14)
      LcCriterioFornec = "cod='" & Item.TextMatrix(a, 1) & "'"
      RsProduto.FindFirst LcCriterioFornec
      If Not RsProduto.NoMatch Then
         RsComissaoRepr("Fornecedor") = RsProduto!fornecedor
      End If
      RsComissaoRepr.Update
Next


'==== Atualiza os dados de credito do cliente

   LcCriterioBusca = "Codigo='" & CodigoCliente.Text & "'"
   RsCliente.FindFirst LcCriterioBusca
   If Not RsCliente.NoMatch Then
       RsCliente.Edit
       If LcNatureza = "A Prazo" Then RsCliente!CreditoUtilizado = RsCliente!CreditoUtilizado + CDbl(TotalOrcamento.Text)
       RsCliente!ULTCOMPRA = Format(emissao.Text, "dd/mm/yy")
       RsCliente.Update
   End If
'End If
'==== Atualizar Contas a Receber

If LcNatureza = "A Prazo" Then
   If GlFaturaSaida Then
      LcNumeroContas = CDbl(DadosOrcamento.Quantidade.Text)
      For a = 1 To LcNumeroContas
          RsContas.AddNew
          RsContas("NF") = Documento.Text & "/" & Right("00" & CStr(a), 2)
          RsContas("CLIENTE") = CodigoCliente.Text
          LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
          If Not RsTipoMonetario.NoMatch Then
             RsContas("TPMONET") = RsTipoMonetario("TPMONET")
          End If
          RsContas("DATA") = CDate(emissao.Text)
          RsContas("VALOR") = CDbl(DadosOrcamento.valor.Text)
          Select Case a
                 Case Is = 1
                      RsContas("DTVENC") = CDate(DadosOrcamento.Vencimento(0).Text)
                 Case Is = 2
                      RsContas("DTVENC") = CDate(DadosOrcamento.Vencimento(1).Text)
                 Case Is = 3
                      RsContas("DTVENC") = CDate(DadosOrcamento.Vencimento(2).Text)
           End Select
           RsContas("TIPORD") = "R"
           RsContas("Acrescimo") = 0
           RsContas.Update
        Next
    End If
    LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
    RsTipoMonetario.FindFirst LcCriterioPes
    If Not RsTipoMonetario.NoMatch Then
       LcTipoMOne = RsTipoMonetario("TPMONET")
    End If
    If GlVendaVista Then
       LcValor = CDbl(TotalOrcamento.Text)
       Call lancacaixa("Receita", Documento.Text, LcTipoMOne, LcValor)
    End If
Else
   If GlVistaSaida Then
      RsContas.AddNew
      RsContas("NF") = LcNumero
      RsContas("CLIENTE") = CodigoCliente.Text
      LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
      If Not RsTipoMonetario.NoMatch Then
         RsContas("TPMONET") = RsTipoMonetario("TPMONET")
      End If
      RsContas("VALOR") = CCur(TotalOrcamento.Text)
      RsContas("DATA") = CDate(emissao.Text)
      RsContas("DTVENC") = CDate(emissao.Text)
      RsContas("DTPAGTO") = CDate(emissao.Text)
      RsContas("VALPAGO") = CCur(TotalOrcamento.Text)
      RsContas("TIPORD") = "R"
      RsContas("Acrescimo") = 0
      RsContas.Update
    End If
          
    If GlCaixaSaida Then
       
      ' RsCaixa.AddNew
      ' RsCaixa("NF") = LcNumero
      ' RsCaixa("RECDESP") = "R"
      ' RsCaixa("CLICRED") = CodigoCliente.Text
       LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
       RsTipoMonetario.FindFirst LcCriterioPes
       If Not RsTipoMonetario.NoMatch Then
          LcTipoMOne = RsTipoMonetario("TPMONET")
       End If
       LcValor = CDbl(TotalOrcamento.Text)
       If GlVendaVista Then
          LcValor = CDbl(TotalOrcamento.Text)
          Call lancacaixa("Receita", Documento.Text, LcTipoMOne, LcValor)
       End If
      ' RsCaixa("VALOR") = CCur(TotalOrcamento.Text)
      ' RsCaixa("DATA") = CDate(emissao.Text)
      ' RsCaixa.Update
     End If
 End If
'=== Fechando as base de dados
SaidaSistema:
Command4.Caption = "Pes&quisa F7"
RsDados.Close
RsCliente.Close
RsProduto.Close
RsContas.Close
RsCaixa.Close
RsComissao.Close
RsTipoMonetario.Close
bb.Close

Set RsDados = Nothing
Set RsCliente = Nothing
Set RsProduto = Nothing
Set RsContas = Nothing
Set RsCaixa = Nothing
Set RsComissao = Nothing
Set RsTipoMonetario = Nothing
Set bb = Nothing
End Function
Function numeronota() As String
On Error Resume Next
Dim RsNota As Recordset, Dbn As Database
Set BbN = OpenDatabase(GLBase, False, False)
Set RsNota = BbN.OpenRecordset("Orcamento", dbOpenTable, dbSeeChanges, dbOptimistic)
RsNota.Index = "doc"
If Not RsNota.EOF Then
   RsNota.MoveLast
   numeronota = Right("000000" & CStr(CDbl(RsNota!doc) + 1), 6)
Else
   numeronota = "000001"
End If
RsNota.Close
BbN.Close
Set RsNota = Nothing
Set BbN = Nothing

End Function

Private Sub codigoproduto_GotFocus()
On Error Resume Next
LcLimpa = True
LcPerguntaPRoduto = False
If GlBuscaProdutoAgora Then Exit Sub
If Len(Trim(LimiteCredito.Text)) = 0 Then LimiteCredito.Text = "0"
If Len(Trim(LimiteUtilizado.Text)) = 0 Then LimiteUtilizado.Text = "0"
If CDbl(LimiteCredito.Text) < CDbl(LimiteUtilizado.Text) Then
   If GlSenha Then
      LiberacaoCli.Show , Me
   End If
   If GlSoOrcamento Then
      Natureza.Text = "Orçamento"
      Natureza.Locked = True
      Exit Sub
   End If
End If

End Sub

Private Sub codigoproduto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   GlCriterioSql = "select * From alid009 where nome like '" & NomeProduto.Text & "*' order by Nome"
   LcPerguntaPRoduto = False
   FrmBuscaProduto.Tag = NomeProduto.Text
   FrmBuscaProduto.Show , Me
End If

If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
If KeyCode = 123 Then UltimasComprasCliente.Show , Me
End Sub

Private Sub CodigoProduto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   codigoproduto.Text = ""
   LcLimpa = False
End If
End Sub

Private Sub CodigoProduto_LostFocus()
On Error Resume Next

Dim bb As Database, RsProduto As Recordset, RsUnidade As Recordset
'===> Verifica se deve Buscar Produto
'If Not LcBuscaCliente Then Exit Sub
If Len(Trim(codigoproduto.Text)) = 0 Then Exit Sub
   
'===> Verifica se o Valor digitado é Númerico
If GLCalculacodigoProduto Then
   If Not IsNumeric(codigoproduto.Text) And Len(codigoproduto.Text) > 0 Then
      MsgBox "O Código do Produto deve ser Numérico...", vbExclamation, "Aviso"
      codigoproduto.SetFocus
      Exit Sub
   End If
   codigoproduto.Text = Right("00000" & codigoproduto.Text, 5)
End If

Set bb = OpenDatabase(GLBase, False, False)
Set RsProduto = bb.OpenRecordset("select * From alid009 where COD='" & codigoproduto.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

If Not RsProduto.EOF Then
   Set RsUnidade = bb.OpenRecordset("select * From alid004 where COD='" & RsProduto!UNIMED & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)   ', dbOpenDynaset)
   NomeProduto.Text = RsProduto!nome
   ipi.Text = RsProduto!ipi & ""
   lcprocuraunidade = "cod='" & RsProduto!UNIMED & "'"
   If Not RsUnidade.EOF Then
      Unidade.Text = RsUnidade!Simbolo
      codigounidade.Text = RsUnidade!cod
   Else
      MsgBox "A unidade deste produto não Foi Cadastrada...", 64, "Aviso"
      Unidade.Text = ""
   End If
   If Len(RsProduto!Ptab) > 0 Then
      Unitario.Text = RsProduto!Ptab
      preconormal.Text = RsProduto!Ptab
   Else
      Unitario.Text = 0
   End If
   
   If Not IsNull(RsProduto!fornecedor) Then
      If Len(RsProduto!fornecedor) > 0 Then
         Industria.Text = RsProduto!fornecedor & ""
      Else
         If GlRepresentante Then MsgBox "Este Produto Não Está Vinculado a um Fabricante, Haverá erro no Lançamento da Comissão do Representante", 64, "Aviso"
      End If
   Else
      If GlRepresentante Then MsgBox "Este Produto Não Está Vinculado a um Fabricante, Haverá erro no Lançamento da Comissão do Representante", 64, "Aviso"
   End If
   If Not IsNull(RsProduto!RsProduto!ComissaoFornecedor) Then
      ComissaoFabricante.Text = RsProduto!ComissaoFornecedor & ""
   Else
      ComissaoFabricante.Text = 0
   End If
   Unidade.SetFocus
Else
'   LcPerguntaCliente = False
 '  FrmPesquisaCliente.Show , Me
    codigoproduto.Text = ""
End If
RsUnidade.Close
RsProduto.Close
bb.Close
End Sub

Private Sub CodigoVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   LcMontaComissao = True
   LcBuscaVendedor = False
   LcPerguntaVendedor = False
   FrmPesquisaFuncionarios.Show , Me
End If


If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
'If KeyCode = 122 Then Exibecomissao.Show , Me
'If KeyCode = 123 Then UltimasComprasCliente.Show , Me

End Sub

Private Sub CodigoVendedor_KeyPress(KeyAscii As Integer)
LcBuscaVendedor = True
End Sub

Private Sub CodigoVendedor_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   Natureza.SetFocus
End If
End Sub


Private Sub CodigoVendedor_LostFocus()
On Error Resume Next
Dim Bbs As Database, RsVendedor As Recordset
'===> Verifica se deve Buscar Produto
If Not LcBuscaVendedor Then Exit Sub

'===> Verifica se o Valor digitado é Númerico
If Not IsNumeric(codigoVendedor.Text) And Len(codigoVendedor.Text) > 0 Then
   MsgBox "O Código do Vendedor deve ser Numérico...", vbExclamation, "Aviso"
   codigoVendedor.SetFocus
   Exit Sub
End If
codigoVendedor.Text = Right("00000" & codigoVendedor.Text, 5)
Set Bbs = OpenDatabase(GLBase, False, False)
Set RsVendedor = Bbs.OpenRecordset("select * From alid200 where CODIGO='" & codigoVendedor.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
If Not RsVendedor.EOF Then
   NomeVendedor.Text = RsVendedor!nome
   If GlVariasComissao Then
      Exibecomissao.Show , Me
      NomeVendedor.SetFocus
   Else
      Comissao.Text = RsVendedor!Comissao
      ComissaoProduto.Text = RsVendedor!comisao
      CodigoCliente.SetFocus
   End If
   
Else
   LcPerguntaVendedor = False
   LcMontaComissao = True
   FrmPesquisaFuncionarios.Show , Me
End If
LcBuscaVendedor = True
RsVendedor.Close
bb.Close

End Sub



Private Sub Command3_Click()
On Error Resume Next
DadosOrcamento.Show , Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Command4.Caption = "Pes&quisa F7" Then
   LcSql1 = "Select * from orcamento"
   FrmPesquisaNota.Show , Me
   Command4.Caption = "&Incluir F7"
   LcPesquisa = True
Else
   Command4.Caption = "Pes&quisa F7"
   limpanota
   LcPesquisa = False
End If

End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub DEscontoItem_GotFocus()
On Error Resume Next
LcLimpa = True
If Len(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   
   Quantidade.SetFocus
   Exit Sub
End If
If CDbl(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   Quantidade.SetFocus
   Exit Sub
End If
End Sub

Private Sub DEscontoItem_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me

If KeyCode = 123 Then UltimasComprasCliente.Show , Me

If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me

End Sub

Private Sub DEscontoItem_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   DEscontoItem.Text = ""
   LcLimpa = False
End If
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub DEscontoItem_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   Unitario.SetFocus
End If
End Sub

Private Sub DEscontoItem_LostFocus()
On Error Resume Next
Dim b As Integer
Dim LCLEtra, LcDesconto As String
Dim LcValor, LcQuantidade, LcIpi, LcTotalProduto, LcGerado As Double
Dim LcAntigo As Double
If Len(DEscontoItem.Text) > 0 Then
   DescricaoDesconto.Enabled = False
   ValorDesconto.Enabled = False
End If

If Len(DEscontoItem.Text) = 0 Then Exit Sub

LcValor = CDbl(Unitario.Text)
LcAntigo = CDbl(Unitario.Text)
For b = 1 To Len(DEscontoItem.Text)
    LCLEtra = Mid(DEscontoItem.Text, b, 1)
    If LCLEtra = "+" Then
       If Len(LcDesconto) > 0 Then
          LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
          LcDesconto = ""
       End If
     Else
        If IsNumeric(LCLEtra) Then
           LcDesconto = LcDesconto & LCLEtra
        Else
           If LCLEtra = "," Then
              LcDesconto = LcDesconto & LCLEtra
           End If
        End If
     End If
Next
If Len(LcDesconto) > 0 Then
   LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
   LcDesconto = ""
End If
LcQuantidade = CDbl(Quantidade.Text)
If GlDescUnit Then
   Unitario.Text = AcertaNumero(CStr(LcValor), GlDecimais)
   'DEscontoItem.Text = 0
End If
total.Text = AcertaNumero(CStr(LcValor * LcQuantidade), GlDecimais)
LcGerado = ((LcAntigo - LcValor) / LcAntigo) * 100
DescontoGerado.Text = AcertaNumero(CStr(LcGerado), 2)

End Sub

Private Sub DescricaoAcrescimo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then MsgBox "Opção Não Disponivel.", 64, "Aviso"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub DescricaoAcrescimo_KeyPress(KeyAscii As Integer)

On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub DescricaoAcrescimo_LostFocus()
On Error Resume Next
Dim LcValor, LcDesconto, LcIpi, LcACrescimo, LcTotal, LcTotalProduto As Double

If Len(DescricaoAcrescimo.Text) = 0 Then Exit Sub
If Not IsNumeric(DescricaoAcrescimo.Text) Then
   MsgBox "O Percentual do Acréscimo deve ser um Valor Numérico...", vbExclamation, "Aviso"
   DescricaoAcrescimo.SetFocus
   Exit Sub
End If

If Len(ValorDesconto.Text) = 0 Then LcDesconto = 0 Else LcDesconto = CDbl(ValorDesconto.Text)
If Len(TotalIpi.Text) = 0 Then LcIpi = 0 Else LcIpi = CDbl(TotalIpi.Text)
If Len(TotalProduto.Text) = 0 Then LcTotalProduto = 0 Else LcTotalProduto = CDbl(TotalProduto.Text)
LcACrescimo = (CDbl(DescricaoAcrescimo.Text) / 100) * (LcTotalProduto - LcDesconto)
LcTotal = LcTotalProduto - LcDesconto + LcACrescimo
If GlIpi Then
   LcTotal = LcTotal + LcIpi
End If
TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)
ValorAcrescimo.Text = AcertaNumero(CStr(LcACrescimo), 2)

End Sub

Private Sub DescricaoDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then MsgBox "Opção Não Disponivel.", 64, "Aviso"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub DescricaoDesconto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub DescricaoDesconto_LostFocus()
On Error Resume Next
Dim a, b As Integer
Dim LCLEtra, LcDesconto As String
Dim LcValor, LcQuantidade, LcIpi, LcTotalProduto As Double
If GlDetalhaDesconto Then
  For a = 1 To Item.Rows - 1
            
      LcValor = CDbl(Item.TextMatrix(a, 11))
      LcQuantidade = CDbl(Item.TextMatrix(a, 5))
      
      For b = 1 To Len(DescricaoDesconto.Text)
          LCLEtra = Mid(DescricaoDesconto.Text, b, 1)
          If LCLEtra = "+" Then
             If Len(LcDesconto) > 0 Then
                LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
                LcDesconto = ""
              End If
          Else
             If IsNumeric(LCLEtra) Then
                LcDesconto = LcDesconto & LCLEtra
              Else
                If LCLEtra = "," Then
                   LcDesconto = LcDesconto & LCLEtra
                End If
              End If
          End If
      Next
      If Len(LcDesconto) > 0 Then
         LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
         LcDesconto = ""
      End If
      Item.TextMatrix(a, 6) = AcertaNumero(CStr(LcValor), GlDecimais)
      Item.TextMatrix(a, 8) = AcertaNumero(CStr(LcValor * LcQuantidade), GlDecimais)
      LcIpi = LcIpi + ((CDbl(Item.TextMatrix(a, 9)) / 100) * (LcValor * LcQuantidade))
      LcTotalProduto = LcTotalProduto + (LcValor * LcQuantidade)
    Next
    If GlIpi Then TotalIpi.Text = AcertaNumero(CStr(LcIpi), 2)
    TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
    TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto + LcIpi), 2)
Else
   If DescricaoDesconto.Text = "0" Or Len(DescricaoDesconto.Text) = 0 Then
      ValorDesconto.Text = 0
   Else
      If Len(ValorDesconto.Text) = 0 Then ValorDesconto.Text = 0
         LcValor = CDbl(TotalProduto.Text)
      For a = 1 To Len(DescricaoDesconto.Text)
      'LcQuantidade = CDbl(Item.TextMatrix(a, 5))
       LCLEtra = Mid(DescricaoDesconto.Text, a, 1)
          If LCLEtra = "+" Then
             If Len(LcDesconto) > 0 Then
                LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
                LcDesconto = ""
              End If
          Else
             If IsNumeric(LCLEtra) Then
                LcDesconto = LcDesconto & LCLEtra
              Else
                If LCLEtra = "," Then
                   LcDesconto = LcDesconto & LCLEtra
                End If
              End If
          End If
     Next
     
     If Len(LcDesconto) > 0 Then
        LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
        LcDesconto = ""
     End If
     If GlIpi Then
        TotalOrcamento.Text = AcertaNumero(CStr(LcValor + CDbl(TotalIpi.Text)), 2)
     Else
        TotalOrcamento.Text = AcertaNumero(CStr(LcValor), 2)
     End If
     ValorDesconto.Text = AcertaNumero(CStr(CDbl(TotalProduto.Text) - LcValor), 2)
   End If
End If

                
End Sub

Private Sub Documento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub emissao_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub
Function MontaUnidade()
On Error Resume Next
Dim bb As Database
LcQUn = 0
Dim LcAchou As Integer
Dim RsUnidade As Recordset
Dim LcPrimeiro As String
Set bb = OpenDatabase(GLBase, False, False)
Set RsUnidade = bb.OpenRecordset("select * From alid004 order By SIMBOLO", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsUnidade.EOF
   Unidade.AddItem RsUnidade!Simbolo
   RsUnidade.MoveNext
Loop
If LcQUn > 0 Then LcQUn = LcQUn - 1
RsUnidade.Close
bb.Close
Set RsUnidade = Nothing
Set bb = Nothing

End Function
Function LimpaTela()
On Error Resume Next
If GlLimpaTelaOrc Then
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Item.Height = 4215
    Item.Top = 2520
    Line2(1).Visible = False
Else
    Item.Height = 3375
    Item.Top = 3360
End If
End Function
Private Sub Form_Load()
On Error Resume Next
Label3(7).Visible = GlRepresentante
Representada.Visible = GlRepresentante
GeraGrid
MontaUnidade
emissao.Text = Format(Date, "dd/mm/yy")
Status = "Em Lançamento"
LimpaTela
'=== seta as variaveis para incio
LcBuscaVendedor = True
LcBuscaCliente = True
LcCalculatotal = True
LcCalculaIpi = True
LcCalculaDesconto = True
LcCalculaAcrescimo = True
LcGeraComissao = True
LcPerguntaVendedor = True
LcPerguntaCliente = True
GlLibera = False
LcItem = 0
ValorDesconto.Enabled = Not GlDetalhaDesconto
ValorAcrescimo.Enabled = Not GlAcrescimo
codigoVendedor.TabStop = GlEsclheVendedor
codigoVendedor.Locked = Not GlEsclheVendedor
NomeVendedor.Locked = Not GlEsclheVendedor
NomeVendedor.TabStop = GlEsclheVendedor
CodigoCliente.TabStop = GlEscolheCliente
NomeCliente.TabStop = GlEscolheCliente
CodigoCliente.Locked = Not GlEscolheCliente
NomeCliente.Locked = Not GlEscolheCliente
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me

If KeyCode = 123 Then UltimasComprasCliente.Show , Me

If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub Natureza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Natureza_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   emissao.SetFocus
End If
End Sub


Private Sub NomeCliente_Change()
 'GlLibera = False
End Sub

Private Sub NomeCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   GlCriterioSql = "where RAZAOSOC like '" & NomeCliente.Text & "*' order by RAZAOSOC"
   LcPerguntaCliente = False
   FrmBuscaCliente.Show , Me
End If

If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub NomeCliente_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then LcBuscaCliente = True
End Sub

Private Sub NomeCliente_LostFocus()

On Error Resume Next
Dim bb As Database, RsCliente As Recordset
'===> Verifica se deve Buscar Produto
If Not LcBuscaCliente Then Exit Sub
If Len(Trim(CodigoCliente.Text)) > 0 Then Exit Sub
'===> Verifica se o Valor digitado é Númerico
Set bb = OpenDatabase(GLBase, False, False)
Set RsCliente = bb.OpenRecordset("select * From alid001 where RAZAOSOC='" & NomeCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
If Not RsCliente.EOF Then
   NomeCliente.Text = RsCliente!razaosoc
   CodigoCliente.Text = RsCliente!codigo
   codigoproduto.SetFocus
Else
   GlCriterioSql = "where RAZAOSOC like '" & NomeCliente.Text & "*' order by RAZAOSOC"
   LcPerguntaCliente = False
   FrmBuscaCliente.Show , Me
End If
RsCliente.Close
bb.Close
End Sub

Private Sub NomeProduto_GotFocus()
LcPerguntaPRoduto = False
End Sub

Private Sub NomeProduto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   GlCriterioSql = "select * From alid009 where nome like '" & NomeProduto.Text & "*' order by Nome"
   LcPerguntaPRoduto = False
   FrmBuscaProduto.Tag = NomeProduto.Text
   FrmBuscaProduto.Show , Me
End If
If KeyCode = 123 Then UltimasComprasCliente.Show , Me
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub NomeProduto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   codigoproduto.SetFocus
End If
End Sub

Private Sub NomeProduto_LostFocus()
On Error Resume Next
Dim bb As Database, RsProduto As Recordset, RsUnidade As Recordset
'===> Verifica se deve Buscar Produto
'If Not LcBuscaCliente Then Exit Sub
If Len(NomeProduto.Text) = 0 Then Exit Sub
If Len(Trim(codigoproduto.Text)) > 0 Then Exit Sub
'===> Verifica se o Valor digitado é Númerico
Set bb = OpenDatabase(GLBase, False, False)
Set RsProduto = bb.OpenRecordset("select * From alid009 where nome='" & NomeProduto.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
Set RsUnidade = bb.OpenRecordset("alid004", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

If Not RsProduto.EOF Then
   NomeProduto.Text = RsProduto!nome
   codigoproduto.Text = RsProduto!cod
   lcprocuraunidade = "cod='" & RsProduto!UNIMED & "'"
   ipi.Text = RsProduto!ipi & ""
   RsUnidade.FindFirst lcprocuraunidade
   If Not RsUnidade.NoMatch Then
      Unidade.Text = RsUnidade!Simbolo
      codigounidade.Text = RsUnidade!cod
   End If
   If Len(RsProduto!Ptab) > 0 Then
      Unitario.Text = RsProduto!Ptab
      preconormal.Text = RsProduto!Ptab
   Else
      Unitario.Text = 0
   End If
   Industria.Text = RsProduto!fornecedor & ""
   If Not IsNull(RsProduto!RsProduto!ComissaoFornecedor) Then
      ComissaoFabricante.Text = RsProduto!ComissaoFornecedor
   Else
      ComissaoFabricante.Text = 0
   End If
   Unidade.SetFocus
   
Else
   GlCriterioSql = "select * From alid009 where nome like '" & NomeProduto.Text & "*' order by Nome"
   LcPerguntaPRoduto = False
   LcPerguntaPRoduto = True
   FrmBuscaProduto.Tag = NomeProduto.Text
   FrmBuscaProduto.Show , Me
End If
RsProduto.Close
bb.Close
End Sub

Private Sub NomeVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 116 Then
   LcMontaComissao = True
   LcBuscaVendedor = False
   LcPerguntaVendedor = False
   FrmPesquisaFuncionarios.Show , Me
End If

If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
'If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub NomeVendedor_LostFocus()
On Error Resume Next
If Not LcPerguntaVendedor Then Exit Sub
If Len(Trim(NomeVendedor.Text)) > 0 Then
  ' If GlVariasComissao Then Exibecomissao.Show , Me
Else
   MsgBox "É Necessário Escolher um Vendedor.", vbExclamation, "Aviso"
   codigoVendedor.SetFocus
End If

End Sub

Private Sub quantidade_Change()
On Error Resume Next
Dim LcQuantidade, LcUnitario, LcTotal As Double

If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub
If Not IsNumeric(Quantidade.Text) Then
   MsgBox "A Quantidade deve ser um Valor Numérico.", vbExclamation, "Aviso"
   Quantidade.Text = ""
   Exit Sub
End If
If Len(Unitario.Text) = 0 Then LcUnitario = 0 Else LcUnitario = Unitario.Text
If Len(Quantidade.Text) = 0 Then LcQuantidade = 0 Else LcQuantidade = Quantidade.Text
LcTotal = LcQuantidade * LcUnitario
total.Text = AcertaNumero(CStr(LcTotal), 2)


End Sub

Private Sub Quantidade_GotFocus()
LcLimpa = True
End Sub


Private Sub Quantidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 123 Then UltimasComprasCliente.Show , Me
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 123 Then DescontoRepresentante.Show , Me
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   Quantidade.Text = ""
   LcLimpa = False
End If

If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Quantidade_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   Acomodacao.SetFocus
End If
End Sub

Private Sub Representada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Status_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub total_GotFocus()
On Error Resume Next
total.Tag = total.Text
LcLimpa = True
LcFechaitem = True
If Len(Unitario.Text) = 0 Then
   MsgBox "É Necessário Informar o Valor Unitário de Venda...", vbInformation, "Aviso"
   Exit Sub
End If
If CDbl(Unitario.Text) = 0 Then
   MsgBox "É Necessário Informar o Valor Unitário de Venda...", vbInformation, "Aviso"
   Exit Sub
End If
total.Text = total.Tag
End Sub

Private Sub Total_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Function ImprimeorcamentoWindows()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset, RsIndustria As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim Lcemail, LcFax, LcCep, Lccelular, Lccelular1 As String
Dim LcI As Integer
AbreBase
LcSqlIn = "Select * from alid002 where codigo='" & Industria.Text & "'"
Set RsIndustria = Dbbase.OpenRecordset(LcSqlIn, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsIndustria.EOF Then
   LCNInd = RsIndustria!razaosoc & ""
End If
RsIndustria.Close
Set RsIndustria = Nothing

LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!bairro & "  " & RsEmpresa!cidade & "  " & RsEmpresa!Cep
   LcFone = RsEmpresa!fone & ""
   Lcemail = RsEmpresa!email & ""
   LcFax = RsEmpresa!Fax & ""
   LcCep = RsEmpresa!Cep & ""
   Lccelular = RsEmpresa!celular & ""
   Lccelular1 = RsEmpresa!celular1 & ""
End If

If GlFuncCodigo Then LcVen = codigoVendedor.Text
If GlFuncNome Then LcVen = NomeVendedor.Text
If GlFuncEmpresa Then LcVen = LcEmpresa

'Abertura do relatório de vendas
    
    
    CryRelatorio.DataFiles(0) = GLBase
    CryRelatorio.ReportFileName = App.Path & "\pedidovendas.rpt"
    LcFormula = "{orcamento.doc}='" & UCase(Documento.Text) & "'"
    CryRelatorio.CopiesToPrinter = 1

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Orçamento"
LcI = 3
CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"

If Len(LcFax) > 0 Then
   CryRelatorio.Formulas(2) = "Fone='Fone " & LcFone & " Fax " & LcFax & "'"
Else
   CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
End If
If Len(Lcemail) > 0 Then
   CryRelatorio.Formulas(LcI) = "email='" & Lcemail & "'"
   LcI = LcI + 1
End If

If Len(Lccelular) > 0 Then
   CryRelatorio.Formulas(LcI) = "celular='" & Lccelular & "'"
   LcI = LcI + 1
End If
If Len(Lccelular1) > 0 Then
   CryRelatorio.Formulas(LcI) = "celular1='" & Lccelular1 & "'"
   LcI = LcI + 1
End If

CryRelatorio.Formulas(3) = "vend='" & LcVen & "'"
CryRelatorio.Formulas(4) = "ind='" & LCNInd & "'"
'CryRelatorio.Formulas(5) = "titulo='Produtos'"
If DadosOrcamento.Visualizar.Value = 1 Then
    LcTipoSaida = 0
Else
    LcTipoSaida = 1
End If
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Function


Private Sub total_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   Rem total.Text = ""
   LcLimpa = False
End If
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub total_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   LcFechaitem = False
   DEscontoItem.SetFocus
End If
End Sub

Private Sub total_LostFocus()
On Error Resume Next
Dim a As Integer
Dim LcRow As Long
Dim LcTotalProduto, LcTotalNota, LcDesconto, LcACrescimo As Double
Dim LcTotalIpi As Double
'Dim LcDesconto As String
Dim LcValor As Double
'==== Verifica a digitação dos dados.
On Error Resume Next
If Not LcFechaitem Then Exit Sub
If Len(codigoproduto.Text) = 0 Then
   MsgBox "É Necessário Informar o Produto..", vbInformation, "Aviso"
   Exit Sub
End If

If Len(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   Exit Sub
End If
If CDbl(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   Exit Sub
End If

If Len(Unitario.Text) = 0 Then
   MsgBox "É Necessário Informar o Valor Unitário de Venda...", vbInformation, "Aviso"
   Exit Sub
End If
If CDbl(Unitario.Text) = 0 Then
   MsgBox "É Necessário Informar o Valor Unitário de Venda...", vbInformation, "Aviso"
   Exit Sub
End If
LcLiberaCalculo = False


'===> Monta o Grid
LcItem = LcItem + 1
Item.Rows = LcItem + 1
LcRow = LcItem
 Item.TextMatrix(LcRow, 0) = Right("000" & LcItem, 3)
   Item.TextMatrix(LcRow, 1) = UCase(codigoproduto.Text)
   Item.TextMatrix(LcRow, 2) = NomeProduto.Text
   Item.TextMatrix(LcRow, 3) = Unidade.Text & ""
   Item.TextMatrix(LcRow, 4) = Acomodacao.Text & ""
   Item.TextMatrix(LcRow, 5) = Quantidade.Text
  


If GlDetalhaDesconto Then
   If Len(DescricaoDesconto.Text) > 0 Then
      '=== Rateia Desconto para Este Item
      LcValor = CDbl(Unitario.Text)
      For a = 1 To Len(DescricaoDesconto.Text)
          LCLEtra = Mid(DescricaoDesconto.Text, a, 1)
          If LCLEtra = "+" Then
             If Len(LcDesconto) > 0 Then
                LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
                LcDesconto = ""
             End If
           Else
             If IsNumeric(LCLEtra) Then
                LcDesconto = LcDesconto & LCLEtra
              Else
                If LCLEtra = "," Then
                   LcDesconto = LcDesconto & LCLEtra
                End If
              End If
           End If
       Next
       If Len(LcDesconto) > 0 Then
          LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
          LcDesconto = ""
       End If
     
     Item.TextMatrix(LcRow, 6) = AcertaNumero(CStr(LcValor), GlDecimais)
     Item.TextMatrix(LcRow, 8) = AcertaNumero(CStr(LcValor * CDbl(Quantidade.Text)), GlDecimais)
   Else
    Item.TextMatrix(LcRow, 6) = AcertaNumero(Unitario.Text, GlDecimais)
    Item.TextMatrix(LcRow, 8) = AcertaNumero(total.Text, GlDecimais)
   End If
Else
   If Len(DescricaoDesconto.Text) > 0 Then
      MsgBox "Foi Incluido um Ítem Na Lista de Produtos," & Chr(13) & "O Desconto dado Anteriormente será Ignorado.", vbInformation, "Aviso"
   End If
   DescricaoDesconto.Text = ""
   ValorDesconto.Text = ""
   DescricaoAcrescimo.Text = ""
   ValorAcrescimo.Text = ""
   Item.TextMatrix(LcRow, 6) = AcertaNumero(Unitario.Text, GlDecimais) & ""
   Item.TextMatrix(LcRow, 8) = AcertaNumero(total.Text, GlDecimais) & ""
End If
If GlRateiaAcrecimo Then
   LcValor = CDbl(Item.TextMatrix(LcRow, 6))
   If Len(DescricaoAcrescimo.Text) > 0 Then
      LcValor = AcertaNumero(CStr(LcValor + (CDbl((DescricaoAcrescimo.Text) / 100) * LcValor)), GlDecimais)
   End If
   Item.TextMatrix(LcRow, 6) = AcertaNumero(CStr(LcValor), GlDecimais)
   Item.TextMatrix(LcRow, 8) = AcertaNumero(CStr(LcValor * CDbl(Quantidade.Text)), GlDecimais)
End If

If GlIpi Then
   Item.TextMatrix(LcRow, 9) = ipi.Text & ""
End If
If Len(ComissaoProduto.Text) = 0 Then ComissaoProduto.Text = "0"
Item.TextMatrix(LcRow, 10) = ComissaoProduto.Text
Item.TextMatrix(LcRow, 11) = preconormal.Text
Item.TextMatrix(LcRow, 7) = DescontoGerado.Text
Item.TextMatrix(LcRow, 12) = acrescimo.Text
Item.TextMatrix(LcRow, 13) = codigounidade.Text
Item.TextMatrix(LcRow, 14) = ComissaoFabricante.Text
Item.TextMatrix(LcRow, 15) = DEscontoItem.Text
'===> Calcula total produto
LcTotalIpi = 0
For a = 1 To Item.Rows - 1
    LcTotalProduto = LcTotalProduto + CDbl(AcertaNumero(Item.TextMatrix(a, 8), GlDecimais))
    LcTotalIpi = LcTotalIpi + ((CDbl(AcertaNumero(Item.TextMatrix(a, 9), GlDecimais)) / 100) * CDbl(AcertaNumero(Item.TextMatrix(a, 8), GlDecimais)))
Next
TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
If GlIpi Then TotalIpi.Text = AcertaNumero(CStr(LcTotalIpi), 2)

If Not GlDetalhaDesconto Then
  If Len(DescricaoDesconto.Text) > 0 Then
      LcValor = CDbl(TotalProduto.Text)
      For a = 1 To Len(DescricaoDesconto.Text)
      'LcQuantidade = CDbl(Item.TextMatrix(a, 5))
          LCLEtra = Mid(DescricaoDesconto.Text, a, 1)
          If LCLEtra = "+" Then
             If Len(LcDesconto) > 0 Then
                LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
                LcDesconto = ""
              End If
          Else
             If IsNumeric(LCLEtra) Then
                LcDesconto = LcDesconto & LCLEtra
              Else
                If LCLEtra = "," Then
                   LcDesconto = LcDesconto & LCLEtra
                End If
              End If
          End If
      Next
      If Len(LcDesconto) > 0 Then
        LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
        LcDesconto = ""
     End If
     If GlIpi Then
         TotalOrcamento.Text = AcertaNumero(CStr(LcValor + CDbl(TotalIpi.Text)), 2)
     Else
         TotalOrcamento.Text = AcertaNumero(CStr(LcValor), 2)
     End If
     ValorDesconto.Text = AcertaNumero(CStr(CDbl(TotalProduto.Text) - LcValor), 2)
   Else
     If Len(ValorDesconto.Text) > 0 Then
        If GlIpi Then
           If Len(TotalIpi.Text) = 0 Then TotalIpi = 0
           TotalOrcamento.Text = AcertaNumero(CStr(TotalProduto + CDbl(ValorDesconto.Text) + CDbl(TotalIpi.Text)), 2)
        Else
           TotalOrcamento.Text = AcertaNumero(CStr(TotalProduto + CDbl(ValorDesconto.Text)), 2)
        End If
     End If
   End If
Else
 If GlIpi Then
    TotalOrcamento.Text = AcertaNumero(CStr(TotalIpi + LcTotalProduto), 2)
 Else
    TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto), 2)
 End If
End If
If Not GlRateiaAcrecimo Then
   If Len(ValorAcrescimo.Text) = 0 Then ValorAcrescimo.Text = 0
   If Len(TotalIpi.Text) = 0 Then TotalIpi.Text = 0
   LcProduto = CDbl(TotalProduto.Text) - CDbl(ValorDesconto.Text)
   LcTotalIpi = CDbl(TotalIpi.Text)
   If Len(DescricaoAcrescimo.Text) > 0 Then
      LcACrescimo = ((CDbl(DescricaoAcrescimo.Text) / 100) * LcProduto)
      LcTotal = LcProduto + LcACrescimo
      If GlIpi Then
         LcTotal = LcTotal + LcTotalIpi
      End If
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)
      ValorAcrescimo.Text = AcertaNumero(CStr(LcACrescimo), 2)
   Else
      If Len(ValorAcrescimo.Text) > 0 Then
        If Len(ValorDesconto.Text) = 0 Then ValorDesconto.Text = 0
        If Len(TotalIpi.Text) = 0 Then TotalIpi.Text = 0
        LcACrescimo = LcTotal + CDbl(ValorAcrescimo.Text)
        LcProduto = CDbl(TotalProduto.Text)
        LcTotal = LcProduto + LcACrescimo
        If GlIpi Then
           LcTotal = LcTotal + LcTotalIpi
        End If
        TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)
      End If
   End If
End If
'DescricaoDesconto.SetFocus
codigoproduto.Text = ""
NomeProduto.Text = ""
Unidade.Text = ""
Acomodacao.Text = ""
Quantidade.Text = ""
Unidade.Text = ""
total.Text = ""
ComissaoFabricante.Text = ""
ipi.Text = ""
Unitario.Text = ""
codigounidade.Text = ""
ComissaoProduto.Text = Comissao.Text
preconormal.Text = ""
acrescimo.Text = ""
DescontoGerado.Text = ""
DEscontoItem.Text = ""
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True
codigoproduto.SetFocus
LcLiberaCalculo = True
End Sub

Private Sub TotalIpi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub TotalOrcamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then MsgBox "Opção Não Disponivel.", 64, "Aviso"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub TotalProduto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub Unidade_GotFocus()
On Error Resume Next
If LcPerguntaPRoduto Then Exit Sub
If Len(codigoproduto.Text) = 0 Then
   MsgBox "É Necessário Informar o Produto..", vbInformation, "Aviso"
   codigoproduto.SetFocus
   Exit Sub
End If

End Sub

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 123 Then UltimasComprasCliente.Show , Me
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub Unidade_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   NomeProduto.SetFocus
End If
End Sub

Private Sub Unitario_Change()
On Error Resume Next
Dim LcQuantidade, LcUnitario, LcTotal As Double

If Len(Trim(Unitario.Text)) = 0 Then Exit Sub
If Not IsNumeric(Unitario.Text) Then
   MsgBox "O Valor Unitário deve ser um Valor Numérico.", vbExclamation, "Aviso"
   Unitario.Text = ""
   Exit Sub
End If
If Len(Unitario.Text) = 0 Then LcUnitario = 0 Else LcUnitario = Unitario.Text
If Len(Quantidade.Text) = 0 Then LcQuantidade = 0 Else LcQuantidade = Quantidade.Text
LcTotal = LcQuantidade * LcUnitario
total.Text = AcertaNumero(CStr(LcTotal), GlDecimais)
End Sub

Private Sub Unitario_GotFocus()
On Error Resume Next
LcLimpa = True
If Len(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   
   Quantidade.SetFocus
   Exit Sub
End If
If CDbl(Quantidade.Text) = 0 Then
   MsgBox "É Necessário Informar a Quantidade de Venda...", vbInformation, "Aviso"
   Quantidade.SetFocus
   Exit Sub
End If
End Sub

Private Sub Unitario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me

If KeyCode = 123 Then UltimasComprasCliente.Show , Me

If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
If KeyCode = 122 Then Exibecomissao.Show , Me
End Sub

Private Sub Unitario_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   Unitario.Text = ""
   LcLimpa = False
End If
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Unitario_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
   Quantidade.SetFocus
End If
End Sub

Private Sub Unitario_LostFocus()
On Error Resume Next
preconormal.Text = Unitario.Text
Unitario.Text = AcertaNumero(Unitario.Text, GlDecimais)

End Sub

Private Sub ValorAcrescimo_Change()
On Error Resume Next
If Not LcLiberaCalculo Then Exit Sub
Dim LcValor, LcDesconto, LcIpi, LcACrescimo, LcTotal, LcTotalProduto As Double
If Len(ValorDesconto.Text) = 0 Then LcDesconto = 0 Else LcDesconto = CDbl(ValorDesconto.Text)
If Len(ValorAcrescimo.Text) = 0 Then LcACrescimo = 0 Else LcACrescimo = CDbl(ValorAcrescimo.Text)
If Len(TotalIpi.Text) = 0 Then LcIpi = 0 Else LcIpi = CDbl(TotalIpi.Text)
If Len(TotalProduto.Text) = 0 Then LcTotalProduto = 0 Else LcTotalProduto = CDbl(TotalProduto.Text)
LcTotal = LcTotalProduto - LcDesconto + LcACrescimo
If GlIpi Then
   LcTotal = LcTotal + LcIpi
End If
TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)
End Sub

Private Sub ValorAcrescimo_GotFocus()
LcLimpa = True
End Sub


Private Sub ValorAcrescimo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then MsgBox "Opção Não Disponivel.", 64, "Aviso"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub

Private Sub ValorAcrescimo_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   ValorAcrescimo.Text = ""
   LcLimpa = False
   DescricaoAcrescimo.Text = ""
End If

If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub ValorDesconto_Change()
On Error Resume Next
If Not LcLiberaCalculo Then Exit Sub
Dim LcValor, LcDesconto, LcIpi, LcACrescimo, LcTotal, LcTotalProduto As Double
If Len(ValorDesconto.Text) = 0 Then LcDesconto = 0 Else LcDesconto = CDbl(ValorDesconto.Text)
If Len(ValorAcrescimo.Text) = 0 Then LcACrescimo = 0 Else LcACrescimo = CDbl(ValorAcrescimo.Text)
If Len(TotalIpi.Text) = 0 Then LcIpi = 0 Else LcIpi = CDbl(TotalIpi.Text)
If Len(TotalProduto.Text) = 0 Then LcTotalProduto = 0 Else LcTotalProduto = CDbl(TotalProduto.Text)
LcTotal = LcTotalProduto - LcDesconto + LcACrescimo
If GlIpi Then
   LcTotal = LcTotal + LcIpi
End If
TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)

End Sub

Private Sub ValorDesconto_GotFocus()
LcLimpa = True
End Sub


Private Sub ValorDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 117 Then DescricaoDesconto.SetFocus
If KeyCode = 120 Then DescricaoAcrescimo.SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
If KeyCode = 119 Then MsgBox "Opção Não Disponivel.", 64, "Aviso"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{G}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{S}"
End Sub
Function GeraGrid()
On Error Resume Next
Item.ColAlignment(0) = 7
Item.ColAlignment(1) = 3
Item.ColAlignment(2) = 1
Item.ColAlignment(3) = 3
Item.ColAlignment(4) = 1
Item.ColAlignment(5) = 3
Item.ColAlignment(6) = 8
Item.ColAlignment(7) = 8
Item.ColAlignment(8) = 3
Item.ColAlignment(9) = 3


Item.ColWidth(0) = 500
Item.ColWidth(1) = 1100
Item.ColWidth(2) = 4600
Item.ColWidth(3) = 550
Item.ColWidth(4) = 800
Item.ColWidth(5) = 900
Item.ColWidth(6) = 1200
Item.ColWidth(7) = 900
Item.ColWidth(8) = 1200
Item.ColWidth(9) = 600
Item.ColWidth(10) = 0
Item.ColWidth(11) = 0
Item.ColWidth(12) = 0
Item.ColWidth(13) = 0
Item.ColWidth(14) = 0
Item.ColWidth(15) = 0

Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Código"
Item.TextMatrix(0, 2) = "Descrição"
Item.TextMatrix(0, 3) = "Unid."
Item.TextMatrix(0, 4) = "Ac."
Item.TextMatrix(0, 5) = "Quant"
Item.TextMatrix(0, 6) = "Unitário"
Item.TextMatrix(0, 7) = "Desconto"
Item.TextMatrix(0, 8) = "Total"
If GlIpi Then
   Item.TextMatrix(0, 9) = "IPI"
Else
   Item.TextMatrix(0, 9) = "ICMS"
End If

End Function

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   ValorDesconto.Text = ""
   LcLimpa = False
   DescricaoDesconto.Text = ""
End If
If KeyAscii = 46 Then KeyAscii = 44
End Sub
Function ExcluiItem(LcItemExcluir As String)
On Error Resume Next
Dim LcTotalProduto, LcTotalNota, LcTotalIpi As Double
Dim LcTotalDesconto, LcTotalAcrescimo As Double
Dim LcAcha, a As Integer

If Len(LcItemExcluir) = 0 Then
   MsgBox "Não Foi Digitado Nenhum item para Excluir...", vbExclamation, "Aviso"
   Exit Function
End If
LcItemExcluir = Right("000" & LcItemExcluir, 3)


'===> Busca o item para excluir
For a = 1 To Item.Rows - 1
    If Item.TextMatrix(a, 0) = LcItemExcluir Then
      If Item.Rows - 1 = 1 Then
        If a = 1 Then
          For lro = 0 To 13
              Item.TextMatrix(1, lro) = ""
          Next
          TotalProduto.Text = ""
          TotalOrcamento.Text = ""
          TotalIpi.Text = ""
          ValorDesconto.Text = ""
          ValorAcrescimo.Text = ""
          DescricaoDesconto.Text = ""
          DescricaoAcrescimo.Text = ""
          LcItem = 0
          Exit Function
       End If
     Else
      Item.RemoveItem (a)
    End If
    Exit For
  End If
Next
LcItem = 1
For a = 1 To Item.Rows - 1
    Item.TextMatrix(a, 0) = Right("000" & LcItem, 3)
    LcItem = LcItem + 1
Next
LcItem = LcItem - 1
LcTotalProduto = 0
If GlDetalhaDesconto Then
   For a = 1 To Item.Rows - 1
       LcTotalProduto = LcTotalProduto + CDbl(Item.TextMatrix(a, 8))
       LcTotalIpi = LcTotalIpi + (CDbl(Item.TextMatrix(a, 9)) / 100) * CDbl(Item.TextMatrix(a, 8))
   Next
   ValorDesconto.Text = "0"
   If GlIpi Then
      TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto + LcTotalIpi), 2)
      TotalIpi.Text = AcertaNumero(CStr(LcTotalIpi), 2)
    Else
      TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      TotalIpi.Text = AcertaNumero(CStr(0), 2)
    End If
Else
   If Len(DescricaoDesconto.Text) > 0 Then
      MsgBox "Foi Incluido um Ítem Na Lista de Produtos," & Chr(13) & "O Desconto dado Anteriormente será Ignorado.", vbInformation, "Aviso"
   End If
   DescricaoDesconto.Text = ""
   ValorDesconto.Text = ""
   DescricaoAcrescimo.Text = ""
   ValorAcrescimo.Text = ""
   For a = 1 To Item.Rows - 1
       LcTotalProduto = LcTotalProduto + CDbl(Item.TextMatrix(a, 8))
       LcTotalIpi = LcTotalIpi + ((CDbl(Item.TextMatrix(a, 9)) / 100) * CDbl(Item.TextMatrix(a, 8)))
   Next
   If Len(DescricaoDesconto.Text) > 0 Then
      TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      LcValor = CDbl(LcTotalProduto)
      For a = 1 To Len(DescricaoDesconto.Text)
      'LcQuantidade = CDbl(Item.TextMatrix(a, 5))
          LCLEtra = Mid(DescricaoDesconto.Text, a, 1)
          If LCLEtra = "+" Then
             If Len(LcDesconto) > 0 Then
                LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
                LcDesconto = ""
              End If
          Else
             If IsNumeric(LCLEtra) Then
                LcDesconto = LcDesconto & LCLEtra
              Else
                If LCLEtra = "," Then
                   LcDesconto = LcDesconto & LCLEtra
                End If
              End If
          End If
      Next
      If Len(LcDesconto) > 0 Then
        LcValor = LcValor - ((CDbl(LcDesconto) / 100) * LcValor)
        LcDesconto = ""
     End If
     If GlIpi Then
         TotalOrcamento.Text = AcertaNumero(CStr(LcValor + CDbl(TotalIpi.Text)), 2)
     Else
         TotalOrcamento.Text = AcertaNumero(CStr(LcValor), 2)
     End If
     ValorDesconto.Text = AcertaNumero(CStr(CDbl(TotalProduto.Text) - LcValor), 2)
   Else
     If Len(ValorDesconto.Text) > 0 Then
        If GlIpi Then
           If Len(TotalIpi.Text) = 0 Then TotalIpi = 0
           TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto + CDbl(ValorDesconto.Text) + CDbl(TotalIpi.Text)), 2)
        Else
           TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto + CDbl(ValorDesconto.Text)), 2)
        End If
        TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
     Else
        
     End If
   End If
End If
If Not GlRateiaAcrecimo Then
     LcProduto = CDbl(TotalProduto.Text)
     LcACrescimo = ((CDbl(DescricaoAcrescimo.Text) / 100) * LcProduto)
      LcTotal = LcProduto + LcACrescimo
      If GlIpi Then
         LcTotal = LcTotal + LcTotalIpi
      End If
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotal), 2)
      ValorAcrescimo.Text = AcertaNumero(CStr(LcACrescimo), 2)
End If
If Not GlRateiaAcrecimo And Not GlDetalhaDesconto Then
    LcTotalIpi = 0
    LcTotalProduto = 0
    For a = 1 To Item.Rows - 1
       LcTotalProduto = LcTotalProduto + CDbl(Item.TextMatrix(a, 8))
       LcTotalIpi = LcTotalIpi + ((CDbl(Item.TextMatrix(a, 9)) / 100) * CDbl(Item.TextMatrix(a, 8)))
    Next
    ValorDesconto.Text = "0"
    If Len(ValorDesconto.Text) = 0 Then LcTotalDesconto = 0 Else LcTotalDesconto = CDbl(ValorDesconto.Text)
    If Len(ValorAcrescimo.Text) = 0 Then LcTotalAcrescimo = 0 Else LcTotalAcrescimo = CDbl(ValorAcrescimo.Text)
    If GlIpi Then
      TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto - LcTotalDesconto + LcTotalAcrescimo + LcTotalIpi), 2)
      TotalIpi.Text = AcertaNumero(CStr(LcTotalIpi), 2)
    Else
      TotalProduto.Text = AcertaNumero(CStr(LcTotalProduto), 2)
      TotalOrcamento.Text = AcertaNumero(CStr(LcTotalProduto - LcTotalDesconto + LcTotalAcrescimo), 2)
      TotalIpi.Text = AcertaNumero(CStr(0), 2)
    End If
End If
codigoproduto.SetFocus
End Function
