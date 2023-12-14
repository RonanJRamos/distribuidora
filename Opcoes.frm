VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Opcoes 
   Caption         =   "Opções do Sistema"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   Icon            =   "Opcoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox MostraMsgClientePedido 
      Caption         =   "Mostra Menssagem atraso e limite de crédito do clienteno pedido"
      Height          =   255
      Left            =   240
      TabIndex        =   134
      Top             =   5400
      Width           =   6975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Projeto\Lids\banco\lidis.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Impressoras"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar F2"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   10
      Tab             =   5
      TabsPerRow      =   5
      TabHeight       =   697
      TabCaption(0)   =   "&Notas Fiscais"
      TabPicture(0)   =   "Opcoes.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LinhasSaltarBFim"
      Tab(0).Control(1)=   "LinhasSaltarInicio"
      Tab(0).Control(2)=   "Intrucao2"
      Tab(0).Control(3)=   "Intrucao3"
      Tab(0).Control(4)=   "Intrucao1"
      Tab(0).Control(5)=   "BoletoCEF"
      Tab(0).Control(6)=   "AproveitamentoICMS"
      Tab(0).Control(7)=   "NImprimeBaseC"
      Tab(0).Control(8)=   "InformacaoNF"
      Tab(0).Control(9)=   "IcmsDiferenciado"
      Tab(0).Control(10)=   "NaoVerificaEstoque"
      Tab(0).Control(11)=   "ImprimeValorCFOPNota"
      Tab(0).Control(12)=   "BoletoA4"
      Tab(0).Control(13)=   "margemnota"
      Tab(0).Control(14)=   "Serv"
      Tab(0).Control(15)=   "PuloFIm"
      Tab(0).Control(16)=   "Estoque"
      Tab(0).Control(17)=   "Nota"
      Tab(0).Control(18)=   "txt"
      Tab(0).Control(19)=   "Boleto"
      Tab(0).Control(20)=   "Label27"
      Tab(0).Control(21)=   "Label26"
      Tab(0).Control(22)=   "Label25"
      Tab(0).Control(23)=   "Label23"
      Tab(0).Control(24)=   "Label22"
      Tab(0).Control(25)=   "Label15"
      Tab(0).Control(26)=   "Label14(0)"
      Tab(0).Control(27)=   "Label7"
      Tab(0).Control(28)=   "Line5"
      Tab(0).Control(29)=   "Line4"
      Tab(0).Control(30)=   "Line1"
      Tab(0).Control(31)=   "Label6"
      Tab(0).Control(32)=   "Line2(0)"
      Tab(0).Control(33)=   "Label5"
      Tab(0).Control(34)=   "Label4"
      Tab(0).Control(35)=   "Label3"
      Tab(0).Control(36)=   "Label2"
      Tab(0).Control(37)=   "Label1"
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "&Confirmações/Preços"
      TabPicture(1)   =   "Opcoes.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Có&digos"
      TabPicture(2)   =   "Opcoes.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Contas a &Pagar"
      TabPicture(3)   =   "Opcoes.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Contas a &Receber"
      TabPicture(4)   =   "Opcoes.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Sis&tema"
      TabPicture(5)   =   "Opcoes.frx":04CE
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label16"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label17"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label24"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame5"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "decimaiss"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Galpoes"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "ImprimeSemLinha"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Frame13"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Sintegra"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Command5"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "UsaEstoqueSeguranca"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "PermitirVendaEstoqueNegativo"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "BaixarEstoquenoPedido"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).ControlCount=   13
      TabCaption(6)   =   "C&omissão"
      TabPicture(6)   =   "Opcoes.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame8"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "GeraPorItem"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "ComissaoBelclean"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Orça&mento"
      TabPicture(7)   =   "Opcoes.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label12"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame7"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame9"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "servidorimpressaoorc"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Frame10"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "limpaorc"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "ExibirLucratividade"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "Men&sagem Pedido"
      TabPicture(8)   =   "Opcoes.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label13"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "msg"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Msg1"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "msg2"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "&Impressoras de Rede"
      TabPicture(9)   =   "Opcoes.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Command3"
      Tab(9).Control(1)=   "impressora"
      Tab(9).ControlCount=   2
      Begin VB.TextBox msg2 
         Height          =   375
         Left            =   -74880
         MaxLength       =   120
         TabIndex        =   135
         Top             =   2640
         Width           =   9975
      End
      Begin VB.CheckBox BaixarEstoquenoPedido 
         Caption         =   "Baixar Estoque no Pedido"
         Height          =   255
         Left            =   480
         TabIndex        =   133
         Top             =   4800
         Width           =   3975
      End
      Begin VB.CheckBox PermitirVendaEstoqueNegativo 
         Caption         =   "Permitir Venda com estoque Negativo"
         Height          =   255
         Left            =   480
         TabIndex        =   132
         Top             =   4440
         Width           =   3975
      End
      Begin VB.CheckBox UsaEstoqueSeguranca 
         Caption         =   "Utilizar Estoque de Segurança"
         Height          =   255
         Left            =   480
         TabIndex        =   131
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox LinhasSaltarBFim 
         Height          =   285
         Left            =   -72600
         TabIndex        =   130
         Text            =   "0"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox LinhasSaltarInicio 
         Height          =   285
         Left            =   -73080
         TabIndex        =   129
         Text            =   "0"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Intrucao2 
         Height          =   285
         Left            =   -69960
         TabIndex        =   125
         Top             =   4750
         Width           =   5055
      End
      Begin VB.TextBox Intrucao3 
         Height          =   285
         Left            =   -69960
         TabIndex        =   124
         Top             =   5040
         Width           =   5055
      End
      Begin VB.TextBox Intrucao1 
         Height          =   285
         Left            =   -69960
         TabIndex        =   123
         Top             =   4440
         Width           =   5055
      End
      Begin VB.CheckBox BoletoCEF 
         Caption         =   "Bolelo Caixa E.F."
         Height          =   255
         Left            =   -74760
         TabIndex        =   122
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CheckBox AproveitamentoICMS 
         Caption         =   "Imprimir aproveitamento de ICMS nos dados adicionais"
         Height          =   375
         Left            =   -69960
         TabIndex        =   121
         Top             =   3960
         Width           =   4935
      End
      Begin VB.CheckBox NImprimeBaseC 
         Caption         =   "Não imprimir base de calculo"
         Height          =   375
         Left            =   -69960
         TabIndex        =   120
         Top             =   3720
         Width           =   4335
      End
      Begin VB.TextBox InformacaoNF 
         Height          =   285
         Left            =   -74640
         TabIndex        =   118
         Top             =   5400
         Width           =   9255
      End
      Begin VB.CheckBox ComissaoBelclean 
         Caption         =   "Habiltar digitação do percentual de comissão com base no lucro"
         Height          =   375
         Left            =   -74640
         TabIndex        =   117
         Top             =   2400
         Width           =   9375
      End
      Begin VB.CheckBox IcmsDiferenciado 
         Caption         =   "Utilizar ICMS diferenciado"
         Height          =   375
         Left            =   -74760
         TabIndex        =   116
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CheckBox ExibirLucratividade 
         Caption         =   "Exibir Lucratividade no Pedido de Venda"
         Height          =   195
         Left            =   -70560
         TabIndex        =   115
         Top             =   3240
         Width           =   3255
      End
      Begin VB.CheckBox NaoVerificaEstoque 
         Caption         =   "Não Verificar Quant. no Estoque"
         Height          =   255
         Left            =   -69960
         TabIndex        =   114
         Top             =   3480
         Width           =   3615
      End
      Begin VB.CheckBox ImprimeValorCFOPNota 
         Caption         =   "Imprime o Valor de cada CFOP na Observação da Nota"
         Height          =   255
         Left            =   -74760
         TabIndex        =   112
         Top             =   4080
         Width           =   5055
      End
      Begin VB.CheckBox GeraPorItem 
         Caption         =   "Diferencia o valor da Comissão por item"
         Height          =   375
         Left            =   -74640
         TabIndex        =   111
         Top             =   1920
         Width           =   3735
      End
      Begin VB.CheckBox BoletoA4 
         Caption         =   "Imprimir Boleto A4"
         Height          =   255
         Left            =   -74760
         TabIndex        =   110
         Top             =   3800
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   255
         Left            =   9600
         TabIndex        =   108
         Top             =   4800
         Width           =   375
      End
      Begin VB.TextBox Sintegra 
         Height          =   285
         Left            =   5280
         TabIndex        =   107
         Top             =   4800
         Width           =   4335
      End
      Begin VB.TextBox margemnota 
         Height          =   285
         Left            =   -74040
         TabIndex        =   104
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox limpaorc 
         Caption         =   "Limpa a Tela de Orçamento e Vendas"
         Height          =   255
         Left            =   -74760
         TabIndex        =   103
         Top             =   4920
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Suporte a Estoque"
         Height          =   495
         Left            =   -74880
         TabIndex        =   102
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame Frame13 
         Caption         =   "Senhas"
         ForeColor       =   &H00FF0000&
         Height          =   2895
         Left            =   5400
         TabIndex        =   95
         Top             =   1560
         Width           =   4455
         Begin VB.TextBox PedidoVendas 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   100
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox AcimaLimiteCredito 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox ClienteDebito 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox SenhaLib 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label21 
            Caption         =   "Liberação do pedido de Vendas"
            Height          =   255
            Left            =   720
            TabIndex        =   101
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label Label20 
            Caption         =   "Liberação Limite de Crédito"
            Height          =   255
            Left            =   720
            TabIndex        =   98
            Top             =   1560
            Width           =   3495
         End
         Begin VB.Label Label19 
            Caption         =   "Liberação Cliente em Débito"
            Height          =   255
            Left            =   720
            TabIndex        =   97
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Liberação Preço Baixo"
            Height          =   195
            Left            =   720
            TabIndex        =   96
            Top             =   360
            Width           =   1605
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Lançamento no Caixa"
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   88
         Top             =   3360
         Width           =   6735
         Begin VB.CheckBox EntradaPrazo 
            Caption         =   "N Entrada de Notas a Prazo"
            Height          =   255
            Left            =   3840
            TabIndex        =   92
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox EntradaVista 
            Caption         =   "Na entrada de Notas a Vista"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox BaixaDespesa 
            Caption         =   "Na Baixa da Despesa"
            Height          =   375
            Left            =   3840
            TabIndex        =   90
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox InclusaoDespesa 
            Caption         =   "Na Inclusão da Despesa"
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Lançamento no Caixa"
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   -74760
         TabIndex        =   84
         Top             =   3240
         Width           =   6615
         Begin VB.CheckBox BaixaCheque 
            Caption         =   "Na Baixa de Cheques Recebidos"
            Height          =   255
            Left            =   3360
            TabIndex        =   94
            Top             =   840
            Width           =   3015
         End
         Begin VB.CheckBox inclusaoCheque 
            Caption         =   "Na Inclusão de Cheques Recebidos"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   840
            Width           =   3015
         End
         Begin VB.CheckBox VendaVista 
            Caption         =   "Na Venda "
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox BaixaReceita 
            Caption         =   "Na Baixa da Receita"
            Height          =   255
            Left            =   3360
            TabIndex        =   86
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox InclusaoReceita 
            Caption         =   "Na Inclusão da Receita"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.CheckBox ImprimeSemLinha 
         Caption         =   "Imprime Relatórios Com Linha"
         Height          =   255
         Left            =   480
         TabIndex        =   83
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Frame Frame10 
         Caption         =   "Impressão do Funcionário"
         Height          =   855
         Left            =   -71280
         TabIndex        =   79
         Top             =   4440
         Width           =   6015
         Begin VB.OptionButton FuncEmpresa 
            Caption         =   "Imprimir Nome Empresa"
            Height          =   255
            Left            =   3720
            TabIndex        =   82
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton FuncNome 
            Caption         =   "Imprimir o Nome"
            Height          =   255
            Left            =   1920
            TabIndex        =   81
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton FuncCodigo 
            Caption         =   "Imprimir o Código"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   390
            Width           =   1575
         End
      End
      Begin VB.CheckBox servidorimpressaoorc 
         Caption         =   "Servidor de Impressão Para o Orçamento."
         Height          =   375
         Left            =   -74760
         TabIndex        =   77
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tipo Bloqueio Atraso de Clientes"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   73
         Top             =   3120
         Width           =   4095
         Begin VB.OptionButton NaoBloqueia 
            Caption         =   "Não Bloqueia"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton SoOrcamento 
            Caption         =   "Bloqueia, Mais não Gera a Venda só Orçamento "
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   600
            Width           =   3735
         End
         Begin VB.OptionButton Senha 
            Caption         =   "Bloqueia e Libera Com Digitação de Senha"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.CheckBox Serv 
         Caption         =   "Servidor de Impressão"
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   3480
         Width           =   3975
      End
      Begin VB.CheckBox Galpoes 
         Caption         =   "Utilizar Armazenamento em Galpões"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox decimaiss 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox PuloFIm 
         Height          =   285
         Left            =   -74040
         TabIndex        =   67
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox Estoque 
         Height          =   315
         Left            =   -70080
         TabIndex        =   64
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Msg1 
         Height          =   375
         Left            =   -74880
         MaxLength       =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   9975
      End
      Begin VB.TextBox msg 
         Height          =   375
         Left            =   -74880
         MaxLength       =   120
         TabIndex        =   56
         Top             =   1680
         Width           =   9975
      End
      Begin VB.Frame Frame7 
         Caption         =   "Orçamento"
         ForeColor       =   &H00C00000&
         Height          =   2055
         Left            =   -74880
         TabIndex        =   42
         Top             =   960
         Width           =   9615
         Begin VB.CheckBox Contrato 
            Caption         =   "Usar Valor de Contrato para Produto"
            Height          =   375
            Left            =   7320
            TabIndex        =   113
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox meiafolha 
            Caption         =   "Imprime em Meia Folha"
            Height          =   255
            Left            =   4680
            TabIndex        =   47
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox DesUnitario 
            Caption         =   "Altera V. Unit Pelo Desc."
            Height          =   195
            Left            =   7320
            TabIndex        =   78
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox Padrao 
            Caption         =   "Imprime Orçamento no Padrão Windows"
            Height          =   255
            Left            =   3960
            TabIndex        =   69
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CheckBox RateiaAcrecimo 
            Caption         =   "Rateia Acrécimo entre os Itens"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox ImprimeDesconto 
            Caption         =   "Imprime Detalhes de Descontos"
            Height          =   255
            Left            =   3960
            TabIndex        =   62
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox Desconto 
            Caption         =   "Rateia Desconto entre os Itens"
            Height          =   255
            Left            =   1440
            TabIndex        =   61
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CheckBox ipi 
            Caption         =   "Imprime IPI"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox colunas 
            Caption         =   "Imprime 40 Colunas"
            Height          =   255
            Left            =   4680
            TabIndex        =   50
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox linhasorcamento 
            Height          =   375
            Left            =   600
            TabIndex        =   49
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox margem 
            Height          =   285
            Left            =   3600
            TabIndex        =   48
            Text            =   "0"
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox Transp 
            Caption         =   "Dados Transp."
            Height          =   255
            Left            =   3000
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox Vendedor 
            Caption         =   "Escolhe Vendedor"
            Height          =   255
            Left            =   7680
            TabIndex        =   45
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Cliente 
            Caption         =   "Escolhe Cliente"
            Height          =   255
            Left            =   7680
            TabIndex        =   44
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox portaorcamento 
            Height          =   315
            Left            =   600
            TabIndex        =   43
            Top             =   360
            Width           =   2175
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   7440
            X2              =   7440
            Y1              =   120
            Y2              =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Porta"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Pular"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Linhas ao Terminar"
            Height          =   195
            Left            =   1440
            TabIndex        =   52
            Top             =   840
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label11 
            Caption         =   "Margem"
            Height          =   255
            Left            =   2880
            TabIndex        =   51
            Top             =   360
            Width           =   735
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000005&
            X1              =   7440
            X2              =   0
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comissão"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -74760
         TabIndex        =   40
         Top             =   1080
         Width           =   2535
         Begin VB.CheckBox comissao 
            Caption         =   "Múltiplas Comissões"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Utilização do Sistema"
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   480
         TabIndex        =   39
         Top             =   1560
         Width           =   4695
         Begin VB.CheckBox Comercio 
            Caption         =   "Comercial"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox Representante 
            Caption         =   "Representação"
            Height          =   255
            Left            =   2400
            TabIndex        =   1
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.ComboBox Nota 
         Height          =   315
         Left            =   -70080
         TabIndex        =   38
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Left            =   -74040
         TabIndex        =   31
         Top             =   2190
         Width           =   1095
      End
      Begin VB.ComboBox Boleto 
         Height          =   315
         Left            =   -70080
         TabIndex        =   30
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Caption         =   "Preços"
         ForeColor       =   &H00FF0000&
         Height          =   2775
         Left            =   -71040
         TabIndex        =   26
         Top             =   1200
         Width           =   2655
         Begin VB.CheckBox AtualizaPrecoLanca 
            Caption         =   "Atualiza Preço de Venda  Na Entrada de Notas Fiscais."
            Height          =   495
            Left            =   120
            TabIndex        =   99
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Altera  Minimo de Venda Na Alteração de Preço "
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Alterar Lucro Na Alteração de Preço"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Altera na Digitação Lucro no cadastro de Produto."
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Confirmações"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   -74520
         TabIndex        =   22
         Top             =   1080
         Width           =   2655
         Begin VB.CheckBox Excluir 
            Caption         =   " Antes de Excluir"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Alterar 
            Caption         =   " Antes de Alterar"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox Incluir 
            Caption         =   "Confirma Antes de Incluir"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   840
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cálculo De Código"
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   -74520
         TabIndex        =   18
         Top             =   1560
         Width           =   5415
         Begin VB.CheckBox CodigoFornecedor 
            Caption         =   "Fornecedor"
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox CodigoCliente 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   2160
            TabIndex        =   20
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox codigoproduto 
            Caption         =   "Produto"
            Height          =   255
            Left            =   720
            TabIndex        =   19
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Lançamentos "
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   -74880
         TabIndex        =   14
         Top             =   1560
         Width           =   6735
         Begin VB.CheckBox CaixaEntrada 
            Caption         =   "Atualizar Caixa Por Nota de Entrada a Vista."
            Height          =   675
            Left            =   4560
            TabIndex        =   17
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox VistaEntrada 
            Caption         =   "Atualizar Por Notas de Entradas a Vista."
            Height          =   615
            Left            =   2400
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox FaturaEntrada 
            Caption         =   "Atualizar  Por Nota de Entrada Faturada."
            Height          =   615
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lançamentos Contas a Receber"
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   -74760
         TabIndex        =   10
         Top             =   1560
         Width           =   6615
         Begin VB.CheckBox CaixaSaida 
            Caption         =   "Atualizar Caixa Por Venda a Vista."
            Height          =   675
            Left            =   4320
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox VistaSaida 
            Caption         =   "Atualizar Por Venda a Vista."
            Height          =   615
            Left            =   2160
            TabIndex        =   12
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox FaturaSaida 
            Caption         =   "Atualizar  Por Venda  Faturada."
            Height          =   615
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1935
         End
      End
      Begin MSDBGrid.DBGrid impressora 
         Bindings        =   "Opcoes.frx":055A
         Height          =   2655
         Left            =   -74520
         OleObjectBlob   =   "Opcoes.frx":056E
         TabIndex        =   59
         Top             =   1200
         Width           =   9375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "3ª Instrução Boleto CEF"
         Height          =   195
         Left            =   -71880
         TabIndex        =   128
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "2ª Instrução Boleto CEF"
         Height          =   195
         Left            =   -71880
         TabIndex        =   127
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "1ª Instrução Boleto CEF"
         Height          =   195
         Left            =   -71880
         TabIndex        =   126
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informação complementar"
         Height          =   195
         Left            =   -74640
         TabIndex        =   119
         Top             =   5160
         Width           =   1830
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento temporario do arquivo Sintegra"
         Height          =   195
         Left            =   5280
         TabIndex        =   109
         Top             =   4560
         Width           =   3990
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margem"
         Height          =   195
         Left            =   -74640
         TabIndex        =   106
         Top             =   3045
         Width           =   570
      End
      Begin VB.Label Label17 
         Caption         =   "Casas Decimais."
         Height          =   255
         Left            =   3480
         TabIndex        =   71
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Valor Unitário com "
         Height          =   255
         Left            =   480
         TabIndex        =   70
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "No Fim da Nota Fiscal"
         Height          =   255
         Left            =   -72840
         TabIndex        =   68
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Pular "
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   66
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Disponivel"
         Height          =   255
         Left            =   -70920
         TabIndex        =   65
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem No Rodapé do Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74760
         TabIndex        =   57
         Top             =   1320
         Width           =   3120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linhas ao Terminar"
         Height          =   195
         Left            =   -70440
         TabIndex        =   55
         Top             =   2460
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Line Line5 
         X1              =   -74760
         X2              =   -67920
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   -74760
         X2              =   -67920
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   -67920
         X2              =   -67920
         Y1              =   1560
         Y2              =   3360
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressão Nota Saída  e Boleto "
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   37
         Top             =   1740
         Width           =   2535
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   -74760
         X2              =   -74760
         Y1              =   1560
         Y2              =   3360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Portas Impressora"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -70800
         TabIndex        =   36
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label Label4 
         Caption         =   "Boleto"
         Height          =   255
         Left            =   -70920
         TabIndex        =   35
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nota"
         Height          =   255
         Left            =   -70920
         TabIndex        =   34
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No inicio da Nota Fiscal"
         Height          =   195
         Left            =   -72840
         TabIndex        =   33
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Pular"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   2220
         Width           =   615
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Pular "
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   105
      Top             =   3360
      Width           =   375
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   7200
      X2              =   7200
      Y1              =   1680
      Y2              =   3360
   End
End
Attribute VB_Name = "Opcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LcAchou As Integer, LcAtivos As Integer

Private a As Integer

Private Sub AcimaLimiteCredito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Alterar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub boleto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub CaixaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub CaixaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub ClienteDebito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub CodigoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub CodigoFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub codigoproduto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub colunas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Comercio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim RsOpcoes As Recordset
Dim bb As Database
Dim LcCap As String
Dim a As Integer
Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsOpcoes = bb.OpenRecordset("alid901", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcCap = Me.Caption
NumeroDoArquivo = FreeFile
NumeroMsg = NumeroDoArquivo + 1
Me.Caption = "Aguarde,Atualizando Informações..."
If RsOpcoes.EOF Then
   RsOpcoes.AddNew
Else
   RsOpcoes.Edit
End If
RsOpcoes("fantasia") = SenhaLib.Text
RsOpcoes("nome") = ClienteDebito.Text
RsOpcoes("END") = AcimaLimiteCredito.Text
RsOpcoes("cidade") = PedidoVendas.Text
RsOpcoes.Update
For a = Len(GLBase) To 1 Step -1
    letra = Mid(GLBase, a, 1)
    If letra = "\" Then
       LcArqMsg = Mid(GLBase, 1, a) & "msg.txt"
       Exit For
    End If
Next
Open LcArqMsg For Output As #NumeroMsg
Print #NumeroMsg, msg.Text
Print #NumeroMsg, Msg1.Text
Print #NumeroMsg, msg2.Text
Close #NumeroMsg

Open App.Path & "\opcao.txt" For Output As #NumeroDoArquivo
If Len(txt.Text) = 0 Then txt.Text = "0"
If Len(PuloFIm.Text) = 0 Then PuloFIm.Text = 0

Print #NumeroDoArquivo, txt.Text '1
Print #NumeroDoArquivo, Nota.Text  '2
Print #NumeroDoArquivo, Boleto.Text '3
Print #NumeroDoArquivo, Incluir  '4
Print #NumeroDoArquivo, Alterar  '5
Print #NumeroDoArquivo, Excluir   '6
Print #NumeroDoArquivo, FaturaSaida '7
Print #NumeroDoArquivo, VistaSaida  '8
Print #NumeroDoArquivo, CaixaSaida      '9
Print #NumeroDoArquivo, FaturaEntrada   '10
Print #NumeroDoArquivo, VistaEntrada    '11
Print #NumeroDoArquivo, CaixaEntrada    '12
Print #NumeroDoArquivo, Check1          '13
Print #NumeroDoArquivo, Check2          '14
Print #NumeroDoArquivo, Check3          '15
Print #NumeroDoArquivo, Comercio        '16
Print #NumeroDoArquivo, Representante   '17
Print #NumeroDoArquivo, codigoproduto   '18
Print #NumeroDoArquivo, portaorcamento  '19
Print #NumeroDoArquivo, colunas         '20
Print #NumeroDoArquivo, linhasorcamento '21
Print #NumeroDoArquivo, CodigoCliente   '22
Print #NumeroDoArquivo, CodigoFornecedor '23
Print #NumeroDoArquivo, comissao         '24
Print #NumeroDoArquivo, margem.Text     '25
Print #NumeroDoArquivo, Transp          '26
Print #NumeroDoArquivo, Vendedor        '27
Print #NumeroDoArquivo, Cliente         '28
Print #NumeroDoArquivo, ipi             '29
Print #NumeroDoArquivo, Desconto        '30
Print #NumeroDoArquivo, ImprimeDesconto '31
Print #NumeroDoArquivo, RateiaAcrecimo  '32
Print #NumeroDoArquivo, Estoque.Text    '33
Print #NumeroDoArquivo, PuloFIm.Text    '34
Print #NumeroDoArquivo, Padrao          '35
Print #NumeroDoArquivo, decimaiss.Text  '36
Print #NumeroDoArquivo, Galpoes         '37
Print #NumeroDoArquivo, Serv            '38
Print #NumeroDoArquivo, DesUnitario     '39
If Senha Then
   Print #NumeroDoArquivo, 1            '40
Else
   Print #NumeroDoArquivo, 0
End If

If SoOrcamento Then                     '41
   Print #NumeroDoArquivo, 1
Else
   Print #NumeroDoArquivo, 0
End If

If NaoBloqueia Then                     '42
   Print #NumeroDoArquivo, 1
Else
   Print #NumeroDoArquivo, 0
End If
Print #NumeroDoArquivo, SenhaLib.Text   '43
Print #NumeroDoArquivo, servidorimpressaoorc '44

If FuncCodigo Then                      '45
   Print #NumeroDoArquivo, 1
Else
   Print #NumeroDoArquivo, 0
End If
If FuncNome Then                        '46
   Print #NumeroDoArquivo, 1
Else
   Print #NumeroDoArquivo, 0
End If
If FuncEmpresa Then                     '47
   Print #NumeroDoArquivo, 1
Else
   Print #NumeroDoArquivo, 0
End If
If Len(margemnota.Text) = 0 Then margemnota.Text = 0
Print #NumeroDoArquivo, ImprimeSemLinha  '48
Print #NumeroDoArquivo, InclusaoReceita  '49
Print #NumeroDoArquivo, BaixaReceita  '50
Print #NumeroDoArquivo, VendaVista  '51
Print #NumeroDoArquivo, 0 ' VendaPrazo  '52
Print #NumeroDoArquivo, InclusaoDespesa  '53
Print #NumeroDoArquivo, BaixaDespesa  '54
Print #NumeroDoArquivo, EntradaVista  '55
Print #NumeroDoArquivo, EntradaPrazo  '56
Print #NumeroDoArquivo, inclusaoCheque  '57
Print #NumeroDoArquivo, BaixaCheque  '58
Print #NumeroDoArquivo, ClienteDebito.Text   '59
Print #NumeroDoArquivo, AcimaLimiteCredito  '60
Print #NumeroDoArquivo, AtualizaPrecoLanca ' 61
Print #NumeroDoArquivo, limpaorc '62
Print #NumeroDoArquivo, margemnota.Text  '63
Print #NumeroDoArquivo, meiafolha.Value '64
Print #NumeroDoArquivo, BoletoA4.Value '65
Print #NumeroDoArquivo, GeraPorItem.Value '66
Print #NumeroDoArquivo, ImprimeValorCFOPNota.Value '67
Print #NumeroDoArquivo, Contrato.Value '68
Print #NumeroDoArquivo, NaoVerificaEstoque.Value '69
Print #NumeroDoArquivo, ExibirLucratividade.Value '70
Print #NumeroDoArquivo, IcmsDiferenciado.Value '71
Print #NumeroDoArquivo, ComissaoBelclean.Value '72
Print #NumeroDoArquivo, InformacaoNF.Text '73
Print #NumeroDoArquivo, NImprimeBaseC.Value '74
Print #NumeroDoArquivo, AproveitamentoICMS.Value '75
Print #NumeroDoArquivo, BoletoCEF.Value '76
Print #NumeroDoArquivo, Intrucao1.Text '77
Print #NumeroDoArquivo, Intrucao2.Text '78
Print #NumeroDoArquivo, Intrucao3.Text '79
Print #NumeroDoArquivo, LinhasSaltarInicio.Text '80
Print #NumeroDoArquivo, LinhasSaltarBFim.Text '81
Print #NumeroDoArquivo, UsaEstoqueSeguranca.Value '82
Print #NumeroDoArquivo, PermitirVendaEstoqueNegativo.Value ' 83
Print #NumeroDoArquivo, BaixarEstoquenoPedido.Value ' 84
Print #NumeroDoArquivo, MostraMsgClientePedido.Value '85

GlMostraMsgClientePedido = MostraMsgClientePedido.Value
GlBaixarEstoquenoPedido = BaixarEstoquenoPedido.Value
GlPermitirVendaEstoqueNegativo = PermitirVendaEstoqueNegativo.Value
GlUsaEstoqueSeguranca = UsaEstoqueSeguranca.Value
GlLinhasSaltarBFim = LinhasSaltarBFim.Text
GlLinhasSaltarInicio = LinhasSaltarInicio.Text
GLInformacaoNF = InformacaoNF.Text
GlIntrucao1 = Intrucao1.Text
GlIntrucao2 = Intrucao2.Text
GlIntrucao3 = Intrucao3.Text
If BoletoCEF.Value = 1 Then GlBoletoCEF = True Else GlBoletoCEF = False
If AproveitamentoICMS.Value = 1 Then GLAproveitamentoICMS = True Else GLAproveitamentoICMS = False
If NImprimeBaseC.Value = 1 Then GLNImprimeBaseC = True Else GLNImprimeBaseC = False
If ComissaoBelclean.Value = 1 Then GlComissaoBelclean = True Else GlComissaoBelclean = False
GlVerificarIcmsDiferenciado = IcmsDiferenciado.Value
GlExibirLucratividade = ExibirLucratividade.Value
GlNaoVerificaEstoque = NaoVerificaEstoque.Value
If NaoVerificaEstoque.Value = 1 Then GlNaoVerificaEstoque = True Else GlNaoVerificaEstoque = False
GlContrato = Contrato.Value
GlboletoA4 = BoletoA4.Value
GlMeiaFolha = meiafolha.Value
GlGeraPorItem = GeraPorItem.Value

Glmargemnota = margemnota.Text
GlLiberaPedidoVendas = PedidoVendas.Text & ""
GlSenhaCredito = AcimaLimiteCredito.Text & ""
GlSenhaDebito = ClienteDebito.Text & ""

If limpaorc = 1 Then GlLimpaTelaOrc = True Else GlLimpaTelaOrc = False
If AtualizaPrecoLanca = 1 Then GlAtualizaPreco = True Else GlAtualizaPreco = False

If ImprimeSemLinha = 1 Then GlImprimeSemLinha = True Else GlImprimeSemLinha = False

If FuncCodigo Then GlFuncCodigo = True Else GlFuncCodigo = False
If FuncNome Then GlFuncNome = True Else GlFuncNome = False
If FuncEmpresa Then GlFuncEmpresa = True Else GlFuncEmpresa = False

If servidorimpressaoorc = 1 Then GlServidorImpressoraOrc = True Else GlServidorImpressoraOrc = False

GlSenha = Senha
GlSoOrcamento = SoOrcamento
GlNaoBloqueia = NaoBloqueia

If Len(txt.Text) > 0 Then GLSaltoLinhaNota = CInt(txt.Text) Else GLSaltoLinhaNota = 0
GlSenhaLiberacao = SenhaLib.Text
GlMsg = msg.Text
GlEstoqueDisponivel = Estoque.Text
GlMsg1 = Msg1.Text
GlMsg2 = msg2.Text
GlPortaNota = Nota.Text
GlPortaBoleto = Boleto.Text
If Len(decimaiss.Text) > 0 Then
   GlDecimais = CLng(decimaiss.Text)
Else
   GlDecimais = 2
End If
GlPuloFim = CInt(PuloFIm.Text)
GlPortaOrcamento = portaorcamento.Text

If Galpoes = 1 Then GlArmazenaGalpao = True Else GlArmazenaGalpao = False

For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then
       Exit For
    End If
Next

Dim Arquivo As String
Arquivo = Mid(GLBase, 1, a)
Arquivo = Arquivo & "Configuracaopdv.txt"
GravaIni "Pdv", "Maquina", Maquina.Text, Arquivo
GravaIni "Pdv", "Sintegra", Sintegra.Text, Arquivo


If Len(margem.Text) > 0 Then GlMargem = CLng(margem.Text) Else GlMargem = 0
If DesUnitario = 1 Then GlDescUnit = True Else GlDescUnit = False
If Serv = 1 Then GlServidorImpressora = True Else GlServidorImpressora = False
If comissao = 1 Then GlVariasComissao = True Else GlVariasComissao = False
If Incluir = 1 Then GLConfirmaNovo = True Else GLConfirmaNovo = False
If Padrao = 1 Then GLPadraoWindows = True Else GLPadraoWindows = False
If Alterar = 1 Then GlConfirmaAlteracao = True Else GlConfirmaAlteracao = False
If Excluir = 1 Then GlConfirmaExclusao = True Else GlConfirmaExclusao = False
If FaturaSaida = 1 Then GlFaturaSaida = True Else GlFaturaSaida = False
If VistaSaida = 1 Then GlVistaSaida = True Else GlVistaSaida = False
If CaixaSaida = 1 Then GlCaixaSaida = True Else GlCaixaSaida = False
If FaturaEntrada = 1 Then GlFaturaEntrada = True Else GlFaturaEntrada = False
If VistaEntrada = 1 Then GlVistaEntrada = True Else GlVistaEntrada = False
If CaixaEntrada = 1 Then GlCaixaEntrada = True Else GlCaixaEntrada = False
If Check1 = 1 Then GlLucroCad = True Else GlLucroCad = False
If Check2 = 1 Then GlLucroAlteracao = True Else GlLucroAlteracao = False
If Check3 = 1 Then GlMinimoAlteracao = True Else GlMinimoAlteracao = False
If Comercio = 1 Then GlComercio = True Else GlComercio = False
If Representante = 1 Then GlRepresentante = True Else GlRepresentante = False
If codigoproduto = 1 Then GLCalculacodigoProduto = True Else GLCalculacodigoProduto = False
If CodigoCliente = 1 Then GLCalculacodigoCliente = True Else GLCalculacodigoCliente = False
If CodigoFornecedor = 1 Then GLCalculacodigoFornecedor = True Else GLCalculacodigoFornecedor = False
If ExibirLucratividade.Value = 1 Then GlExibirLucratividade = True Else GlExibirLucratividade = False
If ipi = 1 Then GlIpi = True Else GlIpi = False
If Desconto = 1 Then GlDetalhaDesconto = True Else GlDetalhaDesconto = False
If ImprimeDesconto = 1 Then GlImprimeDetalhaDesconto = True Else GlImprimeDetalhaDesconto = False

If RateiaAcrecimo = 1 Then GlRateiaAcrecimo = True Else GlRateiaAcrecimo = False

If Transp = 1 Then GlDadosTransportadora = True Else GlDadosTransportadora = False
If Vendedor = 1 Then GlEsclheVendedor = True Else GlEsclheVendedor = False
If Cliente = 1 Then GlEscolheCliente = True Else GlEscolheCliente = False
If colunas = 1 Then Gl40colunas = True Else Gl40colunas = False

If InclusaoReceita = 1 Then GlInclusaoReceita = True Else GlInclusaoReceita = False
If BaixaReceita = 1 Then GlBaixaReceita = True Else GlBaixaReceita = False
If VendaVista = 1 Then GlVendaVista = True Else GlVendaVista = False
'If VendaPrazo = 1 Then GlVendaPrazo = True Else GlVendaPrazo = False
If InclusaoDespesa = 1 Then GlInclusaoDespesa = True Else GlInclusaoDespesa = False
If BaixaDespesa = 1 Then GlBaixaDespesa = True Else GlBaixaDespesa = False
If EntradaVista = 1 Then GlEntradaVista = True Else GlEntradaVista = False
If EntradaPrazo = 1 Then GlEntradaPrazo = True Else GlEntradaPrazo = False
If inclusaoCheque = 1 Then GlInclusaoCheque = True Else GlInclusaoCheque = False
If BaixaCheque = 1 Then GlBaixaCheque = True Else GlBaixaCheque = False
If GeraPorItem = 1 Then GlGeraPorItem = True Else GlGeraPorItem = False

If Len(linhasorcamento.Text) > 0 Then GlSaltoLinhasOrcamento = CInt(linhasorcamento.Text) Else GlSaltoLinhasOrcamento = 0

Close #NumeroDoArquivo
RsOpcoes.Close
bb.Close
Set RsOpcoes = Nothing
Set bb = Nothing
Me.Caption = LcCap
Unload Me
'If GlServidorImpressora Then LogNotaFiscal
If GlServidorImpressoraOrc Then logOrcamento
End Sub
Function CarregaCombo()
'On Error Resume Next
Dim RsImpressoras As Recordset

AbreBase
Set RsImpressoras = Dbbase.OpenRecordset("Impressoras", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcAchou = 0
Do Until RsImpressoras.EOF
   LcAchou = -1
   If err > 0 Then Exit Do
   portaorcamento.AddItem RsImpressoras!impressora
   Nota.AddItem RsImpressoras!impressora
   Boleto.AddItem RsImpressoras!impressora
   
   RsImpressoras.MoveNext
Loop

RsImpressoras.Close
Set RsImpressoras = Nothing


End Function
Function Reconfigura()
txt.Text = GLSaltoLinhaNota
Estoque.Text = GlEstoqueDisponivel
portaorcamento.Text = GlPortaOrcamento
portaorcamento.Text = GlPortaOrcamento
If GlAtualizaPreco Then AtualizaPrecoLanca = 1 Else AtualizaPrecoLanca = 0
margemnota.Text = Glmargemnota & ""

SenhaLib.Text = GlSenhaLiberacao
decimaiss.Text = GlDecimais

If GlMostraMsgClientePedido Then MostraMsgClientePedido.Value = 1 Else MostraMsgClientePedido.Value = 0
If GlBaixarEstoquenoPedido Then BaixarEstoquenoPedido.Value = 1 Else BaixarEstoquenoPedido.Value = 0
If GlUsaEstoqueSeguranca Then UsaEstoqueSeguranca.Value = 1 Else UsaEstoqueSeguranca.Value = 0
If GlPermitirVendaEstoqueNegativo Then PermitirVendaEstoqueNegativo.Value = 1 Else PermitirVendaEstoqueNegativo.Value = 0
If GlComissaoBelclean Then ComissaoBelclean.Value = 1 Else ComissaoBelclean.Value = 0
If GlVerificarIcmsDiferenciado Then IcmsDiferenciado.Value = 1 Else IcmsDiferenciado.Value = 0
If GlGeraPorItem Then GeraPorItem.Value = 1 Else GeraPorItem.Value = 0
If GlContrato Then Contrato.Value = 1 Else Contrato.Value = 0
If GlImprimeValorCFOPNota Then ImprimeValorCFOPNota.Value = 1 Else ImprimeValorCFOPNota.Value = 0
If GlNaoVerificaEstoque Then NaoVerificaEstoque.Value = 1 Else NaoVerificaEstoque.Value = 0
If GlboletoA4 Then BoletoA4.Value = 1 Else BoletoA4.Value = 0
If GlLimpaTelaOrc Then limpaorc = 1 Else limpaorc = 0
If GlFuncCodigo Then FuncCodigo = True Else FuncCodigo = False
If GlFuncNome Then FuncNome = True Else FuncNome = False
If GlFuncEmpresa Then FuncEmpresa = True Else FuncEmpresa = False
If GlDescUnit Then DesUnitario = 1 Else DesUnitario = 0
If GlArmazenaGalpao Then Galpoes = 1 Else Galpoes = 0
If Len(GlPortaNota) > 0 Then Nota.Text = GlPortaNota
If Len(GlPortaBoleto) > 0 Then Boleto.Text = GlPortaBoleto
If GLConfirmaNovo Then Incluir = 1 Else Incluir = 0
If GlConfirmaAlteracao Then Alterar = 1 Else Alterar = 0
If GlConfirmaExclusao Then Excluir = 1 Else Excluir = 0
If GlFaturaSaida Then FaturaSaida = 1 Else FaturaSaida = 0
If GlVistaSaida Then VistaSaida = 1 Else VistaSaida = 0
If GlCaixaSaida Then CaixaSaida = 1 Else CaixaSaida = 0
If GlFaturaEntrada Then FaturaEntrada = 1 Else FaturaEntrada = 0
If GlVistaEntrada Then VistaEntrada = 1 Else VistaEntrada = 0
If GlCaixaEntrada Then CaixaEntrada = 1 Else CaixaEntrada = 0
If GlLucroCad Then Check1 = 1 Else Check1 = 0
If GlLucroAlteracao Then Check2 = 1 Else Check2 = 0
If GlMinimoAlteracao Then Check3 = 1 Else Check3 = 0
If GlComercio Then Comercio = 1 Else Comercio = 0
If GlRepresentante Then Representante = 1 Else Representante = 0
If GLCalculacodigoProduto Then codigoproduto = 1 Else codigoproduto = 0
If Gl40colunas Then colunas = 1 Else colunas = 0
If GLCalculacodigoCliente Then CodigoCliente = 1 Else CodigoCliente = 0
If GLCalculacodigoFornecedor Then CodigoFornecedor = 1 Else CodigoFornecedor = 0
If GlVariasComissao Then comissao = 1 Else comissao = 0
If GlRateiaAcrecimo Then RateiaAcrecimo = 1 Else RateiaAcrecimo = 0
If GlDadosTransportadora Then Transp = 1 Else Transp = 0
If GlEsclheVendedor Then Vendedor = 1 Else Vendedor = 0
If GlEscolheCliente Then Cliente = 1 Else Cliente = 0
If GlServidorImpressora Then Serv = 1 Else Serv = 0
If GlServidorImpressoraOrc Then servidorimpressaoorc = 1 Else servidorimpressaoorc = 0
If GlImprimeSemLinha Then ImprimeSemLinha = 1 Else ImprimeSemLinha = 0

If GlInclusaoReceita Then InclusaoReceita = 1 Else InclusaoReceita = 0
If GlBaixaReceita Then BaixaReceita = 1 Else BaixaReceita = 0
If GlVendaVista Then VendaVista = 1 Else VendaVista = 0
'If GlVendaPrazo = 1 Then VendaPrazo = 1 Else VendaPrazo = 0
If GlInclusaoDespesa Then InclusaoDespesa = 1 Else InclusaoDespesa = 0
If GlBaixaDespesa Then BaixaDespesa = 1 Else BaixaDespesa = 0
If GlEntradaVista Then EntradaVista = 1 Else EntradaVista = 0
If GlEntradaPrazo Then EntradaPrazo = 1 Else EntradaPrazo = 0
If GlInclusaoCheque Then inclusaoCheque = 1 Else inclusaoCheque = 0
If GlBaixaCheque Then BaixaCheque = 1 Else BaixaCheque = 0
If GlMeiaFolha Then meiafolha.Value = 1 Else meiafolha.Value = 0
If GlExibirLucratividade Then ExibirLucratividade.Value = 1 Else ExibirLucratividade.Value = 0
If GLNImprimeBaseC Then NImprimeBaseC.Value = 1 Else NImprimeBaseC.Value = 0
InformacaoNF.Text = GLInformacaoNF
PuloFIm.Text = GlPuloFim
If GLAproveitamentoICMS = True Then AproveitamentoICMS.Value = 1 Else AproveitamentoICMS.Value = 0
If GlBoletoCEF Then BoletoCEF.Value = 1 Else BoletoCEF.Value = 0
If GlIpi Then ipi = 1 Else ipi = 0
If GlDetalhaDesconto Then Desconto = 1 Else Desconto = 0
If GlImprimeDetalhaDesconto Then ImprimeDesconto = 1 Else ImprimeDesconto = 0
If GLPadraoWindows Then Padrao = 1 Else Padrao = 0
Intrucao1.Text = GlIntrucao1
LinhasSaltarInicio.Text = GlLinhasSaltarInicio
LinhasSaltarBFim.Text = GlLinhasSaltarBFim
Intrucao2.Text = GlIntrucao2
Intrucao3.Text = GlIntrucao3
Senha = GlSenha
SoOrcamento = GlSoOrcamento
NaoBloqueia = GlNaoBloqueia

If Len(GlMargem) > o Then
   margem.Text = GlMargem
Else
   GlMargem = 0
End If
msg.Text = GlMsg
Msg1.Text = GlMsg1
msg2.Text = GlMsg2
linhasorcamento.Text = GlSaltoLinhasOrcamento
AcimaLimiteCredito.Text = GlSenhaCredito
ClienteDebito.Text = GlSenhaDebito
PedidoVendas.Text = GlLiberaPedidoVendas & ""
End Function

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command3_Click()
LcSenha = InputBox("Entre com a Senha Para a Liberação Desta Tarefa.", "Acerto de Estoque")
If Len(LcSenha) = 0 Then Exit Sub
If LcSenha <> "mestre" Then Exit Sub
LcCap = Me.Caption
Me.Caption = "Aguarde, Acertando o estoque."
Screen.MousePointer = 11
GlImplanta = True
Call acertaestoque("")
GlImplanta = False
Me.Caption = LcCap
MsgBox "Operação Terminada.", 64, "Aviso"
Screen.MousePointer = 0
End Sub

Private Sub Command5_Click()
On Error Resume Next
configuracaoSintegra.Show , Me

End Sub

Private Sub decimaiss_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Desconto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub DesUnitario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Estoque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Excluir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub FaturaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub FaturaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Form_Load()
Me.Refresh
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Data1.DatabaseName = GLBase
Data1.Refresh
CarregaCombo
Reconfigura
Command3.Visible = Galpoes.Value
End Sub

Private Sub FuncCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub FuncEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub FuncNome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Galpoes_Click()
Command3.Visible = Galpoes.Value
End Sub

Private Sub Galpoes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"

End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub ImprimeDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub ImprimeSemLinha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Incluir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub ipi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub linhasorcamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub margem_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub meiafolha_Click()
On Error Resume Next
If meiafolha.Value = 1 Then ConfiguraMeiaFolha.Show , Me
End Sub

Private Sub meiafolha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub msg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub Msg1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub NaoBloqueia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Nota_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub Padrao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub portaorcamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub PuloFIm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub RateiaAcrecimo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Representante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Senha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub SenhaLib_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Serv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub servidorimpressaoorc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub SoOrcamento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub Transp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub VistaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub VistaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then Call Command1_Click
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
