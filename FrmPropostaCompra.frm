VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmPropostaCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proposta de Compra e Venda / Contrato de Compra e Venda de Mercadorias (P.F. Carnê/Ch)"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3880
      TabIndex        =   216
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   5760
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Height          =   495
      Left            =   120
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   495
      Left            =   2000
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Proposta de Compra e Venda "
      TabPicture(0)   =   "FrmPropostaCompra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label28"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label22(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label21(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label18(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label3(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label18(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label18(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label18(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label18(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label18(7)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label18(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label18(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label18(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label18(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cidade(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label30(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label29(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label13(0)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label12(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label5(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label4(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Line1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Line2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label31(0)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Frame1(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame1(2)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Frame1(1)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Frame3"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Frame2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt(27)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt(28)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt(26)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt(24)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt(22)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt(21)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt(4)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txt(35)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txt(34)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txt(33)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txt(32)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txt(31)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txt(29)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txt(15)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txt(30)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txt(12)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txt(1)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txt(41)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txt(40)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txt(39)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txt(38)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txt(37)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txt(36)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txt(14)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txt(10)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txt(9)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txt(8)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txt(7)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txt(6)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txt(3)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "CryRelatorio"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).ControlCount=   73
      TabCaption(1)   =   "&Dados da Empresa Onde Trabalha"
      TabPicture(1)   =   "FrmPropostaCompra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8(0)"
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(2)=   "Label5(1)"
      Tab(1).Control(3)=   "Label12(1)"
      Tab(1).Control(4)=   "Label13(1)"
      Tab(1).Control(5)=   "Label29(1)"
      Tab(1).Control(6)=   "cidade(1)"
      Tab(1).Control(7)=   "Label18(13)"
      Tab(1).Control(8)=   "Label18(12)"
      Tab(1).Control(9)=   "Label18(11)"
      Tab(1).Control(10)=   "Label8(2)"
      Tab(1).Control(11)=   "Label8(1)"
      Tab(1).Control(12)=   "Label32(1)"
      Tab(1).Control(13)=   "Label32(0)"
      Tab(1).Control(14)=   "txt(11)"
      Tab(1).Control(15)=   "txt(49)"
      Tab(1).Control(16)=   "txt(48)"
      Tab(1).Control(17)=   "txt(47)"
      Tab(1).Control(18)=   "txt(46)"
      Tab(1).Control(19)=   "txt(45)"
      Tab(1).Control(20)=   "txt(44)"
      Tab(1).Control(21)=   "txt(43)"
      Tab(1).Control(22)=   "txt(20)"
      Tab(1).Control(23)=   "txt(17)"
      Tab(1).Control(24)=   "txt(2)"
      Tab(1).Control(25)=   "txt(42)"
      Tab(1).Control(26)=   "txt(13)"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "&Referências Pessoais / Comerciais"
      TabPicture(2)   =   "FrmPropostaCompra.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3(6)"
      Tab(2).Control(1)=   "Label29(3)"
      Tab(2).Control(2)=   "Label29(2)"
      Tab(2).Control(3)=   "Label3(5)"
      Tab(2).Control(4)=   "txt(53)"
      Tab(2).Control(5)=   "txt(52)"
      Tab(2).Control(6)=   "txt(51)"
      Tab(2).Control(7)=   "txt(50)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "&Condições de Compra/ Venda"
      TabPicture(3)   =   "FrmPropostaCompra.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3(17)"
      Tab(3).Control(1)=   "Label3(16)"
      Tab(3).Control(2)=   "Label3(15)"
      Tab(3).Control(3)=   "Label3(14)"
      Tab(3).Control(4)=   "Label3(13)"
      Tab(3).Control(5)=   "Label3(12)"
      Tab(3).Control(6)=   "Label3(11)"
      Tab(3).Control(7)=   "Label3(10)"
      Tab(3).Control(8)=   "Label3(9)"
      Tab(3).Control(9)=   "Label3(8)"
      Tab(3).Control(10)=   "Label3(7)"
      Tab(3).Control(11)=   "Label1"
      Tab(3).Control(12)=   "Label6(0)"
      Tab(3).Control(13)=   "Label7(0)"
      Tab(3).Control(14)=   "Label9(0)"
      Tab(3).Control(15)=   "Label14"
      Tab(3).Control(16)=   "Label15"
      Tab(3).Control(17)=   "Label16"
      Tab(3).Control(18)=   "Label17"
      Tab(3).Control(19)=   "Label19"
      Tab(3).Control(20)=   "Label20"
      Tab(3).Control(21)=   "Label27"
      Tab(3).Control(22)=   "Label33"
      Tab(3).Control(23)=   "Frame1(3)"
      Tab(3).Control(24)=   "txt(64)"
      Tab(3).Control(25)=   "txt(63)"
      Tab(3).Control(26)=   "txt(62)"
      Tab(3).Control(27)=   "txt(61)"
      Tab(3).Control(28)=   "txt(60)"
      Tab(3).Control(29)=   "txt(59)"
      Tab(3).Control(30)=   "txt(58)"
      Tab(3).Control(31)=   "txt(57)"
      Tab(3).Control(32)=   "txt(56)"
      Tab(3).Control(33)=   "txt(55)"
      Tab(3).Control(34)=   "txt(54)"
      Tab(3).Control(35)=   "Frame4"
      Tab(3).Control(36)=   "txt(5)"
      Tab(3).Control(37)=   "txt(16)"
      Tab(3).Control(38)=   "txt(18)"
      Tab(3).Control(39)=   "txt(19)"
      Tab(3).Control(40)=   "txt(23)"
      Tab(3).Control(41)=   "txt(66)"
      Tab(3).Control(42)=   "txt(65)"
      Tab(3).Control(43)=   "txt(68)"
      Tab(3).Control(44)=   "txt(67)"
      Tab(3).Control(45)=   "txt(69)"
      Tab(3).Control(46)=   "txt(25)"
      Tab(3).ControlCount=   47
      TabCaption(4)   =   "D&ados do Cônjuge"
      TabPicture(4)   =   "FrmPropostaCompra.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label34"
      Tab(4).Control(1)=   "Label35"
      Tab(4).Control(2)=   "Label36"
      Tab(4).Control(3)=   "Label37"
      Tab(4).Control(4)=   "Label38"
      Tab(4).Control(5)=   "Label39"
      Tab(4).Control(6)=   "Label40"
      Tab(4).Control(7)=   "Label41"
      Tab(4).Control(8)=   "Label42"
      Tab(4).Control(9)=   "Label43"
      Tab(4).Control(10)=   "Label44"
      Tab(4).Control(11)=   "txt(70)"
      Tab(4).Control(12)=   "txt(71)"
      Tab(4).Control(13)=   "txt(72)"
      Tab(4).Control(14)=   "txt(73)"
      Tab(4).Control(15)=   "txt(74)"
      Tab(4).Control(16)=   "txt(75)"
      Tab(4).Control(17)=   "txt(76)"
      Tab(4).Control(18)=   "txt(77)"
      Tab(4).Control(19)=   "txt(78)"
      Tab(4).Control(20)=   "txt(79)"
      Tab(4).Control(21)=   "txt(80)"
      Tab(4).ControlCount=   22
      TabCaption(5)   =   "D&evedor / Solidário"
      TabPicture(5)   =   "FrmPropostaCompra.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt(81)"
      Tab(5).Control(1)=   "txt(82)"
      Tab(5).Control(2)=   "txt(83)"
      Tab(5).Control(3)=   "txt(84)"
      Tab(5).Control(4)=   "txt(85)"
      Tab(5).Control(5)=   "txt(86)"
      Tab(5).Control(6)=   "Label46"
      Tab(5).Control(7)=   "Label47"
      Tab(5).Control(8)=   "Label48"
      Tab(5).Control(9)=   "Label49"
      Tab(5).Control(10)=   "Label50"
      Tab(5).Control(11)=   "Label51"
      Tab(5).ControlCount=   12
      Begin Crystal.CrystalReport CryRelatorio 
         Left            =   6600
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   81
         Left            =   -72120
         TabIndex        =   119
         Tag             =   "S/T/S/81/N/nomedevsolid"
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   82
         Left            =   -72120
         TabIndex        =   120
         Tag             =   "S/T/S/82/N/cpfdevsolid"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   83
         Left            =   -70440
         TabIndex        =   121
         Tag             =   "S/T/S/83/N/rgdevsolid"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   84
         Left            =   -68160
         TabIndex        =   122
         Tag             =   "S/T/S/84/N/nascdevsolid"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   85
         Left            =   -72000
         TabIndex        =   123
         Tag             =   "S/T/S/85/N/enddevsolid"
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   86
         Left            =   -70920
         TabIndex        =   124
         Tag             =   "S/T/S/86/N/fonedevsolid"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   80
         Left            =   -73680
         TabIndex        =   118
         Tag             =   "S/T/S/80/N/rendabrconjuge"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   79
         Left            =   -67560
         TabIndex        =   116
         Tag             =   "S/T/S/79/N/telempconj"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   78
         Left            =   -72840
         TabIndex        =   114
         Tag             =   "S/T/S/78/N/empresaconjuge"
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   77
         Left            =   -67200
         TabIndex        =   113
         Tag             =   "S/T/S/77/N/emissaoconjuge"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   76
         Left            =   -70560
         TabIndex        =   112
         Tag             =   "S/T/S/76/N/orgaoconjuge"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   75
         Left            =   -73680
         TabIndex        =   111
         Tag             =   "S/T/S/75/N/Rgconjuge"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   74
         Left            =   -67200
         TabIndex        =   110
         Tag             =   "S/T/S/74/N/nascimconjuge"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   73
         Left            =   -69120
         TabIndex        =   109
         Tag             =   "S/T/S/73/N/nacionalidadeconjuge"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   72
         Left            =   -73680
         TabIndex        =   108
         Tag             =   "S/T/S/72/N/naturalidadeconjuge"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   71
         Left            =   -69600
         TabIndex        =   107
         Tag             =   "S/T/S/71/N/naturalidadeconjuge"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   70
         Left            =   -73680
         TabIndex        =   106
         Tag             =   "S/T/S/70/N/nomeconjuge"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   25
         Left            =   -67320
         TabIndex        =   87
         Tag             =   "S/T/S/25/N/taxaam"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   69
         Left            =   -73200
         TabIndex        =   102
         Tag             =   "S/T/S/69/N/descrbem"
         Top             =   4440
         Width           =   5535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   67
         Left            =   -71400
         TabIndex        =   101
         Tag             =   "S/T/S/67/N/ultcheque"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   68
         Left            =   -72720
         TabIndex        =   100
         Tag             =   "S/T/S/68/N/Primcheque"
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   65
         Left            =   -68760
         TabIndex        =   99
         Tag             =   "S/T/S/65/N/Desde"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   66
         Left            =   -70440
         TabIndex        =   98
         Tag             =   "S/T/S/66/N/contacorrente"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   23
         Left            =   -72000
         TabIndex        =   97
         Tag             =   "S/T/S/23/N/Agencia"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   19
         Left            =   -74040
         TabIndex        =   96
         Tag             =   "S/T/S/19/N/Banco"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   18
         Left            =   -68880
         TabIndex        =   86
         Tag             =   "S/T/S/18/N/taxaam"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   16
         Left            =   -70560
         TabIndex        =   85
         Tag             =   "S/T/S/16/N/ultimovenc"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   -73080
         TabIndex        =   84
         Tag             =   "S/T/S/5/N/1vencimento"
         Top             =   2760
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo de Conta"
         Height          =   615
         Left            =   -74640
         TabIndex        =   105
         Top             =   4920
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "Comum"
            Height          =   255
            Left            =   1320
            TabIndex        =   104
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Especial"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   54
         Left            =   -71280
         MaxLength       =   50
         TabIndex        =   73
         Tag             =   "S/T/S/54/N/Tarifa"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   55
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   74
         Tag             =   "S/T/S/55/N/Tabela"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   56
         Left            =   -67200
         MaxLength       =   50
         TabIndex        =   75
         Tag             =   "S/T/S/56/N/NParcelas"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   57
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   76
         Tag             =   "S/T/S/57/N/DataContrato"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   58
         Left            =   -71400
         MaxLength       =   50
         TabIndex        =   77
         Tag             =   "S/T/S/58/N/Carencia"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   59
         Left            =   -68880
         MaxLength       =   50
         TabIndex        =   78
         Tag             =   "S/T/S/59/N/VrCompra"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   60
         Left            =   -67200
         MaxLength       =   50
         TabIndex        =   79
         Tag             =   "S/T/S/60/N/Tar"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   61
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   80
         Tag             =   "S/T/S/61/N/ValorEntrada"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   62
         Left            =   -71280
         MaxLength       =   50
         TabIndex        =   81
         Tag             =   "S/T/S/62/N/VrTarEntrada"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   63
         Left            =   -69120
         MaxLength       =   50
         TabIndex        =   82
         Tag             =   "S/T/S/63/N/ValorPrestacao"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   64
         Left            =   -66960
         MaxLength       =   50
         TabIndex        =   83
         Tag             =   "S/T/S/64/N/Valortotalprazo"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   50
         Left            =   -72240
         MaxLength       =   50
         TabIndex        =   66
         Tag             =   "S/T/S/50/N/NomeRef1"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   51
         Left            =   -68160
         MaxLength       =   20
         TabIndex        =   67
         Tag             =   "S/T/N/51/N/FoneRef1"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   52
         Left            =   -72240
         MaxLength       =   50
         TabIndex        =   68
         Tag             =   "S/T/S/52/N/NomeRef2"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   53
         Left            =   -68160
         MaxLength       =   20
         TabIndex        =   69
         Tag             =   "S/T/N/53/N/FoneRef2"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   13
         Left            =   -73800
         MaxLength       =   20
         TabIndex        =   54
         Tag             =   "S/T/N/13/N/SalarioLiq"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   42
         Left            =   -71400
         MaxLength       =   20
         TabIndex        =   55
         Tag             =   "S/T/N/42/N/TempoServ"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   -69240
         MaxLength       =   20
         TabIndex        =   56
         Tag             =   "S/T/N/2/N/Cargo"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   17
         Left            =   -72960
         MaxLength       =   20
         TabIndex        =   57
         Tag             =   "S/T/N/17/N/CNPJEmpPropria"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   20
         Left            =   -68520
         MaxLength       =   40
         TabIndex        =   60
         Tag             =   "S/T/N/20/N/ComplEmpresa"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   43
         Left            =   -70560
         MaxLength       =   40
         TabIndex        =   59
         Tag             =   "S/T/N/43/N/NumeroEmpresa"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   44
         Left            =   -71280
         MaxLength       =   20
         TabIndex        =   65
         Tag             =   "S/T/N/44/N/DDDTelRamalEmp"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   45
         Left            =   -72960
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "S/T/N/45/N/CepEmp"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   46
         Left            =   -74040
         MaxLength       =   2
         TabIndex        =   63
         Tag             =   "S/T/N/46/N/UfEmpresa"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   47
         Left            =   -70680
         MaxLength       =   50
         TabIndex        =   62
         Tag             =   "S/T/N/47/N/CidadeEmp"
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   48
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   61
         Tag             =   "S/T/N/48/N/BairroEmp"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   49
         Left            =   -74040
         MaxLength       =   40
         TabIndex        =   58
         Tag             =   "S/T/N/49/N/EndEmpresa"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   11
         Left            =   -72840
         MaxLength       =   150
         TabIndex        =   53
         Tag             =   "S/T/N/11/N/EmpresaTrabalha"
         Top             =   1140
         Width           =   5655
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   2640
         MaxLength       =   40
         TabIndex        =   44
         Tag             =   "S/T/N/03/N/END"
         Top             =   4500
         Width           =   3495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   49
         Tag             =   "S/T/N/06/N/BAIRRO"
         Top             =   4980
         Width           =   2415
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   7
         Left            =   720
         MaxLength       =   50
         TabIndex        =   47
         Tag             =   "S/T/N/07/N/CIDADE"
         Top             =   4980
         Width           =   3735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   8
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   48
         Tag             =   "S/T/N/08/N/ESTADO"
         Top             =   4980
         Width           =   375
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   9
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   50
         Tag             =   "S/T/N/09/N/CEP"
         Top             =   4980
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   10
         Left            =   720
         MaxLength       =   20
         TabIndex        =   51
         Tag             =   "S/T/N/10/N/FONE"
         Top             =   5340
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   3960
         TabIndex        =   52
         Tag             =   "S/T/N/14/N/TempoResid"
         Top             =   5340
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   36
         Left            =   720
         MaxLength       =   30
         TabIndex        =   33
         Tag             =   "S/T/N/36/N/CartaoCredito"
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   37
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "S/T/N/37/N/NumeroCartao"
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   38
         Left            =   7320
         MaxLength       =   30
         TabIndex        =   38
         Tag             =   "S/T/N/38/N/QtdeVeiculo"
         Top             =   3900
         Width           =   735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   39
         Left            =   9120
         MaxLength       =   30
         TabIndex        =   40
         Tag             =   "S/T/N/39/N/OutrasPropried"
         Top             =   3900
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   40
         Left            =   6960
         MaxLength       =   40
         TabIndex        =   45
         Tag             =   "S/T/N/40/N/EndNumero"
         Top             =   4500
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   41
         Left            =   9000
         MaxLength       =   40
         TabIndex        =   46
         Tag             =   "S/T/N/41/N/Complemento"
         Top             =   4500
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "S/T/S/01/N/RAZAOSOC"
         Top             =   2100
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   12
         Left            =   480
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "S/T/N/12/N/Rg"
         Top             =   2460
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   30
         Left            =   480
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "S/T/N/30/S/CPF"
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   15
         Left            =   480
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "S/T/N/15/N/Pai"
         Top             =   2820
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   29
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "S/D/N/29/N/DataEmissao"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   31
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   14
         Tag             =   "S/T/N/31/N/Orgao"
         Top             =   2460
         Width           =   735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   32
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "S/T/N/32/N/UFRG"
         Top             =   2460
         Width           =   495
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   33
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   16
         Tag             =   "S/D/N/33/N/DataNacimento"
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   34
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "S/T/N/34/N/Mãe"
         Top             =   2820
         Width           =   4695
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   35
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "S/T/N/35/N/Nacionalidade"
         Top             =   2100
         Width           =   1815
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   480
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "S/T/S/04/N/Loja"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   21
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "S/T/S/21/N/Filial"
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   22
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "S/T/S/22/N/NomeLoja"
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   24
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "S/T/S/24/N/Produto"
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   26
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "S/T/S/26/N/Vendedor"
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   28
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "S/T/S/28/N/FoneLoja"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "S/T/S/00/S/CODIGO"
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   27
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "S/T/S/27/N/NProposta"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Residência"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   4215
         Begin VB.OptionButton Option8 
            Caption         =   "Outros"
            Height          =   255
            Left            =   3120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Financ."
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Alug."
            Height          =   255
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Famil."
            Height          =   255
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Prop."
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado Civil"
         Height          =   615
         Left            =   4440
         TabIndex        =   31
         Top             =   3120
         Width           =   3615
         Begin VB.OptionButton Option12 
            Caption         =   "Outros"
            Height          =   255
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Casado"
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Separ."
            Height          =   255
            Left            =   720
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Solt"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Telefone"
         Height          =   615
         Index           =   1
         Left            =   3720
         TabIndex        =   39
         Top             =   3720
         Width           =   2655
         Begin VB.OptionButton Option17 
            Caption         =   "Não Tem"
            Height          =   255
            Left            =   1560
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option16 
            Caption         =   "Rec"
            Height          =   255
            Left            =   840
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Prop."
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Endereço p/ Corresp."
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   4260
         Width           =   2055
         Begin VB.OptionButton Option19 
            Caption         =   "Trabalho"
            Height          =   195
            Left            =   840
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option18 
            Caption         =   "Resid."
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sexo"
         Height          =   615
         Index           =   0
         Left            =   8280
         TabIndex        =   32
         Top             =   3120
         Width           =   1815
         Begin VB.OptionButton Option14 
            Caption         =   "Masc."
            Height          =   255
            Left            =   960
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Fem."
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Contratação"
         Height          =   615
         Index           =   3
         Left            =   -74640
         TabIndex        =   72
         Top             =   960
         Width           =   2655
         Begin VB.OptionButton Option21 
            Caption         =   "Pós Fixado"
            Height          =   255
            Left            =   1320
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option20 
            Caption         =   "Pré-Fixado"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label46 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72720
         TabIndex        =   215
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label47 
         Caption         =   "CPF"
         Height          =   255
         Left            =   -72720
         TabIndex        =   214
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label48 
         Caption         =   "Rg"
         Height          =   255
         Left            =   -70800
         TabIndex        =   213
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label49 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -69120
         TabIndex        =   212
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label50 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   -72720
         TabIndex        =   211
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "DDD/Telefone/Ramal"
         Height          =   255
         Left            =   -72720
         TabIndex        =   210
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Dados do Comprador "
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   115
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   1080
         X2              =   10080
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   10080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label44 
         Caption         =   "Renda Bruta"
         Height          =   255
         Left            =   -74640
         TabIndex        =   209
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "DDD/Telefone/Ramal"
         Height          =   255
         Left            =   -69480
         TabIndex        =   208
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label42 
         Caption         =   "Empresa Onde Trabalha"
         Height          =   255
         Left            =   -74640
         TabIndex        =   207
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label41 
         Caption         =   "Data Emissão"
         Height          =   255
         Left            =   -68280
         TabIndex        =   206
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Orgão Emissor"
         Height          =   375
         Left            =   -71880
         TabIndex        =   205
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label39 
         Caption         =   "Identidade"
         Height          =   255
         Left            =   -74640
         TabIndex        =   204
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label38 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -68160
         TabIndex        =   203
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Nacionalidade"
         Height          =   255
         Left            =   -70440
         TabIndex        =   202
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Naturalidade"
         Height          =   255
         Left            =   -74640
         TabIndex        =   201
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "CPF"
         Height          =   255
         Left            =   -70320
         TabIndex        =   200
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -74640
         TabIndex        =   199
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label33 
         Caption         =   "Descrição do Bem"
         Height          =   255
         Left            =   -74640
         TabIndex        =   195
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "A"
         Height          =   255
         Left            =   -71640
         TabIndex        =   194
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "Numeração dos Cheques"
         Height          =   255
         Left            =   -74640
         TabIndex        =   193
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -69360
         TabIndex        =   192
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "C/C"
         Height          =   255
         Left            =   -70920
         TabIndex        =   191
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Agência"
         Height          =   255
         Left            =   -72720
         TabIndex        =   190
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Banco"
         Height          =   255
         Left            =   -74640
         TabIndex        =   189
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "%AA"
         Height          =   255
         Left            =   -66360
         TabIndex        =   188
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "%AM"
         Height          =   255
         Index           =   0
         Left            =   -67920
         TabIndex        =   185
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Taxa"
         Height          =   255
         Index           =   0
         Left            =   -69360
         TabIndex        =   184
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Último Vencimento"
         Height          =   255
         Index           =   0
         Left            =   -72000
         TabIndex        =   183
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Primeiro Vencimento"
         Height          =   255
         Left            =   -74640
         TabIndex        =   182
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarifa"
         Height          =   195
         Index           =   7
         Left            =   -71880
         TabIndex        =   180
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela"
         Height          =   195
         Index           =   8
         Left            =   -69840
         TabIndex        =   179
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N. Parc."
         Height          =   195
         Index           =   9
         Left            =   -67800
         TabIndex        =   178
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Contrato"
         Height          =   195
         Index           =   10
         Left            =   -74640
         TabIndex        =   177
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carência"
         Height          =   195
         Index           =   11
         Left            =   -72120
         TabIndex        =   176
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vr.Compra/Serv."
         Height          =   195
         Index           =   12
         Left            =   -70080
         TabIndex        =   175
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tar"
         Height          =   195
         Index           =   13
         Left            =   -67800
         TabIndex        =   174
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V.Entrada"
         Height          =   195
         Index           =   14
         Left            =   -74640
         TabIndex        =   173
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vr.+Tar+Entr."
         Height          =   195
         Index           =   15
         Left            =   -72240
         TabIndex        =   172
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V. Prestação"
         Height          =   195
         Index           =   16
         Left            =   -70080
         TabIndex        =   171
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V.Total Prazo"
         Height          =   195
         Index           =   17
         Left            =   -68040
         TabIndex        =   170
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome 1-"
         Height          =   195
         Index           =   5
         Left            =   -72960
         TabIndex        =   169
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
         Height          =   195
         Index           =   2
         Left            =   -68640
         TabIndex        =   168
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
         Height          =   195
         Index           =   3
         Left            =   -68640
         TabIndex        =   167
         Top             =   2520
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome 2-"
         Height          =   195
         Index           =   6
         Left            =   -72960
         TabIndex        =   166
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label32 
         Caption         =   "Salário Liq."
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   165
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Tempo Serviço"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   164
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Cargo"
         Height          =   255
         Index           =   1
         Left            =   -69840
         TabIndex        =   163
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "CNPJ (Se empr. Prop.)"
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   162
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         Height          =   195
         Index           =   11
         Left            =   -69600
         TabIndex        =   161
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End."
         Height          =   195
         Index           =   12
         Left            =   -74640
         TabIndex        =   160
         Top             =   2640
         Width           =   330
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Index           =   13
         Left            =   -71160
         TabIndex        =   159
         Top             =   2640
         Width           =   555
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
         Index           =   1
         Left            =   -71520
         TabIndex        =   158
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
         Height          =   195
         Index           =   1
         Left            =   -71760
         TabIndex        =   157
         Top             =   3600
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uf"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   156
         Top             =   3600
         Width           =   165
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.E.P.:"
         Height          =   195
         Index           =   1
         Left            =   -73560
         TabIndex        =   155
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   154
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   195
         Index           =   1
         Left            =   -71280
         TabIndex        =   153
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Empresa onde Trabalha"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   152
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   151
         Top             =   4980
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         Height          =   195
         Index           =   0
         Left            =   5280
         TabIndex        =   150
         Top             =   4980
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.E.P.:"
         Height          =   195
         Index           =   0
         Left            =   8280
         TabIndex        =   149
         Top             =   4980
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uf"
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   148
         Top             =   4980
         Width           =   165
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   147
         Top             =   5340
         Width           =   360
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo Resid."
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   146
         Top             =   5340
         Visible         =   0   'False
         Width           =   990
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
         Index           =   0
         Left            =   1560
         TabIndex        =   145
         Top             =   4980
         Width           =   2895
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   4680
         TabIndex        =   144
         Top             =   3300
         Width           =   975
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cartão "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   143
         Top             =   3900
         Width           =   510
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   142
         Top             =   3900
         Width           =   555
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. Veíc."
         Height          =   195
         Index           =   6
         Left            =   6480
         TabIndex        =   141
         Top             =   3900
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outras Propr."
         Height          =   195
         Index           =   7
         Left            =   8160
         TabIndex        =   140
         Top             =   3900
         Width           =   930
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Index           =   8
         Left            =   6240
         TabIndex        =   139
         Top             =   4500
         Width           =   555
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End."
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   138
         Top             =   4500
         Width           =   330
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         Height          =   195
         Index           =   10
         Left            =   7920
         TabIndex        =   137
         Top             =   4500
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidade"
         Height          =   195
         Index           =   5
         Left            =   7200
         TabIndex        =   136
         Top             =   2100
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   135
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rg"
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   2460
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   2100
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pai"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   132
         Top             =   2820
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Emissão"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   131
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orgão"
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   130
         Top             =   2520
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   129
         Top             =   2460
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nascimento"
         Height          =   195
         Index           =   4
         Left            =   7680
         TabIndex        =   128
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mãe"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   117
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label Label21 
         Caption         =   "Loja"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Filial"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   94
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Nome da Loja"
         Height          =   255
         Left            =   2520
         TabIndex        =   93
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Produto"
         Height          =   255
         Left            =   7080
         TabIndex        =   92
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   8520
         TabIndex        =   91
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Telefone da Loja"
         Height          =   375
         Left            =   5160
         TabIndex        =   90
         Top             =   1500
         Width           =   735
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
         Left            =   3120
         TabIndex        =   89
         Top             =   840
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label26 
         Caption         =   "N. Propota"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pré-Fixado"
      Height          =   255
      Index           =   18
      Left            =   960
      TabIndex        =   196
      Top             =   5400
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pós Fixado"
      Height          =   255
      Index           =   19
      Left            =   2160
      TabIndex        =   197
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Contratação"
      Height          =   615
      Index           =   4
      Left            =   840
      TabIndex        =   198
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "%AM"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   187
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Taxa"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   186
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Total Prazo"
      Height          =   195
      Index           =   20
      Left            =   4800
      TabIndex        =   181
      Top             =   3720
      Width           =   960
   End
End
Attribute VB_Name = "FrmPropostaCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer

Private Function DesabilitaCtr()
CmdPrimeiro.Enabled = False
CmdAnterior.Enabled = False
CmdUltimo.Enabled = False
CmdSeguinte.Enabled = False
MnMovimento.Enabled = False
MnRegistro.Enabled = False
CmdExcluir.Enabled = False
End Function
Function BuscaPRop()
'On Error Resume Next
Dim Bbb As Database
Dim RsCompras As Recordset

Set Bbb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsCompras = Bbb.OpenRecordset("PropostaCliente", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterio = "Codigo='" & FrmCliente.Txt(0).Text & "'"
RsCompras.FindFirst LcCriterio


If Not RsCompras.NoMatch Then
        Txt(39).Text = RsCompras!outraspropried & ""
        Txt(71).Text = RsCompras!CPFConjuge & ""
        Txt(0).Text = RsCompras!Codigo & ""
        Txt(1).Text = RsCompras!razaosoc & ""
        Txt(27).Text = RsCompras!nproposta & ""
        Txt(4).Text = RsCompras!loja & ""
        Txt(3).Text = RsCompras!End & ""
        Txt(6).Text = RsCompras!Bairro & ""
        Txt(8).Text = RsCompras!Estado & ""
        Txt(7).Text = RsCompras!Cidade & ""
        Txt(9).Text = RsCompras!Cep & ""
        Txt(10).Text = RsCompras!Fone & ""
        Txt(21).Text = RsCompras!filial & ""
        Txt(22).Text = RsCompras!nomeloja & ""
        Txt(28).Text = RsCompras!foneloja & ""
        Txt(30).Text = RsCompras!cpf & ""
        Txt(12).Text = RsCompras!rg & ""
        Txt(29).Text = RsCompras!dataemissao & ""
        Txt(26).Text = RsCompras!Vendedor & ""
        Txt(24).Text = RsCompras!Produto & ""
        Txt(31).Text = RsCompras!orgao & ""
        Txt(32).Text = RsCompras!ufrg & ""
        Txt(33).Text = RsCompras!DataNacimento & ""
        Txt(15).Text = RsCompras!pai & ""
         Txt(34).Text = RsCompras!mae & ""
        Select Case RsCompras!Residencia & ""
                Case Is = 1
                  Option4.Value = True
                Case Is = 2
                  Option5.Value = True
                Case Is = 3
                  Option6.Value = True
                Case Is = 4
                  Option7.Value = True
                Case Is = 5
                  Option8.Value = True
         End Select
         GlCampo12 = RsCompras!Residencia
        Select Case RsCompras!estadocivil
                Case Is = 1
                  Option99.Value = True
                Case Is = 2
                  Option10.Value = True
                Case Is = 3
                  Option11.Value = True
                Case Is = 4
                  Option12.Value = True
                
        End Select
        GlCampo29 = RsCompras!estadocivil & ""
        Txt(35).Text = RsCompras!nacionalidade & ""
        Txt(86).Text = RsCompras!fonedevsolid & ""
        Select Case RsCompras!sexo
               Case Is = 1
                  Option13.Value = True
               Case Is = 2
                  Option14.Value = True
        End Select
        GlCampo31 = RsCompras!sexo
        Txt(36).Text = RsCompras!cartaocredito & ""
        Txt(37).Text = RsCompras!numerocartao & ""
        Select Case RsCompras!situacaofone
         Case Is = 1
                  Option15.Value = True
                Case Is = 2
                  Option16.Value = True
                Case Is = 3
                  Option17.Value = True
         End Select
        GlCampo15 = RsCompras!situacaofone & ""
        Txt(38).Text = RsCompras!qtdeveiculo & ""
        Txt(38).Text = RsCompras!outraspropried & ""
        Txt(40).Text = RsCompras!endnumero & ""
        Txt(41).Text = RsCompras!Complemento & ""
        Txt(14).Text = RsCompras!temporesid & ""
        Txt(11).Text = RsCompras!empresatrabalha & ""
        Txt(13).Text = RsCompras!salarioliq & ""
        Txt(42).Text = RsCompras!temposerv & ""
        Txt(2).Text = RsCompras!cargo & ""
        Txt(17).Text = RsCompras!cnpjemppropria & ""
        Txt(49).Text = RsCompras!endempresa & ""
        Txt(43).Text = RsCompras!numeroempresa & ""
        Txt(20).Text = RsCompras!complempresa & ""
        Txt(46).Text = RsCompras!ufempresa & ""
        Txt(47).Text = RsCompras!cidadeemp & ""
        Txt(48).Text = RsCompras!bairroemp & ""
        Txt(45).Text = RsCompras!cepemp & ""
        Txt(44).Text = RsCompras!dddtelramalemp & ""
        Txt(50).Text = RsCompras!nomeref1 & ""
        Txt(51).Text = RsCompras!foneref1 & ""
        Txt(52).Text = RsCompras!nomeref2 & ""
        Txt(53).Text = RsCompras!foneref2 & ""
        Select Case RsCompras!tipocontratacao
               Case Is = 1
                  Option20.Value = True
                Case Is = 2
                  Option21.Value = True
         End Select
        GlCampo24 = RsCompras!tipocontratacao & ""
        Txt(54).Text = RsCompras!tarifa & ""
        Txt(55).Text = RsCompras!Tabela & ""
        Txt(56).Text = RsCompras!nparcelas & ""
        Txt(57).Text = RsCompras!datacontrato & ""
        Txt(58).Text = RsCompras!carencia & ""
        Txt(59).Text = RsCompras!vrcompra & ""
        Txt(60).Text = RsCompras!tar & ""
        Txt(61).Text = RsCompras!valorentrada & ""
        Txt(62).Text = RsCompras!vrtarentrada & ""
        Txt(63).Text = RsCompras!valorprestacao & ""
        Txt(64).Text = RsCompras!valortotalprazo & ""
        Txt(5).Text = RsCompras!primvencimento & ""
        Txt(16).Text = RsCompras!ultimovenc & ""
        Txt(18).Text = RsCompras!taxaam & ""
        Txt(25).Text = RsCompras!taxaaa & ""
        Txt(19).Text = RsCompras!banco & ""
        Txt(23).Text = RsCompras!Agencia & ""
        Txt(66).Text = RsCompras!contacorrente & ""
        Txt(65).Text = RsCompras!desde
        Txt(68).Text = RsCompras!primcheque & ""
        Txt(67).Text = RsCompras!ultcheque & ""
        Txt(69).Text = RsCompras!descrbem & ""
        Txt(70).Text = RsCompras!nomeconjuge & ""
        Txt(72).Text = RsCompras!naturalidadeconjuge & ""
        Txt(73).Text = RsCompras!nacionalidadeconjuge & ""
        Txt(74).Text = RsCompras!nascimconjuge & ""
        Txt(75).Text = RsCompras!rgconjuge & ""
        Txt(76).Text = RsCompras!orgaoconjuge & ""
        Txt(77).Text = RsCompras!emissaoconjuge & ""
        Txt(78).Text = RsCompras!empresaconjuge & ""
        Txt(79).Text = RsCompras!telempconj & ""
        Txt(80).Text = RsCompras!rendabrconj & ""
        Txt(81).Text = RsCompras!nomedevsolid & ""
        Txt(82).Text = RsCompras!cpfdevsolid & ""
        Txt(83).Text = RsCompras!rgdevsolid & ""
        Txt(84).Text = RsCompras!nascdevsolid & ""
        Txt(85).Text = RsCompras!enddevsolid & ""
        Select Case RsCompras!EndCorresp
         Case Is = 1
              Option18.Value = True
         Case Is = 2
              Option19.Value = True
        End Select
        GlCampo87 = RsCompras!EndCorresp
        Select Case RsCompras!tipoconta
            Case Is = 1
              Option2.Value = True
            Case Is = 2
              Option3.Value = True
        End Select
        GlCampo88 = RsCompras!tipoconta
        Command1.Enabled = True
 Else
     Txt(0).Text = FrmCliente.Txt(0).Text
 End If
 
 RsCompras.Close
End Function

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Dim Bbb As Database
Dim RsCompras As Recordset

Set Bbb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsCompras = Bbb.OpenRecordset("PropostaCliente", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterio = "Codigo='" & Txt(0).Text & "'"
RsCompras.FindFirst LcCriterio
If Not RsCompras.NoMatch Then
   RsCompras.Edit
Else
   RsCompras.AddNew
End If
        RsCompras!outraspropried = Txt(39).Text
        RsCompras!Codigo = Txt(0).Text
        RsCompras!razaosoc = Txt(1).Text
        RsCompras!nproposta = Txt(27).Text
        RsCompras!loja = Txt(4).Text
        RsCompras!End = Txt(3).Text
        RsCompras!Bairro = Txt(6).Text
        RsCompras!Estado = Txt(8).Text
        RsCompras!Cidade = Txt(7).Text
        RsCompras!Cep = Txt(9).Text
        RsCompras!Fone = Txt(10).Text
        RsCompras!filial = Txt(21).Text
        RsCompras!nomeloja = Txt(22).Text
        RsCompras!foneloja = Txt(28).Text
        RsCompras!cpf = Txt(30).Text
        RsCompras!rg = Txt(12).Text
        If Len(Txt(29).Text) > 0 Then If IsDate(Txt(29).Text) Then RsCompras!dataemissao = CDate(Txt(29).Text)
        RsCompras!Vendedor = Txt(26).Text
        RsCompras!Produto = Txt(24).Text
        RsCompras!orgao = Txt(31).Text
        RsCompras!ufrg = Txt(32).Text
        If Len(Txt(33).Text) > 0 Then RsCompras!DataNacimento = CDate(Txt(33).Text)
        RsCompras!pai = Txt(15).Text
        RsCompras!mae = Txt(34).Text
        If Len(GlCampo12) > 0 Then RsCompras!Residencia = Val(GlCampo12)
        If Len(GlCampo29) > 0 Then RsCompras!estadocivil = Val(GlCampo29)
        RsCompras!nacionalidade = Txt(35).Text
        RsCompras!fonedevsolid = Txt(86).Text
        If Len(GlCampo31) > 0 Then RsCompras!sexo = Val(GlCampo31)
        RsCompras!cartaocredito = Txt(36).Text
        RsCompras!numerocartao = Txt(37).Text
        If Len(GlCampo15) > 0 Then RsCompras!situacaofone = Val(GlCampo15)
        RsCompras!qtdeveiculo = Txt(38).Text
        RsCompras!outraspropried = Txt(38).Text
        RsCompras!endnumero = Txt(40).Text
        RsCompras!Complemento = Txt(41).Text
        RsCompras!temporesid = Txt(14).Text
        RsCompras!empresatrabalha = Txt(11).Text
        If Len(Txt(13).Text) > 0 Then RsCompras!salarioliq = CDbl(Txt(13).Text)
        RsCompras!temposerv = Txt(42).Text
        RsCompras!cargo = Txt(2).Text
        RsCompras!cnpjemppropria = Txt(17).Text
        RsCompras!endempresa = Txt(49).Text
        RsCompras!numeroempresa = Txt(43).Text
        RsCompras!complempresa = Txt(20).Text
        RsCompras!ufempresa = Txt(46).Text
        RsCompras!cidadeemp = Txt(47).Text
        RsCompras!bairroemp = Txt(48).Text
        RsCompras!cepemp = Txt(45).Text
        RsCompras!dddtelramalemp = Txt(44).Text
        RsCompras!nomeref1 = Txt(50).Text
        RsCompras!foneref1 = Txt(51).Text
        RsCompras!nomeref2 = Txt(52).Text
        RsCompras!foneref2 = Txt(53).Text
        If Len(GlCampo24) > 0 Then RsCompras!tipocontratacao = Val(GlCampo24)
        RsCompras!tarifa = Txt(54).Text
        RsCompras!Tabela = Txt(55).Text
        RsCompras!nparcelas = Txt(56).Text
        If Len(Txt(57).Text) > 0 Then RsCompras!datacontrato = CDate(Txt(57).Text)
        RsCompras!carencia = Txt(58).Text
        RsCompras!vrcompra = Txt(59).Text
        RsCompras!tar = Txt(60).Text
        RsCompras!valorentrada = Txt(61).Text
        RsCompras!vrtarentrada = Txt(62).Text
        RsCompras!valorprestacao = Txt(63).Text
        RsCompras!valortotalprazo = Txt(64).Text
        If Len(Txt(5).Text) > 0 Then RsCompras!primvencimento = CDate(Txt(5).Text)
        If Len(Txt(16).Text) > 0 Then RsCompras!ultimovenc = CDate(Txt(16).Text)
        RsCompras!taxaam = Txt(18).Text
        RsCompras!taxaaa = Txt(25).Text
        RsCompras!banco = Txt(19).Text
        RsCompras!Agencia = Txt(23).Text
        RsCompras!contacorrente = Txt(66).Text
        If Len(Txt(65).Text) > 0 Then RsCompras!desde = CDate(Txt(65).Text)
        RsCompras!primcheque = Txt(68).Text
        RsCompras!ultcheque = Txt(67).Text
        RsCompras!descrbem = Txt(69).Text
        If Len(GlCampo88) > 0 Then RsCompras!tipoconta = Val(GlCampo88)
        RsCompras!nomeconjuge = Txt(70).Text
        RsCompras!naturalidadeconjuge = Txt(72).Text
        RsCompras!nacionalidadeconjuge = Txt(73).Text
        If Len(Txt(74).Text) > 0 Then RsCompras!nascimconjuge = CDate(Txt(74).Text)
        RsCompras!rgconjuge = Txt(75).Text
        RsCompras!orgaoconjuge = Txt(76).Text
        If Len(Txt(77).Text) > 0 Then RsCompras!emissaoconjuge = CDate(Txt(77).Text)
        RsCompras!empresaconjuge = Txt(78).Text
        RsCompras!telempconj = Txt(79).Text
        RsCompras!rendabrconj = Txt(80).Text
        RsCompras!nomedevsolid = Txt(81).Text
        RsCompras!cpfdevsolid = Txt(82).Text
        RsCompras!rgdevsolid = Txt(83).Text
        If Len(Txt(84).Text) > 0 Then If IsDate(Txt(84).Text) Then RsCompras!nascdevsolid = CDate(Txt(84).Text)
        RsCompras!enddevsolid = Txt(85).Text
        If Len(GlCampo87) > 0 Then RsCompras!EndCorresp = Val(GlCampo87)
        RsCompras!CPFConjuge = Txt(71).Text
 RsCompras.Update
 RsCompras.Close
 Command1.Enabled = True
End Sub


Private Sub Command1_Click()
        
    CryRelatorio.DataFiles(0) = GLBase
    CryRelatorio.ReportFileName = App.Path & "\PropostaCompra.rpt"
    LcFormula = "{PropostaCliente.Codigo}='" & Txt(0).Text & "'"


'== fim filtro
'== fim filtro

CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Proposta de Compra"
LcTipoSaida = 0
CryRelatorio.SelectionFormula = LcFormula
ination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub

Private Sub Form_Activate()
Txt(27).SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Height = 7290
Me.Width = 11145
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
BuscaPRop

End Sub

Private Sub Option10_Click()
GlCampo29 = 2
End Sub

Private Sub Option10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option11_Click()
GlCampo29 = 3
End Sub

Private Sub Option11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option12_Click()
GlCampo29 = 4
End Sub

Private Sub Option12_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option13_Click()
GlCampo31 = 1
End Sub

Private Sub Option13_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option14_Click()
GlCampo31 = 2
End Sub

Private Sub Option14_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option15_Click()
GlCampo15 = 1
End Sub

Private Sub Option16_Click()
GlCampo15 = 2
End Sub

Private Sub Option17_Click()
GlCampo15 = 3
End Sub

Private Sub Option18_Click()
GlCampo87 = 1
End Sub

Private Sub Option18_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option19_Click()
GlCampo87 = 2
End Sub

Private Sub Option19_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option2_Click()
GlCampo88 = 1
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option20_Click()
GlCampo24 = 1
End Sub

Private Sub Option20_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option21_Click()
GlCampo24 = 2
End Sub

Private Sub Option21_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option3_Click()
GlCampo88 = 2
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option4_Click()
GlCampo12 = 1
End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option5_Click()
GlCampo12 = 2
End Sub

Private Sub Option5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option6_Click()
GlCampo12 = 3
End Sub

Private Sub Option6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option7_Click()
GlCampo12 = 4
End Sub

Private Sub Option7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option8_Click()
GlCampo12 = 5
End Sub

Private Sub Option8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Option9_Click()
GlCampo29 = 1
End Sub

Private Sub Option9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 33 Or Index = 29 Or Index = 84 Or Index = 74 Or Index = 77 Or Index = 57 Then
   If Not IsDate(Txt(Index)) Then
      MsgBox "O Campo deve Ser uma data Válida...", vbInformation, "Aviso"
      Txt(Index).Text = ""
      Txt(Index).SetFocus
   End If
End If


End Sub
