VERSION 5.00
Begin VB.Form FrmPropostaCliente 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   495
      Left            =   3600
      TabIndex        =   173
      Top             =   1200
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   120
         TabIndex        =   174
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   64
      Left            =   10920
      MaxLength       =   50
      TabIndex        =   171
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   63
      Left            =   10920
      MaxLength       =   50
      TabIndex        =   169
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   62
      Left            =   10920
      MaxLength       =   50
      TabIndex        =   167
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   61
      Left            =   10800
      MaxLength       =   50
      TabIndex        =   165
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   60
      Left            =   10560
      MaxLength       =   50
      TabIndex        =   163
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   59
      Left            =   11160
      MaxLength       =   50
      TabIndex        =   161
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   58
      Left            =   10680
      MaxLength       =   50
      TabIndex        =   159
      Tag             =   "S/T/S/55/N/Tabela"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   57
      Left            =   10560
      MaxLength       =   50
      TabIndex        =   157
      Tag             =   "S/T/S/54/N/Tarifa"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   56
      Left            =   10560
      MaxLength       =   50
      TabIndex        =   155
      Tag             =   "S/T/S/55/N/NParcelas"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   55
      Left            =   10560
      MaxLength       =   50
      TabIndex        =   153
      Tag             =   "S/T/S/55/N/Tabela"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   54
      Left            =   10560
      MaxLength       =   50
      TabIndex        =   151
      Tag             =   "S/T/S/54/N/Tarifa"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pré-Fix."
      Height          =   255
      Index           =   17
      Left            =   10080
      TabIndex        =   149
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pós Fix."
      Height          =   255
      Index           =   16
      Left            =   10920
      TabIndex        =   148
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   53
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   146
      Tag             =   "S/T/N/53/N/FoneRef2"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   52
      Left            =   5880
      MaxLength       =   50
      TabIndex        =   143
      Tag             =   "S/T/S/52/N/NomeRef2"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   18
      Left            =   12120
      MaxLength       =   20
      TabIndex        =   142
      Tag             =   "S/T/N/51/N/FoneRef1"
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   51
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   140
      Tag             =   "S/T/N/51/N/FoneRef1"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   50
      Left            =   840
      MaxLength       =   50
      TabIndex        =   138
      Tag             =   "S/T/S/50/N/NomeRef1"
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   49
      Left            =   4320
      MaxLength       =   40
      TabIndex        =   127
      Tag             =   "S/T/N/49/N/EndEmpresa"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   48
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   126
      Tag             =   "S/T/N/48/N/BairroEmp"
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   47
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   125
      Tag             =   "S/T/N/47/N/CidadeEmp"
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   46
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   124
      Tag             =   "S/T/N/46/N/UfEmpresa"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   45
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   123
      Tag             =   "S/T/N/45/N/CepEmp"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   44
      Left            =   600
      MaxLength       =   20
      TabIndex        =   122
      Tag             =   "S/T/N/44/N/DDDTelRamalEmp"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   43
      Left            =   7800
      MaxLength       =   40
      TabIndex        =   121
      Tag             =   "S/T/N/43/N/NumeroEmpresa"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   20
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   120
      Tag             =   "S/T/N/20/N/ComplEmpresa"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   17
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   118
      Tag             =   "S/T/N/17/N/CNPJEmpPropria"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   116
      Tag             =   "S/T/N/2/N/Cargo"
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   42
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   114
      Tag             =   "S/T/N/42/N/TempoServ"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   41
      Left            =   8760
      MaxLength       =   40
      TabIndex        =   110
      Tag             =   "S/T/N/41/N/Complemento"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   40
      Left            =   6720
      MaxLength       =   40
      TabIndex        =   107
      Tag             =   "S/T/N/40/N/EndNumero"
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Trabalho"
      Height          =   255
      Index           =   15
      Left            =   1080
      TabIndex        =   105
      Top             =   4080
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Resid."
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   104
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   39
      Left            =   9120
      MaxLength       =   30
      TabIndex        =   102
      Tag             =   "S/T/N/39/N/OutrasPropried"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   38
      Left            =   7320
      MaxLength       =   30
      TabIndex        =   100
      Tag             =   "S/T/N/38/N/QtdeVeiculo"
      Top             =   3480
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Não Tem"
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   99
      Top             =   3600
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Rec"
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   97
      Top             =   3600
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Prop."
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   96
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   37
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   94
      Tag             =   "S/T/N/36/N/Nacionalidade"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   36
      Left            =   720
      MaxLength       =   30
      TabIndex        =   92
      Tag             =   "S/T/N/36/N/Nacionalidade"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Outros"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   78
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Financ."
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   77
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Alug."
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   76
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Famil."
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   75
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Prop."
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   74
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Masc."
      Height          =   255
      Index           =   10
      Left            =   8880
      TabIndex        =   85
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fem."
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   84
      Top             =   3000
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sexo"
      Height          =   615
      Index           =   0
      Left            =   8040
      TabIndex        =   88
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   35
      Left            =   8280
      MaxLength       =   30
      TabIndex        =   86
      Tag             =   "S/T/N/35/N/Nacionalidade"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Outros"
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   83
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Casado"
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   82
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Separ."
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   81
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Solt."
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   80
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   34
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   72
      Tag             =   "S/T/N/34/N/Mãe"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   33
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   70
      Tag             =   "S/D/N/33/N/DataNacimento"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   32
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   68
      Tag             =   "S/T/N/32/N/UFRG"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   31
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   66
      Tag             =   "S/T/N/31/N/Orgao"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   29
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   64
      Tag             =   "S/D/N/29/N/DataEmisao"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   28
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   62
      Tag             =   "S/T/S/28/N/FoneLoja"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   27
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   60
      Tag             =   "S/T/S/26/N/NProposta"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   26
      Left            =   8880
      MaxLength       =   50
      TabIndex        =   58
      Tag             =   "S/T/S/26/N/Vendedor"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   24
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   56
      Tag             =   "S/T/S/24/N/Produto"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   22
      Left            =   3600
      MaxLength       =   50
      TabIndex        =   54
      Tag             =   "S/T/S/22/N/NomeLoja"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   21
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   52
      Tag             =   "S/T/S/21/N/Filial"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   480
      MaxLength       =   50
      TabIndex        =   50
      Tag             =   "S/T/S/04/N/Loja"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   23
      Left            =   7440
      MaxLength       =   50
      TabIndex        =   15
      Tag             =   "S/T/N/23/N/CondicaoEspecial"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   6600
      TabIndex        =   48
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   10680
      TabIndex        =   46
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   8160
      TabIndex        =   45
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   7440
      TabIndex        =   44
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   9360
      TabIndex        =   43
      Top             =   7680
      Width           =   825
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   16
      Left            =   4440
      TabIndex        =   17
      Tag             =   "S/D/N/16/N/DATAULTIMACOMPRA"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   15
      Left            =   480
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "S/T/N/15/N/Pai"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   3960
      TabIndex        =   12
      Tag             =   "S/T/N/14/N/TempoResid"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   960
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "S/T/N/13/N/SalarioLiq"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   30
      Left            =   480
      MaxLength       =   20
      TabIndex        =   13
      Tag             =   "S/T/N/30/S/CPF"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Tag             =   "S/T/N/25/N/OBSERVACAO"
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   1680
      TabIndex        =   16
      Tag             =   "S/D/N/19/N/DATAULTIMAVISITA"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   480
      MaxLength       =   20
      TabIndex        =   14
      Tag             =   "S/T/N/12/N/Rg"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   7080
      MaxLength       =   20
      TabIndex        =   10
      Tag             =   "S/T/N/11/N/EmpresaTrabalha"
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   720
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "S/T/N/10/N/FONE"
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "S/T/N/09/N/CEP"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "S/T/N/08/N/ESTADO"
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   720
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "S/T/N/07/N/CIDADE"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   6000
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "S/T/N/06/N/BAIRRO"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   4
      Tag             =   "S/T/N/05/N/COMPLEMENTO"
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "S/T/N/03/N/END"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "S/T/S/01/N/RAZAOSOC"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "S/T/S/00/S/CODIGO"
      Top             =   480
      Width           =   2175
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
      Left            =   7800
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   10800
      TabIndex        =   23
      Top             =   6840
      Width           =   825
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   9960
      TabIndex        =   22
      Top             =   6840
      Width           =   825
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   9840
      TabIndex        =   21
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9000
      TabIndex        =   19
      Top             =   6120
      Width           =   825
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   10200
      TabIndex        =   20
      Top             =   7680
      Width           =   1665
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
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Residência"
      Height          =   615
      Left            =   120
      TabIndex        =   89
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Èstado Civil"
      Height          =   615
      Left            =   4320
      TabIndex        =   91
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Telefone"
      Height          =   615
      Index           =   1
      Left            =   3720
      TabIndex        =   98
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereço p/ Corresp."
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   106
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Contratação"
      Height          =   615
      Index           =   3
      Left            =   9960
      TabIndex        =   150
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Total Prazo"
      Height          =   195
      Index           =   17
      Left            =   9960
      TabIndex        =   172
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. Prest."
      Height          =   195
      Index           =   16
      Left            =   9960
      TabIndex        =   170
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr.+Tar+Entr."
      Height          =   195
      Index           =   15
      Left            =   9960
      TabIndex        =   168
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Entrada"
      Height          =   195
      Index           =   14
      Left            =   9960
      TabIndex        =   166
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tar"
      Height          =   195
      Index           =   13
      Left            =   9960
      TabIndex        =   164
      Top             =   3720
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vr.Compra/Serv."
      Height          =   195
      Index           =   12
      Left            =   9960
      TabIndex        =   162
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carencia"
      Height          =   195
      Index           =   11
      Left            =   9960
      TabIndex        =   160
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data "
      Height          =   195
      Index           =   10
      Left            =   9960
      TabIndex        =   158
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N. Parc."
      Height          =   195
      Index           =   9
      Left            =   9960
      TabIndex        =   156
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela"
      Height          =   195
      Index           =   8
      Left            =   9960
      TabIndex        =   154
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarifa"
      Height          =   195
      Index           =   7
      Left            =   9960
      TabIndex        =   152
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label31 
      Caption         =   "Cond. de Compra e Venda"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   147
      Top             =   600
      Width           =   1935
   End
   Begin VB.Line Line5 
      Index           =   2
      X1              =   360
      X2              =   7560
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line2 
      X1              =   9840
      X2              =   9840
      Y1              =   600
      Y2              =   6600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome 2-"
      Height          =   195
      Index           =   6
      Left            =   5160
      TabIndex        =   145
      Top             =   6960
      Width           =   600
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
      Height          =   195
      Index           =   3
      Left            =   8280
      TabIndex        =   144
      Top             =   6960
      Width           =   360
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   141
      Top             =   6960
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome 1-"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   139
      Top             =   6960
      Width           =   600
   End
   Begin VB.Label Label31 
      Caption         =   "Referências Pessoais / Comerciais"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   137
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   240
      X2              =   7440
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   136
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   135
      Top             =   6000
      Width           =   405
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.E.P.:"
      Height          =   195
      Index           =   1
      Left            =   6480
      TabIndex        =   134
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uf"
      Height          =   195
      Index           =   1
      Left            =   5760
      TabIndex        =   133
      Top             =   6000
      Width           =   165
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   132
      Top             =   6360
      Width           =   360
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
      Left            =   3480
      TabIndex        =   131
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      Height          =   195
      Index           =   13
      Left            =   7200
      TabIndex        =   130
      Top             =   5640
      Width           =   555
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End."
      Height          =   195
      Index           =   12
      Left            =   3960
      TabIndex        =   129
      Top             =   5640
      Width           =   330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   128
      Top             =   6000
      Width           =   960
   End
   Begin VB.Label Label8 
      Caption         =   "CNPJ (Se empr. Prop.)"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   119
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Cargo"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   117
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label32 
      Caption         =   "Tempo Serviço"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   115
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label32 
      Caption         =   "Salário Liq."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   113
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Empresa onde Trabalha"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   112
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      Height          =   195
      Index           =   10
      Left            =   7680
      TabIndex        =   111
      Top             =   4080
      Width           =   960
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End."
      Height          =   195
      Index           =   9
      Left            =   2280
      TabIndex        =   109
      Top             =   4080
      Width           =   330
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      Height          =   195
      Index           =   8
      Left            =   6000
      TabIndex        =   108
      Top             =   4080
      Width           =   555
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outras Propr."
      Height          =   195
      Index           =   7
      Left            =   8160
      TabIndex        =   103
      Top             =   3480
      Width           =   930
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde. Veíc."
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   101
      Top             =   3480
      Width           =   825
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      Height          =   195
      Index           =   4
      Left            =   1920
      TabIndex        =   95
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão "
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   93
      Top             =   3480
      Width           =   510
   End
   Begin VB.Label Label31 
      Caption         =   "Dados do Comprador "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   90
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   240
      X2              =   9840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mãe"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   73
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nascimento"
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
      Index           =   4
      Left            =   6840
      TabIndex        =   71
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
      Height          =   195
      Index           =   3
      Left            =   5880
      TabIndex        =   69
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orgão"
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   67
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   65
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pai"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   49
      Top             =   2400
      Width           =   225
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rg"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   32
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nacionalidade"
      Height          =   195
      Index           =   5
      Left            =   7200
      TabIndex        =   87
      Top             =   1680
      Width           =   1020
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
      TabIndex        =   79
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "Telefone da Loja"
      Height          =   255
      Left            =   4800
      TabIndex        =   63
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "N. Propota"
      Height          =   255
      Left            =   4200
      TabIndex        =   61
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   8160
      TabIndex        =   59
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Produto"
      Height          =   255
      Left            =   6720
      TabIndex        =   57
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Nome da Loja"
      Height          =   255
      Left            =   2520
      TabIndex        =   55
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label22 
      Caption         =   "Filial"
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "Loja"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite F5 Para Escolher a Cidade"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5400
      TabIndex        =   47
      Top             =   6480
      Width           =   2340
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
      TabIndex        =   42
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo Resid."
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Width           =   360
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
      TabIndex        =   40
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
      TabIndex        =   26
      Top             =   8640
      Width           =   675
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uf"
      Height          =   195
      Index           =   0
      Left            =   4800
      TabIndex        =   34
      Top             =   4560
      Width           =   165
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.E.P.:"
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   36
      Top             =   4560
      Width           =   495
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
      TabIndex        =   38
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
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      Height          =   195
      Index           =   0
      Left            =   5520
      TabIndex        =   37
      Top             =   4560
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   4560
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9600
      Y1              =   840
      Y2              =   840
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
      TabIndex        =   31
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Proposta de Compra e Venda"
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
Attribute VB_Name = "FrmPropostaCliente"
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

Private LcTamanho As Integer
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
If MovImentacao(enAnterior, cliente) Then VinculaDados
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
If Exclui(cliente) = 1 Then
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
If MovImentacao(enPrimeiro, cliente) Then VinculaDados
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
Call SalvaRegistro(cliente)
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
If MovImentacao(enSeguinte, cliente) Then VinculaDados
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
If MovImentacao(enultimo, cliente) Then VinculaDados
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
   Vendedor.AddItem RsVendedor!Nome
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
      Call AbreBanco(cliente)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(cliente)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
If Not GLCalculacodigoCliente Then
   Txt(0).SetFocus
Else
  Txt(0).Enabled = False
End If
Label1.Caption = Label1.Caption & LcCap
CarregaTelemarketing
End Sub
Function CarreGamatriz()
On Error Resume Next
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.cliente

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
Me.Height = 9000
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
   SalvaRegistro (cliente)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (cliente)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False
GlAlteraCodigo = False
End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(cliente) = 1 Then
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
If MovImentacao(enPrimeiro, cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(cliente)
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, cliente) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, cliente) Then VinculaDados
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
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(cliente)


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
Txt(9).Text = GlCampo9
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
Txt(23).Text = GlCampo23
Vendedor = GlCampo22
Txt(1).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
Exit Function
ErroVinculo:
Resume Next
End Function

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
  Txt(a).Text = ""
Next
Txt(1).SetFocus
CmdSalvar.Enabled = False
Vendedor.Text = " "
End Function

Private Sub Txt_LostFocus(Index As Integer)
'If Index = 7 Then BuscaCidade
If Index = 0 Then
   If Not GLCalculacodigoCliente Then Txt(0).Text = Trim(Txt(0).Text)
End If

If Not GLCalculacodigoCliente Then If VerificaDuplicado(Index) Then Txt(Index).SetFocus
   
End Sub

Private Sub Vendedor_Change()
GlCampo22 = Vendedor.Text

End Sub

Private Sub Vendedor_Click()
On Error Resume Next
Dim a As Integer
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Vendedor.Text Then
       codigo.Text = MtVendedor(a).codigo
       Exit For
    End If
Next
GlCampo22 = Vendedor.Text
CmdSalvar.Enabled = True
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
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
