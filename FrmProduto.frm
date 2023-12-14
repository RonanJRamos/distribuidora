VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProduto 
   BackColor       =   &H00CFD3AF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Produto"
   ClientHeight    =   8160
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ForeColor       =   &H00800000&
   Icon            =   "FrmProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSMask.MaskEdBox UltimaAlteracao 
      Height          =   285
      Left            =   5040
      TabIndex        =   100
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CheckBox Desativado 
      BackColor       =   &H00CFD3AF&
      Caption         =   "Desativado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   99
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox CEST 
      Height          =   285
      Left            =   1080
      TabIndex        =   97
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   14
      Left            =   1080
      TabIndex        =   94
      Top             =   6360
      Width           =   3375
   End
   Begin VB.CommandButton CmdCadastraInpostos 
      Caption         =   "Cadastra Impostos NFe"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   93
      Top             =   3720
      Width           =   2625
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   1080
      TabIndex        =   30
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Seguranca 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   90
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Icms 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   5955
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   5955
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   11
      Top             =   5955
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   84
      Top             =   7905
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox EstCalifornia 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox MinCalifornia 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox EstSanta2 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox MinSanta2 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox EstSanta 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox MinSanta 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Sub Itens"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   76
      Top             =   4440
      Width           =   2625
   End
   Begin VB.CheckBox subi 
      BackColor       =   &H00CFD3AF&
      Caption         =   "Possui Sub Itens"
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
      Left            =   3120
      TabIndex        =   75
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox Mantem 
      BackColor       =   &H00CFD3AF&
      Caption         =   "Mantem Último Registro Digitado."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   68
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E8F0D7&
      Height          =   285
      Index           =   9
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00E8F0D7&
      Height          =   285
      Index           =   20
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   4440
      TabIndex        =   63
      Top             =   7560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Fornecedor 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   7560
      Width           =   3135
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   18
      Left            =   6360
      TabIndex        =   20
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Es&toque Galpões F4"
      Height          =   495
      Left            =   7560
      TabIndex        =   23
      Top             =   2040
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ordenar F12"
      Height          =   495
      Left            =   8880
      TabIndex        =   24
      Top             =   1440
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pes&quisar F11"
      Height          =   495
      Left            =   7560
      TabIndex        =   22
      Top             =   1440
      Width           =   1305
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   16
      Left            =   5040
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   13
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   8880
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   0
      Left            =   1200
      MaxLength       =   13
      TabIndex        =   16
      Tag             =   "codigo"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   8640
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   15
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   5880
      TabIndex        =   8
      Top             =   3690
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   7920
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   7800
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   8400
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "Nome"
      Top             =   1080
      Width           =   6015
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   495
      Left            =   8880
      TabIndex        =   28
      Top             =   3120
      Width           =   1305
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
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
      Left            =   2280
      Top             =   3720
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
      Left            =   6600
      TabIndex        =   42
      Top             =   120
      Width           =   1575
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
      Left            =   8160
      TabIndex        =   41
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   3120
      Width           =   1305
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   495
      Left            =   8880
      TabIndex        =   26
      Top             =   2640
      Width           =   1305
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   840
      Width           =   1305
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   2640
      Width           =   1305
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   13
      Top             =   840
      Width           =   1305
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   5040
      Width           =   2625
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
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   8640
      Width           =   2895
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Tag             =   "Preco"
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Tag             =   "Lucro"
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   3
      Left            =   8400
      TabIndex        =   102
      Tag             =   "Custo"
      Top             =   5640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   6
      Left            =   8400
      TabIndex        =   109
      Tag             =   "CustoTotal"
      Top             =   7560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   88
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   108
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   9
      Left            =   8400
      TabIndex        =   103
      Tag             =   "Custo"
      Top             =   6000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   10
      Left            =   8400
      TabIndex        =   104
      Tag             =   "Custo"
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   11
      Left            =   8400
      TabIndex        =   105
      Tag             =   "Custo"
      Top             =   6510
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   12
      Left            =   8400
      TabIndex        =   106
      Tag             =   "Custo"
      Top             =   6765
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   255
      Index           =   13
      Left            =   8400
      TabIndex        =   107
      Tag             =   "Custo"
      Top             =   7020
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   7560
      TabIndex        =   115
      Top             =   7020
      Width           =   585
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   7560
      TabIndex        =   114
      Top             =   6772
      Width           =   660
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   7560
      TabIndex        =   113
      Top             =   6517
      Width           =   465
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   7560
      TabIndex        =   112
      Top             =   6262
      Width           =   225
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS ST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7560
      TabIndex        =   111
      Top             =   6007
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limite. Venda"
      Height          =   195
      Index           =   6
      Left            =   4440
      TabIndex        =   110
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Última alteração de Preço"
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
      Left            =   2760
      TabIndex        =   101
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7440
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "CEST"
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
      Left            =   120
      TabIndex        =   98
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label LNCM 
      BackStyle       =   0  'Transparent
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
      Left            =   2760
      TabIndex        =   96
      Top             =   6720
      Width           =   7575
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   95
      Top             =   6375
      Width           =   780
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "NCM"
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
      Left            =   120
      TabIndex        =   92
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Segurança"
      Height          =   195
      Left            =   3000
      TabIndex        =   91
      Top             =   4290
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Perc."
      Height          =   195
      Index           =   5
      Left            =   2160
      TabIndex        =   89
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      TabIndex        =   87
      Top             =   5970
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cst"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   86
      Top             =   5970
      Width           =   285
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPI Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   85
      Top             =   5970
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7320
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   -120
      X2              =   7320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "California"
      Height          =   195
      Left            =   120
      TabIndex        =   83
      Top             =   5400
      Width           =   645
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa 2"
      Height          =   195
      Left            =   120
      TabIndex        =   80
      Top             =   5040
      Width           =   555
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa"
      Height          =   195
      Left            =   120
      TabIndex        =   78
      Top             =   4680
      Width           =   420
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   74
      Top             =   2160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compra"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   73
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   72
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   71
      Top             =   3480
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impostos"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   70
      Top             =   5640
      Width           =   630
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Pressione F5 Para Escolher a Unidade"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   69
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   67
      Top             =   7560
      Width           =   495
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   600
      X2              =   7440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Perc. de Custo"
      Height          =   255
      Left            =   2040
      TabIndex        =   66
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   65
      Top             =   2760
      Width           =   150
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Geral"
      Height          =   195
      Left            =   120
      TabIndex        =   64
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão Fornecedor"
      Height          =   255
      Left            =   4440
      TabIndex        =   61
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   720
      X2              =   7440
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   720
      X2              =   7440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Máximo"
      Height          =   195
      Index           =   4
      Left            =   5040
      TabIndex        =   60
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mínimo"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   59
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mim. Venda"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   58
      Top             =   3120
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lucro"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Venda"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6480
      TabIndex        =   55
      Top             =   1560
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   7440
      X2              =   7440
      Y1              =   480
      Y2              =   7440
   End
   Begin VB.Label nomeUnidade 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5880
      TabIndex        =   54
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Unidade 
      BackStyle       =   0  'Transparent
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
      Left            =   2280
      TabIndex        =   53
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem"
      Height          =   255
      Left            =   4080
      TabIndex        =   52
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Unid."
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   7560
      TabIndex        =   50
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   930
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
      TabIndex        =   39
      Top             =   8640
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   49
      Top             =   840
      Visible         =   0   'False
      Width           =   360
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
      Left            =   -480
      TabIndex        =   46
      Top             =   4680
      Width           =   195
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pai"
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
      Left            =   -360
      TabIndex        =   48
      Top             =   5520
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   47
      Top             =   840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   45
      Top             =   1080
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   720
      X2              =   7440
      Y1              =   2040
      Y2              =   2040
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
      TabIndex        =   44
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Produtos"
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
      TabIndex        =   43
      Top             =   120
      Width           =   10875
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
Attribute VB_Name = "FrmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadoFornecedor
        Codigo  As String
        Nome    As String
End Type
Private MtFornecedor() As DadoFornecedor
Private LcTam          As Long

Private Type MtEsp
     Codigo     As Long
     Descricao  As String
End Type

Private LcEspecie()     As MtEsp
Private LcTa            As Long
Private LcEscolheGasto  As Boolean
Private LcCarregado     As Integer
Private a               As Integer
Private ValorAuterado   As Boolean
Sub BuscaNCM()
Dim RsAtualPMcm As ADODB.Recordset
If Len(Txt(12).Text) > 0 Then
   Set RsAtualPMcm = AbreRecordset("Select * from ncm where ncm='" & Txt(12).Text & "'", True)
   If Not RsAtualPMcm.EOF Then
       LNCM.ForeColor = vbBlack
       CEST.Text = RsAtualPMcm!CEST & ""
       LNCM.Caption = RsAtualPMcm!Descricao & ""
   Else
       CEST.Text = ""
       LNCM.Caption = "NCM não Encontrado no cadastro"
       LNCM.ForeColor = vbRed
  End If
End If
End Sub
Sub CalculaCusto()
Dim LcCompra As Currency
Dim LcSt As Currency
Dim LcIpi As Currency
Dim LcFrete As Currency
Dim LcSeguro As Currency
Dim LcOutras As Currency
Dim LcCusto As Currency
If IsNumeric(valor(3).Text) Then LcCompra = valor(3).Text Else LcCompra = 0
If IsNumeric(valor(9).Text) Then LcSt = valor(9).Text Else LcSt = 0
If IsNumeric(valor(10).Text) Then LcIpi = valor(10).Text Else LcIpi = 0
If IsNumeric(valor(11).Text) Then LcFrete = valor(11).Text Else LcFrete = 0
If IsNumeric(valor(12).Text) Then LcSeguro = valor(12).Text Else LcSeguro = 0
If IsNumeric(valor(13).Text) Then LcOutras = valor(13).Text Else LcOutras = 0
LcCusto = LcCompra + LcSt + LcIpi + LcFrete + LcSeguro + LcOutras

valor(6).Text = FormatNumber(LcCusto, 2)

End Sub
Function carregaFornecedor()
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim RsEmpresa As Recordset
AbreBase
''abreconexao
LcTam = 0
Set RsEmpresa = Dbbase.OpenRecordset("Select * from alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Do Until RsEmpresa.EOF
    ReDim Preserve MtFornecedor(LcTam)
    If Not IsNull(RsEmpresa!RazaoSoc) Then
        MtFornecedor(LcTam).Codigo = RsEmpresa!Codigo
        MtFornecedor(LcTam).Nome = RsEmpresa!RazaoSoc
        fornecedor.AddItem RsEmpresa!RazaoSoc
        LcTam = LcTam + 1
    End If
    RsEmpresa.MoveNext
Loop

If LcTam > 0 Then LcTam = LcTam - 1
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set dbbasee = Nothing

End Function

Private Function Desabilitatodos()
Dim a As Integer
For a = 0 To 30
    Txt(a).Enabled = False
Next
End Function
Function CarregaEspecie()
Dim RsEspecie As ADODB.Recordset
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Exit Function
'abreconexao
Set RsEspecie = AbreRecordset("select * from * especie", True) ', dbOpenTable, dbSeeChanges, dbOptimistic)
LcTa = 0
Do Until RsEspecie.EOF
   ReDim Preserve LcEspecie(LcTa)
   LcEspecie(LcTa).Codigo = RsEspecie!Codigo
   LcEspecie(LcTa).Descricao = RsEspecie!Especie
   cbo.AddItem RsEspecie!Especie
   RsEspecie.MoveNext
   LcTa = LcTa + 1
Loop
RsEspecie.Close

End Function




Private Sub CmdAnterior_Click()
On Error Resume Next
'abreconexao
Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
RsAtualP.Find "Codigo=" & Txt(0).Text
If Not RsAtualP.BOF Then
   RsAtualP.MovePrevious
   If Not RsAtualP.BOF Then
      VinculaDados RsAtualP!Codigo
   Else
      MsgBox "Este é o primeiro registro.", 64, "Aviso"
   End If
Else
  MsgBox "Este é o primeiro registro.", 64, "Aviso"
End If
   
'GlMov = True
'If MovImentacao(enAnterior, Produto) Then VinculaDados
'GlMov = False
'LcRegAtual = False
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdCadastraInpostos_Click()
'On Error Resume Next
'Dim LcImposto As New Decisao_NFE_FrmCadastroIcmsProduto
'LcImposto.Nome_do_Projeto = GlNomeProjeto
'LcImposto.Sistema_Implementado = GlSistemaImplementado

'LcImposto.Nome_do_Produto = txt(3).Text
'LcImposto.codigoproduto = txt(0).Text

'LcImposto.Show
On Error Resume Next
Dim LcCommando As String
LcCommando = App.Path & "\"
LcCommando = LcCommando & "ChamaImposto.exe 14 Usuario Lidis Lidis " & Txt(0).Text & " " & Replace(Txt(3).Text, " ", "§")
Shell LcCommando, vbNormalNoFocus

End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
'GlTab = "produtos"
Dim LcResp As Integer
LcResp = MsgBox("Confirma a Exclusão do Registro?", vbCritical + vbYesNo, "Excluindo Registro")
If LcResp = vbNo Then Exit Sub

LcSql = "delete from produtos where codigo=" & Txt(0).Text
'abreconexao
total = ExecutaSql(LcSql)
If total > 0 Then
      Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
      RsAtualP.Resync adAffectCurrent
      LcCodigo = CInt(Txt(0).Text + 1)
      RsAtualP.Find "Codigo=" & LcCodigo
      If Not RsAtualP.EOF Then
         VinculaDados RsAtualP!Codigo
      Else
         LcCodigo = CInt(Txt(0).Text - 1)
         RsAtualP.Find "Codigo=" & LcCodigo
         If Not RsAtualP.EOF Then
            VinculaDados RsAtualP!Codigo
         Else
            limpa
         End If
      End If
End If
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Cmdfechar_Click()
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

Private Sub CmdPrimeiro_Click()
On Error Resume Next
'GlMov = True
'If MovImentacao(enPrimeiro, Produto) Then VinculaDados
'abreconexao
Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
If Not RsAtualP.BOF Then
    RsAtualP.MoveFirst
    VinculaDados RsAtualP!Codigo
Else
   MsgBox "Este é o primeiro Registro.", 64, "Aviso"
End If
'GlMov = False
'LcRegAtual = False

End Sub

Private Sub CmdPrimeiro_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSalvar_Click()

Call Salv

'Call SalvaRegistro(Produto)
'VinculaDados
'LcRegAtual = False
'NovoReg
'If LcTipoDados = 1 Then
'   If Mantem = 0 Then limpa
'   txt(0).Text = ""
'End If
'If txt(0).Enabled Then
'   txt(0).SetFocus
'Else
'   txt(1).SetFocus
'End If
End Sub
Function Salv()
LcP = GLPesquisa
LcI = LcIndice
Dim LcIncluir As Boolean
Dim StrSql    As String
Dim LcSeguranca As Double
On Error GoTo errSlvar
'===> Se for Consulta não Salva
If LcTipoDados = 3 Then Exit Function
If LcTipoDados = 2 Then GLPesquisa = True
If LcTipoDados = 1 Then GLPesquisa = False: LcIndice = "CODIGO"

If Len(EstSanta.Text) = 0 Then EstSanta.Text = 0
If Len(MinSanta.Text) = 0 Then MinSanta.Text = 0

If Len(EstSanta2.Text) = 0 Then EstSanta2.Text = 0
If Len(MinSanta2.Text) = 0 Then MinSanta2.Text = 0

If Len(EstCalifornia.Text) = 0 Then EstCalifornia.Text = 0
If Len(MinCalifornia.Text) = 0 Then MinCalifornia.Text = 0

If Len(Txt(16).Text) = 0 Then Txt(16).Text = 0

If Len(Txt(1).Text) = 0 Then Txt(1).Text = "."
Txt(1).Text = Replace(Txt(1).Text, "'", "")
If Len(Trim(Txt(20).Text)) = 0 Then Txt(20).Text = 0
If Len(Trim(Txt(7).Text)) = 0 Then Txt(7).Text = 0
If Len(Trim(valor(2).Text)) = 0 Then valor(2).Text = 0
If Len(Trim(valor(4).Text)) = 0 Then valor(4).Text = 0
If Len(Trim(Txt(9).Text)) = 0 Then Txt(9).Text = 0
'rsatualp!quantUnidade = CDbl(Txt(9).Text)
If Len(Trim(Txt(10).Text)) = 0 Then Txt(10).Text = 0
If Len(Trim(GlCampo18)) = 0 Then GlCampo18 = 0
If Len(Txt(18).Text) = 0 Then Txt(18).Text = 0

   
If Len(Trim(valor(3).Text)) = 0 Then valor(3).Text = 0
If Len(Trim(valor(1).Text)) = 0 Then valor(1).Text = 0
If Len(Trim(valor(0).Text)) = 0 Then valor(0).Text = 0
If Len(Trim(Txt(16).Text)) = 0 Then Txt(16).Text = 1
If Len(Trim(valor(6).Text)) = 0 Then valor(6).Text = 0
If Len(Trim(valor(7).Text)) = 0 Then valor(7).Text = 0
If Len(Trim(valor(8).Text)) = 0 Then valor(8).Text = 0

If Len(Trim(valor(9).Text)) = 0 Then valor(9).Text = 0
If Len(Trim(valor(10).Text)) = 0 Then valor(10).Text = 0
If Len(Trim(valor(11).Text)) = 0 Then valor(11).Text = 0
If Len(Trim(valor(12).Text)) = 0 Then valor(12).Text = 0
If Len(Trim(valor(13).Text)) = 0 Then valor(13).Text = 0

If Len(Trim(icms.Text)) = 0 Then icms.Text = 0

If Len(Seguranca.Text) > 0 Then
   LcSeguranca = CDbl(Seguranca.Text) * CDbl(Txt(16).Text)

End If
'Call AbreBanco(produto) '===> Abre a Tabela
'abreconexao
Set RsAtualP = AbreRecordset("select * from produtos " & Command2.Tag, True)
'===> Verifica se é inclusão
If Len(Txt(0).Text) = 0 Then
   '==> Verifica se já cadastrou o nome do produto
   RsAtualP.Find "nome='" & Txt(1).Text & "'"
   If Not RsAtualP.EOF Then
      MsgBox "O produto " & Txt(1).Text & " já foi cadastrado com o código:" & RsAtualP!Codigo, 64, "Aviso"
      Exit Function
   End If
   'RsAtualP.AddNew
   LcIncluir = True
Else
   If GlConfirmaAlteracao Then
      Resposta = MsgBox("Confirma a Alteração deste registro?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Aviso")
   Else
      Resposta = vbYes
   End If
   If Resposta = 7 Then GoTo Saida
   '===> Pesquisa o produto
   LcPes = "Codigo=" & Txt(0).Text
   RsAtualP.Find LcPes
   If RsAtualP.EOF Then
     ' RsAtualP.AddNew
     LcIncluir = True
   End If
End If

If Len(Txt(0).Text) = 0 Then
    StrSql = "Insert into produtos (ComissaoFornecedor,Nome,cst,Fornecedor," & _
             "ipi,percentualcusto,minimoVenda,MinimoEst,maximoEstoque,icms," & _
             "Custo,UnidMedida,PRECO,Lucro,QtdMedida,CustoTotal," & _
             "santa1,Santa2,California,multiplositens,per,EstoqueSeguranca,ClassificacaoFiscal,SubItem,CEST,ultimaAlteracao,Desativado,LimiteVenda,"
   StrSql = StrSql & "ValorST,ValorIpi,valorFrete,valorSeguro,ValorDespesas) Values(" & _
             Replace(Replace(Replace(Replace(Txt(18).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ",'" & _
             Replace(Replace(Replace(Txt(1).Text, Chr(34), ""), ",", ""), "'", "") & "','" & _
             Replace(Replace(Replace(Txt(3).Text, Chr(34), ""), ",", ""), "'", "") & "','" & _
             Replace(Replace(Replace(Nome.Text, Chr(34), ""), ",", ""), "'", "") & "'," & _
             Replace(Replace(Replace(Replace(Txt(7).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ",'" & _
             Replace(Replace(Replace(valor(5).Text, Chr(34), ""), ",", "."), "'", "") & "'," & _
             Replace(Replace(Replace(Replace(valor(2).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(valor(4).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(Txt(10).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(icms.Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(valor(3).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ",'" & _
             UCase(Txt(13).Text) & "'," & _
             Replace(Replace(Replace(Replace(valor(1).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(valor(0).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(Txt(16).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(valor(6).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(CDbl(EstSanta.Text) * CDbl(Txt(16).Text) + CDbl(MinSanta.Text), " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(CDbl(EstSanta2.Text) * CDbl(Txt(16).Text) + CDbl(MinSanta2.Text), " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(CDbl(EstCalifornia.Text) * CDbl(Txt(16).Text) + CDbl(MinCalifornia.Text), " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
             CInt(subi.Value) * -1 & "," & _
             Replace(Replace(Replace(Replace(valor(7).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(CStr(LcSeguranca), " ", ""), "$", ""), "R", ""), ",", ".") & ",'" & Txt(12).Text & "',"
    StrSql = StrSql & "'" & Replace(Txt(14).Text, "'", "''") & "',"
    StrSql = StrSql & "'" & Replace(CEST.Text, "'", "''") & "',"
    If IsDate(UltimaAlteracao.Text) Then
       StrSql = StrSql & "'" & Format(CDate(UltimaAlteracao.Text), "yyyy-mm-dd") & "',"
    Else
       StrSql = StrSql & "null,"
    End If
    If Desativado.Value Then
       StrSql = StrSql & "-1,"
    Else
      StrSql = StrSql & "0,"
    End If
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(8).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(9).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(10).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(11).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(12).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & Replace(Replace(Replace(Replace(valor(13).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ")"
    
Else
    StrSql = "Update Produtos Set " & _
           "ComissaoFornecedor =" & Replace(Replace(Replace(Replace(Txt(18).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "Nome ='" & Replace(Replace(Replace(Txt(1).Text, Chr(34), ""), ",", ""), "'", "") & "'," & _
           "Cst ='" & UCase(Txt(3).Text) & "'," & _
           "Fornecedor ='" & Replace(Replace(Replace(UCase(Nome.Text), Chr(34), ""), ",", ""), "'", "") & "'," & _
           "ipi =" & Replace(Replace(Replace(Replace(Txt(7).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "percentualcusto ='" & Replace(Replace(Replace(valor(5).Text, Chr(34), ""), ",", "."), "'", "") & "'," & _
           "minimoVenda =" & Replace(Replace(Replace(Replace(valor(2).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "MinimoEst = " & Replace(Replace(Replace(Replace(valor(4).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "maximoEstoque =" & Replace(Replace(Replace(Replace(Txt(10).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "Icms =" & Replace(Replace(Replace(Replace(icms.Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "Custo = " & Replace(Replace(Replace(Replace(valor(3).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "UnidMedida ='" & UCase(Txt(13).Text) & "'," & _
           "PRECO =" & Replace(Replace(Replace(Replace(valor(1).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "LimiteVenda =" & Replace(Replace(Replace(Replace(valor(8).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "Lucro =" & Replace(Replace(Replace(Replace(valor(0).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "per =" & Replace(Replace(Replace(Replace(valor(7).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "QtdMedida = " & Replace(Replace(Replace(Replace(Txt(16).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "CustoTotal =" & Replace(Replace(Replace(Replace(valor(6).Text, " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "santa1 =" & Replace(Replace(Replace(Replace(CDbl(EstSanta.Text) * CDbl(Txt(16).Text) + CDbl(MinSanta.Text), " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
           "Santa2 =" & Replace(Replace(Replace(Replace(CDbl(EstSanta2.Text) * CDbl(Txt(16).Text) + CDbl(MinSanta2.Text), " ", ""), "$", ""), "R", ""), ",", ".") & "," & _
            "CEST='" & Replace(CEST.Text, "'", "''") & "'," & _
           "california =" & Replace(Replace(Replace(Replace(CDbl(EstCalifornia.Text) * CDbl(Txt(16).Text) + CDbl(MinCalifornia.Text), " ", ""), "$", ""), "R", ""), ",", ".") & _
           ",EstoqueSeguranca=" & Replace(Replace(Replace(Replace(CStr(LcSeguranca), " ", ""), "$", ""), "R", ""), ",", ".")
    StrSql = StrSql & ",subitem='" & Replace(Txt(14).Text, "'", "''") & "'"
    If IsDate(UltimaAlteracao.Text) Then
       StrSql = StrSql & ",ultimaAlteracao='" & Format(CDate(UltimaAlteracao.Text), "yyyy-mm-dd") & "'"
    Else
       StrSql = StrSql & ",ultimaAlteracao=null"
    End If
    If Desativado.Value Then
       StrSql = StrSql & ",Desativado=-1"
    Else
      StrSql = StrSql & ",Desativado=0"
    End If
    StrSql = StrSql & ",ValorST =" & Replace(Replace(Replace(Replace(valor(9).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & "ValorIpi=" & Replace(Replace(Replace(Replace(valor(10).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & "valorFrete=" & Replace(Replace(Replace(Replace(valor(11).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & "valorSeguro=" & Replace(Replace(Replace(Replace(valor(12).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & "ValorDespesas=" & Replace(Replace(Replace(Replace(valor(13).Text, " ", ""), "$", ""), "R", ""), ",", ".") & ","
    StrSql = StrSql & "ClassificacaoFiscal='" & Txt(12).Text & "'"
    StrSql = StrSql & " Where Codigo=" & Txt(0).Text
End If
afetados = ExecutaSql(StrSql)
If LcTipoDados = 1 Then
   If Mantem.Value = 1 Then
      Txt(0).Text = ""
   Else
      limpa
   End If
End If
RetornaCorFundo
Txt(1).SetFocus
CmdSalvar.Enabled = False
Exit Function
Saida:
GLPesquisa = LcP
LcIndice = LcI
RsAtualP.Close
'Dbbase.Close
Set RsAtualP = Nothing
'Set Dbbase = Nothing

errSlvar:
MsgBox err.Description & " Nº: " & err.Number
'Resume 0
Exit Function
End Function
Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
'abreconexao
Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
RsAtualP.Find "Codigo=" & Txt(0).Text
If Not RsAtualP.EOF Then
   
   RsAtualP.MoveNext
   If Not RsAtualP.EOF Then
      VinculaDados RsAtualP!Codigo
   Else
      MsgBox "Este é o Último registro.", 64, "Aviso"
   End If
Else
  MsgBox "Este é o Último registro.", 64, "Aviso"
End If

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
'abreconexao
Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
If Not RsAtualP.EOF Then
    RsAtualP.MoveLast
    VinculaDados RsAtualP!Codigo
Else
   MsgBox "Este é o Último Registro.", 64, "Aviso"
End If
End Sub



Private Sub CmdUltimo_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()

End Sub

Private Sub Command1_Click()
On Error Resume Next
Load frmPesquisaProduto
Command2.Tag = "Order by Nome"
SetaBarra
DoEvents
'abreconexao
LcSql = "Select * from produtos order by nome"
Set RsAtualP = AbreRecordset(LcSql, True)
frmPesquisaProduto.CriaLista Me, RsAtualP, "Nome"
frmPesquisaProduto.Show , Me
LcRegAtual = False
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub Command2_Click()
On Error Resume Next
FrmOrdena.Show , Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub Command3_Click()
On Error Resume Next
Load MostraEstoqueGalpao
MostraEstoqueGalpao.Tag = LcTipoDados
If LcTipoDados = 3 Then
    LcResp = MsgBox("Você esta em modo Consulta, Não poderá alterar o valor do estoque nos galpões." & Chr(13) & "Deseja Entrar mesmo assim ?", vbInformation + vbYesNo, "Aviso")
  Exit Sub
End If
MostraEstoqueGalpao.Show , Me
CmdSalvar.Enabled = True
LcRegAtual = False
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub Command4_Click()
GlCodigoProduto = Txt(0).Text
GastosFixos.Show , Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
subitensproduto.Show , Me
End Sub

Sub BuscaNomeGalpao()
Dim db As Database
Dim Rs As Recordset
Dim a As Integer

Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid012 order by codigo")
Do Until Rs.EOF
  If a = 3 Then Exit Do
  If a = 0 Then
     Label28.Caption = Rs!Nome
  End If
  If a = 1 Then
     Label30.Caption = Rs!Nome
  End If
  If a = 2 Then
     Label32.Caption = Rs!Nome
  End If
  a = a + 1
  Rs.MoveNext
Loop
Set db = Nothing
Set Rs = Nothing

End Sub

Private Sub Desativado_Click()
On Error Resume Next
CmdSalvar.Enabled = True
End Sub

Private Sub Form_Activate()
On Error Resume Next
If GlUsaEstoqueSeguranca Then
   Seguranca.Visible = True
   Label27.Visible = True
Else
  Seguranca.Visible = False
  Label27.Visible = False
End If
Set GlFormA = Me
If LcTipoDados < 3 Then
   GlMov = False
   LcRegAtual = False
End If
SetaBarra
If LcCarregado Then Exit Sub
Mantem.Visible = False
CarregaEspecie
'abreconexao
Select Case LcTipoDados
   Case Is = 1
        LcCap = "   <<Modo Inclusão>>"
        Command3.Enabled = False
        DesabilitaCtr
        Mantem.Visible = True
   Case Is = 2
      LcCap = "   <<Modo Alteração>>"
     ' Call AbreBanco(Produto)
      'abreconexao
      Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
      VinculaDados RsAtualP!Codigo
   Case Is = 3
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
      CmdExcluir.Enabled = False
      VinculaDados RsAtualP!Codigo
 End Select
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False

CarreGamatriz
LcCarregado = True
If Not GLCalculacodigoProduto Then
   Txt(0).SetFocus
Else
  Txt(0).Enabled = False
End If
Command3.Enabled = GlArmazenaGalpao
Txt(20).Enabled = Not GlArmazenaGalpao
Txt(9).Enabled = Not GlArmazenaGalpao
End Sub
Function CriaMascara() As String
'#,##0.00;(#,##0.00)
Dim LcPrimeiraparte, LcSegundaParte As String
Dim a As Integer
If Len(GlDecimais) = 0 Then GlDecimais = 2
For a = 1 To GlDecimais
    LcPrimeiraparte = LcPrimeiraparte & "#"
    LcSegundaParte = LcSegundaParte & "0"
Next
LcMask = "#," & LcPrimeiraparte & "0." & LcSegundaParte
LcMask = LcMask & ";(#," & LcPrimeiraparte & "0." & LcSegundaParte & ")"
CriaMascara = LcMask

End Function
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.produto
For a = 0 To 31
   MtPesquisa(a).campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).Tipo = ""
Next
For a = 0 To 30
  If Len(Trim(Txt(a).Tag)) <> 0 And Txt(a).Visible Then
    LcNome = Mid$(Txt(a).Tag, 12)
    LcTipo = Mid$(Txt(a).Tag, 3, 1)
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).Tipo = LcTipo
    Select Case LcNome
           Case Is = "quantUnidade"
                MtPesquisa(a).campo = "ESTOQUE UNIT."
           Case Is = "UNIMED"
                MtPesquisa(a).campo = "UNIDADE"
           Case Is = "QTDUNIMED"
                MtPesquisa(a).campo = "COM"
           Case Is = "PMU"
                MtPesquisa(a).campo = "PREÇO DE CUSTO"
           Case Is = "PTAB"
                MtPesquisa(a).campo = "PREÇO DE VENDA"
           Case Is = "MPVENDA"
                MtPesquisa(a).campo = "MENOR PREÇO VENDA"
           Case Is = "COD"
                MtPesquisa(a).campo = "CÓDIGO"
           Case Else
                MtPesquisa(a).campo = LcNome
                
      End Select
    End If
 Next
 
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If LcTipoDados = 3 Then
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'Me.Height = 6885
'Me.Width = 10425
BuscaNomeGalpao
Txt(9).Visible = FrmPrincipal.MnSaida.Visible
Label24.Visible = FrmPrincipal.MnSaida.Visible
Command2.Tag = "Order by Codigo"
'Valor(1).Format = CriaMascara
'Valor(2).Format = CriaMascara
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
'Top = Screen.Height / 2 - Height / 2
'Left = Screen.Width / 2 - Width / 2
'Left = Screen.Width / 2 - Width / 2
LcIndice = "NOME"
Label6.Visible = GlRepresentante
Label11.Visible = GlRepresentante
'Label10.Visible = GlRepresentante
Txt(18).Visible = GlRepresentante
Txt(19).Visible = GlRepresentante
fornecedor.Visible = GlRepresentante
carregaFornecedor
End Sub
Function AlteraPreco(LcIndex As Integer)
On Error Resume Next
Dim LcLucro         As Double
Dim LcPrecoVenda    As Double
Dim LcMinimo        As Double
Dim LcCusto         As Double
Dim LcDesp          As Double
Dim LcCustoFinal    As Double
Dim LPercent        As Double
Dim LCLEtra         As String
Dim LcPerc          As String
Dim a               As Integer
If GlDecimais = 0 Then GlDecimais = 2
LPercent = CStr((CDbl(valor(2).Text) - CDbl(valor(6).Text)) / CDbl(valor(6).Text))
Select Case LcIndex
    Case Is = 3
       LcCusto = CDbl(valor(3).Text)
       '=== Calcula o Valor do Custo, Buscando o Percentual
       For a = 1 To Len(valor(5).Text)
           LCLEtra = Mid(valor(5).Text, a, 1)
           If LCLEtra = "+" Then
           '==== Encontrou o separador, Calcula O Perc.
              LcCusto = LcCusto + (LcCusto * (CDbl(LcPerc) / 100))
              LcPerc = ""
           Else
             LcPerc = LcPerc & LCLEtra
           End If
       Next
       If Len(Trim(LcPerc)) > 0 Then
          LcCusto = LcCusto + (LcCusto * (CDbl(LcPerc) / 100))
       Else
          LcPerc = (CStr((CDbl(valor(6).Text) * 100) / CDbl(valor(3).Text)) - 100)
          valor(5).Text = LcPerc
       End If
       If Len(valor(7).Text) = 0 Then valor(7).Text = 0
       valor(1).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(0).Text) / 100))), GlDecimais)
       'Valor(2).Text = AcertaNumero(CStr(CDbl(Valor(6).Text) * (LPercent + 1)), GlDecimais)
        valor(2).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(7).Text) / 100))), GlDecimais)
 
    Case Is = 5
       LcCusto = CDbl(valor(3).Text)
       '=== Calcula o Valor do Custo, Buscando o Percentual
       For a = 1 To Len(valor(5).Text)
           LCLEtra = Mid(valor(5).Text, a, 1)
           If LCLEtra = "+" Then
           '==== Encontrou o separador, Calcula O Perc.
              LcCusto = LcCusto + (LcCusto * (CDbl(LcPerc) / 100))
              LcPerc = ""
           Else
             LcPerc = LcPerc & LCLEtra
           End If
       Next
       If Len(Trim(LcPerc)) > 0 Then
          LcCusto = LcCusto + (LcCusto * (CDbl(LcPerc) / 100))
       End If
       valor(6).Text = AcertaNumero(CStr(LcCusto), GlDecimais)
       valor(1).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(0).Text) / 100))), GlDecimais)
       'Valor(2).Text = AcertaNumero(CStr(CDbl(Valor(6).Text) * (LPercent + 1)), GlDecimais)
       valor(2).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(7).Text) / 100))), GlDecimais)

    Case Is = 6
       valor(1).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(0).Text) / 100))), GlDecimais)
       'Valor(2).Text = AcertaNumero(CStr(CDbl(Valor(6).Text) * (LPercent + 1)), GlDecimais)
      valor(2).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(7).Text) / 100))), GlDecimais)

    Case Is = 0
       valor(1).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(0).Text) / 100))), GlDecimais)
      ' Valor(2).Text = AcertaNumero(CStr(CDbl(Valor(6).Text) * (LPercent + 1)), GlDecimais)
    Case Is = 1
      ' valor(0).Text = AcertaNumero(CStr(((CDbl(valor(1).Text) - CDbl(valor(6).Text)) / CDbl(valor(6).Text)) * 100), GlDecimais)
       valor(0).Text = (CStr((CDbl(valor(1).Text) * 100) / CDbl(valor(6).Text)) - 100)
    Case Is = 7
       valor(2).Text = AcertaNumero(CStr(CDbl(valor(6).Text) + (CDbl(valor(6).Text) * (CDbl(valor(7).Text) / 100))), GlDecimais)
      ' Valor(2).Text = AcertaNumero(CStr(CDbl(Valor(6).Text) * (LPercent + 1)), GlDecimais)
    Case Is = 2
      ' valor(0).Text = AcertaNumero(CStr(((CDbl(valor(1).Text) - CDbl(valor(6).Text)) / CDbl(valor(6).Text)) * 100), GlDecimais)
       valor(7).Text = AcertaNumero(CStr(((CDbl(valor(2).Text) * 100) / CDbl(valor(6).Text)) - 100), 2)

End Select
    
Exit Function




If Not GlLucroCad Then Exit Function
LcLucro = CCur(valor(0).Text)
LcPrecoVenda = CCur(valor(1).Text)
LcMinimo = CCur(valor(2).Text)
LcCusto = CCur(valor(3).Text)
If Len(valor(5).Text) > 0 Then
   LcDesp = CCur(valor(5).Text)
Else
   LcDesp = 0
End If
LcCusto = ((LcDesp / 100) * LcCusto) + LcCusto

Select Case LcIndex
    Case Is = 0
       LcPrecoVenda = ((LcLucro / 100) + 1) * LcCusto
       valor(1).Text = LcPrecoVenda
    Case Is = 1
       LcLucro = ((LcPrecoVenda / LcCusto) - 1) * 100
       valor(0).Text = LcLucro
    Case Is = 3
        LcLucro = ((LcPrecoVenda / LcCusto) - 1) * 100
        valor(0).Text = LcLucro
    Case Is = 5
       LcPrecoVenda = ((LcLucro / 100) + 1) * LcCusto
       valor(1).Text = LcPrecoVenda
       valor(6).Text = LcCusto
End Select
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaBanco

'If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
'   GlPergunta = True
'   Call Salv
'End If
'If (LcTipoDados = 2) And LcAlterado Then Call Salv
'FechaBanco
GlStringBase = ""
FrmPrincipal.Visible = True
LcCarregado = False
GlAlteraCodigo = False
FrmPrincipal.SetFocus
End Sub

Private Sub Fornecedor_Change()
If LcTipoDados <> 3 Then
   
    On Error Resume Next
    Dim a As Integer
    Nome.Text = ""
    For a = 0 To LcTam
       If fornecedor.Text = MtFornecedor(a).Nome Then
          If err.Number > 0 Then Exit For
          GlCampo19 = MtFornecedor(a).Codigo & ""
          Nome.Text = MtFornecedor(a).Codigo & ""
          Alterado
          Exit For
        End If
        If err.Number > 0 Then Exit For
    Next
    CmdSalvar.Enabled = True
End If
End Sub

Private Sub fornecedor_Click()
Dim a As Integer
For a = 0 To LcTam
   If fornecedor.Text = MtFornecedor(a).Nome Then
           GlCampo19 = MtFornecedor(a).Codigo & ""
           Nome.Text = MtFornecedor(a).Codigo & ""
           Alterado
           Exit For
    End If
Next


End Sub

Private Sub Fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   If KeyCode = 16 And Index <> 13 Then
   Else
      Call Teclas(KeyCode)
   End If
End If

End Sub

Private Sub Fornecedor_LostFocus()
On Error Resume Next
Dim a As Integer
Nome.Text = ""
For a = 0 To LcTam
   If fornecedor.Text = MtFornecedor(a).Nome Then
      If err.Number > 0 Then Exit For
      GlCampo19 = MtFornecedor(a).Codigo & ""
      Nome.Text = MtFornecedor(a).Codigo & ""
      Alterado
      Exit For
    End If
    If err.Number > 0 Then Exit For
Next
'CmdSalvar.Enabled = True
End Sub

Private Sub Minimo_Change(Index As Integer)

End Sub

Private Sub Icms_Change()
On Error Resume Next
If LcTipo <> 3 Then CmdSalvar.Enabled = True
   
End Sub

Private Sub Icms_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   If KeyCode = 16 And Index <> 13 Then
   Else
      Call Teclas(KeyCode)
     ' If Index = 13 Then txt(16).SetFocus
   End If
End If
If Index = 0 Then
   If Not GlAlteraCodigo Then
          GlAlteraCodigo = True
          GlCodigoAnterior = GlCampo0
   End If
End If

End Sub

Private Sub MnAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, produto) Then VinculaDados RsAtualP!Codigo
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnExcluir_Click()
On Error Resume Next
If Exclui(produto) = 1 Then
      VinculaDados RsAtualP!Codigo
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
If MovImentacao(enPrimeiro, produto) Then VinculaDados RsAtualP!Codigo
GlMov = False
LcRegAtual = False
End Sub

Private Sub MnSair_Click()
Unload Me
End Sub

Private Sub MnSalvar_Click()
Call SalvaRegistro(produto)
VinculaDados RsAtualP!Codigo
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
End Sub

Private Sub MnUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, produto) Then VinculaDados RsAtualP!Codigo
GlMov = False
LcRegAtual = False
End Sub

Private Sub MSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, produto) Then VinculaDados RsAtualP!Codigo
GlMov = False
LcRegAtual = False
End Sub



Private Sub Seguranca_KeyDown(KeyCode As Integer, Shift As Integer)
Call Alterado
End Sub

Private Sub subi_Click()
If LcTipoDados <> 3 Then
    Command5.Enabled = subi.Value
    'If subi.Value Then Salv
    CmdSalvar.Enabled = True
End If
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
End Function
Function VinculaDados(LcCod As Integer)
On Error Resume Next
Dim LcVal1  As String
Dim LcVal2  As String
Dim LcVal3  As String
Dim LcVal4  As String
Dim LcVal5  As String
Dim Lcval6  As String
Dim LcVal7  As String
Dim Estoque As ControleDb
ValorAuterado = False
Set RsAtualP = AbreRecordset("Select * from produtos " & Command2.Tag, True)
RsAtualP.MoveFirst
RsAtualP.Find "codigo=" & LcCod, 0, adSearchForward
RetornaCorFundo
If LcTipoDados = 1 Then
   If Mantem = 0 Then NovoReg
   
Else
   'Call RegistroAtual(produto)
   Set Estoque = New ControleDb
   
End If
If IsNumeric(RsAtualP!Desativado) Then
   If RsAtualP!Desativado Then
      Desativado.Value = 1
   Else
      Desativado.Value = 0
   End If
Else
   Desativado.Value = 0
End If
'Desativado.Value = IIf(IsNumeric(RsAtualP!Desativado), 1, 0)
If IsDate(RsAtualP!UltimaAlteracao) Then
  UltimaAlteracao.Text = Format(CDate(RsAtualP!UltimaAlteracao), "dd/mm/yy")
Else
  UltimaAlteracao.Text = "  /  /  "
End If
'UltimaAlteracao.Text = IIf(IsDate(RsAtualP!UltimaAlteracao), Format(CDate(RsAtualP!UltimaAlteracao), "dd/mm/yy"), "  /  /  ")
CEST.Text = RsAtualP!CEST & ""
Txt(0).Text = RsAtualP!Codigo & ""
Txt(1).Text = RsAtualP!Nome & "" 'VerificaTipo(1, GlCampo1)
Txt(3).Text = RsAtualP!cst & "" ' VerificaTipo(3, GlCampo3)
Txt(7).Text = RsAtualP!ipi & ""
Txt(5).Text = RsAtualP!percentualcusto & ""
valor(2).Text = RsAtualP!MinimoVenda & ""
valor(4).Text = RsAtualP!MinimoEst & ""
Txt(10).Text = RsAtualP!maximoEstoque & ""
Txt(18).Text = RsAtualP!ComissaoFornecedor & ""
valor(3).Text = RsAtualP!Custo & ""
Txt(13).Text = RsAtualP!UnidMedida & ""
valor(1).Text = RsAtualP!Preco & ""
valor(0).Text = RsAtualP!Lucro & ""
valor(7).Text = RsAtualP!per & ""
valor(8).Text = RsAtualP!LimiteVenda & ""
Txt(12).Text = RsAtualP!ClassificacaoFiscal & ""

'===> Impostos
valor(9).Text = RsAtualP!ValorST & ""
valor(10).Text = RsAtualP!ValorIpi & ""
valor(11).Text = RsAtualP!valorFrete & ""
valor(12).Text = RsAtualP!valorSeguro & ""
valor(13).Text = RsAtualP!ValorDespesas & ""

BuscaNCM
Txt(16).Text = RsAtualP!QtdMedida & ""
valor(5).Text = RsAtualP!percentualcusto & ""
valor(6).Text = RsAtualP!CustoTotal & ""
Nome.Text = RsAtualP!fornecedor & ""
icms.Text = RsAtualP!icms & ""
Txt(14).Text = RsAtualP!subitem & ""
If LcTipoDados <> 1 Then
    Estoque.ArmazenaEmGalpao = True
    Estoque.CodProduto = Txt(0).Text
    Txt(20).Text = Estoque.EstoqueTotalFechado
    Txt(9).Text = Estoque.EstoqueTotalUnitario
    EstSanta.Text = Estoque.Santa1Fechado
    MinSanta.Text = Estoque.Santa1Unitario
    EstSanta2.Text = Estoque.Santa2Fechado
    MinSanta2.Text = Estoque.Santa2Unitario
    EstCalifornia.Text = Estoque.QuantidadeCaliforniaFechado
    MinCalifornia.Text = Estoque.QuantidadeCaliforniaUnitario
    Seguranca = Estoque.EstoqueSegurancaTotalFechado
End If



If RsAtualP!multiplositens Then subi.Value = 1 Else subi.Value = 0

Command5.Enabled = subi.Value
BuscaFornecedor
'Fornecedor.Text = rsatualp!Fornecedor & ""
BuscaUnidade
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
If Len(Trim(Txt(0).Text)) > 0 Then
  CmdCadastraInpostos.Enabled = True
Else
 CmdCadastraInpostos.Enabled = False
End If
If Not IsEmpty(GlIniceAtual) Then
   Txt(GlIniceAtual).SetFocus
Else
   Txt(0).SetFocus
End If
LcRegAtual = False
Exit Function
ErroVinculo:
Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
On Error Resume Next

If Index = 6 Then
   Txt(8).Text = CCur(Txt(5).Text) - CCur(Txt(6).Text)
End If
Call Alterado
GlCampo8 = valor(2).Text
GlCampo22 = valor(4).Text
GlCampo14 = valor(1).Text
GlCampo12 = valor(3).Text
GlCampo17 = valor(0).Text
GlCampo21 = valor(5).Text
GlCampo23 = valor(6).Text
End Sub


Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If LcTipoDados <> 3 Then
    
    If KeyCode = 13 Then
       SendKeys "{TAB}"
    Else
       If KeyCode = 16 And Index <> 13 Then
       Else
          Call Teclas(KeyCode)
         ' If Index = 13 Then txt(16).SetFocus
       End If
    End If
    If Index = 0 Then
       If Not GlAlteraCodigo Then
              GlAlteraCodigo = True
              GlCodigoAnterior = GlCampo0
       End If
    End If
End If
'Call MoveTecla(Index, KeyCode)
End Sub
Function limpa()
Dim a As Long
On Error Resume Next
For a = 0 To 36
  Txt(a).Text = ""
Next
For a = 0 To 13
  valor(a).Text = ""
Next
EstSanta.Text = ""
MinSanta.Text = ""

EstSanta2.Text = ""
MinSanta2.Text = ""

EstCalifornia.Text = ""
 MinCalifornia.Text = ""

Txt(16).Text = ""

Txt(0).SetFocus
CmdSalvar.Enabled = False
 CmdCadastraInpostos.Enabled = False
fornecedor.Text = " "
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 10 Then
   If KeyAscii = 46 Then KeyAscii = 44
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
If LcTipoDados <> 3 Then

    If Index = 0 Then
       If Not GLCalculacodigoProduto Then Txt(0).Text = Trim(Txt(0).Text)
    End If
    If Index = 12 Then
       If Len(Txt(12).Text) > 0 Then
           BuscaNCM
       End If
    End If
    If LcTipoDados = 1 Then
      VerificaDuplicado (Index)
    End If
    If Not GLCalculacodigoProduto Then If VerificaDuplicado(Index) Then Txt(Index).SetFocus
    If Index = 13 Then
      If Not IsNumeric(Txt(13).Text) Then
       If Len(Txt(Index).Text) = 0 Then Exit Sub
       MsgBox "Digite o Código da Unidade" & Chr(13) & "Ou Pressione F5 para Selecionar.", 64, "Aviso"
       Txt(13).Text = ""
       Txt(13).SetFocus
       Exit Sub
     End If
    End If
    If Index = 16 Or Index = 7 Or Index = 10 Or Index = 20 Or Index = 9 Or Index = 18 Then
      If Not IsNumeric(Txt(Index).Text) Then
       If Len(Txt(Index).Text) = 0 Then Exit Sub
       MsgBox "Digite um Valor Numérico.", 64, "Aviso"
       Txt(Index).Text = ""
       Txt(Index).SetFocus
       Exit Sub
     End If
     
    End If
    
    
    If Index = 13 Then BuscaUnidade
    'txt(Index).Text = AcertaNumero(Index, txt(Index))
End If
End Sub

Function BuscaUnidade()

Dim RsUnidade As Recordset
AbreBase
If Len(Trim(Txt(13).Text)) = 0 Then Exit Function
''abreconexao
Set RsUnidade = Dbbase.OpenRecordset("select * from alid004 order by nome", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Txt(13).Text = Right("00" & Txt(13).Text, 2)
LcCriterioCidade = "cod='" & Txt(13).Text & "'"
RsUnidade.FindFirst LcCriterioCidade
If Not RsUnidade.NoMatch Then
   Unidade.Caption = RsUnidade!Nome
   LcDesCidade = RsUnidade!Nome
Else
   'MsgBox "O código da Unidade não foi encontrado...,", 64, "Aviso"
   LcDesCidade = ""
   Unidade.Caption = ""
End If
RsUnidade.Close
Set RsUnidade = Nothing

End Function
Function BuscaFornecedor()
Dim RsUnidade As Recordset
''abreconexao
AbreBase
'If Len(Trim(txt(13).Text)) = 0 Then Exit Function
Set RsUnidade = Dbbase.OpenRecordset("select * from alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterioCidade = "codigo='" & Nome.Text & "'"
RsUnidade.FindFirst LcCriterioCidade
If Not RsUnidade.NoMatch Then
   fornecedor.Text = RsUnidade!RazaoSoc
Else
   fornecedor.Text = ""
   Nome.Text = ""
  ' MsgBox "O código da Unidade não foi encontrado...,", 64, "Aviso"
End If
RsUnidade.Close
Set RsUnidade = Nothing

End Function

Private Sub UltimaAlteracao_Change()
On Error Resume Next
CmdSalvar.Enabled = True
End Sub

Private Sub valor_Change(Index As Integer)
If LcRegAtual Then Exit Sub
Call Alterado
GlCampo8 = valor(2).Text
GlCampo22 = valor(4).Text
GlCampo14 = valor(1).Text
GlCampo12 = valor(3).Text
GlCampo17 = valor(0).Text
GlCampo21 = valor(5).Text
GlCampo23 = valor(6).Text
If Index = 5 Then
   If Not ValorAuterado Then
      ValorAuterado = True
      valor(5).Text = Replace(valor(5).Text, ".", ",")
   End If
End If

End Sub

Private Sub valor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If LcTipoDados <> 3 Then
    
    If KeyCode = 13 Then
       SendKeys "{TAB}"
    Else
       If KeyCode = 16 And Index <> 13 Then
       Else
          Call Teclas(KeyCode)
       End If
    End If
    If Index = 1 Or Index = 2 Or Index = 8 Then UltimaAlteracao = Format(Date, "dd/mm/yy")
End If
End Sub

Private Sub valor_KeyPress(Index As Integer, KeyAscii As Integer)

On Error Resume Next
If LcTipoDados <> 3 Then
 If KeyAscii = 46 Then KeyAscii = 44
End If
End Sub

Private Sub valor_LostFocus(Index As Integer)
If LcTipoDados <> 3 Then
    If Index <> 5 Then
     If Not IsNumeric(valor(Index).Text) Then
       If Len(valor(Index).Text) = 0 Then Exit Sub
       MsgBox "Digite um Valor Numérico.", 64, "Aviso"
       valor(Index).Text = ""
       valor(Index).SetFocus
       Exit Sub
     End If
    End If
    If Index >= 9 And Index <= 13 Then CalculaCusto
    If Index >= 3 Then CalculaCusto
    
    If Index = 5 Then ValorAuterado = False
    AlteraPreco (Index)
 End If
End Sub
Function SetaBarra()
StatusBar1.Panels(1).Width = Me.Width
StatusBar1.Panels(1).Text = "Ordem Atual:" & Replace(UCase(Command2.Tag), "ORDER BY", "")
StatusBar1.Font.Bold = True
End Function
