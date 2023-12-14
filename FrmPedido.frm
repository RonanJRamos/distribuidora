VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido"
   ClientHeight    =   7425
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
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
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
      Index           =   18
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
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
      Index           =   17
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox ComissaoProduto 
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
      Left            =   1680
      TabIndex        =   54
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ComissaoFabrica 
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
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox tam 
      Height          =   375
      Left            =   8520
      TabIndex        =   52
      Top             =   2880
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
      Left            =   7320
      TabIndex        =   51
      Top             =   0
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
      Height          =   405
      Index           =   16
      Left            =   10080
      TabIndex        =   48
      Top             =   3000
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
      Index           =   15
      Left            =   10080
      TabIndex        =   47
      Top             =   2160
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
      Height          =   525
      Index           =   14
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   45
      Top             =   6600
      Width           =   6975
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
      Left            =   120
      TabIndex        =   42
      Top             =   6600
      Width           =   2055
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
      Left            =   2400
      TabIndex        =   41
      Top             =   6600
      Width           =   2055
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
      Left            =   5160
      TabIndex        =   40
      Top             =   0
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
      Left            =   5880
      TabIndex        =   39
      Top             =   0
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
      Left            =   6600
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   615
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
      Index           =   10
      Left            =   1200
      TabIndex        =   3
      Top             =   960
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
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   6855
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
      TabIndex        =   12
      Top             =   2520
      Width           =   810
   End
   Begin VB.ComboBox Unidade 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
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
      Left            =   10080
      TabIndex        =   36
      Top             =   240
      Width           =   1575
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
      ItemData        =   "FrmPedido.frx":0000
      Left            =   7320
      List            =   "FrmPedido.frx":000A
      TabIndex        =   2
      Text            =   "VENDAS A VISTA"
      Top             =   570
      Width           =   2055
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
      Left            =   8520
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   12
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1695
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
      Left            =   6600
      TabIndex        =   31
      Top             =   120
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
      Left            =   6840
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   2295
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   10
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
      Top             =   600
      Width           =   1215
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
      Left            =   10080
      TabIndex        =   26
      Top             =   615
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
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
      Left            =   10080
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
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
      Left            =   10080
      TabIndex        =   24
      Top             =   990
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   6855
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
      Index           =   8
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
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
      Index           =   6
      Left            =   8520
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
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
      Index           =   5
      Left            =   7200
      TabIndex        =   14
      Top             =   2520
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
      TabIndex        =   13
      Top             =   2520
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
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
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para dar Desconto pressione F6"
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
      Left            =   4920
      TabIndex        =   55
      Top             =   2880
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Pedido"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   10080
      TabIndex        =   50
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Produtos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   10080
      TabIndex        =   49
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Informações Complementares"
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
      Left            =   4560
      TabIndex        =   46
      Top             =   6360
      Width           =   2085
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   44
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor Desconto"
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
      Left            =   2400
      TabIndex        =   43
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      TabIndex        =   37
      Top             =   960
      Width           =   1095
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
      TabIndex        =   35
      Top             =   3120
      Width           =   2925
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
      Left            =   2520
      TabIndex        =   34
      Top             =   2040
      Width           =   2850
   End
   Begin VB.Label Label3 
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
      Left            =   6000
      TabIndex        =   33
      Top             =   600
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   9960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   0
      X2              =   9960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   9960
      X2              =   9960
      Y1              =   0
      Y2              =   3480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Produtos Já Lançados"
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
      TabIndex        =   29
      Top             =   3600
      Width           =   1590
   End
   Begin VB.Label Label2 
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
      TabIndex        =   27
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Cod. Fornec."
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
      TabIndex        =   23
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      TabIndex        =   21
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
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
      TabIndex        =   19
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      TabIndex        =   17
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
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
      TabIndex        =   16
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "FrmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadosEntradaPedido
     item As Long
     CodPro As String
     Produto As String
     Und As String
     Qut As Long
     VUnit As Currency
     Vtotal As Currency
     Venda1 As Currency
     Venda2 As Currency
     Venda3 As Currency
     cst As String
     precomim As Currency
     icms As Long
     com As Long
     Comissao As Currency
End Type
Private LcItem As Long, LcTam As Long
Private FnunNota, FnunBoleto
Private LcNota, LcBoleto, LcEspC As String
Private LcFocus, a As Integer
Dim LcPrecoVelho, LcTotal As Currency
Private LcLinha As String
Private RsOpcoes As Recordset, RsClientes As Recordset
Private RsCidade As Recordset
Private LcMat() As DadosEntradaPedido


Private Sub CmdExcluir_Click()
FrmExcluiItem.Show
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 116 Then Teclas (KeyCode)
  Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdFechar_Click()
Unload Me
ReDim LcMat(0)
LcTam = 0
LcItem = 0
End Sub



Function CalculaValores()
Dim LcTotal As Currency, LcQuant As Long, LcUnit As Currency
On Error Resume Next

'=== Converte os Valores
LcQuant = CLng(Txt(3).Text)
LcUnit = CCur(Txt(5).Text)
LcTotal = LcQuant * LcUnit
Txt(6).Text = LcTotal

End Function
Function GeraGrid()
item.ColAlignment(0) = 7
item.ColAlignment(1) = 3
item.ColAlignment(2) = 1
item.ColAlignment(3) = 3
item.ColAlignment(4) = 1
item.ColAlignment(5) = 3
item.ColAlignment(6) = 3
item.ColAlignment(7) = 3
item.ColAlignment(8) = 3
item.ColAlignment(9) = 3

item.ColWidth(0) = 500
item.ColWidth(1) = 700
item.ColWidth(2) = 4600
item.ColWidth(3) = 500
item.ColWidth(4) = 1000
item.ColWidth(5) = 900
item.ColWidth(6) = 1200
item.ColWidth(7) = 1200
item.ColWidth(8) = 600
item.ColWidth(9) = 0

item.TextMatrix(0, 0) = "Item"
item.TextMatrix(0, 1) = "Código"
item.TextMatrix(0, 2) = "Descrição"
item.TextMatrix(0, 3) = "CST"
item.TextMatrix(0, 4) = "Unidade"
item.TextMatrix(0, 5) = "Quant"
item.TextMatrix(0, 6) = "Unitário"
item.TextMatrix(0, 7) = "Total"
item.TextMatrix(0, 8) = "ICMS"

LcTamanhoGrid = 1
End Function
Function montagrid()
Dim LcAchou As Integer
On Error Resume Next
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
If Len(Trim(Txt(5).Text)) = 0 Or Txt(5).Text = "0" Then
   MsgBox "Necessário Informar o Valor Unitario do Item.", 48, "Aviso"
   Txt(5).SetFocus
   Exit Function
End If
LcItem = LcItem + 1
ReDim Preserve LcMat(LcTam)
LcMat(LcTam).item = LcItem
LcMat(LcTam).CodPro = Txt(1).Text
LcMat(LcTam).Produto = Txt(2).Text
LcMat(LcTam).Qut = CLng(Txt(3).Text)
LcMat(LcTam).Und = Unidade.Text
LcMat(LcTam).com = Txt(4).Text
LcMat(LcTam).VUnit = CCur(Txt(5).Text)
LcMat(LcTam).Vtotal = CCur(Txt(6).Text)
LcMat(LcTam).Venda1 = CCur(Custo.Text)
LcMat(LcTam).cst = cst.Text
LcMat(LcTam).icms = icms.Text
If Not IsNull(ComissaoProduto.Text) Then
   LcMat(LcTam).Comissao = (CCur(ComissaoProduto.Text) / 100) * CCur(Txt(6).Text)
Else
   If Not IsNull(ComissaoFabrica.Text) Then
      LcMat(LcTam).Comissao = (CCur(ComissaoFabrica.Text) / 100) * CCur(Txt(6).Text)
   End If
End If
LcTam = LcTam + 1
EscreveGrid

For a = 1 To 6
   Txt(a).Text = ""
Next
Txt(3).Text = " "
Txt(5).Text = " "
Txt(6).Text = " "
Custo.Text = "0"
icms.Text = "0"
cst.Text = "0"
minimo.Text = "0"
ComissaoFabrica.Text = ""
ComissaoProduto.Text = ""
Txt(1).SetFocus
End Function
Function EscreveGrid()
Dim b As Integer
b = 1
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       item.Rows = b + 1
       item.TextMatrix(b, 0) = LcMat(a).item
       item.TextMatrix(b, 1) = LcMat(a).CodPro
       item.TextMatrix(b, 2) = LcMat(a).Produto
       item.TextMatrix(b, 3) = LcMat(a).cst
       item.TextMatrix(b, 4) = LcMat(a).Und & " C/" & LcMat(a).com
       item.TextMatrix(b, 5) = LcMat(a).Qut
       item.TextMatrix(b, 6) = Format(LcMat(a).VUnit, "Currency")
       item.TextMatrix(b, 7) = Format(LcMat(a).Vtotal, "Currency")
       item.TextMatrix(b, 8) = LcMat(a).icms
       b = b + 1
    End If
Next
CalculaIcms
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True

End Function
Function CalculaIcms()
Dim LcBaseCalculo, LcIcms, LcPRodutos, LcNota As Currency
Dim LcItem As String, LcComp As String
Dim LcQuantItemSubs As Integer
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
         LcItem = LcItem & Right("00" & CStr(LcMat(a).item), 2)
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
   'Txt(14).Text = LcComp
End If

'Txt(13).Text = Format(LcBaseCalculo, "Currency")
'Txt(11).Text = Format(LcIcms, "Currency")
Txt(15).Text = Format(LcPRodutos, "Currency")
Txt(16).Text = Format(LcNota, "Currency")

End Function
Function VerificaDesconto()

Dim LcCaracter, LcPalavra, LcDesconto As String
Dim Lct As Long
LcTotal = CCur(Txt(15).Text)
Lct = Len(Txt(13).Text)

For a = 1 To Lct
    LcCaracter = Mid(Txt(13).Text, a, 1)
    If LcCaracter = "+" Then
       CalculaDesconto (CCur(LcDesconto))
       LcDesconto = ""
    Else
       LcDesconto = LcDesconto & LcCaracter
    End If
Next
If Len(LcDesconto) > 0 Then CalculaDesconto (CCur(LcDesconto))
    
       
    
End Function
Function CalculaDesconto(LcDesconto As Double)

Dim LcValorDes As Currency
LcValorDes = (LcDesconto / 100) * LcTotal
Txt(16).Text = Format((LcTotal - LcValorDes), "currency")
LcTotal = CCur(Txt(16).Text)
Txt(11).Text = Format((CCur(Txt(15).Text) - CCur(Txt(16).Text)), "currency")

End Function
Function RemontaIndice()

LcItem = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      LcItem = LcItem + 1
      LcMat(a).item = LcItem
   End If
Next


End Function
Function CarregaCboUnidade()
On Error Resume Next
Dim LcAchou As Integer
Dim RsUnidade As Recordset
Dim LcPrimeiro As String
AbreBase
Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 order By SIMBOLO")
Do Until RsUnidade.EOF

   Unidade.AddItem RsUnidade!Simbolo
   RsUnidade.MoveNext
Loop
RsUnidade.Close
Dbbase.Close
Set RsUnidade = Nothing
Set Dbbase = Nothing


End Function
Function BuscaProduto()
On Error Resume Next
If Len(Txt(2).Text) > 0 And Len(Txt(1).Text) = 0 Then Exit Function
Dim LcAchou As Integer
Dim RsProduto As Recordset, RsUnidade As Recordset
AbreBase
Set RsProduto = Dbbase.OpenRecordset("select * From alid009 where cod='" & Txt(1).Text & "'") ', dbOpenDynaset)
If Not RsProduto.EOF Then
   'LcCriterio = "Cod='" & RsProduto!Unimed & "'"
   Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!Unimed & "'")  ', dbOpenDynaset)

   'RsUnidade.FindFirst LcCriterio
   If Not RsUnidade.EOF Then
      LcUnidade = RsUnidade!Simbolo
   End If
   Txt(1).Text = RsProduto!cod
   Txt(2).Text = RsProduto!Nome
   Unidade.Text = LcUnidade
   Txt(4).Text = RsProduto!QTDUNIMED
   Txt(5).Text = RsProduto!Ptab
   cst.Text = RsProduto!cst
   LcPrecoVelho = RsProduto!Ptab
   If Not IsNull(RsProduto!ComissaoFornecedor) Then
      ComissaoProduto.Text = RsProduto!ComissaoFornecedor
   End If
   minimo.Text = RsProduto!MPVENDA
   If Val(cst.Text) = 60 Or Val(cst.Text) = 160 Or Val(cst.Text) = 260 Then
      icms.Text = "00"
   Else
     icms.Text = "18"
   End If
   
   'Custo.Text = RsProduto!Custo
   LcAchou = True
Else
   Txt(1).Text = ""
   Txt(2).Text = ""
   LcAchou = False
   Txt(1).SetFocus
   MsgBox "Código Não Encontrado...", 64, "Aviso"
End If
If LcAchou Then SendKeys "{TAB}"
RsProduto.Close
RsUnidade.Close
Set RsProduto = Nothing
Set RsUnidade = Nothing

End Function
Function BuscaVendendor()
'On Error Resume Next
Dim LcAchou As Integer
Dim RsVendedor As Recordset
AbreBase
If Len(Comissao.Text) = 0 Then Comissao.Text = 0
Txt(10).Text = Right("00000" & Txt(10).Text, 5)
Set RsVendedor = Dbbase.OpenRecordset("select * From alid200 where codigo='" & Txt(10).Text & "'") ', dbOpenDynaset)
If Not RsVendedor.EOF Then
   Txt(7).Text = RsVendedor!Nome
   If CLng(Comissao.Text) <> 1 Then
      Comissao.Text = RsVendedor!Comissao
   End If
   LcAchou = True
Else
   LcAchou = False
   Txt(10).Text = ""
   Comissao.Text = 0
   Txt(7).Text = ""
   Txt(10).SetFocus
   MsgBox "Vendedor Não Encontrado...", 64, "Aviso"
End If
If LcAchou Then SendKeys "{TAB}"
RsVendedor.Close
Set RsVendedor = Nothing


End Function
Function BuscaFornecedor(LcBusca As Integer)
On Error Resume Next
Dim LcAchou As Integer
LcSql = "select * from alid002 where CODIGO='" & Right("00000" & Txt(8).Text, 5) & "'"
LcMsg = "Codigo do Fornecedor Não Cadastrado..."

AbreBase

Set RsAtual = Dbbase.OpenRecordset(LcSql)
If Not RsAtual.EOF Then
   Txt(9).Text = RsAtual!razaosoc
   If Not IsNull(RsAtual!COMISSAOREPRESENTANTE) Then
      ComissaoFabrica.Text = RsAtual!COMISSAOREPRESENTANTE
   End If
   Txt(18).SetFocus
Else
   MsgBox LcMsg, 64, "Aviso"
   Txt(8).Text = ""
   Txt(8).SetFocus
End If
RsAtual.Close
Set RsAtual = Nothing

End Function
Function BuscaCliente(LcBusca As Integer)
On Error Resume Next
Dim LcAchou As Integer
If GLCalculacodigoCliente Then
    LcSql = "select * from alid001 where CODIGO='" & Right("00000" & Txt(18).Text, 5) & "'"
Else
   LcSql = "select * from alid001 where CODIGO='" & Txt(18).Text & "'"
End If
LcMsg = "Codigo do Cliente Não Cadastrado..."

AbreBase

Set RsAtual = Dbbase.OpenRecordset(LcSql)
If Not RsAtual.EOF Then
   Txt(17).Text = RsAtual!razaosoc
   Txt(1).SetFocus
Else
   MsgBox LcMsg, 64, "Aviso"
   Txt(18).Text = ""
   Txt(18).SetFocus
End If
RsAtual.Close
Set RsAtual = Nothing

End Function

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 116 Then Teclas (KeyCode)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub


Function limpanota()
On Error Resume Next
Liberado = False
LcTam = 0
LcItem = 0
ReDim LcMat(0)
item.Rows = 1
For a = 0 To 18
   Txt(a).Text = ""
   'valor.Text = ""
Next
limite.Text = 0
utilizado.Text = 0
CalculaNumeroNota
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
Txt(7).SetFocus
End Function
Function imprimepedido()
'On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If

'Abertura do relatório de vendas
    
    
   CryRelatorio.DataFiles(0) = GLBase
   CryRelatorio.ReportFileName = App.Path & "\Pedido.rpt"
   LcFormula = "{Pedido.NUMNF}='" & Txt(0).Text & "'"
   
      
   CryRelatorio.CopiesToPrinter = 0

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Pedido de Compra"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
CryRelatorio.Formulas(5) = "titulo='Pedido'"
 
 
LcTipoSaida = 0

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
Private Sub CmdSalvar_Click()
Dim RsProduto As Recordset, RsEntrada As Recordset
Dim LcCodigo As String
On Error Resume Next

AbreBase
Set RsProduto = Dbbase.OpenRecordset("Produto", dbOpenTable, dbSeeChanges, dbOptimistic)
Set RsEntrada = Dbbase.OpenRecordset("SaidaProduto", dbOpenTable, dbSeeChanges, dbOptimistic)
RsProduto.Index = "Codigo"
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       LcCodigo = LcMat(a).CodPro
       RsProduto.Seek "=", LcCodigo
       If Not RsProduto.NoMatch Then
          RsProduto.Edit
          RsProduto("Estoque") = RsProduto!Estoque - LcMat(a).Qut
                 
          RsProduto.Update
       End If
       RsEntrada.AddNew
       RsEntrada("Doc") = Txt(0).Text
       RsEntrada("Produto") = LcCodigo
       RsEntrada("Descricao") = LcMat(a).Produto
       RsEntrada("Quantidade") = LcMat(a).Qut
       RsEntrada("ValorUnitario") = LcMat(a).VUnit
       RsEntrada("ValorTotal") = LcMat(a).Vtotal
       RsEntrada("DataSaida") = CDate(Format(Txt(12).Text, "dd/mm/yy"))
       RsEntrada("CodigoCliente") = CInt(Txt(8).Text)
       RsEntrada("NomeCliente") = Txt(9).Text
       RsEntrada("Unidade") = LcMat(a).Und
       RsEntrada("custo") = LcMat(a).Venda1
       RsEntrada.Update
    End If
Next
MsgBox "Os Dados Foram Salvos Com Sucesso...", 48, "Aviso"
ReDim LcMat(0)
LcTam = 0
LcItem = 0
For a = 0 To 30
   Txt(a).Text = ""
Next
item.Rows = 1
Txt(3).Text = "0"
Txt(5).Text = "0"
Txt(6).Text = "0"
Custo.Text = "0"

Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True

Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Txt(0).SetFocus
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Command1_Click()
FrmPesquisaCliente.Show , Me
End Sub

Private Sub Command2_Click()
FrmPesquisaProdutos.Show , Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
tam.Text = LcTam
DadosPedido.Show , Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 116 Then Teclas (KeyCode)
  Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If Not GlCarregado Then
   Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
   GlCarregado = True
End If

End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
Me.Height = 7800
Me.Width = 11970
GlEscolhe = 1
CalculaNumeroNota
CarregaCboUnidade

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
GlCarregado = False
End Sub

Private Sub Natureza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  SendKeys "{TAB}"
Else
  If KeyCode <> 116 Then Teclas (KeyCode)
End If

End Sub



Private Sub txt_Change(Index As Integer)
On Error Resume Next
If Index = 3 Or Index = 5 Then CalculaValores

End Sub

Private Sub txt_GotFocus(Index As Integer)
If Index = 8 Then
   If Len(Trim(Txt(10).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Vendedor Responsável.", 64, "Aviso"
      Txt(10).SetFocus
   End If
End If
If Index = 1 Then
   If Len(Trim(Txt(8).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Cliente para a Nota Fiscal.", 64, "Aviso"
      Txt(8).SetFocus
   End If
End If
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 117 Then
   Txt(13).SetFocus
   Exit Sub
End If
If KeyCode = 38 Then
   VoltaCampo (KeyCode)
End If
If KeyCode = 116 Then
   If Index = 18 Or Index = 17 Then
      GlEscolhe = 3  'Exibe Clientes
      If Len(Trim(Txt(9).Text)) > 0 Then
            FrmPesquisaCliente.Txt.Text = Txt(17).Text
            GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt(17).Text) & "*'  order by RAZAOSOC"
         Else
            GlCriterioSql = ""
         End If
      Teclas (KeyCode)
   End If
   If Index = 8 Or Index = 9 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(9).Text)) > 0 Then
            FrmPesquisaFornecedores.Txt.Text = Txt(9).Text
            GlCriterioSql = "select * From alid002 where RAZAOSOC like '" & UCase(Txt(9).Text) & "*'  order by RAZAOSOC"
         Else
            GlCriterioSql = ""
         End If
      Teclas (KeyCode)
   Else
      If Index = 1 Or Index = 2 Then 'Exibe Produtos
         GlEscolhe = 2
         If Len(Trim(Txt(2).Text)) > 0 Then
            FrmPesquisaProdutos.Txt.Text = Txt(2).Text
            GlCriterioSql = "select * From alid009 where nome like '" & UCase(Txt(2).Text) & "*'  order by nome"
         Else
            GlCriterioSql = ""
         End If
         Teclas (KeyCode)
      End If
    End If
Else
 
  Teclas (KeyCode)
End If

   
End Sub
Function VoltaCampo(LcIndex As Integer)
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

Private Sub Txt_LostFocus(Index As Integer)
If Index = 6 And GlLibera Then montagrid
If Index = 1 Then
   If Len(Trim(Txt(1).Text)) > 0 Then
      If GLCalculacodigoProduto Then Txt(1).Text = Right("00000" & Txt(1).Text, 5)
      BuscaProduto
   End If
End If
If Index = 2 Then BuscaProduto
If Index = 8 Then
   If Len(Txt(8).Text) > 0 Then
      Txt(8).Text = Right("00000" & Txt(8).Text, 5)
      If Len(Trim(Txt(8).Text)) > 0 Then BuscaFornecedor (1)
   End If
End If
If Index = 18 Then
   If Len(Txt(18).Text) > 0 Then
      If GLCalculacodigoCliente Then Txt(18).Text = Right("00000" & Txt(18).Text, 5)
      If Len(Trim(Txt(18).Text)) > 0 Then BuscaCliente (1)
   End If
End If
'If Index = 9 Then BuscaFornecedor (2)

If Index = 5 Then
   ConferePreco
End If
If Index = 10 And Len(Trim(Txt(Index).Text)) <> 0 Then BuscaVendendor
If Index = 13 Then VerificaDesconto
End Sub
Function ConferePreco()
On Error Resume Next

Dim LcPreconovo, LcPRecoAntigo As Currency
GlLibera = False
If Len(minimo.Text) = 0 Then minimo.Text = 0
LcPreconovo = CCur(Txt(5).Text)

LcPRecoAntigo = CCur(minimo.Text)

If LcPreconovo < LcPRecoAntigo Then
    Liberacao.Show
    GlLibera = False
    GlEscolha = True
     Do Until Not GlEscolha
        DoEvents
     Loop
     If GlLibera Then
        Comissao.Text = 1
     Else
        Txt(5) = LcPrecoVelho
        Txt(5).SetFocus
     End If
Else
  GlLibera = True
  If Len(Comissao.Text) = 0 Then Comissao.Text = 0
  If CLng(Comissao.Text) <> 1 Then
     Comissao.Text = 2
  End If
End If

End Function
Function SalvaPedido()
Dim RsPedido As Recordset, RsItens As Recordset
Dim RsFornecedor As Recordset

LcSql1 = "Select * from pedido"
LcSql2 = "Select * from dadospedido"
LcSql3 = "Select * from Alid002"
AbreBase
Set RsPedido = Dbbase.OpenRecordset(LcSql1)
Set RsItens = Dbbase.OpenRecordset(LcSql2)
Set RsFornecedor = Dbbase.OpenRecordset(LcSql3)

'==== Grava Os dados da Nota Fiscal
RsPedido.AddNew
RsPedido("NUMNF") = Txt(0).Text
RsPedido("DTEMIS") = CDate(Format(Txt(12).Text, "dd/mm/yy"))
RsPedido("NATUREZA") = "PC"
RsPedido("CLiente") = Txt(18).Text
RsPedido("Fornecedor") = Txt(8).Text
RsPedido("TRANSP") = DadosPedido.Txt(0).Text
RsPedido("TIPOTRANS") = Mid(DadosPedido.Tipo.Text, 1, 1)
RsPedido("FoneTransp") = DadosPedido.Txt(10).Text
RsPedido("PLACATRANS") = DadosPedido.Placa.Text
RsPedido("UFTRANS") = DadosPedido.Txt(1).Text
RsPedido("CGCCPFTRAN") = DadosPedido.Txt(2).Text
RsPedido("ENDTRANS") = DadosPedido.Txt(3).Text
RsPedido("MUNICTRANS") = DadosPedido.Txt(4).Text
RsPedido("UFMUNIC") = DadosPedido.Txt(5).Text
RsPedido("INSCEST") = DadosPedido.Txt(6).Text
RsPedido("OBS02") = DadosPedido.Txt(7).Text
RsPedido("OBS03") = DadosPedido.Txt(8).Text
RsPedido("OBS04") = DadosPedido.Txt(9).Text
RsPedido("TotalProduto") = Txt(15).Text
RsPedido("TotalGeral") = Txt(16).Text
If Len(Txt(11).Text) > 0 Then RsPedido("TotalDesconto") = Txt(11).Text
RsPedido("Desconto") = Txt(13).Text
RsPedido("CondPag") = DadosPedido.TipoMonetario.Text
RsPedido.Update

'==== Grava os itens da nota fiscal

For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         RsItens.AddNew
         'RsItens("Vendedor") = txt(10).Text
         RsItens("NPedido") = Txt(0).Text
        ' RsItens("ITEM") = LcMat(a).CodPro
         RsItens("Quant") = LcMat(a).Qut
         RsItens("Unit") = LcMat(a).VUnit
         'RsItens("UNIMED") = LcMat(a).Und
         'RsItens("QTDUM") = CLng(LcMat(a).com)
         RsItens("Total") = LcMat(a).Vtotal
         RsItens("CodigoProduto") = LcMat(a).CodPro
         'RsItens("CLIENTE") = txt(8).Text
         RsItens.Update
     End If
Next
'==== Atualiza Dados Cliente
LcCriterioPes = "codigo='" & Txt(8).Text & "'"
RsFornecedor.FindFirst LcCriterioPes
If Not RsFornecedor.NoMatch Then
   RsFornecedor.Edit
   RsFornecedor("ULTCOMPRA") = CDate(Format(Txt(12).Text, "dd/mm/yy"))
   RsFornecedor.Update
End If

'=== Fecha as Bases
RsPedido.Close
RsItens.Close
RsFornecedor.Close
Dbbase.Close
Set RsPedido = Nothing
Set RsComissao = Nothing
Set RsFornecedor = Nothing

Set Dbbase = Nothing
         
End Function
Function SalvaOrcamento()
Dim Rsorcamento As Recordset, RsItens As Recordset
Dim RsFornecedor As Recordset

LcSql1 = "Select * from pedido"
LcSql2 = "Select * from dadospedido"
LcSql3 = "Select * from Alid002"
AbreBase
Set RsPedido = Dbbase.OpenRecordset(LcSql1)
Set RsItens = Dbbase.OpenRecordset(LcSql2)
Set RsFornecedor = Dbbase.OpenRecordset(LcSql3)

'==== Grava Os dados da Nota Fiscal
Rsorcamento.AddNew
Rsorcamento("NUMNF") = Txt(0).Text
Rsorcamento("DTEMIS") = CDate(Format(Txt(12).Text, "dd/mm/yy"))
Rsorcamento("NATUREZA") = "PC"
Rsorcamento("CLiente") = Txt(18).Text
Rsorcamento("Fornecedor") = Txt(8).Text
Rsorcamento("TRANSP") = DadosPedido.Txt(0).Text
Rsorcamento("TIPOTRANS") = Mid(DadosOrcamento.Tipo.Text, 1, 1)
Rsorcamento("FoneTransp") = DadosOrcamento.Txt(10).Text
Rsorcamento("PLACATRANS") = DadosOrcamento.Placa.Text
Rsorcamento("UFTRANS") = DadosOrcamento.Txt(1).Text
Rsorcamento("CGCCPFTRAN") = DadosOrcamento.Txt(2).Text
Rsorcamento("ENDTRANS") = DadosOrcamento.Txt(3).Text
Rsorcamento("MUNICTRANS") = DadosOrcamento.Txt(4).Text
Rsorcamento("UFMUNIC") = DadosOrcamento.Txt(5).Text
Rsorcamento("INSCEST") = DadosOrcamento.Txt(6).Text
Rsorcamento("OBS02") = DadosOrcamento.Txt(7).Text
Rsorcamento("OBS03") = DadosOrcamento.Txt(8).Text
Rsorcamento("OBS04") = DadosOrcamento.Txt(9).Text
Rsorcamento.Update

'==== Grava os itens da nota fiscal

For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         RsItens.AddNew
         'RsItens("Vendedor") = txt(10).Text
         RsItens("NPedido") = Txt(0).Text
        ' RsItens("ITEM") = LcMat(a).CodPro
         RsItens("Quant") = LcMat(a).Qut
         RsItens("Unit") = LcMat(a).VUnit
         'RsItens("UNIMED") = LcMat(a).Und
         'RsItens("QTDUM") = CLng(LcMat(a).com)
         RsItens("Total") = LcMat(a).Vtotal
         RsItens("CodigoProduto") = LcMat(a).CodPro
         'RsItens("CLIENTE") = txt(8).Text
         RsItens.Update
     End If
Next
'==== Atualiza Dados Cliente
LcCriterioPes = "codigo='" & Txt(8).Text & "'"
RsFornecedor.FindFirst LcCriterioPes
If Not RsFornecedor.NoMatch Then
   RsFornecedor.Edit
   RsFornecedor("ULTCOMPRA") = CDate(Format(Txt(12).Text, "dd/mm/yy"))
   RsFornecedor.Update
End If

'=== Fecha as Bases
RsPedido.Close
RsItens.Close
RsFornecedor.Close
Dbbase.Close
Set RsPedido = Nothing
Set RsComissao = Nothing
Set RsFornecedor = Nothing

Set Dbbase = Nothing
         
End Function
Function ExcluiItem(LcNItem As Integer)
Dim a, b As Integer

For a = 0 To LcTam - 1
    If LcMat(a).item = LcNItem Then
       LcMat(a).CodPro = ""
       LcAchou = True
       Exit For
       
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
LcSql = "Select * from pedido order by NUMNF"
AbreBase
Set RsNota = Dbbase.OpenRecordset(LcSql)
If RsNota.EOF Then
   LcNumeroNota = "000001"
Else
   RsNota.MoveLast
   
   If IsNull(RsNota("NUMNF")) Then
       LcNumeroNota = "000001"
   Else
       LcNumeroNota = Right("000000" & CStr(Val(RsNota("NUMNF")) + 1), 6)
   End If
End If
Txt(0).Text = LcNumeroNota

RsNota.Close
Dbbase.Close
Set RsNota = Nothing
Set Dbbase = Nothing

End Function

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 117 Then
   Txt(13).SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
  SendKeys "{TAB}"
Else
  If KeyCode <> 116 Then Teclas (KeyCode)
End If
End Sub

Function GeraComissao()
Dim RsComissao As Recordset
LcSql = "Select * from ComissaoPedido"
AbreBase
Set RsComissao = Dbbase.OpenRecordset(LcSql)

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
         RsComissao("COMISSAO") = LcMat(a).Comissao
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
