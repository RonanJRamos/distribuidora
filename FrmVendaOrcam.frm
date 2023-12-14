VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmVendaOrcam 
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
   Begin VB.TextBox dados 
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
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox fone 
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
      Left            =   4680
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox transportadora 
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
      Left            =   5400
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox controle 
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
      Left            =   0
      TabIndex        =   72
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Status 
      Height          =   285
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox unitreal 
      Height          =   285
      Left            =   6600
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   19
      Left            =   6600
      TabIndex        =   60
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox Acrecimo 
      Height          =   495
      Left            =   8760
      TabIndex        =   61
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   11
      Left            =   240
      TabIndex        =   64
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox ipi 
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
      Left            =   0
      TabIndex        =   63
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox desconto 
      Height          =   495
      Left            =   4320
      TabIndex        =   59
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox comisVenda 
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
      Left            =   4080
      TabIndex        =   57
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
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
      Left            =   10080
      TabIndex        =   56
      Top             =   1116
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
      Index           =   18
      Left            =   1200
      TabIndex        =   7
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
      Index           =   17
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   7335
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
      Left            =   5520
      TabIndex        =   53
      Top             =   0
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
      Left            =   4680
      TabIndex        =   52
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox tam 
      Height          =   375
      Left            =   8400
      TabIndex        =   51
      Top             =   120
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
      TabIndex        =   50
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
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2880
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
      TabIndex        =   46
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
      Left            =   11280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
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
      Left            =   2160
      TabIndex        =   41
      Top             =   7080
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
      Left            =   5400
      TabIndex        =   40
      Top             =   120
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
      Left            =   6120
      TabIndex        =   39
      Top             =   120
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
      Width           =   7335
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
      Top             =   2280
      Width           =   810
   End
   Begin VB.ComboBox Unidade 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
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
      Top             =   0
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
      ItemData        =   "FrmVendaOrcam.frx":0000
      Left            =   6120
      List            =   "FrmVendaOrcam.frx":000D
      TabIndex        =   2
      Text            =   "Orçamento"
      Top             =   570
      Width           =   1215
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
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
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
      Height          =   3255
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   11
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
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
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
      Top             =   744
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
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
      TabIndex        =   25
      Top             =   360
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
      Top             =   1485
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
      Visible         =   0   'False
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
      Index           =   6
      Left            =   8520
      TabIndex        =   20
      Top             =   2280
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
      Top             =   2280
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
      Top             =   2280
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
      Top             =   2280
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   7440
      TabIndex        =   71
      Top             =   600
      Width           =   615
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
      TabIndex        =   69
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para dar Acrécimo pressione F9"
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
      TabIndex        =   68
      Top             =   2880
      Width           =   2250
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Acrecimo R$"
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
      Left            =   8760
      TabIndex        =   66
      Top             =   6840
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total IPI"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   65
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
      TabIndex        =   62
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Detalhar o Produto Pressione F8"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   5640
      TabIndex        =   58
      Top             =   2640
      Width           =   3225
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   1320
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
      Left            =   3120
      TabIndex        =   54
      Top             =   2640
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Orçamento"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   10080
      TabIndex        =   49
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Produtos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   10080
      TabIndex        =   48
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   6600
      TabIndex        =   45
      Top             =   6840
      Width           =   825
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   10
      Left            =   2160
      TabIndex        =   43
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   4440
      TabIndex        =   42
      Top             =   6840
      Width           =   945
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
      Top             =   2640
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
      Top             =   1680
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
      Left            =   5280
      TabIndex        =   33
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   9960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   0
      X2              =   9960
      Y1              =   2040
      Y2              =   2040
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
      Left            =   240
      TabIndex        =   29
      Top             =   4200
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
      Top             =   2040
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
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unid. / Embal."
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
      Top             =   2040
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
      Top             =   2040
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
      Left            =   1080
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido Cliente"
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
      Width           =   5415
   End
End
Attribute VB_Name = "FrmVendaOrcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type DadosEntradaPedido
     item As String
     CodPro As String
     produto As String
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
     ipi As Long
     Com As Long
     comissao As Currency
     UnitarioAntigo As Currency
     controle As Long
End Type
Private LcItem As Long, LcTam, LcLinhaAtual As Long
Private Fnum, FnunNota, LcItensImpressos, LcCalculadoDesconto, LcCalculadoAcrecimo As Integer
Private LcNota, LcBoleto, LcEspC As String
Private LcFocus As Integer
Dim LcPrecoVelho, LcTotal, LcValorDes, LcValorAcre As Currency
Private LcLinha As String
Private RsOpcoes As Recordset, RsClientes As Recordset, RsEmpresa As Recordset
Private RsCidade As Recordset, RsI As Recordset
Private LcMat() As DadosEntradaPedido
Private LcPesquisa, LcVerificaAcrecimo As Integer
Private MtPedido(500) As String
Private LcTamanhoPedido As Long, LcPermissaoImpressao As Long
Private LcMargem As String


Private Sub Acrecimo_Change()
On Error Resume Next
Dim LcPRodutos As Currency
If LcVerificaAcrecimo Then Exit Sub
If Len(Txt(16).Text) = 0 Then
   'MsgBox "Não Foi Cadastrado Nenhum Item para dar Desconto...", 64, "Aviso"
   
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Sub
End If
LcPRodutos = CCur(Txt(15).Text)
Txt(16).Text = AcertaNumero(CCur(LcPRodutos + CCur(Acrecimo.Text) - CCur(Desconto.Text) + CCur(Txt(11).Text)), 2)
End Sub

Private Sub Acrecimo_GotFocus()
LcVerificaAcrecimo = False
End Sub

Private Sub Acrecimo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me

End Sub

Private Sub Acrecimo_LostFocus()
On Error Resume Next
Dim LcPrimeiro As Integer
Dim LcPercentual As Long
LcPrimeiro = True
If acrescimo.Text = "0" Then Exit Sub
If Len(acrescimo.Text) = 0 Then Exit Sub
If Len(Txt(16).Text) = 0 Then
   'MsgBox "Não Foi Cadastrado Nenhum Item para dar Desconto...", 64, "Aviso"
   
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Sub
End If
If Not GlRateiaAcrecimo Then
   Txt(16).Text = CCur(CCur(Txt(15).Text) - CCur(Desconto.Text)) + CCur(acrescimo.Text)
Else
   LcPercentual = (CCur(acrescimo.Text) / CCur(Txt(16).Text)) * 100
   Txt(19).Text = AcertaNumero(CStr(LcPercentual), 2)
   RateiaAcrescimo
End If
End Sub

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
If Natureza.Text <> "Emprestimo" Then
   LcQuant = CLng(Txt(3).Text)
   LcUnit = CCur(Txt(5).Text)
   LcTotal = LcQuant * LcUnit
   Txt(6).Text = LcTotal
Else
   LcQuant = CLng(Txt(3).Text)
   LcUnit = 0
   LcTotal = LcQuant * LcUnit
   Txt(6).Text = LcTotal
End If
End Function
Function GeraGrid()
item.ColAlignment(0) = 7
item.ColAlignment(1) = 3
item.ColAlignment(2) = 1
item.ColAlignment(3) = 3
item.ColAlignment(4) = 1
item.ColAlignment(5) = 3
item.ColAlignment(6) = 8
item.ColAlignment(7) = 8
item.ColAlignment(8) = 3
item.ColAlignment(9) = 3

item.ColWidth(0) = 500
item.ColWidth(1) = 1100
item.ColWidth(2) = 4600
item.ColWidth(3) = 500
item.ColWidth(4) = 1000
item.ColWidth(5) = 900
item.ColWidth(6) = 1200
item.ColWidth(7) = 1200
item.ColWidth(8) = 600
item.ColWidth(9) = 0
item.ColWidth(10) = 0
item.TextMatrix(0, 0) = "Item"
item.TextMatrix(0, 1) = "Código"
item.TextMatrix(0, 2) = "Descrição"
item.TextMatrix(0, 3) = "CST"
item.TextMatrix(0, 4) = "Unidade"
item.TextMatrix(0, 5) = "Quant"
item.TextMatrix(0, 6) = "Unitário"
item.TextMatrix(0, 7) = "Total"
If GlIpi Then
   item.TextMatrix(0, 8) = "IPI"
Else
   item.TextMatrix(0, 8) = "ICMS"
End If
LcTamanhoGrid = 1
End Function
Function montagrid()
Dim LcAchou As Integer
Dim x As Integer
On Error Resume Next
'==== Verifica se Foi digitados todos os campos
If Natureza.Text <> "Emprestimo" Then
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
Else
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
    Txt(5).Text = 0
End If
LcItem = LcItem + 1
ReDim Preserve LcMat(LcTam)
LcMat(LcTam).item = Right("000" & LcItem, 3)
LcMat(LcTam).CodPro = Txt(1).Text
LcMat(LcTam).produto = Txt(2).Text
LcMat(LcTam).Qut = CDbl(Txt(3).Text)
LcMat(LcTam).Und = Unidade.Text
LcMat(LcTam).Com = Txt(4).Text
LcMat(LcTam).VUnit = CDbl(Txt(5).Text)
LcMat(LcTam).Vtotal = CCur(Txt(6).Text)
LcMat(LcTam).Venda1 = CCur(Custo.Text)
LcMat(LcTam).cst = cst.Text
LcMat(LcTam).icms = icms.Text
LcMat(LcTam).controle = CLng(controle.Text)
LcMat(LcTam).ipi = CCur(ipi.Text)
LcMat(LcTam).UnitarioAntigo = CDbl(Txt(5).Text)
If Len(Trim(ComissaoProduto.Text)) <> 0 Then
   LcMat(LcTam).comissao = (CDbl(ComissaoProduto.Text) / 100) * CDbl(Txt(6).Text)
Else
   If Len(Trim(ComissaoFabrica.Text)) <> 0 Then
      LcMat(LcTam).comissao = (CDbl(ComissaoFabrica.Text) / 100) * CDbl(Txt(6).Text)
   Else
      LcMat(LcTam).comissao = (CDbl(comisVenda.Text) / 100) * CDbl(Txt(6).Text)
   End If
End If
LcTam = LcTam + 1
EscreveGrid

For x = 1 To 6
   Txt(x).Text = ""
Next
Txt(3).Text = " "
Txt(5).Text = " "
Txt(6).Text = " "
Custo.Text = "0"
icms.Text = "0"
ipi.Text = "0"
cst.Text = "0"
minimo.Text = "0"
ComissaoFabrica.Text = ""
ComissaoProduto.Text = ""
unitreal.Text = ""
controle.Text = ""
If GlComissaoVelha > 0 Then
   comisVenda.Text = GlComissaoVelha
End If
GlComissaoVelha = 0
Txt(1).SetFocus
End Function
Function EscreveGrid()
Dim b As Integer
Dim x As Integer
b = 1
item.Rows = 1
For x = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       item.Rows = b + 1
       item.TextMatrix(b, 0) = LcMat(a).item
       item.TextMatrix(b, 1) = LcMat(a).CodPro
       
       item.TextMatrix(b, 2) = LcMat(a).produto
       item.TextMatrix(b, 3) = LcMat(a).cst
       item.TextMatrix(b, 4) = LcMat(a).Und & " C/" & LcMat(a).Com
       item.TextMatrix(b, 5) = LcMat(a).Qut
       item.TextMatrix(b, 6) = LcMat(a).VUnit
       item.TextMatrix(b, 7) = Format(LcMat(a).Vtotal, "Currency")
       If GlIpi Then
           item.TextMatrix(b, 8) = LcMat(a).ipi
       Else
           item.TextMatrix(b, 8) = LcMat(a).icms
       End If
       b = b + 1
    End If
Next
CalculaIcms

If Not GlDetalhaDesconto Then
   VerificaDesconto
Else
  RateiaDescontoPerc
End If

'VerificaAcrecimo
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True
End Function
Function CalculaIcms()
Dim LcBaseCalculo, LcIcms, LcPRodutos, LcNota As Currency
Dim LcItemCst As String, LcComp As String
Dim LcIpi As Currency
Dim a As Integer
Dim LcQuantItemSubs As Integer
'LcItem = 0
LcQuantItemSubs = 0
LcPRodutos = 0
LcNota = 0
LcIpi = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      If LcMat(a).icms > 0 Then
         LcBaseCalculo = LcBaseCalculo + LcMat(a).Vtotal
         LcIcms = LcIcms + ((LcMat(a).icms / 100) * LcMat(a).Vtotal)
      Else
         LcQuantItemSubs = LcQuantItemSubs + 1
         If LcQuantItemSubs > 1 Then
            LcItemCst = LcItemCst & ", "
         End If
         LcItemCst = LcItemCst & Right("000" & CStr(LcMat(a).item), 3)
      End If
      If LcMat(a).ipi > 0 Then
         LcIpi = LcIpi + ((LcMat(a).ipi / 100) * LcMat(a).Vtotal)
      End If
      LcPRodutos = LcPRodutos + LcMat(a).Vtotal
      LcNota = LcNota + LcMat(a).Vtotal
   End If
Next

If LcQuantItemSubs > 0 Then
   If LcQuantItemSubs > 1 Then
      LcComp = "Itens " & LcItemCst & " ICMS cobrado por subst. Tributária."
   Else
      LcComp = "Item " & LcItemCst & " ICMS cobrado por subst. Tributária."
   End If
   'Txt(14).Text = LcComp
End If

'Txt(13).Text = Format(LcBaseCalculo, "Currency")
'Txt(11).Text = Format(LcIcms, "Currency")
 If GlIpi Then
    Txt(11).Text = AcertaNumero(CStr(LcIpi), 2)
    LcNota = LcNota + LcIpi
 End If
 Txt(15).Text = AcertaNumero(CStr(LcPRodutos), 2)
 Txt(16).Text = AcertaNumero(CStr(LcNota), 2)
 If Len(Txt(15).Text) > 0 Then
    If Txt(15).Text = 0 Then
       Txt(13).Text = ""
       Desconto.Text = ""
       CmdExcluir.Enabled = False
       Command3.Enabled = False
       CmdSalvar.Enabled = False
    End If
 Else
    Txt(13).Text = ""
    Desconto.Text = ""
    CmdExcluir.Visible = False
    Command3.Visible = False
    CmdSalvar.Visible = True
 End If
 
End Function
Function VerificaDesconto()

Dim LcCaracter, LcPalavra, LcDesconto As String
Dim Lct As Long
Dim a As Integer
'If Item.Rows = 1 Then Exit Function
LcValorDes = 0
If Len(Txt(15).Text) = 0 Then
   'MsgBox "Não Existe nenhum item para dar Desconto...", 64, "Aviso"
   Txt(13).Text = ""
   Desconto.Text = ""
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Function
End If
If (Txt(15).Text) = "0" Then
   'MsgBox "Não Existe nenhum item para dar Desconto...", 64, "Aviso"
   Txt(13).Text = ""
   Desconto.Text = ""
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Function
End If
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
RecalculaIpi
       
    
End Function
Function RateiaDescontoPerc()
Dim LcPrimeiro As Integer
Dim a As Integer
If LcCalculadoDesconto Then Exit Function

LcPrimeiro = True
If Len(Txt(13).Text) = 0 Then Exit Function
'If Txt(13).Text = "0" Then Exit Function
LcCalculadoDesconto = True
LcCalculadoAcrecimo = True
Lct = Len(Txt(13).Text)
For a = 1 To Lct
    LcCaracter = Mid(Txt(13).Text, a, 1)
    If LcCaracter = "+" Then
       Call RecalculaDesconto(CLng(LcDesconto), LcPrimeiro)
       LcPrimeiro = False
       LcDesconto = ""
    Else
       LcDesconto = LcDesconto & LcCaracter
    End If
Next
If Len(LcDesconto) > 0 Then Call RecalculaDesconto(CLng(LcDesconto), LcPrimeiro)
RecalculaIpi
RateiaAcrescimo
Desconto.Text = ""
LcCalculadoDesconto = False
LcCalculadoAcrecimo = False
End Function
Function RecalculaDesconto(LcPercentual As Long, LcPrimeiro As Integer)
Dim LcTotalP As Currency
Dim a As Integer

For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       If LcPrimeiro Then
          LcTotalP = LcMat(a).UnitarioAntigo
       Else
          LcTotalP = LcMat(a).VUnit
       End If

       
       LcMat(a).VUnit = LcTotalP - ((LcPercentual / 100) * LcTotalP)
       LcMat(a).Vtotal = LcMat(a).VUnit * LcMat(a).Qut
       LcMat(a).comissao = LcMat(a).comissao - ((LcPercentual / 100) * LcMat(a).comissao)
    End If
Next
EscreveGrid
End Function
Function RateiaAcrescimo()
Dim LcPrimeiro As Integer
Dim LcTotalP As Currency
Dim a As Integer
If LcCalculadoAcrecimo Then Exit Function
If Len(Txt(19).Text) = 0 Then Exit Function
LcCalculadoAcrecimo = True
LcCalculadoDesconto = True
LcPercentual = CLng(Txt(19).Text)
For a = 0 To LcTam - 1
    LcTotalP = LcMat(a).VUnit
    LcMat(a).VUnit = LcTotalP + ((LcPercentual / 100) * LcTotalP)
    LcMat(a).Vtotal = LcMat(a).VUnit * LcMat(a).Qut
    LcMat(a).comissao = LcMat(a).comissao + ((LcPercentual / 100) * LcMat(a).comissao)
Next
RecalculaIpi
EscreveGrid
Acrecimo.Text = ""
LcCalculadoAcrecimo = False
LcCalculadoDesconto = False
End Function

Function VerificaAcrecimo()
On Error GoTo errAcrecimo
Dim LcCaracter, LcPalavra, LcDesconto As String
Dim Lct As Long
Dim a As Integer
Dim LcValorIpi As Currency
LcValorAcre = 0
LcVerificaAcrecimo = True
'If Item.Rows = 1 Then Exit Function

If Len(Txt(15).Text) = 0 Then
   'MsgBox "Não Existe nenhum item para dar Desconto...", 64, "Aviso"
   Txt(19).Text = ""
   Acrecimo.Text = ""
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Function
End If
If (Txt(15).Text) = "0" Then
   'MsgBox "Não Existe nenhum item para dar Desconto...", 64, "Aviso"
   Txt(19).Text = ""
   Acrecimo.Text = ""
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Function
End If
LcTotal = CCur(Txt(15).Text)
Lct = Len(Txt(19).Text)
Acrecimo.Text = ""
For a = 1 To Lct
    LcCaracter = Mid(Txt(19).Text, a, 1)
    If LcCaracter = "+" Then
       CalculaAcrecimo
       LCAcrecimo = ""
    Else
       LCAcrecimo = LCAcrecimo & LcCaracter
    End If
Next
If Len(Txt(11).Text) = 0 Then
    LcValorIpi = 0
Else
   LcValorIpi = CCur(Txt(11).Text)
End If
LcDesconto = CCur(Desconto.Text)

If LCAcrecimo > 0 Then CalculaAcrecimo
   
'If Len(LCAcrecimo) > 0 Then CalculaAcrecimo (CCur(CalculaAcrecimo(CDbl(LCAcrecimo))))
Txt(16).Text = AcertaNumero(CCur(LcTotal) - LcDesconto + LcValorIpi, 2)
'Acrecimo.Text = AcertaNumero(CCur(LCAcrecimo / 100) * LcTotal)
LcVerificaAcrecimo = False
'RecalculaIpi
Exit Function
errAcrecimo:
'MsgBox Err.Description & Err.Number
Resume Next
End Function
Function TotalDesc(LcPerc, LcValor As Currency) As Currency
Dim LcTo As Currency
LcTo = LcValor - ((LcPerc / 100) * LcValor)
TotalDesc = LcTo
End Function
Function AcertaIpi(LcV As Currency) As Currency
Dim a As Integer
Dim LCLEtra As String
Dim LcTotalIpi, LcVIpi As Currency
If Len(Txt(13).Text) = 0 Then Exit Function
LcTotalIpi = LcV
For a = 1 To Len(Txt(13).Text)
    LcCaracter = Mid(Txt(13).Text, a, 1)
    If Not IsNumeric(LcCaracter) Then
       LcTotalIpi = TotalDesc(CCur(LcDesconto), CCur(LcTotalIpi))
       LcDesconto = ""
    Else
       LcDesconto = LcDesconto & LcCaracter
    End If
Next
LcTotalIpi = TotalDesc(CCur(LcDesconto), CCur(LcTotalIpi))
AcertaIpi = LcTotalIpi
End Function
Function RecalculaIpi()
On Error Resume Next
Dim LcIpit As Double
Dim LcPer As Double
Dim a As Integer
If GlDetalhaDesconto Then Exit Function

For a = 0 To LcTam - 1
    LcPer = (LcMat(a).ipi / 100) * LcMat(a).Vtotal
    LcIpit = LcIpit + LcPer
Next
LcValorIpi = LcIpit
If LcValorIpi = 0 Then Exit Function
Txt(11).Text = AcertaNumero(CStr(LcValorIpi), 2)
Txt(16).Text = AcertaNumero(CStr(CCur(Txt(15).Text) - CCur(Desconto.Text) + CCur(Acrecimo.Text) + CCur(Txt(11).Text)), 2)
End Function
Function CalculaAcrecimo()
On Error Resume Next

If Len(Txt(15).Text) = 0 Then
   MsgBox "Não foi Digitado Nenhum item para dar Acrecimo...", 64, "Aviso"
   Txt(13).Text = ""
   Acrecimo.Text = " 0"
   Txt(1).SetFocus
   Exit Function
End If
If Txt(15).Text = "0" Then
   MsgBox "Não foi Digitado Nenhum item para dar Acrecimo...", 64, "Aviso"
   Txt(13).Text = ""
   Acrecimo.Text = "0"
   Txt(1).SetFocus
   Exit Function
End If

LcTotal = CCur(Txt(15).Text) - CCur(Desconto.Text)
LcValorAcre = (CCur(Txt(19).Text) / 100) * LcTotal
Txt(16).Text = CCur(AcertaNumero(CStr((CCur(LcTotal) + AcertaNumero(CStr(LcValorAcre), 2))), 2))
Txt(16).Text = CCur(Txt(16).Text) + CCur(Txt(11).Text)
Acrecimo.Text = AcertaNumero(CStr(LcValorAcre), 2)




End Function
Function CalculaDesconto(LcDesconto As Double)
'On Error Resume Next
Dim LcValorIpi As Currency
If Len(Txt(15).Text) = 0 Then
   MsgBox "Não foi Digitado Nenhum item para dar Desconto...", 64, "Aviso"
   Txt(13).Text = ""
   Desconto.Text = "0"
   Txt(1).SetFocus
   Exit Function
End If
If Txt(15).Text = "0" Then
   MsgBox "Não foi Digitado Nenhum item para dar Desconto...", 64, "Aviso"
   Txt(13).Text = ""
   Desconto.Text = "0"
   Txt(1).SetFocus
   Exit Function
End If
If Len(Acrecimo.Text) > 0 Then
   LcValorAcrecimo = CCur(Acrecimo.Text)
Else
   LcValorAcrecimo = 0
End If
If Len(Txt(11).Text) = 0 Then
   LcValorIpi = 0
Else
   LcValorIpi = CCur(Txt(11).Text)
End If
LcValorDes = LcValorDes + (LcDesconto / 100) * LcTotal
'txt(16).Text = CCur(AcertaNumero(CStr((LcTotal - LcValorDes + LcValorAcrecimo + LcValorIpi)), 2))
Desconto.Text = AcertaNumero(CCur((LcValorDes)), 2)
LcTotal = CDbl(Txt(15).Text) - LcValorDes
Txt(16).Text = AcertaNumero(LcTotal + LcValorAcrecimo + LcValorIpi, 2)
End Function
Function RemontaIndice()
On Error Resume Next
LcItem = 0
Dim a As Integer
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      LcItem = LcItem + 1
      LcMat(a).item = Right("000" & LcItem, 3)
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
   Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UNIMED & "'")  ', dbOpenDynaset)

   'RsUnidade.FindFirst LcCriterio
   If Not RsUnidade.EOF Then
      LcUnidade = RsUnidade!Simbolo
   End If
   Txt(1).Text = RsProduto!cod
   Txt(2).Text = RsProduto!Nome
   Unidade.Text = LcUnidade
   Txt(4).Text = RsProduto!QTDUNIMED
   Txt(5).Text = RsProduto!Ptab
   ipi.Text = RsProduto!ipi
   unitreal.Text = RsProduto!Ptab
   cst.Text = RsProduto!cst
   LcPrecoVelho = RsProduto!Ptab
   If Not IsNull(RsProduto!ComissaoFornecedor) Then
    If RsProduto!ComissaoFornecedor <> 0 Then
      ComissaoProduto.Text = RsProduto!ComissaoFornecedor
    End If
   End If
   minimo.Text = RsProduto!MPVENDA
   
   If Val(cst.Text) = 6 Or Val(cst.Text) = 16 Or Val(cst.Text) = 26 Then
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
On Error Resume Next
Dim LcAchou As Integer
Dim RsVendedor As Recordset
AbreBase
If Len(comissao.Text) = 0 Then comissao.Text = 0
Txt(10).Text = Right("00000" & Txt(10).Text, 5)
Set RsVendedor = Dbbase.OpenRecordset("select * From alid200 where codigo='" & Txt(10).Text & "'") ', dbOpenDynaset)
If Not RsVendedor.EOF Then
  If err.Number > 0 Then Exit Function
   Txt(7).Text = RsVendedor!Nome
   If CLng(comisVenda.Text) <> 1 Then
      comisVenda.Text = RsVendedor!comissao
   End If
   LcAchou = True
Else
   LcAchou = False
   Txt(10).Text = ""
   comisVenda.Text = 0
   Txt(7).Text = ""
   Txt(10).SetFocus
   MsgBox "Vendedor Não Encontrado...", 64, "Aviso"
End If
If LcAchou Then SendKeys "{TAB}"
RsVendedor.Close
Set RsVendedor = Nothing
If GlVariasComissao Then Exibecomissao.Show , Me

End Function
Function buscafornecedor(LcBusca As Integer)
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
On Error Resume Next
If KeyCode <> 116 Then Teclas (KeyCode)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub



Function Imprimeorcamento()
On Error Resume Next
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
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!fone
End If


'Abertura do relatório de vendas
    
    
    CryRelatorio.DataFiles(0) = GLBase
    CryRelatorio.ReportFileName = App.Path & "\pedidovendas.rpt"
    LcFormula = "{orcamento.doc}='" & UCase(Txt(0).Text) & "'"
    CryRelatorio.CopiesToPrinter = 1

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Orçamento"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
'CryRelatorio.Formulas(3) = "Versiculo='" & LcVer & "'"
'CryRelatorio.Formulas(4) = "Versiculo1='" & LcVer1 & "'"
'CryRelatorio.Formulas(5) = "titulo='Produtos'"
 
LcTipoSaida = 1

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

Function ImprimeNota()
Dim a As Integer
On Error GoTo Errimpr
If GLPadraoWindows Then
   Imprimeorcamento
   Exit Function
End If

Dim item, Descricao, cst, icms, Unidade As String
Dim quant, Unitario, total As Long
Dim LcImpressoes As Integer
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(18).Text & "'")
Set RsEmpresa = Dbbase.OpenRecordset("select * from empresa")
LcEspaco = ""

LcMargem = ""
For a = 1 To GlMargem
    LcMargem = LcMargem & " "
Next
LcSalto = Val(GLSaltoLinhaNota)

'Salta linhas no inicio da nota
'If Gl40colunas Then Print #FnunNota, Chr(15)

cabecalhonota
For az = 0 To (LcTam - 1)
    Call imprimeitem(CLng(az))
Next
FechaImpressao (LcImpressoes)
ImprimeSpool
Me.Caption = "Orçamento e Vendas"
Exit Function
Errimpr:
If err = 76 Then
   MsgBox "A Porta de Impressão " & GlPortaOrcamento & " Não Foi encontrada," & Chr(13) & "Verifique se a impressora está em linha e o cabo Conectado, ou a conexão da Rede.", 64, "Aviso"
   Exit Function
Else
   MsgBox err.Description & err.Number
   Resume 0
End If

End Function

Function ImprimeSpool()
Dim RsImpressoras As Recordset, RsSpool As Recordset
On Error Resume Next
Dim LcNomeMaquina As String * 255
Dim Lct As Long
Dim LcSpool As String
FnunNota = FreeFile
FnunBoleto = FreeFile + 1
If IsNull(GlPortaOrcamento) Then GlPortaOrcamento = "LPT1"
'If IsNull(LcBoleta) Then LcBoleta = "LPT2"
AbreBase

Set RsImpressoras = Dbbase.OpenRecordset("Select * from impressoras")
Set RsSpool = Dbbase.OpenRecordset("Select * from LogImpressao")


LcCri = "Impressora='" & GlPortaOrcamento & "'"
RsImpressoras.FindFirst LcCri
RsSpool.AddNew
RsSpool("Impressora") = GlPortaOrcamento
RsSpool("endereco") = RsImpressoras("EnderecoLocal")
LcPermissaoImpressao = RsSpool("Sequencia")
RsSpool.Update
'RsSpool.Close

'Set RsSpool = Dbbase.OpenRecordset("LogImpressao", dbOpenTable, dbSeeChanges, dbOptimistic)
'RsSpool.Index = "Sequencia"
'RsSpool.MoveFirst
'Do While RsSpool!Sequencia <> LcPermissaoImpressao
'   DoEvents
'   RsSpool.MoveFirst
'Loop

LcSpool = RsImpressoras!EnderecoLocal & ""
If Len(LcSpool) = 0 Then
   LcSpool = "Lpt1"
End If
LcImpressoes = 0
Open LcSpool For Output Access Write As #FnunNota  'Abre Porta Nf
For a = 0 To LcTamanhoPedido
    Print #FnunNota, MtPedido(a)
Next
RsSpool.Delete
RsSpool.Close
LcTamanhoPedido = 0
Close #FnunNota
End Function
Function FechaImpressao(Linhas As Integer)
Dim RsOrc As Recordset
AbreBase
Set RsOrc = Dbbase.OpenRecordset("orcamento", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcCri = "doc='" & Txt(0).Text & "'"
RsOrc.FindFirst LcCri
Dim lcLinhasSalto As Integer
For q = 1 To 45
    LcEspaco = LcEspaco & " "
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)

For a = 0 To 79 - CLng(GlMargem)
    LcSepara = LcSepara + "="
Next



LcLinha = LcEspaco & "Total Produtos: " & Right("            " & AcertaNumero(CStr(RsOrc!TotalProduto), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
If Not GlImprimeDetalhaDesconto Then
   LcLinha = LcEspaco & "Desconto      : " & Right("            " & AcertaNumero(CStr(0), 2), 10)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
Else
   LcLinha = LcEspaco & "Desconto      : " & Right("            " & AcertaNumero(CStr(RsOrc!TotalDesconto), 2), 10)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
End If
LcLinha = LcEspaco & "Acrecimo      : " & Right("           " & AcertaNumero(CStr(RsOrc!Acrecimo), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
If GlIpi Then
   If Len(Txt(11).Text) = 0 Then Txt(11).Text = 0
   LcLinha = LcEspaco & "IPI           : " & Right("            " & AcertaNumero(CStr(Txt(11).Text), 2), 10)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
End If
LcLinha = LcEspaco & "Total Pagar   : " & Right("           " & AcertaNumero(CStr(RsOrc!TotalGeral), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & "Condicoes de Pag.: " & DadosOrcamento.TipoPag & "     Forma de Pag.:" & DadosOrcamento.TipoMonetario & Chr(13) & ""
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
If GlImprimeDetalhaDesconto Then
   If Len(Txt(13).Text) > 0 Then
      LcLinha = "Descricao do Desconto: " & Txt(13).Text & " << Em Percentual >>"
      LcTamanhoPedido = LcTamanhoPedido + 1
      MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
      LcTamanhoPedido = LcTamanhoPedido + 1
      MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
      LcTamanhoPedido = LcTamanhoPedido + 1
      MtPedido(LcTamanhoPedido) = Chr(13)
    End If
End If

If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
  Select Case DadosOrcamento.Quantidade.Text
       Case Is = "1"
           LcLinha = "Vencimento:" & DadosOrcamento.Vencimento(0).Text
           LcLinha = LcLinha & "  Valor : " & Right("       " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
       Case Is = 2
           LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
       Case Is = 3
           LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
        Case Is = 4
           LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "4 Vencimento:" & DadosOrcamento.Vencimento(3).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
        Case Is = 5
           LcLinha = "1 Vencimento:" & DadosOrcamento.Vencimento(0).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "2 Vencimento:" & DadosOrcamento.Vencimento(1).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "3 Vencimento:" & DadosOrcamento.Vencimento(2).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "4 Vencimento:" & DadosOrcamento.Vencimento(3).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
           LcLinha = "5 Vencimento:" & DadosOrcamento.Vencimento(4).Text
           LcLinha = LcLinha & "  Valor : " & Right("         " & AcertaNumero(CStr(DadosOrcamento.valor.Text), 2), 7)
           LcTamanhoPedido = LcTamanhoPedido + 1
           MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

           
  End Select
End If
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
   MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
   
End If

LcCri = "doc='" & Txt(0).Text & "'"
RsOrc.FindFirst LcCri
If Not RsOrc.NoMatch Then
  If Len(RsOrc!Transp) > 0 Then
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = Chr(13)
    LcLinha = "Transportadora: " & RsOrc!Transp & "    Fone: " & RsOrc!FoneTransp & ""
    LcTamanhoPedido = LcTamanhoPedido + 1
    MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13) & Chr(13)
      
  End If
  'If Len(RsOrc!FoneTransp) > 0 Then
  ' LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = Chr(13)
  ' LcLinha = "Fone:" & RsOrc!FoneTransp
  ' LcTamanhoPedido = LcTamanhoPedido + 1
  ' MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
  'End If

  If Len(RsOrc!obs) > 0 Then
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)

   LcLinha = "Dados Complementares:"
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13) & Chr(13)
   LcLinha = RsOrc!obs
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)
  End If
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
lctammsg = Len(GlMsg)
lcspa = ""
For a = 1 To (((80 - CLng(GlMargem)) / 2) - (lctammsg / 2))
    lcspa = lcspa & " "
Next
LcSepara = ""
For a = 0 To 79 - CLng(GlMargem)
    LcSepara = LcSepara + "-"
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = lcspa & GlMsg & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcSepara & Chr(13)
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
  MtPedido(LcTamanhoPedido) = " "

End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)



End Function
Function imprimeitem(az As Long) As Integer

If Len(LcMat(az).CodPro) = 0 Then Exit Function
LcLinha = Left(LcMat(az).CodPro & "             ", 9)
For b = 1 To 2
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & Right("    " & LcMat(az).Qut, 4)
For b = 1 To 2
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & Left(LcMat(az).produto & "                            ", 38)

lctp = Len(LcLinha)
For b = 1 To 40 - CLng(lctp)
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & Right("            " & AcertaNumero(CStr(LcMat(az).VUnit), GlDecimais), 7)
For b = 1 To 2
   LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & Right("            " & AcertaNumero(CStr(LcMat(az).Vtotal), 2), 7)
If GlIpi Then
   For b = 1 To 2
      LcLinha = LcLinha & " "
   Next
   LcLinha = LcLinha & Right("            " & AcertaNumero(CStr(LcMat(az).ipi), 2), 5)
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha + Chr(13)

LcLinha = ""


End Function
Function imprimeimtemOlinto(az As Long)
Dim LcTamanhodes, LcImp As Long
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
Function cabecalhonota()
Dim RsCidade As Recordset
Dim LcSepara As String
LcSepara = ""
For a = 0 To 79 - CLng(GlMargem)
    LcSepara = LcSepara + "="
Next
AbreBase
Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!fone
End If
RsEmpresa.Close
If GlEscolheCliente Then
   LcCriterio = "cod='" & RsClientes!cidade & "'"
   RsCidade.FindFirst LcCriterio
   If Not RsCidade.NoMatch Then
      LcCidade = RsCidade!Nome
   End If
End If
Set RsEmpresa = Nothing
'=== Imprime Cabecalho Nota
'Print #FnunNota,

LcTamanhoPedido = -1
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcEmpresa + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcSepara + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcEndereco + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & "Fone:  " + LcFone + Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)

If Natureza.Text = "Orçamento" Then
   LcLinha = "Orcamento N: " & Txt(0).Text
Else
   LcLinha = Left(Natureza.Text & "      ", 10) & ": " & Txt(0).Text
End If
LcLinha = LcLinha & "                         Data de Emissao: " & Txt(12).Text
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
If Not GlEsclheVendedor Then
   LcLinha = "Vendedor  : " & LcEmpresa
Else
   LcLinha = "Vendedor  : " & Txt(7).Text
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

If GlEscolheCliente Then
   LcLinha = "Cliente   : " & Txt(17).Text
Else
  LcLinha = "Cliente   : Consumidor Final"
End If
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

If GlEscolheCliente Then
   LcLinha = "Endereco  : " & RsClientes!End
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinha = "Bairro    : " & Left(RsClientes!Bairro, 40) & "  Cidade:" & LcCidade & ""
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinha = "UF        : " & RsClientes!estado & "    CEP:" & RsClientes!Cep & ""
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinha = ""
   If Len(RsClientes!fone1) > 0 Then LcLinha = "Fone      : " & RsClientes!fone1
   If Len(RsClientes!Fax) > 0 Then
     If Len(LcLinha) > 0 Then
        LcLinha = LcLinha & "  Fax:" & RsClientes!Fax & ""
     Else
        LcLinha = "Fax: " & RsClientes!Fax & ""
     End If
   End If
   LcLinha = "C.G.C.    : "
   LcLinha = LcLinha & RsClientes!cgc
   LcLinha = LcLinha & "    Insc. Estadual : "
   LcLinha = LcLinha & RsClientes!INSCEST & ""

   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinha = "Fone      : "
   LcLinha = LcLinha & RsClientes!fone1 & ""
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
End If
'====Monta e imprime Titulos dos Itens
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = Chr(13)
LcLinha = "Codigo"
For a = 1 To 5
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & " Qut"
For a = 1 To 2
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "Produto"
For a = 1 To 33
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "V.Unt"
For a = 1 To 4
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "Total"
For a = 1 To 4
    LcLinha = LcLinha & " "
Next
LcLinha = LcLinha & "IPI"
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
lcmenos = ""
For a = 1 To 80 - CLng(GlMargem)
   lcmenos = lcmenos & "-"
Next
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & lcmenos & Chr(13)



End Function
Private Sub CmdSalvar_Click()
On Error GoTo ErrSalvar
Dim RsProduto As Recordset, RsOrca As Recordset, RsItens As Recordset
Dim LcCodigo As String
On Error Resume Next
CalculaNumeroNota
AbreBase
Set RsProduto = Dbbase.OpenRecordset("Produto", dbOpenTable, dbSeeChanges, dbOptimistic)
Set RsOrca = Dbbase.OpenRecordset("Orcamento", dbOpenTable, dbSeeChanges, dbOptimistic)
RsOrca.Index = "doc"
'RsItens.Index = "Pesquisa"
'RsProduto.Index = "Codigo"
LcCriterio = "doc='" & Txt(0).Text & "'"
RsOrca.Seek "=", Txt(0).Text

If Not RsOrca.NoMatch Then
   RsOrca.Edit
Else
   RsOrca.AddNew
End If
RsOrca("DOC") = Txt(0).Text
RsOrca("DTEMIS") = CDate(Txt(12).Text)
RsOrca("NATUREZA") = Left(Natureza.Text, 2)
RsOrca("CLiente") = Txt(18).Text
'+Rsorcamento("TRANSP") = DadosOrcamento.txt(0).Text
RsOrca("TIPOTRANS") = Mid(DadosOrcamento.Tipo.Text, 1, 1)
'Rsorcamento("FoneTransp") = DadosOrcamento.txt(10).Text
RsOrca("PLACATRANS") = DadosOrcamento.Placa.Text
RsOrca("UFTRANS") = DadosOrcamento.Txt(1).Text
RsOrca("CGCCPFTRAN") = DadosOrcamento.Txt(2).Text
RsOrca("ENDTRANS") = DadosOrcamento.Txt(3).Text
RsOrca("MUNICTRANS") = DadosOrcamento.Txt(4).Text
RsOrca("UFMUNIC") = DadosOrcamento.Txt(5).Text
'RsOrcamento("INSCEST") = DadosOrcamento.txt(6).Text
RsOrca("OBS02") = DadosOrcamento.Txt(7).Text
RsOrca("OBS03") = DadosOrcamento.Txt(8).Text
RsOrca("OBS04") = DadosOrcamento.Txt(9).Text
RsOrca("OrcVenda") = Natureza.Text
RsOrca("Vendedor") = Txt(10).Text
RsOrca("cidade") = DadosOrcamento.Txt(6).Text
RsOrca("cep") = DadosOrcamento.Txt(12).Text
RsOrca("CondPag") = DadosOrcamento.TipoPag.Text
If Len(Txt(13).Text) > 0 Then RsOrca!Desconto = Txt(13).Text
If Len(Desconto.Text) > 0 Then RsOrca!TotalDesconto = CCur(Desconto.Text)
If Len(Txt(16).Text) > 0 Then RsOrca!TotalGeral = CCur(Txt(16).Text)
If Len(Txt(15).Text) > 0 Then RsOrca!TotalProduto = CCur(Txt(15).Text)
RsOrca!formapag = DadosOrcamento.TipoMonetario.Text
RsOrca!Dias = DadosOrcamento.Txt(11).Text
RsOrca!Acrecimo = CCur(Acrecimo.Text)
RsOrca!DetalhaDesconto = Txt(13).Text
RsOrca!DetahaAcrecimo = Txt(19).Text



If DadosOrcamento.Vencimento(0) <> "  /  /  " Then
   RsOrca!Vencimento1 = DadosOrcamento.Vencimento(0).Text
   RsOrca!vencimento2 = DadosOrcamento.Vencimento(1).Text
   RsOrca!vencimento3 = DadosOrcamento.Vencimento(2).Text
   RsOrca!vencimento4 = DadosOrcamento.Vencimento(3).Text
   RsOrca!vencimento5 = DadosOrcamento.Vencimento(4).Text
End If
RsOrca.Update
AbreBase
LcCriterio = "select * from DadosOrcamento where doc='" & Txt(0).Text & "'"
Set RsItens = Dbbase.OpenRecordset(LcCriterio)

Do Until RsItens.EOF
    RsItens.Delete
    RsItens.MoveNext
  '  MsgBox RsItens!doc
Loop
RsItens.Close
AbreBase
Set RsItens = Dbbase.OpenRecordset("DadosOrcamento", dbOpenTable, dbSeeChanges, dbOptimistic)
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       LcCodigo = LcMat(a).CodPro
       'RsProduto.Seek "=", LcCodigo
       'If Not RsProduto.NoMatch Then
       '   RsProduto.Edit
       '   RsProduto("Estoque") = RsProduto!ESTOQUE - LcMat(a).Qut
                 
       '   RsProduto.Update
       'End If
       'RsItens.Seek "=", txt(0).Text, LcCodigo
       'If Not RsItens.NoMatch Then
       '   RsItens.Edit
       'Else
          RsItens.AddNew
       'End If
       RsItens("Doc") = Txt(0).Text
       RsItens("CodigoProduto") = LcCodigo
       RsItens("Descricao") = LcMat(a).produto
       RsItens("Quant") = LcMat(a).Qut
       RsItens("Unit") = LcMat(a).VUnit
       RsItens("Total") = LcMat(a).Vtotal
       RsItens("unid") = LcMat(a).Und
       RsItens("com") = LcMat(a).Com
       RsItens("item") = LcMat(a).item
       RsItens("ipi") = LcMat(a).ipi
       RsItens("valorUnitarioReal") = LcMat(a).UnitarioAntigo
       RsItens("comissao") = LcMat(a).comissao
       RsItens.Update
    End If
Next
ConfirmaOrcamento.Show , Me
'Resposta = MsgBox("Imprime o Orçamento ?", vbInformation + vbYesNo, "Impressão")
'If Resposta = 6 Then ImprimeNota
'MsgBox "Os Dados Foram Salvos Com Sucesso...", 48, "Aviso"
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
Txt(11).Text = ""
ipi.Text = "0"
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
limpanota
'CalculaNumeroNota
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Txt(0).SetFocus
Exit Sub
ErrSalvar:
MsgBox err.Description & err.Number
'Stop
Resume Next
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
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
DadosOrcamento.Show , Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode <> 116 Then Teclas (KeyCode)
  Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Command4.Caption = "Pes&quisa F7" Then
   LcSql1 = "Select * from orcamento"
   AbreBase
   Set Rsorcamento = Dbbase.OpenRecordset(LcSql1)
   LcCriterio = "Doc='" & Txt(0).Text & "'"
   Rsorcamento.FindFirst LcCriterio
   If Not Rsorcamento.NoMatch Then
      Rsorcamento.Delete
   End If
   Rsorcamento.Close
   FrmPesquisaNota.Show , Me
   Command4.Caption = "&Incluir F7"
   LcPesquisa = True
Else
   Command4.Caption = "Pes&quisa F7"
   limpanota
   LcPesquisa = False
End If

End Sub

Private Sub desconto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 117 Then Txt(13).SetFocus
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 114 Then SendKeys "%+{P}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 118 Then SendKeys "%+{Q}"
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me

End Sub

Private Sub desconto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub desconto_LostFocus()
On Error Resume Next
Dim LcPrimeiro As Integer
Dim LcPercentual As Long
LcPrimeiro = True
If Desconto.Text = "0" Then Exit Sub
If Len(Desconto.Text) = 0 Then Exit Sub
If Len(Txt(16).Text) = 0 Then
   'MsgBox "Não Foi Cadastrado Nenhum Item para dar Desconto...", 64, "Aviso"
   
   If Len(Txt(10).Text) = 0 Then Txt(10).SetFocus
   Exit Sub
End If
If Not GlDetalhaDesconto Then
   Txt(16).Text = CCur(CCur(Txt(15).Text) - CCur(Desconto.Text))
Else
   LcPercentual = (CCur(Desconto.Text) / CCur(Txt(16).Text)) * 100
   Call RecalculaDesconto(CLng(LcPercentual), LcPrimeiro)
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If Not GlCarregado Then
   Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
   GlCarregado = True
End If
Txt(10).Enabled = GlEsclheVendedor
Txt(7).Enabled = GlEsclheVendedor
Txt(17).Enabled = GlEscolheCliente
Txt(18).Enabled = GlEscolheCliente
If GlVariasComissao Then
    Label11.Visible = True
Else
   Label11.Visible = False
End If

   
End Sub
Function InprimeNotaOLinto()

On Error GoTo Errimpr
LcMargem = ""

Dim item, Descricao, cst, icms, Unidade As String
Dim quant, Unitario, total, LcITemsImpressos As Long
Dim LcImpressoes, LcSegunda As Integer
For a = 0 To 500
   MtPedido(a) = ""
Next
'CalculaNumeroNota
AbreBase
LcSegunda = False
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(18).Text & "'")
Set RsEmpresa = Dbbase.OpenRecordset("select * from empresa")
Set RsI = Dbbase.OpenRecordset("select * from DadosOrcamento where doc='" & Txt(0).Text & "' order by descricao")
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1

If IsNull(GlPortaOrcamento) Then GlPortaOrcamento = "LPT1"
'If IsNull(LcBoleta) Then LcBoleta = "LPT2"

LcImpressoes = 0
'Open GlPortaOrcamento For Output Access Write As #FnunNota 'Abre Porta Nf
'===== Salta linhas no inicio da nota
'    If Gl40colunas Then Print #FnunNota, Chr(15)

cabecalhoOlinto

'=== Determina a Quantidade de Itens disponiveis

For az = 0 To (LcTam - 1)
    If Len(LcMat(az).CodPro) > 0 Then
       LcTotalItem = LcTotalItem + 1
    End If
Next
LcITemsImpressos = 0
Do Until RsI.EOF
     Call imprimeimtemOlinto(CLng(az))
     LcITemsImpressos = LcITemsImpressos + 1
     'If LcLinhaAtual >= 53 Then
     '   If LcItensImpressos < LcTotalitem Then
        '   FechaImpressaoolinto
       '    cabecalhoOlinto.
     '   End If
     'End If
     'MsgBox LcLinhaAtual
     If LcITemsImpressos = 25 Then
        LcSegunda = True
        FechaImpressaoolinto
        cabecalhoOlinto
        LcITemsImpressos = 0
     End If
     RsI.MoveNext
Loop
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)
If LcSegunda Then
   LcLinhaAtual = LcLinhaAtual + 2
Else
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(0).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(DadosOrcamento.valor, 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(1).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(1).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(DadosOrcamento.valor, 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(2).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(2).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(DadosOrcamento.valor, 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(3).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(3).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(DadosOrcamento.valor, 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
If DadosOrcamento.Vencimento(4).Text <> "  /  /  " Then
   LcLinha = "                "
   LcLinha = LcLinha & DadosOrcamento.TipoMonetario & " : " & DadosOrcamento.Vencimento(4).Text
   LcLinha = LcLinha & " Valor de " & AcertaNumero(DadosOrcamento.valor, 2)
   LcTamanhoPedido = LcTamanhoPedido + 1
   MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
   LcLinhaAtual = LcLinhaAtual + 1
End If
FechaImpressaoolinto
'Close #FnunNota
ImprimeSpool

Exit Function
Errimpr:
If err = 76 Then
   MsgBox "A Porta de Impressão " & GlPortaOrcamento & " Não Foi encontrada," & Chr(13) & "Verifique se a impressora está em linha e o cabo Conectado, ou a conexão da Rede.", 64, "Aviso"
   Exit Function
Else
   MsgBox err.Description & err.Number
   Resume Next
   
End If
End Function
Function FechaImpressaoolinto()
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
LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(Txt(15)), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(Desconto.Text), 2), 10)
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)
LcLinha = LcEsp & Right("           " & AcertaNumero(CStr(Txt(16)), 2), 10)
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
Function BuscaNota(LcNumeroOrc As String)
On Error Resume Next
Dim RsOrc As Recordset, RsItem As Recordset
Dim RsProduto As Recordset, RsCliente As Recordset
Dim RsVendedor As Recordset
Dim LcSql1, LcSql2, LcSql3, LcSql4, LcSql5 As String
LcPesquisa = True
LcSql1 = "Select * from orcamento where doc='" & LcNumeroOrc & "'"
LcSql2 = "Select * from DadosOrcamento where doc='" & LcNumeroOrc & "' order by item"
LcSql3 = "Select * from ALid001"
LcSql5 = "Select * from ALid200"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1)
Set RsItem = Dbbase.OpenRecordset(LcSql2)
Set RsCliente = Dbbase.OpenRecordset(LcSql3)
Set RsVendedor = Dbbase.OpenRecordset(LcSql5)
'==== Preenchendo a Nota

If RsOrc.EOF Then
   MsgBox "O Orçamento Nº: " & LcNumeroOrc & " Não foi encontrado..."
   Command4.Caption = "Pes&quisa F7"
   Txt(10).SetFocus
   Exit Function
End If
   
Txt(0).Text = RsOrc!doc
Txt(12).Text = RsOrc!DTEMIS
If RsOrc!Natureza = "Or" Then
   Natureza.Text = "Orçamento"
Else
   Natureza.Text = "Venda"
   
End If
'If Natureza.Text = "Orçamento" Then
'   Command3.Visible = False
'   CmdSalvar.Visible = True
'Else
'  Command3.Visible = True
'  CmdSalvar.Visible = False
'End If
Txt(10).Text = RsOrc!Vendedor & ""
LcCriterio = "Codigo='" & RsOrc!Vendedor & "'"
RsVendedor.FindFirst LcCriterio
If Not RsVendedor.NoMatch Then
   Txt(7).Text = RsVendedor!Nome
Else
  Txt(7).Text = ""
End If
Txt(18).Text = RsOrc!cliente
LcCriterio = "Codigo='" & RsOrc!cliente & "'"
RsCliente.FindFirst LcCriterio
If Not RsCliente.NoMatch Then
   Txt(17).Text = RsCliente!razaosoc
End If
Txt(15).Text = RsOrc!TotalProduto
Txt(16).Text = RsOrc!TotalGeral
Txt(13).Text = RsOrc!DetalhaDesconto
Txt(19).Text = RsOrc!DetahaAcrecimo
Status.Text = RsOrc!Status
Acrecimo.Text = RsOrc!Acrecimo
If Len(RsOrc!Desconto) > 0 Then Txt(13).Text = RsOrc!Desconto Else Txt(13).Text = ""
If Len(RsOrc!TotalDesconto) > 0 Then Desconto.Text = RsOrc!TotalDesconto Else Desconto.Text = ""
'===== Escreve dados Grid
LcItem = 0
LcTam = 0
ReDim LcMat(LcTam)
Do Until RsItem.EOF
    LcItem = LcItem + 1
    ReDim Preserve LcMat(LcTam)
    If Len(RsItem!item) > 0 Then LcMat(LcTam).item = Right("000" & RsItem!item, 3)
    LcMat(LcTam).CodPro = RsItem!codigoproduto
    LcMat(LcTam).produto = RsItem!Descricao
    LcMat(LcTam).Qut = RsItem!quant
    LcMat(LcTam).Und = RsItem!unid
    LcMat(LcTam).Com = RsItem!Com
    LcMat(LcTam).VUnit = RsItem!Unit
    LcMat(LcTam).UnitarioAntigo = RsItem!valorUnitarioReal
    LcMat(LcTam).comissao = RsItem!comissao
    LcMat(LcTam).Vtotal = RsItem!total
    LcMat(LcTam).ipi = RsItem!ipi
    LcMat(LcTam).ipi = RsItem!ipi
    'LcMat(LcTam).icms = icms.Text
     LcTam = LcTam + 1
    EscreveGrid
    RsItem.MoveNext
    LcAchou = True
Loop
 If LcAchou Then
    Txt(1).SetFocus
 Else
    Txt(10).SetFocus
    'CmdSalvar.Visible = True
    'Command3.Visible = False
 End If
 
 RsOrc.Close
 RsItem.Close
 RsCliente.Close
 RsVendedor.Close
 LcCalculadoDesconto = False
End Function
Function cabecalhoOlinto()
Dim RsCidade As Recordset
Dim LcSepara As String
Dim LcSalto As Long
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

AbreBase
Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
'Me.Caption = "Aguarde, Gerando o Relatório..."
LcCriterio = "cod='" & RsClientes!cidade & "'"
RsCidade.FindFirst LcCriterio
If Not RsCidade.NoMatch Then
   LcCidade = RsCidade!Nome
End If
Set RsEmpresa = Nothing
'=== Imprime Cabecalho Nota
For a = 1 To Len(Txt(7).Text)
    LCLEtra = Mid$(Txt(7).Text, a, 1)
    If LCLEtra = " " Then Exit For
    LcVend = LcVend & LCLEtra
Next


LcLinha = Left(Txt(17).Text & LcEspa, 47) & _
Left(LcVend & LcEspa, 12) & Txt(0).Text '=== O Nome do Cliente e o Contato
LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = LcMargem & LcLinha & Chr(13)

LcTamanhoPedido = LcTamanhoPedido + 1
MtPedido(LcTamanhoPedido) = Chr(13)


LcLinha = Left(RsClientes!End & LcEspa, 44) & _
RsClientes!Bairro '=== O endereço e o Bairro
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
For x = 1 To 5 - lcvezes
    LcLinha = LcLinha & "    "
Next

LcLinha = LcLinha & Right("                " & Txt(12).Text, 10)
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
Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Txt(11).Enabled = GlIpi
Label3(15).Enabled = GlIpi
Status.Text = "Em Lançamento"
GeraGrid
GlComissaoVelha = 0
Me.Height = 8115
Me.Width = 11970
GlEscolhe = 1
'CalculaNumeroNota
CarregaCboUnidade

End Sub

Private Sub Form_Unload(Cancel As Integer)
GlCarregado = False
End Sub

Private Sub Natureza_Click()
On Error Resume Next
'If Natureza.Text = "Orçamento" Then
 '  Command3.Visible = False
 '  CmdSalvar.Visible = True
'Else
'  Command3.Visible = True
 ' CmdSalvar.Visible = False
'End If
   
End Sub

Private Sub Natureza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  SendKeys "{TAB}"
Else
  If KeyCode <> 116 Then Teclas (KeyCode)
End If

End Sub



Private Sub Txt_Change(Index As Integer)
On Error Resume Next
If Index = 3 Or Index = 5 Then CalculaValores
If Index = 5 Then
   unitreal.Text = Txt(5).Text
End If

End Sub

Private Sub Txt_GotFocus(Index As Integer)
On Error Resume Next
If Index = 8 Then
   If Len(Trim(Txt(10).Text)) = 0 And GlEsclheVendedor Then
      MsgBox "É Necessário Escolher o Vendedor Responsável.", 64, "Aviso"
      Txt(10).SetFocus
   End If
End If
If Index = 1 Then
   If Len(Trim(Txt(18).Text)) = 0 And GlEscolheCliente Then
      MsgBox "É Necessário Escolher o Cliente para a Nota Fiscal.", 64, "Aviso"
      Txt(1).SetFocus
   End If
End If
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 118 Then SendKeys "%+{Q}"
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me
'==== Vai Detalhar o Produto
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 120 Then Txt(19).SetFocus
If KeyCode = 117 Then
   Txt(13).SetFocus
   Exit Sub
End If
If KeyCode = 38 Then
   VoltaCampo (KeyCode)
End If
If KeyCode = 116 Then
   If Index = 18 Or Index = 17 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(9).Text)) > 0 Then
            FrmPesquisaCliente.Txt.Text = Txt(9).Text
            GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt(9).Text) & "*'  order by RAZAOSOC"
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
If Index = 5 Then
   If KeyAscii = 46 Then KeyAscii = 44
End If
If Index = 13 Then
   If KeyAscii = 46 Then KeyAscii = 44
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 6 And GlLibera Then montagrid
If Index = 1 Then
   If Len(Trim(Txt(1).Text)) > 0 Then
      If GLCalculacodigoProduto Then
         Txt(1).Text = Right("00000" & Txt(1).Text, 5)
      Else
         Txt(1).Text = Trim(Txt(1).Text)
      End If
      BuscaProduto
   End If
End If

If Index = 8 Then
   If Len(Txt(8).Text) > 0 Then
      Txt(8).Text = Right("00000" & Txt(8).Text, 5)
      If Len(Trim(Txt(8).Text)) > 0 Then buscafornecedor (1)
   End If
End If
If Index = 18 Then
   If Len(Txt(18).Text) > 0 Then
      
      If GLCalculacodigoCliente Then Txt(18).Text = Right("00000" & Txt(18).Text, 5)
      If Len(Trim(Txt(18).Text)) > 0 Then BuscaCliente (1)
   End If
End If

If Index = 5 Then
   ConferePreco
End If
If Index = 10 And Len(Trim(Txt(Index).Text)) <> 0 Then BuscaVendendor
If Index = 13 Then
If Not GlDetalhaDesconto Then
      VerificaDesconto
   Else
      RateiaDescontoPerc
   End If
End If
If Index = 19 Then
   If GlRateiaAcrecimo Then
       RateiaAcrescimo
   Else
      CalculaAcrecimo
   End If
End If
End Sub
Function ConferePreco()
On Error Resume Next

Dim LcPreconovo, LcPRecoAntigo As Currency
GlLibera = False
If Len(minimo.Text) = 0 Then minimo.Text = 0
LcPreconovo = CDbl(Txt(5).Text)

LcPRecoAntigo = CDbl(minimo.Text)

If LcPreconovo < LcPRecoAntigo And Natureza.Text = "VENDA" Then
    Liberacao.Show
    GlLibera = False
    GlEscolha = True
     Do Until Not GlEscolha
        DoEvents
     Loop
     If GlLibera Then
        comissao.Text = 1
     Else
        Txt(5) = LcPrecoVelho
        Txt(5).SetFocus
     End If
Else
  GlLibera = True
  If Len(comissao.Text) = 0 Then comissao.Text = 0
  If CLng(comissao.Text) <> 1 Then
     comissao.Text = 2
  End If
End If

End Function

Function SalvaOrcamento()
On Error Resume Next
Dim Rsorcamento As Recordset, RsItens As Recordset
Dim RsCliente As Recordset, RsProduto As Recordset
Dim LcNovo As Integer

If Len(Txt(0).Text) = 0 Then LcNovo = True Else LcNovo = False

LcSql1 = "Select * from orcamento"
LcSql2 = "Select * from DadosOrcamento where doc='" & Txt(0).Text & "'"
LcSql3 = "Select * from Alid001"
LcSql4 = "Select * from Alid009"
CalculaNumeroNota
AbreBase

Set Rsorcamento = Dbbase.OpenRecordset(LcSql1)
Set RsItens = Dbbase.OpenRecordset(LcSql2)
Set RsCliente = Dbbase.OpenRecordset(LcSql3)
Set RsProduto = Dbbase.OpenRecordset(LcSql4)

'==== Grava Os dados da Nota Fiscal

LcCriterio = "Doc='" & Txt(0).Text & "'"
Rsorcamento.FindFirst LcCriterio
If Not Rsorcamento.NoMatch Then
   Rsorcamento.Edit
Else
   Rsorcamento.AddNew
End If
Rsorcamento("DOC") = Txt(0).Text
Rsorcamento("DTEMIS") = CDate(Txt(12).Text)
Rsorcamento("NATUREZA") = Left(Natureza.Text, 2)
Rsorcamento("CLiente") = Txt(18).Text

Rsorcamento("TIPOTRANS") = Mid(DadosOrcamento.Tipo.Text, 1, 1)
Rsorcamento("PLACATRANS") = DadosOrcamento.Placa.Text
Rsorcamento("UFTRANS") = DadosOrcamento.Txt(1).Text
Rsorcamento("CGCCPFTRAN") = DadosOrcamento.Txt(2).Text
Rsorcamento("ENDTRANS") = DadosOrcamento.Txt(3).Text
Rsorcamento("MUNICTRANS") = DadosOrcamento.Txt(4).Text
Rsorcamento("UFMUNIC") = DadosOrcamento.Txt(5).Text
'RsOrcamento("INSCEST") = DadosOrcamento.txt(6).Text
Rsorcamento("OBS02") = DadosOrcamento.Txt(7).Text
Rsorcamento("OBS03") = DadosOrcamento.Txt(8).Text
Rsorcamento("OBS04") = DadosOrcamento.Txt(9).Text
Rsorcamento("OrcVenda") = Natureza.Text
Rsorcamento("Vendedor") = Txt(10).Text
Rsorcamento("Status") = "Confirmado"
Rsorcamento("cidade") = DadosOrcamento.Txt(6).Text
Rsorcamento("cep") = DadosOrcamento.Txt(12).Text
Rsorcamento("CondPag") = DadosOrcamento.TipoPag.Text
If Len(Txt(13).Text) > 0 Then Rsorcamento!Desconto = Txt(13).Text
If Len(Desconto.Text) > 0 Then Rsorcamento!TotalDesconto = CCur(Desconto.Text)
If Len(Txt(16).Text) > 0 Then Rsorcamento!TotalGeral = CCur(Txt(16).Text)
If Len(Txt(15).Text) > 0 Then Rsorcamento!TotalProduto = CCur(Txt(15).Text)
Rsorcamento!formapag = DadosOrcamento.TipoMonetario.Text
Rsorcamento!Dias = DadosOrcamento.Txt(11).Text
Rsorcamento!Acrecimo = CCur(Acrecimo.Text)
Rsorcamento!DetalhaDesconto = Txt(13).Text
Rsorcamento!DetahaAcrecimo = Txt(19).Text
If LcNovo Then
  Rsorcamento("TRANSP") = Transportadora.Text
  Rsorcamento("FoneTransp") = fone.Text
  Rsorcamento("obs") = dados.Text
  LcNovo = False
End If

If DadosOrcamento.Vencimento(0).Text <> "  /  /  " Then
   Rsorcamento!Vencimento1 = DadosOrcamento.Vencimento(0).Text
   Rsorcamento!vencimento2 = DadosOrcamento.Vencimento(1).Text
   Rsorcamento!vencimento3 = DadosOrcamento.Vencimento(2).Text
   Rsorcamento!vencimento4 = DadosOrcamento.Vencimento(3).Text
   Rsorcamento!vencimento5 = DadosOrcamento.Vencimento(4).Text
End If
Rsorcamento.Update

'==== Grava os itens da nota fiscal
Do Until RsItens.EOF
    RsItens.Delete
    RsItens.MoveNext
Loop
RsItens.Close

LcSql2 = "Select * from DadosOrcamento"
Set RsItens = Dbbase.OpenRecordset(LcSql2)
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       LcCodigo = LcMat(a).CodPro
       LcCriterio = "Cod='" & LcCodigo & "'"
       If Natureza.Text <> "Orçamento" Then
          RsProduto.FindFirst LcCriterio
          If Not RsProduto.NoMatch Then
             RsProduto.Edit
             If Len(RsProduto!QuantEstoque) > 0 Then
                RsProduto("QuantEstoque") = RsProduto!QuantEstoque - LcMat(a).Qut
             Else
                RsProduto("QuantEstoque") = 0 - LcMat(a).Qut
             End If
             RsProduto.Update
          End If
       End If
       
       LcCriterio = "Doc='" & Txt(0).Text & "' and CodigoProduto='" & LcCodigo & "' AND Descricao='" & LcMat(a).produto & "'"
       RsItens.FindFirst LcCriterio
       If Not RsItens.NoMatch Then
          RsItens.Edit
       Else
          RsItens.AddNew
       End If
       RsItens("Doc") = Txt(0).Text
       RsItens("CodigoProduto") = LcCodigo
       RsItens("Descricao") = LcMat(a).produto
       RsItens("Quant") = LcMat(a).Qut
       RsItens("Unit") = LcMat(a).VUnit
       RsItens("Total") = LcMat(a).Vtotal
       RsItens("unid") = LcMat(a).Und
       RsItens("com") = LcMat(a).Com
       RsItens("item") = LcMat(a).item
       RsItens("Ipi") = LcMat(a).ipi
       RsItens("valorUnitarioReal") = LcMat(a).UnitarioAntigo
       
       RsItens.Update
     End If
Next
'==== Atualiza Dados Cliente
If Natureza.Text <> "Orçamento" Then
   If GlEscolheCliente Then
      LcCriterioPes = "codigo='" & Txt(18).Text & "'"
      RsCliente.FindFirst LcCriterioPes
      If Not RsCliente.NoMatch Then
         RsCliente.Edit
         RsCliente("ULTCOMPRA") = CDate(Txt(12).Text)
         If DadosOrcamento.TipoPag.Text = "A Prazo" Then
            If Len(RsCliente!CreditoUtilizado) > 0 Then
               RsCliente!CreditoUtilizado = RsCliente!CreditoUtilizado + CCur(Txt(16).Text)
            Else
               RsCliente!CreditoUtilizado = CCur(Txt(16).Text)
            End If
        End If
        RsCliente.Update
      End If
   End If
End If
'=== Fecha as Bases
Rsorcamento.Close
RsItens.Close
RsCliente.Close
RsProduto.Close
Dbbase.Close
Set Rsorcamento = Nothing
Set RsComissao = Nothing
Set RsCliente = Nothing
Set RsProduto = Nothing
Set Dbbase = Nothing
         
End Function

Function ExcluiItem(LcNItem As Integer)
On Error Resume Next
Dim a, b As Integer

For a = 0 To LcTam - 1
    If LcMat(a).item = LcNItem Then
       LcMat(a).CodPro = ""
       LcMat(a).item = 0
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
Txt(1).SetFocus
End Function
Function CalculaNumeroNota()
On Error Resume Next
If Len(Txt(0).Text) > 0 Then Exit Function
Dim LcSql As String, LcNumeroNota As String
Dim RsNota As Recordset
LcSql = "Select * from orcamento order by doc"
AbreBase
Set RsNota = Dbbase.OpenRecordset(LcSql)
If RsNota.EOF Then
   LcNumeroNota = "000001"
Else
   RsNota.MoveLast
   LcNumeroNota = Right("000000" & CStr(Val(RsNota("doc")) + 1), 6)
End If
Txt(0).Text = LcNumeroNota
RsNota.AddNew
RsNota("DOC") = LcNumeroNota
RsNota.Update
RsNota.Close
Dbbase.Close
Set RsNota = Nothing
Set Dbbase = Nothing

End Function

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 118 Then SendKeys "%+{Q}"
If KeyCode = 119 Then FrmDescicaoProduto.Show , Me
If KeyCode = 117 Then
If KeyCode = 120 Then Txt(19).SetFocus
If KeyCode = 122 Then If GlVariasComissao Then Exibecomissao.Show , Me


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
On Error Resume Next
Dim RsComissao As Recordset, RsFuncionario As Recordset
Dim RsProdutos As Recordset
Dim LcPercDesc, LcPercAcres, LcDifAcr, LcDifDes As Double

LcSql = "Select * from Alid201 where NF='" & Txt(0).Text & "'"
LcSq2 = "Select * from Alid200 where codigo='" & Txt(10).Text & "'"
AbreBase
Set RsComissao = Dbbase.OpenRecordset(LcSql)
Set RsFuncionario = Dbbase.OpenRecordset(LcSq2)
If Len(Desconto.Text) > 0 Then
   LcPercDesc = CDbl(Desconto.Text) / CDbl(Txt(15).Text)
Else
   LcPercDesc = 0
End If
If Len(Acrecimo.Text) > 0 Then
   LcPercAcres = CDbl(Acrecimo.Text) / CDbl(Txt(15).Text)
Else
   LcPercDesc = 0
End If
Do Until RsComissao.EOF
   RsComissao.Delete
   RsComissao.MoveNext
Loop
RsComissao.Close
LcSql = "Select * from Alid201"
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
         If comissao.Text = "1" Then Ibaixo = True Else Ibaixo = False
         RsComissao("ITEMBAIXO") = Ibaixo
         LcDifAcr = LcPercAcres * LcMat(a).comissao
         LcDifDes = LcPercDesc * LcMat(a).comissao
         RsComissao("COMISSAO") = LcMat(a).comissao + LcDifAcr - LcDifDes
         RsComissao("DATAVENDA") = CDate(Txt(12).Text)
         RsComissao("CLIENTE") = Txt(18).Text
         LcCriterioFornec = "select * from alid009 where cod='" & LcMat(a).CodPro & "'"
         Set RsProdutos = Dbbase.OpenRecordset(LcCriterioFornec)
         If Not RsProdutos.EOF Then
            RsComissao("Fornecedor") = RsProdutos!Fornecedor
         End If
         RsComissao.Update
         RsProdutos.Close
         Set RsProdutos = Nothing
     End If
Next
RsComissao.Close
Dbbase.Close
Set RsComissao = Nothing
Set Dbbase = Nothing
         
End Function
Function limpanota()
On Error Resume Next
Desconto.Text = ""
Liberado = False
LcTam = 0
LcItem = 0
ReDim LcMat(0)
item.Rows = 1
For a = 0 To 19
   Txt(a).Text = ""
   'valor.Text = ""
Next
limite.Text = 0
utilizado.Text = 0
comisVenda.Text = 0
'CalculaNumeroNota
Acrecimo.Text = ""
Desconto.Text = ""
Transportadora.Text = ""
fone.Text = ""
dados.Text = ""
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
Command4.Caption = "Pes&quisa F7"
Txt(7).SetFocus
ipi.Text = ""
GlComissaoVelha = 0
Status.Text = "Em Lançamento"
End Function
Function Atualizacaixa(LcNumeroContas As Integer)
On Error Resume Next
Dim RsContasReceber As Recordset, RsCaixa As Recordset
Dim RsTipoMonetario As Recordset

LcSql1 = "Select * from Alid015"
LcSql2 = "Select * from Alid016"
LcSql3 = "Select * from Alid008"

AbreBase
Set RsContasReceber = Dbbase.OpenRecordset(LcSql1)
Set RsCaixa = Dbbase.OpenRecordset(LcSql2)
Set RsTipoMonetario = Dbbase.OpenRecordset(LcSql3)
Select Case DadosOrcamento.TipoPag.Text
    Case Is = "A Vista"
         If GlVistaSaida Then
            RsContasReceber.AddNew
            RsContasReceber("NF") = Txt(0).Text
            RsContasReceber("CLIENTE") = Txt(18).Text
            LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
            If Not RsTipoMonetario.NoMatch Then
               RsContasReceber("TPMONET") = RsTipoMonetario("TPMONET")
            End If
            RsContasReceber("VALOR") = CCur(Txt(16).Text)
            RsContasReceber("DATA") = CDate(Txt(12).Text)
            RsContasReceber("DTVENC") = CDate(Txt(12).Text)
            RsContasReceber("DTPAGTO") = CDate(Txt(12).Text)
            RsContasReceber("VALPAGO") = CCur(Txt(16).Text)
            RsContasReceber("TIPORD") = "R"
            RsContasReceber("Acrescimo") = 0
            RsContasReceber.Update
          End If
          
         If GlCaixaSaida Then
            RsCaixa.AddNew
            RsCaixa("NF") = Txt(0).Text
            RsCaixa("RECDESP") = "R"
            RsCaixa("CLICRED") = Txt(18).Text
            LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
            If Not RsTipoMonetario.NoMatch Then
               RsCaixa("TPMONET") = RsTipoMonetario("TPMONET")
            End If
            RsCaixa("VALOR") = CCur(Txt(16).Text)
            RsCaixa("DATA") = CDate(Txt(12).Text)
            RsCaixa.Update
          End If
    Case Is = "A Prazo"
         If GlFaturaSaida Then
            For a = 1 To LcNumeroContas
                RsContasReceber.AddNew
                RsContasReceber("NF") = Txt(0).Text & "/" & Right("00" & CStr(a), 2)
                RsContasReceber("CLIENTE") = Txt(18).Text
                LcCriterioPes = "XTPMONET='" & DadosOrcamento.TipoMonetario.Text & "'"
                If Not RsTipoMonetario.NoMatch Then
                    RsContasReceber("TPMONET") = RsTipoMonetario("TPMONET")
                End If
                RsContasReceber("DATA") = CDate(Txt(12).Text)
                RsContasReceber("VALOR") = CCur(DadosOrcamento.valor.Text)
                Select Case a
                    Case Is = 1
                         RsContasReceber("DTVENC") = CDate(DadosOrcamento.Vencimento(0).Text)
                    Case Is = 2
                         RsContasReceber("DTVENC") = CDate(DadosOrcamento.Vencimento(1).Text)
                    Case Is = 3
                         RsContasReceber("DTVENC") = CDate(DadosOrcamento.Vencimento(2).Text)
                End Select
                RsContasReceber("TIPORD") = "R"
                RsContasReceber("Acrescimo") = 0
                RsContasReceber.Update
            Next
          End If
        
         
End Select

End Function
