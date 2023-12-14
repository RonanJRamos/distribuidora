VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSaidaProdutoAlternativo 
   BackColor       =   &H00E6DB9D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saída de Estoque"
   ClientHeight    =   8010
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11715
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdReimprime 
      Caption         =   "&Reimprimir"
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
      TabIndex        =   70
      Top             =   1620
      Width           =   1575
   End
   Begin VB.TextBox CodVales 
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
      Left            =   120
      TabIndex        =   69
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox proposta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8160
      TabIndex        =   68
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   240
      TabIndex        =   29
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   8040
      TabIndex        =   62
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox santamaria1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   60
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox santamaria 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   59
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox california 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   58
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox almox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   57
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox CFOP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmSaidaProdutoAlternativo.frx":0000
      Left            =   8520
      List            =   "FrmSaidaProdutoAlternativo.frx":0022
      TabIndex        =   56
      Text            =   "5.102"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox utilizado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   53
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox limite 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   52
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   0
      Left            =   7200
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox tam 
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
      Left            =   5400
      TabIndex        =   50
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Comissao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      TabIndex        =   46
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   14
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   7320
      Width           =   4935
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox cst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox icms 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox minimo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5520
      TabIndex        =   11
      Top             =   2640
      Width           =   810
   End
   Begin VB.ComboBox Unidade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2610
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
      ItemData        =   "FrmSaidaProdutoAlternativo.frx":0063
      Left            =   5160
      List            =   "FrmSaidaProdutoAlternativo.frx":0082
      TabIndex        =   2
      Text            =   "VENDAS A VISTA"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Custo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   3135
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Left            =   10080
      TabIndex        =   22
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   8655
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8880
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6360
      TabIndex        =   12
      Top             =   2640
      Width           =   810
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Buscar Pedido F2"
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
      TabIndex        =   67
      Top             =   1245
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Pesquisar F7"
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
      TabIndex        =   54
      Top             =   870
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
      Left            =   10080
      TabIndex        =   24
      Top             =   495
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Fechar Nota F3"
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
      TabIndex        =   34
      Top             =   120
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
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Ver Últimas Compras do Cliente Pressione F12 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto F11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   3240
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   64
      Top             =   7080
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
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
      Left            =   7200
      TabIndex        =   63
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS"
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
      Left            =   5040
      TabIndex        =   61
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
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
      Index           =   14
      Left            =   7920
      TabIndex        =   55
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Nota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   10080
      TabIndex        =   48
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   10080
      TabIndex        =   47
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Complementares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   6720
      TabIndex        =   44
      Top             =   7080
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Calc. ICMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   42
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   4440
      TabIndex        =   41
      Top             =   7080
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Vendedor"
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
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Produto Digite Seu Código,  Nome ou pressione F5 Para Detalhar Produto Pressione F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   3000
      Width           =   9150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Cliente pressione F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1200
      TabIndex        =   32
      Top             =   1320
      Width           =   3450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Natureza"
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
      Index           =   8
      Left            =   4320
      TabIndex        =   31
      Top             =   600
      Width           =   780
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
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   9960
      X2              =   9960
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Cliente"
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
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Emissão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   20
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. Total"
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
      Index           =   5
      Left            =   8640
      TabIndex        =   51
      Top             =   2400
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. Unit."
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
      Index           =   4
      Left            =   7200
      TabIndex        =   19
      Top             =   2400
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unid. / Com"
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
      Index           =   3
      Left            =   4845
      TabIndex        =   18
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
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
      Index           =   2
      Left            =   6360
      TabIndex        =   17
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Notas Fiscais de Saída"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "FrmSaidaProdutoAlternativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoUnid
      Codigo As String
      Descricao As String
      Simbolo As String
      Quantidade As Double
End Type
Private Type Tipo50
    icms As String
    valor As Double
End Type
Private LcSql As String
Private LcSq As String
Private LcItem As Long, LcTam, LcQUn, LcQuantiImpressao, LcQuantiImpressaoBoleto As Long
Private FnunNota, FnunBoleto
Private LcNota, LcBoleto, LcEspC As String
Private LcFocus, LcCalculado, LcSalto  As Integer
Private LcPrecoVelho As Currency
Private ComNormal As Long, ComAlterado As Long, LcQuantNesc As Long, LcQtSta1 As Long, LcQtSta As Long, LcQtCal As Long
Private LcLinha As String
Private RsOpcoes As Recordset, RsClientes As Recordset
Private RsCidade As Recordset
Private LcValor1 As Double, LcValor2 As Double, LcValor3 As Double, LcUltimo As Double
Private LcAlteradoCliente, LcAlteradoProduto, LcAlteradoFuncionario As Integer
Private LcMat() As DadosEntrada, LcLimpa As Integer
Private Liberado, LcBuscaCliente, LcBuscaNota As Integer
Private MtUnidade() As TipoUnid, MtImpressao(), MtBoleto() As String
Private LcImpressoes As Double, LcProximo, LcLimpaValor, LcPesquisaCli As Integer
Private LcSaldoCaixa As Double, LcSaldoUnit As Double
Private TotalCaixa As Double, TotalUnitario As Double
Private LcFechaitem, a As Integer
Private LcQSanta As Double
Private LcQSanta1 As Double
Private LcQCalifornia As Double
Private LcQUnSanta As Double
Private LcQUnSanta1 As Double
Private LcQUnCalifornia As Double
Private LcMargem As String
Private Estoque As ControleDb
Private Rel As New CrysVendasAlternativo
Function Reescreveicms()
Dim a As Integer
If Item.Rows < 1 Then Exit Function
If Len(Txt(5).Text) = 0 Then Exit Function
ClienteForaEstado = ClienteEstado

For a = 0 To UBound(LcMat)
    'GlVerificarIcmsDiferenciado = True
    If GlVerificarIcmsDiferenciado Then
        If Not ClienteForaEstado Then
            If Len(Txt(5).Text) = 0 Then
               
            Else
                If LcMat(a).cst = "060" Then
                    LcMat(a).icms = 0
                Else
                    'LcMat(a).cst = cst.Text
                    LcMat(a).icms = Txt(5).Text
                End If
            End If
        Else
            If Natureza.Text <> "ORG PUBL. EST." Then LcMat(a).icms = Txt(5).Text Else LcMat(a).icms = 0
        End If
    Else
        If Natureza.Text <> "ORG PUBL. EST." Then LcMat(a).icms = Txt(5).Text Else LcMat(a).icms = 0
    End If
Next

'====>Escreve o grid
For a = 1 To Item.Rows - 1
  If GlVerificarIcmsDiferenciado Then
    If Item.TextMatrix(a, 8) <> "0" Or ClienteForaEstado Then
       Item.TextMatrix(a, 8) = Txt(5).Text
    End If
  Else
     Item.TextMatrix(a, 8) = Txt(5).Text
  End If
Next
CalculaIcms
End Function

Function BaixaEstoque(CodigoP As String, QB As Double, cb As Double, UnB As String)
Dim db      As Database
Dim Rsp     As ADODB.Recordset
Dim RsG     As Recordset
Dim RsCG    As Recordset
Dim Rsun    As Recordset
Dim LcSql   As String
Dim LcSql1  As String
Dim LcSql2  As String
Dim LcSql3  As String
Dim LcNome  As String
Dim LcCoG   As String
Dim LcUR    As String
Dim LcCaixa As Double
Dim LcQUn   As Double
Dim LcQunB  As Double
Dim LcUnP   As String
Dim LcSSan  As Double
Dim LcSSanu As Double
Dim LcSa1   As Double
Dim LcSa1u  As Double
Dim LcSc    As Double
Dim LcScu   As Double
Dim LcVS    As Double
Dim LcVsu   As Double
'===> Criando Sql's
LcSql = "Select * from produtos where codigo=" & CodigoP
LcSql1 = "Select * from alid013 where item='" & CodigoP & "' order by codigogalpao"
LcSql2 = "Select * from alid012"
LcSql3 = "Select * from alid004 where SIMBOLO='" & UnB & "'"
'==> Setando o banco e as tabelas

Set db = OpenDatabase(GLBase)
Set Rsp = AbreRecordset(LcSql) ' dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsG = db.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Rsun = db.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCG = db.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'===> Busca os Dados do Produto
If Not Rsp.EOF Then
   LcNome = Rsp!Nome
   LcQUn = Rsp!QtdMedida
   LcUnP = Rsp!Unimed
   
End If
If Not Rsun.EOF Then
   LcUR = Rsun!cod
Else
   LcUR = ""
End If
LcSSan = 0
LcSSanu = 0
LcSa1 = 0
LcSa1u = 0
LcSc = 0
LcScu = 0

'===> Verifica se Existe cadastro do Produto em todos os Galpoes.
If Not RsG.EOF Then
   '===> Procura o galpao california
   LcPes = "ALMOX='CALIFORNIA'"
   RsG.FindFirst LcPes
   If RsG.NoMatch Then
      RsG.AddNew
      RsG!almox = "CALIFORNIA"
      RsG!Item = CodigoP
      RsG!Estoque = 0
      RsG!Descricao = LcNome
      RsG!QuantUnidade = 0
      '===> Busca o Codigo do Galpao
      lcpes1 = "nome='CALIFORNIA'"
      RsCG.FindFirst lcpes1
      If Not RsCG.NoMatch Then
         LcCoG = RsCG!Codigo
      Else
         LcCoG = ""
      End If
      RsG!CODIGOGALPAO = LcCoG
      RsG.Update
   End If
   '===> Procura o Galpao Santa Maria
   LcPes = "ALMOX='SANTA MARIA'"
   RsG.FindFirst LcPes
   If RsG.NoMatch Then
      RsG.AddNew
      RsG!almox = "SANTA MARIA"
      RsG!Item = CodigoP
      RsG!Estoque = 0
      RsG!Descricao = LcNome
      RsG!QuantUnidade = 0
      
      '===> Busca o Codigo do Galpao
      lcpes1 = "nome='SANTA MARIA'"
      RsCG.FindFirst lcpes1
      If Not RsCG.NoMatch Then
         LcCoG = RsCG!Codigo
      Else
         LcCoG = ""
      End If
      RsG!CODIGOGALPAO = LcCoG
      RsG.Update
   End If
      '===> Procura o Galpao Santa Maria 2
   LcPes = "ALMOX='SANTA MARIA 2'"
   RsG.FindFirst LcPes
   If RsG.NoMatch Then
      RsG.AddNew
      RsG!almox = "SANTA MARIA 2"
      RsG!Item = CodigoP
      RsG!Estoque = 0
      RsG!Descricao = LcNome
      RsG!QuantUnidade = 0
      '===> Busca o Codigo do Galpao
      lcpes1 = "nome='SANTA MARIA 2'"
      RsCG.FindFirst lcpes1
      If Not RsCG.NoMatch Then
         LcCoG = RsCG!Codigo
      Else
         LcCoG = ""
      End If
      RsG!CODIGOGALPAO = LcCoG
      RsG.Update
   End If
End If
'===> Vamos Verificar se a unidade vendida é igual a Unidade Principal
'===> Pode Ter Havido Modificações no Estoque do Galpao, então Vamos Fecha-lo e Reabri-lo
RsG.Close
Set RsG = db.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'===> Busca a Quanidade antes da Baixa
Do Until RsG.EOF
    Select Case RsG!almox
        Case Is = "SANTA MARIA"
            LcSSan = RsG!Estoque
            LcSSanu = RsG!QuantUnidade
        Case Is = "SANTA MARIA 2"
            LcSa1 = RsG!Estoque
            LcSa1u = RsG!QuantUnidade
            
        Case Is = "CALIFORNIA"
            LcSc = RsG!Estoque
            LcScu = RsG!QuantUnidade
    End Select
    RsG.MoveNext
Loop

RsG.MoveFirst
If (LcUR = LcUnP) And (LcQUn = cb) Then
   '==> Otimo, é a Mesma Unidade
   LcCaixa = QB
   Do Until RsG.EOF
    If LcCaixa > 0 Then
       
       LcVS = RsG!Estoque
       '===> Verifica se a quantidade do galpao e maior ou igual a qunt. vendida
       If RsG!Estoque >= LcCaixa Then
          '==> é Maior, Pode Baixar normalmente
          RsG.Edit
          RsG!Estoque = RsG!Estoque - LcCaixa
          RsG.Update
          LcCaixa = 0
       Else
         '==> Não é
         LcCaixa = LcCaixa - RsG!Estoque
         RsG.Edit
         RsG!Estoque = 0
         RsG.Update
       End If
    End If
    RsG.MoveNext
   Loop
Else
  '===> A Quantidade é Diferente
  '===> Vamos Verificar se a quantidade Vendida em Outra Unidade é superior a unidade Cadastrada
  If (cb * QB) >= LcQUn Then
        LcCaixa = 0
        '==> é maior, Vamos Ver Quantas Caixas Vamos Baixar
        LcCaixa = Int((cb * QB) / LcQUn)
        LcQunB = (cb * QB) - LcQUn
  Else
        LcCaixa = 0
        LcQunB = QB * cb
  End If
  '===> Vamos Verificar se a quantidade em unidade e disponivel

  
  '===> Vamnos Baixar Nos Galpoes a caixas
  Do Until RsG.EOF

     If LcCaixa > 0 Then
        '===> Verifica se a quantidade do galpao e maior ou igual a qunt. vendida
       'If RsG!quantUnidade < (cb * QB) Then
       '     RsG.Edit
        '    RsG!Estoque = RsG!Estoque - 1
        '    RsG!quantUnidade = RsG!quantUnidade + LcQUn
        '    RsG.Update
      ' End If

        If RsG!Estoque >= LcCaixa Then
           '==> é Maior, Pode Baixar normalmente
           RsG.Edit
           RsG!Estoque = RsG!Estoque - LcCaixa
           RsG.Update
           LcCaixa = 0
        Else
           '==> Não é
          ' LcCaixa = LcCaixa - RsG!Estoque
          ' RsG.Edit
          ' RsG!Estoque = 0
          ' RsG.Update
        End If
     End If
     RsG.MoveNext
  Loop
  '===> Agora Vamos Baixar as Unidades
  RsG.MoveFirst
  Do Until RsG.EOF
     If LcQunB > 0 Then
        '===> Verifica se a quantidade em unidade do galpao e maior ou igual a qunt. vendida
        If RsG!QuantUnidade >= LcQunB Or (RsG!Estoque > 0) Then
           '==> é Maior, Pode Baixar normalmente
           '==> Vamos Verificar se a Quantidade Unitaria é Superior
           If RsG!QuantUnidade >= LcQunB Then
            RsG.Edit
            RsG!QuantUnidade = RsG!QuantUnidade - LcQunB
            RsG.Update
            LcCaixa = 0
           Else
            '===> Vamos Abrir Uma Caixa
            RsG.Edit
            RsG!Estoque = RsG!Estoque - 1
            RsG!QuantUnidade = (RsG!QuantUnidade + LcQUn) - LcQunB
            RsG.Update
            LcQunB = 0
          End If
        Else
           '==> Não é
           'LcQunB = LcQunB - RsG!quantUnidade
           'RsG.Edit
          ' RsG!quantUnidade = 0
           'RsG.Update
        End If
    End If
    RsG.MoveNext
 Loop
End If
'===> Agora, Vamos Acertar o Cadastro Geral doss Produtos
RsG.MoveFirst
LcQunB = 0
LcCaixa = 0
'==> Agora, Vê quanto tirou de cada GalpAO
Do Until RsG.EOF
    Select Case RsG!almox
        Case Is = "SANTA MARIA"
            LcQSanta = LcSSan - RsG!Estoque
            LcQUnSantas = LcSSanu - RsG!QuantUnidade
        Case Is = "SANTA MARIA 2"
            LcQSanta1 = LcSa1 - RsG!Estoque
            LcQUnSanta1 = LcSa1u - RsG!QuantUnidade
            
        Case Is = "CALIFORNIA"
            LcQCalifornia = LcSc - RsG!Estoque
            LcQUnCalifornia = LcScu - RsG!QuantUnidade
    End Select
    RsG.MoveNext
Loop
RsG.MoveFirst
Do Until RsG.EOF
   LcCaixa = LcCaixa + RsG!Estoque
   LcQunB = LcQunB + RsG!QuantUnidade
   RsG.MoveNext
Loop
'Rsp.Edit
Rsp!QuantEstoque = LcCaixa
Rsp!QuantUnidade = LcQunB
Rsp.Update

Rsp.Close
RsG.Close
RsCG.Close
Rsun.Close
db.Close

Set Rsp = Nothing
Set RsG = Nothing
Set RsCG = Nothing
Set Rsun = Nothing
Set db = Nothing
End Function
Function BaixaVale()

LcSql = "Update Vales SET baixado=1 where cliente='" & Txt(8).Text & "' and baixado=0 and marca='X'"
''abreconexao
ExecutaSql LcSql


End Function

Function LancaVendaEstado()
'Dim Rs As adodb.Recordset
Dim Rsp As ADODB.Recordset
Dim LcValor As Currency
Dim LcTem As Boolean
'On Error GoTo erroLacaV
'AbreBase

If Natureza.Text <> "ORG PUBL. EST." Then Exit Function
''abreconexao
LcValor = 0
For a = 1 To Item.Rows - 1
    Set Rsp = AbreRecordset("Select * from produtos where codigo=" & Item.TextMatrix(a, 1), True)  ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
    If Not Rsp.EOF Then
        If Val(Rsp!cst) = 60 Or Val(Rsp!cst) = 160 Or Val(Rsp!cst) = 260 Then
           LcValor = LcValor + CCur(Item.TextMatrix(a, 7))
           LcTem = True
        End If
    End If
    Set Rsp = Nothing
Next
If LcTem Then
    LcValors = Replace(LcValor, ",", ".")
    ExecutaSql "delete from VendaSubstEstado where numnf='" & Txt(0).Text & "'"
    'Set Rs = AbreRecordset("Select * from VendaSubstEstado where numnf='" & Txt(0).Text & "'", Rs)
    LcSql = "insert into vendaSubstEstado (NumNf,Dtemis,Cliente,ValorNota) values ('" & Txt(0).Text & "','"
    LcSql = LcSql & Format(Txt(12).Text, "yyyy-mm-dd") & "','" & Txt(8).Text & "'," & LcValors & ")"
    'MsgBox LcSql
    ExecutaSql LcSql
End If

'Rs.AddNew
'Rs!numnf = Txt(0).Text
'Rs!DTEMIS = Format(Txt(12).Text, "dd/mm/yy")
'Rs!Natureza = Natureza.Text
'Rs!CLIENTE = Txt(8).Text
'Rs!ValorNota = LcValor
'Rs.Update
'Set Rs = Nothing
Exit Function
'erroLacaV:
'MsgBox err.Description & err.Number
'Resume 0
End Function

Function logErro(LcNumero As Variant, LcDesc As String, LcComentario As String)
On Error Resume Next
Dim LcRepete, LcIcone As Integer, msg, lctitulo, LcNomeArquivo As String
Dim LcExibemsg As Integer
Dim LcDiretorio As String
Dim LcGrifa     As String
LcGrifa = String(80, "-")
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then Exit For
Next
LcDiretorio = Mid(GLBase, 1, a)
LcIcone = 64
LcNumero = FreeFile

LcNomeArquivo = LcDiretorio & "ErrosNota.txt"

Open LcNomeArquivo For Append As #LcNumero      ' Open file for output.
 Write #LcNumero, "Data:" & Date & "  Hora:" & Time & " Maquina:" & GlNomeMaquina & "  Usuário:" & GlUsuario
 Write #LcNumero, "          Descrição:" & LcDesc
 Write #LcNumero, "          Nº do Erro:" & LcNumero
 Write #LcNumero, "          Comentario:" & LcComentario
 Write #LcNumero, LcGrifa
Close #LcNumero

End Function
Function LogAtualiza(LcNota As String, LComen As String)

Dim LcRepete, LcIcone As Integer, msg, lctitulo, LcNomeArquivo As String
Dim LcExibemsg, LcNumero As Integer
Dim LcDiretorio As String
Dim LcGrifa     As String
LcGrifa = String(80, "-")
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then Exit For
Next
LcDiretorio = Mid(GLBase, 1, a)
LcIcone = 64
LcNumero = FreeFile

LcNomeArquivo = LcDiretorio & "ErrosNota.txt"

Open LcNomeArquivo For Append As #LcNumero      ' Open file for output.
 Write #LcNumero, "Data:" & Date & "  Hora:" & Time & " Maquina:" & GlNomeMaquina & "  Usuário:" & GlUsuario
 Write #LcNumero, "          Descrição:" & LcDesc
 Write #LcNumero, "          Nº do Erro:" & LcNumero
 Write #LcNumero, "          Comentario:" & LcComentario
 Write #LcNumero, LcGrifa
Close #LcNumero

End Function

Function BuscaProposta(LcNumeroOrc As String)
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
LcSql6 = "Select * from produtos"

LcBuscaNota = True
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsItem = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set rsCliente = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVendedor = Dbbase.OpenRecordset(LcSql5, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = AbreRecordset(LcSql6) ' dbOpenDynaset, dbSeeChanges, dbOptimistic)

'==== Preenchendo a Nota

If RsOrc.EOF Then
   MsgBox "A Nota Fiscal Nº: " & LcNumeroOrc & " Não foi encontrado..."
   Command4.Caption = "Pes&quisa F7"
   Txt(10).SetFocus
   Exit Function
End If
Txt(0).Text = ""
Txt(12).Text = Date
Txt(6).Text = RsOrc!Status
Txt(5).Text = RsOrc!icms & ""
Txt(17).Text = ""
If Not IsNull(RsOrc!Desconto) Then
  If RsOrc!Desconto <> 0 Then
     Txt(17).Text = RsOrc!Desconto
  End If
End If

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
        Natureza.Text = "DEVOLUCAO"
    Case Is = "OR"
        Natureza.Text = "ORG PUBL. EST."
End Select
If Len(RsOrc!Comissao) > 0 Then
   Comissao.Text = RsOrc!Comissao
Else
   Comissao.Text = "1.5"
End If
If Len(RsOrc!CFOP) > 0 Then
   CFOP.Text = RsOrc!CFOP
Else
   CFOP.Text = "512"
End If
Txt(10).Text = RsOrc!Vendedor & ""
LcCriterio = "Codigo='" & RsOrc!Vendedor & "'"
RsVendedor.FindFirst LcCriterio
If Not RsVendedor.NoMatch Then
   Txt(7).Text = RsVendedor!Nome
Else
  Txt(7).Text = ""
End If
Txt(8).Text = RsOrc!Cliente
LcCriterio = "Codigo='" & RsOrc!Cliente & "'"
rsCliente.FindFirst LcCriterio
If Not rsCliente.NoMatch Then
   Txt(9).Text = rsCliente!razaosoc
   Txt(8).Text = rsCliente!Codigo
End If

Txt(15).Text = RsOrc!ValorProduto
Txt(16).Text = RsOrc!ValorNota
'If Len(RsOrc!desconto) > 0 Then Txt(13).Text = RsOrc!desconto Else Txt(13).Text = ""
'If Len(RsOrc!TotalDesconto) > 0 Then desconto.Text = RsOrc!TotalDesconto Else desconto.Text = ""
'===== Escreve dados Grid
LcItem = 0
LcTam = 0
'ReDim LcMat(LcTam)
Do Until RsItem.EOF
    LcItem = LcItem + 1
    '===> verifica o Estoque
    
    If Not VerificaEstoquedisponivel(CStr(RsItem("codProd")), CDbl(RsItem("QTDE")), CDbl(RsItem("QTDUM"))) Then
       MsgBox "O Produto " & Chr(13) & RsItem("descricao") & Chr(13) & " Não Possui Estoque Suficiente para a Venda." & Chr(13) & " Entre com Mais Produto para a Liberação do Pedido.", vbCritical, "Estoque Insuficiente"
       limpanota
       GoTo Saida
    End If
    
    ReDim Preserve LcMat(LcTam)
    If Len(RsItem!Item) > 0 Then LcMat(LcTam).Item = RsItem!Item
      LcCriterio = "codigo=" & RsItem("codProd")
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
      LcMat(LcTam).cst = cst.Text
      LcMat(LcTam).icms = icms.Text
      LcMat(LcTam).CodPro = RsItem("codProd")
      LcMat(LcTam).Qut = RsItem("QTDE")
      LcMat(LcTam).VUnit = RsItem("VALUNIT")
      LcMat(LcTam).Und = RsItem("UNIMED")
      LcMat(LcTam).com = RsItem("QTDUM")
      LcMat(LcTam).Produto = RsItem("descricao")
      LcMat(LcTam).QuanTidadeBaixa = RsItem("QTDE")
    
      LcMat(LcTam).Vtotal = LcMat(LcTam).Qut * LcMat(LcTam).VUnit
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
    CmdSalvar.Visible = True
    Command3.Visible = False
 End If
 If Not verificacredito(CDbl(Txt(16).Text), CStr(Txt(8).Text)) Then
    MsgBox "Limite Não Liberado", 64, "Aviso"
    limpanota
 End If
 'If VerificaAtraso(Txt(8).Text) Then
 '   MsgBox "Cliente em Atraso não Liberado.", 64, "Aviso"
 '   limpanota
 'End If
' verificavale
Saida:
 RsOrc.Close
 RsItem.Close
 rsCliente.Close
 RsVendedor.Close
 LcBuscaNota = False
 
 Exit Function
 
ErroBuscaNota:
 'MsgBox err.Description & err.Number
 Resume Next
End Function
Function verificacredito(LcTotal As Double, LcCliente As String) As Boolean
Dim rsCliente As Recordset
AbreBase
Set rsCliente = Dbbase.OpenRecordset("select * from alid001 where codigo='" & LcCliente & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not rsCliente.EOF Then
   If (LcTotal + rsCliente!CreditoUtilizado) > rsCliente!LimiteCredito Then
       GlUtilizado = LcTotal + rsCliente!CreditoUtilizado
       GlCredito = rsCliente!LimiteCredito
       LiberacaoCli.Show
       GlLibera = False
       GlEscolha = True
       Do Until Not GlEscolha
          DoEvents
       Loop
   Else
      GlLibera = True
   End If
End If
verificacredito = GlLibera
rsCliente.Close
Dbbase.Close
End Function
Function VerificaEstoquedisponivel(LcCodigo As String, LcQuantidade As Double, LcCom As Double) As Boolean
'Dim Dba1 As Database
'Dim rs As Recordset
Dim LcTotalProduto As Double
Dim LcTotalEstoque As Double
Dim Cest As ControleDb
If GlNaoVerificaEstoque Then
    VerificaEstoquedisponivel = True
    Exit Function
End If
Set Cest = New ControleDb

'Set Dba1 = OpenDatabase(GLBase, False, False)
'Set rs = Dba1.OpenRecordset("select * from alid009 where cod='" & LcCodigo & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Cest.CodProduto = LcCodigo
If Cest.EstoqueGeral < (LcQuantidade * LcCom) Then
    VerificaEstoquedisponivel = False
Else
   VerificaEstoquedisponivel = True
End If
End Function
Private Sub CFOP_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 122 Then Txt(17).SetFocus
If KeyCode = 113 Then SendKeys "%+{B}"
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
FrmExcluiItem.Show (1), Me
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 If KeyCode <> 116 Then Teclas (KeyCode)
  Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"
  End If
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
If Len(Trim(Txt(3).Text)) > 0 Then
   LcQuant = CCur(Txt(3).Text)
Else
   LcQuant = 1
End If
If CCur(Txt(3).Text) > 0 Then
   LcQuant = CCur(Txt(3).Text)
Else
   LcQuant = 1
End If
'MsgBox Txt(3).Text
LcUnit = CCur(valor(0).Text)

LcTotal = CCur(AcertaNumero(CStr(LcQuant), 2)) * CCur(AcertaNumero(CStr(LcUnit), GlDecimais))
valor(1).Text = LcTotal

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

Item.ColWidth(0) = 500
Item.ColWidth(1) = 700
Item.ColWidth(2) = 4600
Item.ColWidth(3) = 500
Item.ColWidth(4) = 1000
Item.ColWidth(5) = 900
Item.ColWidth(6) = 1200
Item.ColWidth(7) = 1200
Item.ColWidth(8) = 800
Item.ColWidth(9) = 0

Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Código"
Item.TextMatrix(0, 2) = "Descrição"
Item.TextMatrix(0, 3) = "CST"
Item.TextMatrix(0, 4) = "Unidade"
Item.TextMatrix(0, 5) = "Quant"
Item.TextMatrix(0, 6) = "Unitário"
Item.TextMatrix(0, 7) = "Total"
Item.TextMatrix(0, 8) = "ICMS"

LcTamanhoGrid = 1
End Function
Function montagrid()
Dim LcAchou As Boolean, a As Integer
On Error Resume Next
'==== Verifica se Foi digitados todos os campos
If Not LcFechaitem Then Exit Function
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
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
End If
If Not Liberado And CCur(limite.Text) <= (CCur(utilizado.Text) + CCur(valor(1).Text)) Then
    If Natureza.Text <> "TRANSFERENCIA" And Natureza.Text <> "DEVOLUCAO" And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
        GlUtilizado = CCur(utilizado.Text) + CCur(valor(1).Text)
        GlCredito = CCur(limite.Text)
        LiberacaoCli.Show
        GlLibera = False
        GlEscolha = True
        Do Until Not GlEscolha
           DoEvents
        Loop
    Else
        GlLibera = True
    End If
    If Not GlLibera Then
       Txt(1).SetFocus
       Exit Function
    Else
      utilizado.Text = CCur(utilizado.Text) + CCur(valor(1).Text)
      Liberado = True
    End If
Else
  utilizado.Text = CCur(utilizado.Text) + CCur(valor(1).Text)
End If
VerificaEstoque (CLng(Txt(4).Text) * CCur(Txt(3).Text))

'===> Verifica se o Item ja Exite
LcAchou = False
err.Number = 0
For a = 0 To UBound(LcMat)
    If err.Number > 0 Then Exit For
    If LcMat(a).CodPro = Txt(1).Text Then
       LcAchou = True
       Exit For
    End If
Next
err.Number = 0
If Not LcAchou Then
    LcItem = LcItem + 1
    ReDim Preserve LcMat(LcTam)
    LcMat(LcTam).Item = Right("00" & LcItem, 2)
    LcMat(LcTam).CodPro = Txt(1).Text
    LcMat(LcTam).Produto = Txt(2).Text
    LcMat(LcTam).Qut = CCur(Txt(3).Text)
    LcMat(LcTam).Und = Unidade.Text
    LcMat(LcTam).com = Txt(4).Text
    LcMat(LcTam).VUnit = CCur(valor(0).Text)
    LcMat(LcTam).Vtotal = CCur(valor(0).Text) * CCur(Txt(3).Text)
    LcMat(LcTam).Venda1 = CCur(Custo.Text)
   ' GlVerificarIcmsDiferenciado = True
   ClienteForaEstado = ClienteEstado
    If GlVerificarIcmsDiferenciado Then
        If Not ClienteForaEstado Then
            If Len(Txt(5).Text) = 0 Then
                LcMat(LcTam).cst = cst.Text
                If Natureza.Text <> "ORG PUBL. EST." Then LcMat(LcTam).icms = icms.Text Else LcMat(LcTam).icms = 0
            Else
                If cst.Text = "060" Then
                    LcMat(LcTam).cst = cst.Text
                    LcMat(LcTam).icms = 0
                Else
                    LcMat(LcTam).cst = cst.Text
                    LcMat(LcTam).icms = Txt(5).Text
                End If
            End If
        Else
            If Natureza.Text <> "ORG PUBL. EST." Then
              If Len(Txt(5).Text) = 0 Then
                  LcMat(LcTam).icms = icms.Text
              Else
                LcMat(LcTam).icms = Txt(5).Text
              End If
            Else
               LcMat(LcTam).icms = 0
            End If
        End If
    Else
       If Len(Txt(5).Text) = 0 Then
          LcMat(LcTam).cst = cst.Text
          If Natureza.Text <> "ORG PUBL. EST." Then LcMat(LcTam).icms = icms.Text Else LcMat(LcTam).icms = 0
       Else
         LcMat(LcTam).icms = Txt(5)
         LcMat(LcTam).cst = "000"
       End If
    End If
    LcMat(LcTam).almox = almox.Text
    LcMat(LcTam).california = CLng(california.Text)
    LcMat(LcTam).santamaria = CLng(santamaria.Text)
    LcMat(LcTam).santamaria1 = CLng(santamaria1.Text)
    LcMat(LcTam).QuanTidadeBaixa = CCur(Txt(3).Text)
    If CLng(Comissao.Text) = 1 Then
       LcMat(LcTam).ItemBaixo = True
    Else
       LcMat(LcTam).ItemBaixo = False
    End If
    LcTam = LcTam + 1
Else
   '===> Vamos Verificar se as Unidades São Iguais
   LcQut = CCur(Txt(3).Text)
   LcUni = CCur(valor(0).Text)
   If (LcMat(a).Und <> Unidade.Text) Or (LcMat(a).com <> Txt(4).Text) Then
      LcMat(a).Und = "UN"
      LcMat(a).VUnit = CDbl(LcMat(a).VUnit) / CDbl(LcMat(a).com)
      LcMat(a).Qut = CDbl(LcMat(a).Qut) * CDbl(LcMat(a).com)
      LcMat(a).com = 1
      LcUni = CCur(valor(0).Text) / (Txt(4).Text)
      LcQut = CDbl(Txt(3).Text) * CDbl(Txt(4).Text)
      LcMat(a).QuanTidadeBaixa = LcMat(a).QuanTidadeBaixa + LcQut '(ccur(txt(3).Text) * CDbl(Txt(4).Text))
      LcMat(a).Qut = LcMat(a).Qut + LcQut
   Else
      LcMat(a).QuanTidadeBaixa = LcMat(a).QuanTidadeBaixa + CCur(Txt(3).Text)
      LcMat(a).Qut = LcMat(a).Qut + CCur(Txt(3).Text)
   End If
   LcMat(a).VUnit = LcUni
   LcMat(a).Vtotal = CCur(LcMat(a).VUnit) * CLng(LcMat(a).Qut)
End If
EscreveGrid
LcLimpaValor = True
For a = 1 To 6
   If a <> 5 Then
      Txt(a).Text = ""
   End If
   valor(a).Text = ""
Next
LcLimpaValor = False
If GlGeraPorItem Then Comissao.Text = 0
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
Txt(2).SetFocus
End Function
Function limpanota()
On Error Resume Next
Dim a As Integer
Liberado = False
GlUtilizado = 0
GlCredito = 0
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
proposta.Text = ""
CodVales.Text = ""
limite.Text = 0
utilizado.Text = 0
Comissao.Text = 0
'CalculaNumeroNota
Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
almox.Text = ""
CmdReimprime.Enabled = False
Txt(6).Text = "EM LANCAMENTO"
Txt(0).Locked = True
Txt(12).SetFocus
End Function
Function verificavale()
On Error Resume Next
Exit Function
Dim RsV As ADODB.Recordset
''abreconexao
'MsgBox txt(8).Text
Set RsV = AbreRecordset("Select * from vales where cliente='" & Txt(8).Text & "' and baixado=0", True)
RsV.Requery
If Not RsV.EOF Then
   lcres = MsgBox("O Cliente " & Txt(9).Text & " Tem Vales." & Chr(13) & "Deseja inclui-los na Nota Fiscal.", vbYesNo, "Vales Encontrados.")
   If lcres = 6 Then
      Load MostraVales
      MostraVales.SetaVales (Txt(8).Text)
      MostraVales.Show , Me
   End If
End If
Set RsV = Nothing
End Function
Function LancaVales()
Dim Rs As ADODB.Recordset
Dim RsI As ADODB.Recordset
Dim Rsp As ADODB.Recordset
On Error Resume Next

Dim LcAchou As Boolean
AbreBase
''abreconexao
Set Rs = AbreRecordset("Select * from Vales where cliente='" & Txt(8).Text & "'" _
& " And baixado=0 and Marca='X'")

LcTam = UBound(LcMat)
If err.Number > 0 Then
   LcItem = 1
   LcTam = 0
   ReDim LcMat(0)
   err.Number = 0
Else
   LcItem = LcMat(LcTam).Item + 1
   If LcItem = 0 Then LcItem = 1
   LcTam = LcTam + 1
End If

Do Until Rs.EOF
    LcAchou = False
    Set RsI = AbreRecordset("Select * from valesprodutos where numnf='" & Rs!numnf & "'", True)
        '===> Vamos Verificar Se Já Existe o Produto no Lançamento
      Do Until RsI.EOF
        LcQut = 0
        For a = 0 To UBound(LcMat)
            If LcMat(a).CodPro = RsI!codProd Then
                LcAchou = True
                Exit For
            End If
        Next
        If Not LcAchou Then
            Set Rsp = AbreRecordset("Select * from produtos where Codigo=" & RsI!codProd, True) '& "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            If Not IsNull(Rsp!cst) Then lccst = Rsp!cst Else lccst = "000"
            If Val(lccst) = 60 Or Val(lccst) = 160 Or Val(lccst) = 260 Then
                icms.Text = "00"
            Else
               If Len(Txt(5).Text) = 0 Then
                  icms.Text = "18"
               Else
                  icms.Text = Txt(5).Text
               End If
           End If

            If Not Rsp.EOF Then
               LcCusto = Rsp!Custo
               LcSt = Rsp!cst & ""
            Else
               LcCusto = 0
               LcSt = ""
            End If
            VerificaEstoque (CLng(RsI!Qtde) * CLng(RsI!QTDUM))
            ReDim Preserve LcMat(LcTam)
            LcMat(LcTam).Item = Right("00" & LcItem, 2)
            LcMat(LcTam).CodPro = RsI!codProd
            LcMat(LcTam).Produto = RsI!Descricao
            LcMat(LcTam).Qut = CLng(RsI!Qtde)
            LcMat(LcTam).Und = RsI!Unimed
            LcMat(LcTam).com = CInt(RsI!QTDUM)
            LcMat(LcTam).VUnit = CCur(RsI!VALUNIT)
            LcMat(LcTam).Vtotal = CCur(RsI!VALUNIT) * CLng(RsI!Qtde)
            LcMat(LcTam).Venda1 = CCur(LcCusto)
            LcMat(LcTam).cst = lccst
            If Len(icms.Text) > 0 Then LcMat(LcTam).icms = icms.Text
            LcMat(LcTam).almox = almox.Text
            LcMat(LcTam).california = CLng(california.Text)
            LcMat(LcTam).santamaria = CLng(santamaria.Text)
            LcMat(LcTam).santamaria1 = CLng(santamaria1.Text)
            LcMat(LcTam).NumeroVale = RsI!numnf
            LcTam = LcTam + 1
            LcItem = LcItem + 1
            
        Else
            '===> Vamos Verificar se as Unidades São Iguais

            If LcMat(a).Und <> RsI!Unimed Then
               LcMat(a).Und = "UN"
               LcMat(a).VUnit = CDbl(LcMat(a).VUnit) / CDbl(LcMat(a).com)
               LcMat(a).Qut = CDbl(LcMat(a).Qut) * CDbl(LcMat(a).com)
               LcMat(a).QuanTidadeBaixa = CDbl(LcMat(a).QuanTidadeBaixa) * CDbl(LcMat(a).com)
               LcMat(a).com = 1
               LcUni = (RsI!VALUNIT) / (RsI!QTDUM)
               LcQut = CDbl(RsI!Qtde) * CDbl(RsI!QTDUM)

               'RsI!qtdum = 1
               'RsI.Update
            Else
              LcQut = RsI!Qtde
            End If
            LcMat(a).Qut = LcMat(a).Qut + LcQut
            LcMat(a).Vtotal = CCur(LcMat(a).VUnit) * CLng(LcMat(a).Qut)
            LcMat(a).NumeroVale = RsI!numnf
        End If
        RsI.MoveNext
        Set Rsp = Nothing
        
    Loop
    Set RsI = Nothing
    Rs.MoveNext
Loop
EscreveGrid
End Function
Function EscreveGrid(Optional EExclusao As Boolean)
Dim b, a As Integer
b = 1
If EExclusao Then
    Item.Rows = 1
End If
For a = 0 To UBound(LcMat)
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item.Rows = b + 1
       Item.TextMatrix(b, 0) = Right("00" & LcMat(a).Item, 2)
       Item.TextMatrix(b, 1) = LcMat(a).CodPro
       Item.TextMatrix(b, 2) = LcMat(a).Produto
       Item.TextMatrix(b, 3) = LcMat(a).cst
       Item.TextMatrix(b, 4) = LcMat(a).Und & " C/" & LcMat(a).com
       Item.TextMatrix(b, 5) = LcMat(a).Qut
       Item.TextMatrix(b, 6) = Format(LcMat(a).VUnit, "Currency")
       Item.TextMatrix(b, 7) = Format(LcMat(a).Vtotal, "Currency")
       Item.TextMatrix(b, 8) = LcMat(a).icms
       b = b + 1
    End If
Next
CalculaIcms
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True

End Function
Function CalculaIcms()
On Error Resume Next
Dim LcBaseCalculo As Double, LcIcms As Double, LcPRodutos As Double, LcNota As Currency
Dim LcItem As String, LcComp As String
Dim LcQuantItemSubs As Double, a As Integer
Dim LcValorTotalSubst   As Double
Dim LcPercIcms          As Double
Dim ValorCst000         As Double
Dim ValorCst060         As Double

'LcItem = 0
LcQuantItemSubs = 0
LcValorTotalSubst = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      If LcMat(a).icms > 0 Then
         LcBaseCalculo = LcBaseCalculo + LcMat(a).Vtotal
         LcIcms = LcIcms + ((LcMat(a).icms / 100) * LcMat(a).Vtotal)
         LcPercIcms = CDbl(LcMat(a).icms)
         ValorCst000 = ValorCst000 + LcMat(a).Vtotal
      Else
         LcValorTotalSubst = LcValorTotalSubst + LcMat(a).Vtotal
         LcQuantItemSubs = LcQuantItemSubs + 1
         If LcQuantItemSubs > 1 Then
            LcItem = LcItem & ", "
         End If
         LcItem = LcItem & Right("00" & CStr(LcMat(a).Item), 2)
         ValorCst060 = ValorCst060 + LcMat(a).Vtotal
      End If
      LcPRodutos = LcPRodutos + LcMat(a).Vtotal
      LcNota = LcNota + LcMat(a).Vtotal
   End If
Next
If Natureza.Text <> "ORG PUBL. EST." Then
    If LcQuantItemSubs > 0 Then
       If LcQuantItemSubs > 1 Then
          LcComp = "Itens " & LcItem & " ICMS ja recolhido na operacao anterior" & Chr(13)
          LcComp = LcComp & "por substituicao tributaria."
       Else
          LcComp = "Item " & LcItem & " ICMS ja recolhido na operacao anterior" & Chr(13)
          LcComp = LcComp & "por substituicao tributaria."
       End If
       If GlImprimeValorCFOPNota Then
          '==> Descreve a obs do CST
          LcComp = "base de calculo:"
          If ValorCst000 > 0 Then LcComp = LcComp & "bc-000 -" & Format(ValorCst000, "Currency")
          If ValorCst060 > 0 Then
             If Len(LcComp) > 0 Then LcComp = LcComp & Chr(10) & Chr(13)
             LcComp = LcComp & "bc-060-" & Format(ValorCst060, "Currency")
          End If
       End If
       If Natureza.Text <> "TRANSFERENCIA" And Natureza.Text <> "DEVOLUCAO" Then Txt(14).Text = LcComp
    Else
        If GlImprimeValorCFOPNota Then
          '==> Descreve a obs do CST
          If ValorCst000 > 0 Then LcComp = "Valor CST 000:" & Format(ValorCst000, "Currency")
          If ValorCst060 > 0 Then
             If Len(LcComp) > 0 Then LcComp = LcComp & Chr(10) & Chr(13)
             LcComp = LcComp & "Valor CST 060:" & Format(ValorCst060, "Currency")
          End If
       End If
       If Natureza.Text <> "TRANSFERENCIA" And Natureza.Text <> "DEVOLUCAO" Then Txt(14).Text = LcComp
     
    End If
End If
'LcPerDesconto = CCur(txt(17).Text) / CCur(txt(16).Text)
LcPercDivicao = LcIcms / LcBaseCalculo
If Len(Trim(Txt(17).Text)) = 0 Or Txt(17).Text = "0" Then
   Txt(13).Text = Format(LcNota - LcValorTotalSubst, "Currency")
   Txt(11).Text = Format((LcNota - LcValorTotalSubst) * LcPercDivicao, "Currency")
   Txt(15).Text = Format(LcPRodutos, "Currency")
   Txt(16).Text = Format(LcNota, "Currency")
Else
   Txt(13).Text = Format(((LcNota - CCur(Txt(17).Text)) - LcValorTotalSubst), "Currency")
   Txt(11).Text = Format(((LcNota - CCur(Txt(17).Text)) - LcValorTotalSubst) * LcPercDivicao, "Currency")
   Txt(15).Text = Format(LcPRodutos, "Currency")
   Txt(16).Text = Format((LcNota - CCur(Txt(17).Text)), "Currency")
End If

If CDbl(Txt(13).Text) < 0 Then Txt(13).Text = 0

'If Natureza.Text = "ORG PUBL. EST." Then
'  If Len(Txt(11).Text) = 0 Then Txt(11).Text = 0
'  Txt(16).Text = Format(AcertaNumero(CStr(CCur(Txt(15).Text) - (CCur(Txt(15).Text) * 0.18)), 2), "Currency")
'End If
End Function
Function RemontaIndice()
On Error Resume Next
LcItem = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      LcItem = LcItem + 1
      LcMat(a).Item = Right("00" & LcItem, 2)
   End If
Next


End Function
Function CarregaCboUnidade()
On Error Resume Next
LcQUn = 0
Dim LcAchou As Integer
Dim RsUnidade As Recordset
Dim LcPrimeiro As String
AbreBase
Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 order By SIMBOLO", dbOpenDynaset, dbSeeChanges, dbOptimistic)
err.Number = 0
Do Until RsUnidade.EOF
  
   ReDim Preserve MtUnidade(LcQUn)
   MtUnidade(LcQUn).Codigo = RsUnidade!cod & ""
   MtUnidade(LcQUn).Descricao = RsUnidade!Nome & ""
   MtUnidade(LcQUn).Simbolo = RsUnidade!Simbolo & ""
   If IsNull(RsUnidade!Quantidade) Then
      MtUnidade(LcQUn).Quantidade = ""
   Else
      MtUnidade(LcQUn).Quantidade = RsUnidade!Quantidade
   End If
   Unidade.AddItem RsUnidade!Simbolo & ""
   
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

valor(0).Text = CLng(Txt(4).Text) * PrecoVendaNormal

minimo.Text = CLng(Txt(4).Text) * CCur(AcertaNumero(CStr(PrecoMimimodeVendaAlterado), GlDecimais))
End Function
Function BuscaProduto(LcTipo As Integer)
On Error Resume Next
Dim Rs As ADODB.Recordset
On Error GoTo errBuscaFor
Dim RsProduto As ADODB.Recordset
Dim LcValorDigitado
Dim LcCodigo As String
If Not LcAlteradoProduto Then Exit Function
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then Exit Function
AbreBase
''abreconexao
'Set RsProduto = Dbbase.OpenRecordset("select * from alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = AbreRecordset("select * from produtos", True)
LcCalculado = True
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODigo=" & Txt(1).Text
         RsProduto.Find LcCriterioCli
         If Not RsProduto.EOF Then
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            Txt(1).Text = RsProduto!Codigo
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
            
                
            'Custo.Text = RsProduto!Custo
            LcAchou = True
            SendKeys "{TAB}"
         Else
            Txt(2).Text = ""
         End If
        ' verificavale
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(2).Text
        If Len(Txt(2).Text) = 0 Then Exit Function
        
        lcchave = Txt(2).Text
        If IsNumeric(lcchave) Then
            LcCriterioCli = "CODigo=" & lcchave
        Else
           LcCriterioCli = "nome='" & UCase(lcchave) & "'"
        End If
        RsProduto.Find LcCriterioCli
        If Not RsProduto.EOF Then
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)   ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            Txt(1).Text = RsProduto!Codigo
            Txt(2).Text = RsProduto!Nome
            Unidade.Text = LcUnidade
            Txt(4).Text = RsProduto!QtdMedida
            LcAchou = 0
            If GlContrato Then
                LcSql = "SELECT ContratoDados.Valor,ContratoDados.CodProduto,ContratoFornecimento.DataI,ContratoFornecimento.DataF"
                LcSql = LcSql & " FROM ContratoDados INNER JOIN ContratoFornecimento ON ContratoDados.CodContrato = ContratoFornecimento.Codigo"
                LcSql = LcSql & " WHERE ContratoFornecimento.Cliente='" & Txt(9).Text & "'"
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
            'SendKeys "{TAB}"
           ' verificavale
        Else
            Txt(2).Text = LcValorDigitado
            GlCriterioSql = "select * From produtos where nome like '" & UCase(Txt(2).Text) & "%'  order by nome"
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
Function VerificaEstoque(LcQuantidade As Long)
On Error Resume Next
LcQtCal = 0
LcQtSta1 = 0
LcQtSta = 0
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then Exit Function
If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Natureza.Text = "DEVOLUCAO" Then Exit Function
LcQuantNesc = LcQuantidade

End Function
Function VerificaCalifornia(LcItem As String) As Integer
On Error Resume Next
Dim RsEstoque As Recordset
'LcQuantNesc
LcSql = "Select * from alid013 where ITEM='" & LcItem & "' and ALMOX='CALIFORNIA'"
Set RsEstoque = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsEstoque.EOF Then
   VerificaCalifornia = False
Else
   If RsEstoque!Estoque > LcQuantNesc Then
      LcQtCal = LcQuantNesc
      LcQuantNesc = 0
      VerificaCalifornia = True
   Else
      LcQtCal = RsEstoque!Estoque
      LcQuantNesc = LcQuantNesc - RsEstoque!Estoque
      VerificaCalifornia = False
   End If
End If
      
End Function
Function VerificaSanta1(LcItem As String)
On Error Resume Next
Dim RsEstoque As Recordset
'LcQuantNesc
LcSql = "Select * from alid013 where ITEM='" & LcItem & "' and ALMOX='SANTA MARIA 2'"
Set RsEstoque = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsEstoque.EOF Then
   VerificaSanta1 = False
Else
   If RsEstoque!Estoque > LcQuantNesc Then
      LcQtSta1 = LcQuantNesc
      LcQuantNesc = 0
      VerificaSanta1 = True
   Else
      LcQtSta1 = RsEstoque!Estoque
      LcQuantNesc = LcQuantNesc - RsEstoque!Estoque
      VerificaSanta1 = False
   End If
End If

End Function
Function VerificaSanta(LcItem As String)
On Error Resume Next
Dim RsEstoque As Recordset
'LcQuantNesc
LcSql = "Select * from alid013 where ITEM='" & LcItem & "' and ALMOX='SANTA MARIA'"
Set RsEstoque = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsEstoque.EOF Then
   VerificaSanta = False
Else
   If RsEstoque!Estoque > LcQuantNesc Then
      LcQtSta = LcQuantNesc
      LcQuantNesc = 0
      VerificaSanta = True
   Else
      LcQtSta = RsEstoque!Estoque
      LcQuantNesc = LcQuantNesc - RsEstoque!Estoque
      VerificaSanta = False
   End If
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
       LcLinha = Left(LcMat(a).Produto & "                                      ", 40)
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
       LcLinha = Left(LcMat(a).Produto & "                                      ", 40)
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
        LcCriterioCli = "nome='" & lcchave & "'"
        RsVendedor.FindFirst LcCriterioCli
        If Not RsVendedor.NoMatch Then
            Txt(7).Text = RsVendedor!Nome
            Txt(10).Text = RsVendedor!Codigo
            
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
            FrmPesquisaFuncionarios.Txt.Text = Txt(7).Text
            GlCriterioSql = "select * From alid200 where nome like '" & UCase(Txt(7).Text) & "*'  order by nome"
            If LcAlteradoFuncionario Then
               FrmPesquisaFuncionarios.Show , Me
               LcAlteradoFuncionario = False
            End If
            'Data(1).SetFocus
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
Function ClienteEstado()
Dim Rs As Recordset
Dim db As Database
Dim StrSql As String
Dim ClienteForaEstado As Boolean
On Error GoTo errCli
Set db = OpenDatabase(GLBase)

StrSql = "Select * from alid001 where CODIGO='" & Txt(8).Text & "'"
Set Rs = db.OpenRecordset(StrSql)
If Not Rs.EOF Then
    If UCase(Rs!Estado) = "MG" Then
        ClienteForaEstado = False
    Else
       ClienteForaEstado = True
    End If
End If
ClienteEstado = ClienteForaEstado
Exit Function
errCli:
'MsgBox err.Description
ClienteEstado = True
End Function
Function BuscaCliente(LcTipo As Integer)
On Error GoTo errBuscaFor
Dim rsCliente       As Recordset
Dim RsVend          As Recordset
Dim LcValorDigitado As String
Dim LcCodigo        As String
Dim LcCredito       As Currency
Dim LcUtilizado     As Currency

If LcAlteradoCliente Then Exit Function
AbreBase

GlLibera = False
LcAlteradoCliente = True
Set rsCliente = Dbbase.OpenRecordset("select * from alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVend = Dbbase.OpenRecordset("select * from alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & Txt(8).Text & "'"
         rsCliente.FindFirst LcCriterioCli
         If Not rsCliente.NoMatch Then
            If UCase(rsCliente!Estado) = "MG" Then
                ClienteForaEstado = False
            Else
                ClienteForaEstado = True
            End If
            Txt(9).Text = rsCliente!razaosoc
            LcDesCidade = rsCliente!razaosoc
            LcBusV = "Nome='" & rsCliente!TelemarketingAtende & "'"
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               Txt(7).Text = RsVend!Codigo & ""
            Else
               Txt(7).Text = rsCliente!TelemarketingAtende & ""
            End If
            BuscaVendendor (2)
            If Not IsEmpty(rsCliente!LimiteCredito) And (Not IsNull(rsCliente!LimiteCredito)) Then LcCredito = rsCliente!LimiteCredito Else LcCredito = 0
            If Not IsEmpty(rsCliente!CreditoUtilizado) And (Not IsNull(rsCliente!CreditoUtilizado)) Then LcUtilizado = rsCliente!CreditoUtilizado Else LcUtilizado = 0
            If LcCredito <= LcUtilizado Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   Txt(9).Text = ""
                   Txt(9).SetFocus
                Else
                   Liberado = True
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
            End If
            'SendKeys "{TAB}"
            If Len(Txt(9).Text) > 0 And LcAlteradoCliente Then
               If VerificaAtraso(Txt(8).Text) Then
                  Txt(9).Text = ""
                  Txt(9).SetFocus
               End If
            End If

         Else
            Txt(9).Text = ""
         End If
         verificavale
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
            If UCase(rsCliente!Estado) = "MG" Then
                ClienteForaEstado = False
            Else
                ClienteForaEstado = True
            End If
            Txt(9).Text = rsCliente!razaosoc
            Txt(8).Text = rsCliente!Codigo
            LcDesCidade = rsCliente!razaosoc
            LcBusV = "Nome='" & rsCliente!TelemarketingAtende & "'"
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               Txt(7).Text = RsVend!Codigo & ""
            Else
               Txt(7).Text = rsCliente!TelemarketingAtende & ""
            End If
            BuscaVendendor (2)
            If Not IsEmpty(rsCliente!LimiteCredito) And (Not IsNull(rsCliente!LimiteCredito)) Then LcCredito = rsCliente!LimiteCredito Else LcCredito = 0
            If Not IsEmpty(rsCliente!CreditoUtilizado) And (Not IsNull(rsCliente!CreditoUtilizado)) Then LcUtilizado = rsCliente!CreditoUtilizado Else LcUtilizado = 0
            If LcCredito <= LcUtilizado Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   Txt(9).Text = ""
                   Txt(9).SetFocus
                Else
                   Liberado = True
                   limite.Text = LcCredito
                   utilizado.Text = LcUtilizado
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
            End If
            'SendKeys "{TAB}"
            If Len(Txt(9).Text) > 0 And LcAlteradoCliente Then
               If VerificaAtraso(Txt(8).Text) Then
                  Txt(9).Text = ""
                  Txt(9).SetFocus
               End If
            End If
            verificavale
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

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode <> 116 Then Teclas (KeyCode)
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"

  End If
End Sub

Private Sub CmdReimprime_Click()
ImprimeAlternativo
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Screen.MousePointer = 11
LcCap = Me.Caption
Me.Caption = "Aguarde, salvando a nota..."
SalvaNota
Me.Caption = LcCap
Screen.MousePointer = 0
MsgBox "Nota salva com sucesso.", 64, "Aviso"

Exit Sub
On Error Resume Next
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
       RsEntrada("DataSaida") = CDate(Txt(12).Text)
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
   valor(a).Text = ""
Next
Item.Rows = 1
Txt(3).Text = "0"
valor(0).Text = "0"
valor(1).Text = "0"
Custo.Text = "0"

Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True

Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Txt(0).SetFocus
End Sub

Private Sub Command1_Click()
On Error Resume Next
FrmPesquisaCliente.Show , Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
FrmPesquisaProdutos.Show , Me
End Sub

Private Sub Command3_Click()
'On Error Resume Next

If Not IsDate(Txt(12).Text) Then
   MsgBox "Informe uma data válida para emissão da nota.", 64, "Aviso"
   Txt(12).SetFocus
   SendKeys "+{home}+{End}"
   Exit Sub
End If
If Len(CFOP.Text) = 0 Then
   MsgBox "Informe o CFOP da nota de saída.", 64, "Aviso"
   CFOP.SetFocus
   SendKeys "+{home}+{End}"
   Exit Sub
End If
If Len(Txt(9).Text) = 0 Then
   MsgBox "Informe o Cliente da nota de saída.", 64, "Aviso"
   Txt(9).SetFocus
   SendKeys "+{home}+{End}"
   Exit Sub
End If
If Natureza.Text <> "COMPLEMENTO ICMS" Then
    If Item.Rows = 1 Then
       MsgBox "É nescessario no minimo um item para a emissão da nota fiscal.", 64, "Aviso"
       Txt(2).SetFocus
       SendKeys "+{home}+{End}"
       Exit Sub
    End If
End If
If Not ValidaEntradaSintegra Then
   Exit Sub
End If
If LiberacaoICms Then
    tam.Text = LcTam
    DadosTransp.Show , Me
    DadosTransp.Imprime.Enabled = True
Else
    Txt(9).SetFocus
End If
End Sub
Function GravaNumeroNota()
Dim Rs As ADODB.Recordset
Dim LcIncluir As Boolean
LcNumeroAtual = 0
''abreconexao
 Set Rs = AbreRecordset("select * from numeronotaAlternativo", True)
If Rs.EOF Then
'   Rs.AddNew
    LcIncluir = True
Else
    LcNumeroAtual = Rs("Numero")
End If
Set Rs = Nothing
LcNumeroAtual = Replace(LcNumeroAtual, "AL", "")
LcNumeroAtual = CDbl(LcNumeroAtual)
If Natureza.Text <> "SERIE D" Then
   If LcIncluir Then
      LcSq = "insert into numeronotaalternativo (Numero) values ('" & Txt(0).Text & "')"
   Else
      If Len(Txt(0).Text) > 0 Then
            If CDbl(Replace(CStr(LcNumeroAtual), "AL", "")) < CDbl(Replace(Txt(0).Text, "AL", "")) Then
               LcSq = "UPDATE numeronotaalternativo SET Numero='" & Replace(Txt(0).Text, "AL", "") & "'"
            End If
      Else
         LcSq = "UPDATE numeronotaalternativo SET Numero='" & LcNumeroAtual & "'"
      End If
   End If
      
  ' Rs!numeronota = txt(0).Text
Else
   If LcIncluir Then
      LcSq = "insert into numeronotaalternativo (numerosd) values ('" & Txt(0).Text & "')"
   Else
      
     LcSq = "UPDATE numeronotaalternativo SET numerosd='" & Txt(0).Text & "'"
      
   End If

  ' Rs!numerosd = txt(0).Text
End If
'MsgBox LcSq

If Len(LcSq) > 0 Then total = ExecutaSql(LcSq)
'Rs.Update
'Set Rs = Nothing
End Function
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode <> 116 Then Teclas (KeyCode)
  Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"
  End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Command4.Caption = "&Pesquisar F7" Then
   FrmPesquisaNota.Show , Me
   Command4.Caption = "&Incluir F7"
   LcPesquisa = True
   Txt(0).Locked = False
Else
   Command4.Caption = "&Pesquisar F7"
   limpanota
   LcPesquisa = False
End If

End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"

End Sub

Private Sub Command5_Click()
On Error Resume Next
'FrmBuscaProposta.Show , Me
FrmFaturaProposta.Show , Me
End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"

End Sub



Private Sub Form_Activate()
On Error Resume Next

Set GlFormA = Me

If Not GlCarregado Then
   Txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
   GlCarregado = True
End If

End Sub
Function BuscaNota(LcNumeroOrc As String)
On Error GoTo ErroBuscaNota
Dim RsOrc As ADODB.Recordset, RsItem As ADODB.Recordset
Dim RsProduto As ADODB.Recordset, rsCliente As Recordset
Dim RsVendedor As Recordset
Dim LcSql6 As String
Dim TemItem As Boolean

Dim LcSql1 As String, LcSql2 As String, LcSql3 As String, LcSql4 As String, LcSql5 As String
LcPesquisa = True
LcSql1 = "Select * from saidas where NUMNF='" & LcNumeroOrc & "'"
LcSql2 = "Select * from saidasdados where NUMNF='" & LcNumeroOrc & "' order by item"
LcSql3 = "Select * from ALid001"
LcSql5 = "Select * from ALid200"
LcSql6 = "Select * from produtos"
LcCap = Me.Caption
LcBuscaNota = True
AbreBase
''abreconexao
Me.Caption = "Aguarde, estabelecendo conexão com o banco de dados mysql."
DoEvents
Set RsOrc = AbreRecordset(LcSql1, True)
Set RsItem = AbreRecordset(LcSql2, True)
Me.Caption = "Aguarde, estabelecendo conexão com o banco de dados Access."
'Set RsOrc = Dbbase.OpenRecordset(LcSql1)
'Set RsItem = Dbbase.OpenRecordset(LcSql2)
Set rsCliente = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVendedor = Dbbase.OpenRecordset(LcSql5, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = AbreRecordset(LcSql6, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)

'==== Preenchendo a Nota

If RsOrc.EOF Then
   Me.Caption = LcCap
   MsgBox "A Nota Fiscal Nº: " & LcNumeroOrc & " Não foi encontrado..."
   Command4.Caption = "Pes&quisa F7"
   Natureza.SetFocus
   
   Exit Function
End If
Me.Caption = "Buscando dados da nota fiscal."
DoEvents
Txt(0).Text = RsOrc!numnf
Txt(12).Text = Format(RsOrc!DtEmis, "dd/mm/yy")
Txt(6).Text = RsOrc!Status
If Not IsNull(RsOrc!Desconto) Then
  If RsOrc!Desconto <> 0 Then
     Txt(17).Text = RsOrc!Desconto
  End If
End If
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
        Natureza.Text = "DEVOLUCAO"
    Case Is = "OR"
        Natureza.Text = "ORG PUBL. EST."
End Select


'If Len(RsOrc!comissao) > 0 Then
'   comissao.Text = RsOrc!comissao
'Else
'   comissao.Text = "1.5"
'End If
If Len(RsOrc!CFOP) > 0 Then
   CFOP.Text = RsOrc!CFOP
Else
   CFOP.Text = "512"
End If
Txt(10).Text = RsOrc!Vendedor & ""
LcCriterio = "Codigo='" & RsOrc!Vendedor & "'"
RsVendedor.FindFirst LcCriterio
If Not RsVendedor.NoMatch Then
   Txt(7).Text = RsVendedor!Nome
Else
  Txt(7).Text = ""
End If
Txt(8).Text = RsOrc!Cliente
LcCriterio = "Codigo='" & RsOrc!Cliente & "'"
rsCliente.FindFirst LcCriterio
If Not rsCliente.NoMatch Then
   Txt(9).Text = rsCliente!razaosoc
   Txt(8).Text = rsCliente!Codigo
End If
Txt(5).Text = RsOrc!icms & ""
Txt(15).Text = RsOrc!ValorProduto
Txt(16).Text = RsOrc!ValorNota
Txt(13).Text = RsOrc!BaseIcms & ""
Txt(11).Text = RsOrc!valorIcms & ""
'If Len(RsOrc!desconto) > 0 Then Txt(13).Text = RsOrc!desconto Else Txt(13).Text = ""
'If Len(RsOrc!TotalDesconto) > 0 Then desconto.Text = RsOrc!TotalDesconto Else desconto.Text = ""
'===== Escreve dados Grid
LcItem = 0
LcTam = 0
'ReDim LcMat(LcTam)
On Error Resume Next
'MsgBox "Passar pelo loop"
Do Until RsItem.EOF
    If err.Number <> 0 Then Exit Do
    Me.Caption = "Buscando dados dos itens da nota fiscal"
    DoEvents
    LcItem = LcItem + 1
    ReDim Preserve LcMat(LcTam)
    If Len(RsItem!Item) > 0 Then LcMat(LcTam).Item = RsItem!Item
      LcCriterio = "CODigo=" & RsItem("codProd")
      RsProduto.MoveFirst
      RsProduto.Find LcCriterio
      If Not RsProduto.EOF Then
            cst.Text = RsProduto!cst
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = RsProduto!Preco / RsProduto!QtdMedida Else PrecoVendaNormal = 0
            ComNormal = RsProduto!QtdMedida
            minimo.Text = RsProduto!MinimoVenda & ""
            If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MinimoVenda / RsProduto!QtdMedida Else PrecoMimimodeVendaAlterado = 0
            If Val(cst.Text) = 60 Or Val(cst.Text) = 16 Or Val(cst.Text) = 26 Then
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
      TemItem = True
      LcMat(LcTam).cst = cst.Text
      LcMat(LcTam).icms = icms.Text
      LcMat(LcTam).CodPro = RsItem("codProd")
      LcMat(LcTam).Qut = RsItem("QTDE")
      LcMat(LcTam).VUnit = RsItem("VALUNIT")
      LcMat(LcTam).Und = RsItem("UNIMED")
      LcMat(LcTam).com = RsItem("QTDUM")
      LcMat(LcTam).Produto = RsItem("descricao")
      '===> a quantidade Baixada deve ser igual a quantidade
      LcMat(LcTam).QuanTidadeBaixa = RsItem("QTDE")
    
      LcMat(LcTam).Vtotal = LcMat(LcTam).Qut * LcMat(LcTam).VUnit
      LcTam = LcTam + 1
      
      RsItem.MoveNext
    LcAchou = True
Loop
'MsgBox "Escrever Grid"
If TemItem Then
    EscreveGrid
End If

 'MsgBox "Seta focus"
 If LcAchou Then
   
    
    FrmSaidaProduto.SetFocus
    Txt(2).SetFocus
 Else
    'txt(10).SetFocus
    CmdSalvar.Visible = True
    Command3.Visible = False
 End If
' MsgBox "Finaliza"
 Me.Caption = "Fechando banco de dados"
 DoEvents
 RsOrc.Close
 RsItem.Close
 rsCliente.Close
 RsVendedor.Close
 LcBuscaNota = False
 If Txt(6).Text = "EMITIDA" Then
    Command3.Enabled = False
    CmdExcluir.Enabled = False
 Else
    Command3.Enabled = True
    CmdExcluir.Enabled = True
 End If
 If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
 Me.Caption = LcCap
 If UCase(Txt(6).Text) = "EMITIDA" Then CmdReimprime.Enabled = True
 Exit Function
 
ErroBuscaNota:
 MsgBox err.Description & err.Number
 'Resume Next
' Resume 0
End Function
Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
Me.Height = 7800
Me.Width = 11970
GlEscolhe = 1
CarregaCboUnidade
Txt(6).Text = "EM LANCAMENTO"
Txt(0).Locked = True
'abreconexao

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlCarregado = False
FrmPrincipal.SetFocus
End Sub

Private Sub Item_DblClick()
On Error Resume Next
Dim LcColuna As Integer
If Item.Rows = 1 Then Exit Sub
LcColuna = Item.Col
GlLiberaSenhaAlteraPr = False
linha = Item.Row
'If UCase(Txt(6).Text) = "EMITIDA" Then Exit Sub
If LcColuna = 5 Then
    FrmExibeSenha = True
    SenhaAteraPreco.Show , Me
    Do While FrmExibeSenha
        DoEvents
    Loop
    If GlLiberaSenhaAlteraPr Then
        LcValor = InputBox("Entre com a Nova Quantidade o Produto", "Produto: " & Item.TextMatrix(linha, 2))
        If Len(LcValor) > 0 Then
            If CDbl(LcValor) > 0 Then
                LcValor = Replace(LcValor, ".", ",")
                LcItemc = Item.TextMatrix(linha, 0)
                For z = 0 To UBound(LcMat)
                    If LcMat(z).Item = LcItemc Then
                        If Not VerificaEstoquedisponivel(CStr(Item.TextMatrix(linha, 1)), CDbl(LcValor), CDbl(LcMat(z).com)) Then
                            MsgBox "O Produto " & Chr(13) & Item.TextMatrix(linha, 0) & Chr(13) & " Não Possui Estoque Suficiente para a Venda.", vbCritical, "Estoque Insuficiente"
                            Exit Sub
                        End If
                        LcMat(z).Qut = CCur(LcValor)
                        LcMat(z).Vtotal = CCur(LcMat(z).VUnit) * CCur(LcValor)
                        EscreveGrid
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End If
If Item.Col = 6 Then
    FrmExibeSenha = True
    SenhaAteraPreco.Show , Me
    Do While FrmExibeSenha
        DoEvents
    Loop
    If GlLiberaSenhaAlteraPr Then
       LcValor = InputBox("Entre com o novo Valor Unitario para o Produto", "Produto: " & Item.TextMatrix(linha, 2))
       
       If Len(LcValor) > 0 Then
          LcValor = Replace(LcValor, ".", ",")
          LcItemc = Item.TextMatrix(linha, 0)
         ' For a = 1 To Item.Rows - 1
             For z = 0 To UBound(LcMat)
                If LcMat(z).Item = LcItemc Then
                   LcMat(z).VUnit = CCur(LcValor)
                   LcMat(z).Vtotal = CCur(LcValor) * CCur(LcMat(z).Qut)
                   EscreveGrid
                   Exit For
                 End If
             Next
         ' Next
       End If
    End If
End If
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{B}"
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Natureza_Click()
If Natureza = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Then
   Txt(5).Text = 0
Else
   Txt(5).Text = ""
End If
If Natureza.Text = "COMPLEMENTO ICMS" Then
   Txt(13).Locked = False
   Txt(11).Locked = False
Else
   Txt(13).Locked = True
   Txt(11).Locked = True

End If
CalculaIcms
End Sub

Private Sub Natureza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{B}"
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{C}"
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
If Natureza.Text = "COMPLEMENTO ICMS" Then
   Command3.Enabled = Len(Txt(11).Text) + Len(Txt(13).Text)
End If
End Sub
Function CalculaDesconto()
On Error Resume Next
If Len(Trim(Txt(17).Text)) = 0 Then Exit Function

If Not IsNumeric(Txt(17).Text) Then
   MsgBox "Digite o Desconto Em Valor Numérico...", 64, "Aviso"
   Txt(17).SetFocus
   Exit Function
End If

'Txt(17).Text = AcertaNumero(Txt(17).Text)
'Txt(16).Text = CCur(Txt(15).Text) - CCur(Txt(17).Text)
'Txt(16).Text = AcertaNumero(Txt(16).Text)

CalculaIcms

End Function
Private Sub txt_GotFocus(Index As Integer)
Dim a   As Integer
On Error Resume Next
If Index = 3 Then
   For a = 0 To LcTam - 1
     If LcMat(a).CodPro = Txt(1).Text And Len(CodVales.Text) = 0 Then
        MsgBox "O Produto " & Txt(2).Text & " Já está selecionado...", vbInformation, "Item Duplicado."
        Txt(2).Text = ""
        Txt(1).Text = ""
        Txt(2).SetFocus
     End If
   Next
End If
LcLimpa = True
If Index = 9 Then
'   If Len(Trim(txt(7).Text)) = 0 Then
'      LcPesquisaCli = False
'      MsgBox "É Necessário Escolher o Vendedor Responsável.", 64, "Aviso"
'      txt(7).SetFocus
'     End If
Else
  LcPesquisaCli = True
End If
If Index = 1 Then
   If Len(Trim(Txt(8).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Cliente para a Nota Fiscal.", 64, "Aviso"
      
   End If
End If

If Index = 9 Then LcAlteradoCliente = False
If Index = 2 Then LcAlteradoProduto = False
If Index = 7 Then LcAlteradoFuncionario = False
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Index = 9 Then LcAlteradoCliente = False
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 122 Then Txt(17).SetFocus
If Index <> 7 And Index <> 5 And Index <> 6 Then
    If KeyCode = 123 Then UltimasComprasCliente.Show , Me
End If


If KeyCode = 113 Then SendKeys "%+{B}"
If KeyCode = 114 Then SendKeys "%+{F}"
If KeyCode = 115 Then SendKeys "%+{E}"
If KeyCode = 118 Then Call Command4_Click
If KeyCode = 121 Then SendKeys "%+{C}"

If KeyCode = 38 Then
   VoltaCampo (KeyCode)
End If
If KeyCode = 117 Then FrmDescicaoProduto.Show , Me
If KeyCode = 116 Then
   If Index = 8 Or Index = 9 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(9).Text)) > 0 Then
           ' FrmPesquisaCliente.txt.Text = txt(9).Text
            GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt(9).Text) & "*'  order by RAZAOSOC"
         Else
            GlCriterioSql = ""
         End If
         
     FrmBuscaCliente.Show , Me
     Me.Txt(7).SetFocus
   Else
      If Index = 1 Or Index = 2 Then 'Exibe Produtos
         GlEscolhe = 2
         If Len(Trim(Txt(2).Text)) > 0 Then
            GlCriterioSql = "select * From produtos where nome like '" & UCase(Txt(2).Text) & "%'  order by nome"
            FrmPesquisaProdutos.Txt.Text = Txt(2).Text
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
If Index = 0 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then Exit Sub
If LcLimpa Then
   If Index <> 12 And Index <> 7 Then Txt(Index).Text = ""
   LcLimpa = False
End If
If Index = 5 Then
   If KeyAscii = 46 Then KeyAscii = 44
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
If Index = 2 Then Call BuscaProduto(2)

If Index = 4 Then calculaunitario

If Index = 3 Then VerificaDisponivel

If Index = 5 Then
   ConferePreco
   Reescreveicms
End If
If Index = 17 Then CalculaDesconto
If Index = 9 Then
   If LcPesquisaCli And Len(Txt(9).Text) > 0 Then BuscaCliente (2): SendKeys "{tab}"
End If
'If Index = 7 Then BuscaVendendor (2)
If Index = 2 Then BuscaProduto (2)
'If Index = 10 And Len(Trim(txt(Index).Text)) <> 0 Then BuscaVendendor (2)

End Sub
Function VerificaDisponivel()

On Error Resume Next

Dim LcSql As String, LcNumeroNota As String
Dim LcCom As Double
Dim LcSaldoVenda As Double
Dim RsNota As Recordset, rsnidade As Recordset
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then Exit Function
If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Natureza.Text = "DEVOLUCAO" Then Exit Function
If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Len(Trim(Txt(3).Text)) = 0 Then Exit Function
Set Estoque = New ControleDb
Estoque.CodProduto = Txt(1).Text
If Len(Txt(4).Text) > 0 Then LcCom = CDbl(Txt(4).Text) Else LcCom = 1
If Len(Txt(3).Text) > 0 Then LcSaldoVenda = CDbl(Txt(3).Text) Else LcSaldoVenda = 0

LcSaldoVenda = LcSaldoVenda * LcCom
If Not GlNaoVerificaEstoque Then
    If LcSaldoVenda > CDbl(Estoque.EstoqueGeral) Then
       MsgBox "Não Exite Quantidade Disponivel em Estoque." & Chr(13) & "A quantidade Atual é :" & Estoque.EstoqueTotalFechado & " e " & Estoque.EstoqueTotalUnitario & " Unidade(s).", 64, "Aviso"
       Txt(3).SetFocus
    End If
End If
'RsNota.Close
'Set RsNota = Nothing
'Dbbase.Close
'Set Dbbase = Nothing

End Function

Function ConferePreco()
On Error Resume Next

Dim LcPreconovo, LcPRecoAntigo As Currency
GlLibera = False
If Len(minimo.Text) = 0 Then minimo.Text = 0



If Len(Trim(valor(0).Text)) = 0 Then
   valor(0).Text = 0
End If
LcPreconovo = CCur(valor(0).Text)
GlEscolha = True
LcPRecoAntigo = CCur(minimo.Text)

If IsEmpty(LcPreconovo) Then LcPreconovo = 0
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
        valor(0) = LcPrecoVelho
        valor(0).SetFocus
     End If
Else
  GlLibera = True
  If Len(Comissao.Text) = 0 Then Comissao.Text = 0
  If CLng(Comissao.Text) <> 1 Then
     Comissao.Text = "1.5"
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
       Exit For
       
    End If
Next
If Not LcAchou Then
   MsgBox "Item Não encontrado...", 48, "Aviso"
Else
  RemontaIndice
  Call EscreveGrid(True)
End If
End Function
Function CalculaNumeroNota()
'On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim RsNota As ADODB.Recordset
Dim RsN As ADODB.Recordset
Dim LcNumero As String
''abreconexao
If Natureza.Text <> "SERIE D" Then
  If Len(Txt(0).Text) = 0 Then
     '===> Vamos Ver se tem o Numero da nota na tb numero de nota
     LcSql = "Select * from numeronotaAlternativo"
     Set RsN = AbreRecordset(LcSql)
     If Not RsN.EOF Then
        If Not IsNull(RsN!Numero) Then
            LcNumero = RsN("numero")
            LcNumero = Replace(LcNumero, "AL", "")
            LcNumeroNota = "AL" & Right("000000" & CStr(Val(LcNumero) + 1), 6)
        Else
             LcSql = "Select * from saidas WHERE NATUREZA<>'SD' order by NUMNF"
             'AbreBase
             Set RsNota = AbreRecordset(LcSql)
             If RsNota.EOF Then
                LcNumeroNota = "AL000001"
             Else
                RsNota.MoveLast
                LcNumeroNota = "AL" & Right("000000" & CStr(Val(RsNota("NUMNF")) + 1), 6)
             End If
                 
             'RsNota.Close
             'Dbbase.Close
             Set RsNota = Nothing
       End If
     Else
       LcSql = "Select * from saidas WHERE NATUREZA<>'SD' order by NUMNF"
             'AbreBase
       Set RsNota = AbreRecordset(LcSql)
       If RsNota.EOF Then
          LcNumeroNota = "AL000001"
       Else
          RsNota.MoveLast
          LcNumeroNota = "AL" & Right("000000" & CStr(Val(RsNota("NUMNF")) + 1), 6)
       End If
                 
             'RsNota.Close
             'Dbbase.Close
       Set RsNota = Nothing
    End If
    Set RsN = Nothing
    Txt(0).Text = LcNumeroNota
    ' Set Dbbase = Nothing
  Else
     Txt(0).Text = Right("000000" & Txt(0).Text, 6)
  End If
Else
  If Len(Txt(0).Text) = 0 Then
     LcSql = "Select * from numeronotaAlternativo"
     Set RsN = AbreRecordset(LcSql)
     If Not RsN.EOF Then
        If Not IsNull(RsN!numerosd) Then
           LcNumeroNota = "SD" & Right("0000" & CStr(Val(Mid(RsN("numerosd"), 3)) + 1), 4)
        Else
           LcSql = "Select * from saidas WHERE NATUREZA='SD' order by NUMNF"
               'AbreBase
           Set RsNota = AbreRecordset(LcSql)
           If RsNota.EOF Then
              LcNumeroNota = "SD0001"
           Else
              RsNota.MoveLast
              LcNumeroNota = "SD" & Right("0000" & CStr(Val(Mid(RsNota("NUMNF"), 3)) + 1), 4)
           End If
       End If
     Else
       LcSql = "Select * from saidas WHERE NATUREZA='SD' order by NUMNF"
               'AbreBase
       Set RsNota = AbreRecordset(LcSql)
       If RsNota.EOF Then
          LcNumeroNota = "SD0001"
       Else
          RsNota.MoveLast
          LcNumeroNota = "SD" & Right("0000" & CStr(Val(Mid(RsNota("NUMNF"), 3)) + 1), 4)
       End If

     End If
     Set RsN = Nothing
     Txt(0).Text = LcNumeroNota

     'RsNota.Close
     'Dbbase.Close
     Set RsNota = Nothing
     'Set Dbbase = Nothing
  Else
     Txt(0).Text = Right("SD0000" & Txt(0).Text, 6)
  End If
End If
End Function

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  SendKeys "{TAB}"
Else
  If KeyCode = 122 Then Txt(17).SetFocus
  If KeyCode = 117 Then FrmDescicaoProduto.Show , Me
  If KeyCode = 123 Then UltimasComprasCliente.Show , Me
  If KeyCode <> 116 Then Teclas (KeyCode)
  If KeyCode = 113 Then SendKeys "%+{B}"
  If KeyCode = 114 Then SendKeys "%+{F}"
  If KeyCode = 115 Then SendKeys "%+{E}"
  If KeyCode = 118 Then Call Command4_Click
  If KeyCode = 121 Then SendKeys "%+{C}"
End If
End Sub

Function GeraComissao()
On Error Resume Next
Dim RsComissao As Recordset
Dim a As Integer
'If GlGeraPorItem Then
   GeraComissaoNova
Exit Function
'End If
LcSql = "Select * from Alid201"
AbreBase
LcSql = "Select * from Alid201 where nf='" & Txt(0).Text & "'"
Set RsComissao = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsComissao.EOF
   If err.Number <> 0 Then Exit Do
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
Dim LcValor As Double
Dim LcResto As Double
LcValor = 0
LcResto = 0
If Len(Txt(17).Text) > 0 Then
    If CDbl(Txt(17).Text) > 0 Then
        LcValor = AcertaNumero(CCur(Txt(17).Text) / (Int(Item.Rows) - 1), 2)
        LcResto = AcertaNumero(LcValor * (Item.Rows - 1), 2)
        LcResto = CCur(Txt(17).Text) - CCur(LcResto)
    End If
End If
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         Dim LcTotal As Double
         RsComissao.AddNew
         RsComissao("Vendedor") = Txt(10).Text
         RsComissao("NF") = Txt(0).Text
         RsComissao("Produto") = LcMat(a).CodPro
         RsComissao("QUANTIDADE") = LcMat(a).Qut
         RsComissao("VALORUNIT") = LcMat(a).VUnit
         If a = LcTam - 1 Then
            LcTotal = (CCur(LcMat(a).Vtotal) - CCur(LcValor)) + CCur(LcResto)
         Else
            LcTotal = CCur(LcMat(a).Vtotal) - CCur(LcValor)
         End If
         RsComissao("VALORTOTAL") = CCur(LcTotal)
         If Comissao.Text = "1" Then Ibaixo = True Else Ibaixo = False
         RsComissao("ITEMBAIXO") = Ibaixo
         If Len(Comissao.Text) = 0 Then Comissao.Text = "0"
         If Ibaixo Then
            RsComissao("COMISSAO") = (1 / 100) * (LcTotal - (LcPercDesc * LcMat(a).Vtotal))
         Else
            RsComissao("COMISSAO") = (1.5 / 100) * (LcTotal - (LcPercDesc * LcTotal))
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
Function GeraComissaoNova()
Dim StrSql              As String
Dim LcPercDesc          As Double
Dim PercentualComissao  As Double
Dim a                   As Integer
Dim ValorComissao       As Double
Dim ValorDesconto       As Double

'==> Deleta as Comissoes LAncadas para esta nf
If Len(Txt(0).Text) > 0 Then
   StrSql = "Delete from Alid201 where nf='" & Txt(0).Text & "'"
   afetados = ExecutaSql(StrSql)
End If
'==> Calcula o valor do percentual de desconto dado na nota
If Len(Txt(17).Text) > 0 Then
   LcPercDesc = CDbl(Txt(17).Text) / CDbl(Txt(15).Text)
Else
   LcPercDesc = 0
End If
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       '==> Veirfica se é a baixo do minimo
       If E_Baixo_Minimo(CLng(LcMat(a).CodPro), CDbl(LcMat(a).VUnit), CDbl(LcMat(a).com)) Then
          PercentualComissao = 1
       Else
          PercentualComissao = 1.5
       End If
    
      '==> Calcula a comissao
      ValorComissao = (PercentualComissao / 100) * (LcMat(a).Vtotal - (LcPercDesc * LcMat(a).Vtotal))
      ValorDesconto = LcPercDesc * LcMat(a).Vtotal
      '==>Gera a String Para Gravar a Comissao
      StrSql = "Insert into Alid201 (Vendedor,NF,Produto,QUANTIDADE,VALORUNIT,VALORTOTAL,ITEMBAIXO," & _
               "COMISSAO,DATAVENDA,CLIENTE,NomeCliente,NomeVendedor,DescontUnitario) Values('" & _
               Txt(10).Text & "','" & _
               Txt(0).Text & "','" & _
               LcMat(a).CodPro & "'," & _
               Replace(LcMat(a).Qut, ",", ".") & "," & _
               Replace(LcMat(a).VUnit, ",", ".") & "," & _
               Replace(LcMat(a).Vtotal, ",", ".") & "," & _
               LcMat(a).ItemBaixo & "," & _
               Replace(ValorComissao, ",", ".") & ",#" & _
               Format(Txt(12).Text, "mm/dd/yy") & "#," & _
               Txt(8).Text & ",'" & _
               Replace(Txt(9).Text, "'", "''") & "','" & _
               Replace(Txt(7).Text, "'", "''") & "'," & _
               Replace(ValorDesconto, ",", ".") & ")"

       'Debug.Print StrSql
       afetados = ExecutaSql(StrSql)
    End If
Next

End Function

Function E_Baixo_Minimo(codigoproduto As Long, ValorUnitarioVenda As Double, QuantidadeUnidade As Double) As Boolean
On Error Resume Next
Dim RsProduto As ADODB.Recordset
Dim StrSql As String
Dim LcPRecoAntigo As Currency
Dim MinimoVenda As Currency
Dim QuantUnidade As Double


'==> abre a tabela produtos para saber o minimo
StrSql = "Select * from PRodutos where codigo=" & codigoproduto
Set RsProduto = AbreRecordset(StrSql, True)
If Not RsProduto.EOF Then
   LcPRecoAntigo = RsProduto!Preco
   MinimoVenda = RsProduto!MinimoVenda
   QuantUnidade = RsProduto!QtdMedida
Else
   E_Baixo_Minimo = False
   Exit Function
End If

'==>Calcula o Valor unitario da unidade
If QuantUnidade <> QuantidadeUnidade Then
   MinimoVenda = (MinimoVenda / QuantUnidade)
   ValorUnitarioVenda = ValorUnitarioVenda / QuantidadeUnidade
End If

If ValorUnitarioVenda < MinimoVenda Then
   E_Baixo_Minimo = True
Else
   E_Baixo_Minimo = False
End If
Set RsProduto = Nothing
End Function
Function processanota()
Dim Lcq As Double
Dim LCN As String
conexaoAdo.BeginTrans
 CalculaNumeroNota
 ProcessaSintegra
 
   If Natureza.Text = "SERIE D" Then CalculaNumeroNota
   If SalvaNota() Then
        If Natureza.Text <> "TRANSFERENCIA" And Natureza.Text <> "DEVOLUCAO" Then
           Lcq = DadosTransp.Quantidade
           If Atualizacaixa(DadosTransp.Quantidade) Then
              GeraComissao
           Else
              GoTo desfaz
           End If
       End If
       ImprimeAlternativo
       'If Natureza.Text <> "SERIE D" Then
       '   lcobt = ObtemNomeBanco
       '   If UCase(lcobt) = "LIDIS" Then
       '   '  ImprimeNota
       '   End If
       '   If UCase(lcobt) = "BELCLEAN" Then
       '    '     ImprimeNotaBel
       '   End If
       '   If UCase(lcobt) = "CASANOVA" Then
       '     '    ImprimeNotaBel
       '   End If
       '   If UCase(lcobt) = "DHEL" Then
       '      '   ImprimeNotaDhel
       '   End If
       '   If UCase(lcobt) = "" Then
       '      '   ImprimeNotaBel
       '   End If
       '  ' If GlServidorImpressora Then LogNotaFiscal
       'End If
       LCN = Txt(0).Text
       limpanota
   Else
      GoTo desfaz
  End If
'GravaNumeroNota

conexaoAdo.CommitTrans
'conexaoAdo.RollbackTrans
MsgBox "Lançamento Salvo com numero: " & LCN
Exit Function

desfaz:
conexaoAdo.RollbackTrans
MsgBox "Ocorreu um erro no lançamento da nota, Nota não Lançada.", 64, "Aviso"

End Function
Sub ImprimeAlternativo()
Dim StrSql As String
Dim Rs As ADODB.Recordset

LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11

StrSql = "SELECT saidas.NUMNF, saidas.DTEMIS, saidas.CLIENTE, saidasdados.ITEM, saidasdados.QTDE, saidasdados.VALUNIT, saidasdados.UNIMED, saidasdados.QTDUM, saidasdados.descricao, saidasdados.codProd, saidas.valorproduto, saidas.ValorNota, saidas.Vendedor, saidas.formapag, saidas.dias, saidas.DESCONTO, saidas.localentrega, saidas.oc, saidas.enderecoentrega, saidas.DESCONTO, saidas.OBS02, saidas.OBS03, saidas.OBS04"
StrSql = StrSql & " FROM saidas INNER JOIN saidasdados ON saidas.NUMNF = saidasdados.NUMNF"
StrSql = StrSql & " WHERE (((saidas.NUMNF)='" & Txt(0).Text & "'));"

'StrSql = "SELECT saidas.NUMNF, saidas.DTEMIS, ALID001.RAZAOSOC, ALID001.END, ALID001.BAIRRO, ALID005.NOME, ALID001.Numero, ALID001.ESTADO, ALID001.CEP, ALID001.FONE1, ALID001.CGC, ALID001.INSCEST, alid200.NOME, saidasdados.ITEM, saidasdados.codProd, saidasdados.descricao, saidasdados.QTDE, saidasdados.UNIMED, saidasdados.QTDUM, saidasdados.VALUNIT, saidas.enderecoentrega, saidas.oc, saidas.formapag, saidas.dias, saidas.valorproduto, saidas.DESCONTO, saidas.ValorNota, saidas.OBS02, saidas.OBS03, saidas.OBS04 "
'StrSql = StrSql & " FROM ((((ALID001 INNER JOIN ALID005 ON ALID001.CIDADE = ALID005.COD) INNER JOIN saidas ON ALID001.CODIGO = saidas.CLIENTE) INNER JOIN alid200 ON saidas.Vendedor = alid200.CODIGO) INNER JOIN saidasdados ON saidas.NUMNF = saidasdados.NUMNF) INNER JOIN produtos ON saidasdados.codProd = produtos.codigo"
'StrSql = StrSql & " WHERE (((saidas.NUMNF)=" & Txt(0).Text & "));"
Debug.Print StrSql

Set Rs = AbreRecordset(StrSql, True)
'MsgBox Rs!ValorNota
'MsgBox DEscricaoErro

'MsgBox Rs.RecordCount
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
     setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Screen.MousePointer = vbDefault
Me.Caption = LcCap

End Sub
Sub setaformula()
Dim a As Integer
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset
Dim rsCliente As Recordset
Dim RsVendedor As Recordset
Dim RsCidade As Recordset
Dim RsOpcao As Recordset
Dim LcValor As Double
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
Dim lctitulo As String
Dim StrSql As String
Dim bb     As Database
Dim LcNomeCliente As String
Dim LcEnderecoCliente As String
Dim LcNumeroCliente As String
Dim LcBairroCliente As String
Dim LcCidadeCliente As String
Dim LcEstadoCliente As String
Dim LcCepCliente As String
Dim FoneCliente As String
Dim CnpjCliente As String
Dim InscricaoCliente As String
Dim NomeVendedor As String

Set db = OpenDatabase(GLBase)
Set RsEmpresa = db.OpenRecordset("Select * from EMPRESA")

If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Numero & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & ""
   Lcemail = RsEmpresa!Email & ""
   LcCelular = RsEmpresa!Celular & ""
End If
Set RsEmpresa = Nothing
StrSql = "Select * from Alid001 where codigo='" & Txt(8).Text & "'"
Set rsCliente = db.OpenRecordset(StrSql)
'==> Abrindo vendedor
StrSql = "Select * from Alid200 where codigo='" & Txt(10).Text & "'"
Set RsVendedor = db.OpenRecordset(StrSql)
'==> Busca os dados do cliente
If Not rsCliente.EOF Then
     LcNomeCliente = rsCliente!razaosoc & ""
     LcEnderecoCliente = rsCliente!End & ""
     LcNumeroCliente = rsCliente!Numero & ""
     LcBairroCliente = rsCliente!Bairro & ""
     LcEstadoCliente = rsCliente!Estado & ""
     LcCepCliente = rsCliente!Cep & ""
     FoneCliente = rsCliente!Fone1 & ""
     CnpjCliente = rsCliente!CGC & ""
     InscricaoCliente = rsCliente!INSCEST & ""
     CnpjCliente = Replace(CnpjCliente, ".", "")
     CnpjCliente = Replace(CnpjCliente, ",", "")
     CnpjCliente = Replace(CnpjCliente, "-", "")
     CnpjCliente = Replace(CnpjCliente, "/", "")
     CnpjCliente = Replace(CnpjCliente, "\", "")
     If Len(CnpjCliente) = 0 Then
        CnpjCliente = rsCliente!cpf & ""
     Else
       CnpjCliente = rsCliente!CGC & ""
     End If
     '==> Busca a Cidade
     StrSql = "Select * from alid005 where cod='" & rsCliente!Cidade & "'"
     Set RsCidade = db.OpenRecordset(StrSql)
     If Not RsCidade.EOF Then
        LcCidadeCliente = RsCidade!Nome & ""
     End If
     RsCidade.Close
End If
rsCliente.Close
'==> Busca o nome do vendedor
If Not RsVendedor.EOF Then
   NomeVendedor = RsVendedor!Nome & ""
End If
RsVendedor.Close

   LcFormaPAg = DadosTransp.TipoMonetario.Text
   LCDias = DadosTransp.Dias(0).Text
   LCEndEntrega = DadosTransp.Txt(10).Text
   LcOc = DadosTransp.Txt(11).Text
   Obs = DadosTransp.Txt(7).Text
   obs1 = DadosTransp.Txt(8).Text
   obs2 = DadosTransp.Txt(9).Text
   Desc = Me.Txt(17).Text


'lctitulo = "Relatorio de Produtos (Posicao Balanco)"
With Rel
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("obs1") Then .FormulaFields(a).Text = "totext('" & obs1 & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("obs2") Then .FormulaFields(a).Text = "totext('" & obs2 & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("obs") Then .FormulaFields(a).Text = "totext('" & Obs & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("desc") Then .FormulaFields(a).Text = "totext('" & Desc & "')"

        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("TipoM") Then .FormulaFields(a).Text = "totext('" & LcFormaPAg & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Dia") Then .FormulaFields(a).Text = "totext('" & LCDias & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EndEntrega") Then .FormulaFields(a).Text = "totext('" & LCEndEntrega & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Oc") Then .FormulaFields(a).Text = "totext('" & LcOc & "')"


        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("email") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Celular") Then .FormulaFields(a).Text = "totext('" & LcCelular & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("NomeCliente") Then .FormulaFields(a).Text = "totext('" & LcNomeCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EnderecoCliente") Then .FormulaFields(a).Text = "totext('" & LcEnderecoCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("BairroCliente") Then .FormulaFields(a).Text = "totext('" & LcBairroCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EstadoCliente") Then .FormulaFields(a).Text = "totext('" & LcEstadoCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("cepCliente") Then .FormulaFields(a).Text = "totext('" & LcCepCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("NumeroCliente") Then .FormulaFields(a).Text = "totext('" & LcNumeroCliente & "')"
      
         If UCase(.FormulaFields(a).FormulaFieldName) = UCase("FoneCliente") Then .FormulaFields(a).Text = "totext('" & FoneCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("CnpjCliente") Then .FormulaFields(a).Text = "totext('" & CnpjCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("InscricaoCliente") Then .FormulaFields(a).Text = "totext('" & InscricaoCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Nome_cidade") Then .FormulaFields(a).Text = "totext('" & LcCidadeCliente & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("NomeVendedor") Then .FormulaFields(a).Text = "totext('" & NomeVendedor & "')"
    Next
End With
End Sub
Function SalvaNota() As Boolean
On Error GoTo ErrSalva
Dim RsNotaFiscal As ADODB.Recordset, RsItens As ADODB.Recordset
Dim rsCliente As Recordset, RsProduto As Recordset
Dim RsEstoque As Recordset, RsG As Recordset
Dim RsCheck As ADODB.Recordset
Dim RsProposta As Recordset
Dim LcCom, a As Long
Dim LcSaldoUnit As Double
Dim LcComentario As String
Dim LcSql2 As String
Dim LcTraMdb As Boolean
Dim LcSq As String
LcSql1 = "Select * from saidas"
'LcSql2 = "Select * from saidasdados"
LcSql2 = "Select * from saidasdados where NUMNF = '" & Txt(0).Text & "'"
LcSql3 = "Select * from Alid001 where codigo='" & Txt(8).Text & "'"
LcSql4 = "Select * from produtos"
LcSql5 = "Select * From alid013"
LcComentario = "Abrindo Banco de Dados"
Set Estoque = New ControleDb
'conexaoAdo.BeginTrans

AbreBase
LcComentario = "Abrindo a Tabela :"
'Area.BeginTrans
If Natureza.Text = "ORG PUBL. EST." Then CalculaIcms
LcComentario = LcComentario & "Saidas"
'Set RsNotaFiscal = Dbbase.OpenRecordset(LcSql1)
LcComentario = LcComentario & "saidasdados"
'Set RsItens = Dbbase.OpenRecordset(LcSql2)
LcComentario = LcComentario & "Alid001"
Set rsCliente = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcComentario = LcComentario & "Produtos"
'Set RsProduto = Dbbase.OpenRecordset(LcSql4, dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcComentario = "Buscando dados da Proposta"
If Len(proposta.Text) > 0 Then
   LcSql6 = "Select * From proposta where NUMNF='" & proposta.Text & "'"
   Set RsProposta = Dbbase.OpenRecordset(LcSql6, dbOpenDynaset, dbSeeChanges, dbOptimistic)
   If Not RsProposta.EOF Then
      RsProposta.Edit
      RsProposta("faturado") = True
      RsProposta.Update
   End If
   RsProposta.Close
End If
'==== Verificando se é inclusao ou alteracao da nota fiscao


'==== Grava Os dados da Nota Fiscal
LcComentario = "Gravando dados na Nota de Saida"
'Lccr = "NUMNF='" & Txt(0).Text & "'"
'RsNotaFiscal.FindFirst Lccr
'If Not RsNotaFiscal.NoMatch Then
'   RsNotaFiscal.Edit
'Else
'   RsNotaFiscal.AddNew
'End If

Select Case Natureza.Text
    Case Is = "VENDAS A VISTA"
         LcNatureza = "VV"
    Case Is = "VENDAS A PRAZO"
         LcNatureza = "VP"
    Case Is = "EMPENHO"
         LcNatureza = "EM"
     Case Is = "TRANSFERENCIA"
         LcNatureza = "TR"
     Case Is = "DEVOLUCAO"
        LcNatureza = "DE"
     Case Is = "SERIE D"
        LcNatureza = "SD"
     Case Is = "ORG PUBL. EST."
        LcNatureza = "OR"
Case Is = "RESSARCIMENTO DO ICMS S.T"
        LcNatureza = "RI"
End Select

If Len(Txt(17).Text) = 0 Then
   Txt(17).Text = 0
End If
''abreconexao
'conexaoAdo.BeginTrans
'la = conexaoAdo.BeginTrans
'Set RsNotaFiscal = AbreRecordset("select * from saidas where numnf='" & Txt(0).Text & "'", RsNotaFiscal)

'LcNovo = RsNotaFiscal.EOF
'Set RsNotaFiscal = Nothing

If Len(Txt(0).Text) > 0 Then '===> É Alteracao
    LcSq = "Delete from saidas where numnf='" & Txt(0).Text & "'"
    totalExcl = ExecutaSql(LcSq)
    LcSq = "Delete from saidasDados where Numnf='" & Txt(0).Text & "'"
    totalExcl = ExecutaSql(LcSq)
Else
   'RsNotaFiscal.AddNew
End If
LcValorProduto = CStr(CDbl(Txt(15).Text))
LcValorProduto = Replace(LcValorProduto, ",", ".")

LcValorTotal = CStr(CDbl(Txt(16).Text))
LcValorTotal = Replace(LcValorTotal, ",", ".")

If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
LcValorDesconto = CStr(CDbl(Txt(17).Text))
LcValorDesconto = Replace(LcValorDesconto, ",", ".")

LcSq = "INSERT INTO saidas (Numnf,natureza,Dtemis,Status,CFOP,Cliente,Transp,TipoTrans,PlacaTrans,"
LcSq = LcSq & "Uftrans,CGCCPFTRAN,endtrans,munictrans,ufmunic,inscest,ValorProduto,ValorNota"

If Not IsNull(DadosTransp.Vencimento(0).Text) Then
   If IsDate(DadosTransp.Vencimento(0).Text) Then
       LcSq = LcSq & ",vencimento1"
    End If
End If
If Not IsNull(DadosTransp.Vencimento(1).Text) Then
   If IsDate(DadosTransp.Vencimento(1).Text) Then
       LcSq = LcSq & ",vencimento2"
   End If
End If
If Not IsNull(DadosTransp.Vencimento(2).Text) Then
   If IsDate(DadosTransp.Vencimento(2).Text) Then
       LcSq = LcSq & ",vencimento3"
   End If
End If
'MsgBox LcSq
LcSq = LcSq & ",Vendedor,icms,desconto,BaseIcms,ValorIcms,EnderecoEntrega,Oc) values "
'MsgBox LcSq

'LcSq = LcSq & " values "

LcSq = LcSq & "('" & Txt(0).Text & "','" & LcNatureza & "','" & Format(Txt(12).Text, "yyyy-mm-dd") & "',"
LcSq = LcSq & "'EMITIDA','" & CFOP.Text & "','" & Txt(8).Text & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(0).Text) & "','"
LcSq = LcSq & Mid(DadosTransp.Tipo.Text, 1, 1) & "','" & Trim(DadosTransp.Placa.Text) & "','"
LcSq = LcSq & Estoque.RetiraCaracter(DadosTransp.Txt(1).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(2).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(3).Text)
LcSq = LcSq & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(4).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(5).Text) & "','" & Estoque.RetiraCaracter(DadosTransp.Txt(6).Text)
LcSq = LcSq & "'," & LcValorProduto & "," & LcValorTotal

'MsgBox LcSq
If Not IsNull(DadosTransp.Vencimento(0).Text) Then
   If IsDate(DadosTransp.Vencimento(0).Text) Then
       LcSq = LcSq & ",'" & Format(DadosTransp.Vencimento(0).Text, "yyyy-mm-dd") & "'"
    End If
End If
If Not IsNull(DadosTransp.Vencimento(1).Text) Then
   If IsDate(DadosTransp.Vencimento(1).Text) Then
       LcSq = LcSq & ",'" & Format(DadosTransp.Vencimento(1).Text, "yyyy-mm-dd") & "'"
   End If
End If
If Not IsNull(DadosTransp.Vencimento(2).Text) Then
   If IsDate(DadosTransp.Vencimento(2).Text) Then
       LcSq = LcSq & ",'" & Format(DadosTransp.Vencimento(2).Text, "yyyy-mm-dd") & "'"
   End If
End If
If Len(Txt(13).Text) = 0 Then Txt(13).Text = 0
If Len(Txt(11).Text) = 0 Then Txt(11).Text = 0

LcSq = LcSq & ",'" & Right("00000" & Txt(10).Text, 5) & "','" & Txt(5).Text & "'," & LcValorDesconto & ","
LcSq = LcSq & Replace(Replace(Replace(Replace(Replace(Txt(13).Text, ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & ","
LcSq = LcSq & Replace(Replace(Replace(Replace(Replace(Txt(11).Text, ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & ","
LcSq = LcSq & "'" & DadosTransp.Txt(10).Text & "','" & DadosTransp.Txt(11).Text & "')"
LcRegistrosAfetados = AcessoAdo.ExecutaSql(LcSq)
'If LcRegistrosAfetados < 1 Then
 '   err.Raise vbObjectError + 513, "Nâo foi efetuada a Atualização.", "Atualização do item " & LcMat(a).CodPro & "Não foi Realizada."
'    GoTo ErrSalva
'End If

LcCap1 = DadosTransp.Caption
If Natureza.Text <> "COMPLEMENTO ICMS" Then
For a = 0 To UBound(LcMat)
        DadosTransp.Caption = "Processando o Item: " & LcMat(a).Produto
        DoEvents
        If Len(LcMat(a).com) > 0 Then
           LcCom = LcMat(a).com
        Else
          LcCom = 1
        End If
        If Len(Trim(LcMat(a).CodPro)) > 0 Then
             LcComentario = "Gravando Ficha de Estoque"
             '===> Setando as informações da classe
             Estoque.CodProduto = LcMat(a).CodPro
             Estoque.CodClien_forn = Txt(8).Text
             Estoque.NF = Txt(0).Text
             
             LcComentario = "Gravando poroduto " & LcMat(a).CodPro & " em Item da nota"
             
             LcQuantidade = CDbl(LcMat(a).Qut)
             LcQuantidade = Replace(LcQuantidade, ",", ".")
             LcUnitario = CDbl(LcMat(a).VUnit)
             LcUnitario = Replace(LcUnitario, ",", ".")
             LcCom = CDbl(LcMat(a).com)
             LcCom = Replace(LcCom, ",", ".")
             LcSq = "Insert into saidasdados (Numnf,item,codprod,qtde,valunit,unimed,qtdum,descricao,icms,cst) values "
             LcSq = LcSq & "('" & Txt(0).Text & "','" & Right("00" & a + 1, 2) & "','"
             LcSq = LcSq & LcMat(a).CodPro & "'," & LcQuantidade & "," & LcUnitario & ",'"
             LcSq = LcSq & LcMat(a).Und & "'," & LcCom & ",'" & Estoque.RetiraCaracter(LcMat(a).Produto) & "',"
             LcSq = LcSq & Replace(LcMat(a).icms, ",", ".") & ",'"
             LcSq = LcSq & Replace(LcMat(a).cst, ",", ".") & "')"
             'MsgBox LcSq
             LcCodLancamento = ""
             LcRegistrosAfetados = ExecutaSql(LcSq)
             'MsgBox DEscricaoErro
             If LcRegistrosAfetados < 1 Then
                err.Raise vbObjectError + 513, "Nâo foi efetuada a Atualização.", "Atualização do item " & LcMat(a).CodPro & "Não foi Realizada."
                GoTo ErrSalva
             Else
                '===> Busca o Numero da Atualização para atualizar o Historico
                Set RsCheck = AbreRecordset("Select * from saidasdados where NumNf='" & Txt(0).Text & "' and CodProd=" & LcMat(a).CodPro, True)
                If Not RsCheck.EOF Then
                   LcCodLancamento = RsCheck("Codigo")
                Else
                  LcCodLancamento = ""
                End If
                Set RsCheck = Nothing
            End If
    
             
             If Natureza.Text <> "DEVOLUCAO" And Natureza.Text <> "TRANSFERENCIA" Then
                LcQSanta = 0
                LcQSanta1 = 0
                LcqCanifornia = 0
                LcQUnSanta = 0
                LcQUnSantas = 0
                LcComentario = "Atualiza o saldo em Estoque "
                'Call BaixaEstoque(LcMat(a).CodPro, CDbl(LcMat(a).Qut), CDbl(LcMat(a).com), LcMat(a).Und)
                LcQSanta = (Estoque.Santa1Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa1Unitario
                LcQSanta1 = (Estoque.Santa2Fechado * Estoque.QuantidadeDaUnidade) + Estoque.Santa2Unitario
                LcqCanifornia = (Estoque.QuantidadeCaliforniaFechado * Estoque.QuantidadeDaUnidade) + Estoque.QuantidadeCaliforniaUnitario
                
                'Call BaixaPorNota(LcMat(a).CodPro, CDbl(LcMat(a).QuanTidadeBaixa), CDbl(LcMat(a).Com), LcMat(a).Und, CStr(LcMat(a).Com))
                If LcMat(a).QuanTidadeBaixa > 0 Then
                   If Not Estoque.BaixaEstoque(CDbl(LcMat(a).QuanTidadeBaixa), CDbl(LcMat(a).VUnit), LcMat(a).Und, CDbl(LcMat(a).com), True) Then
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
    Next
End If
'==== Atualiza Dados Cliente
LcComentario = "Atualiza Lançamentos do Clientes"



If UCase(Natureza.Text) <> UCase("Complemento icms") And UCase(Natureza.Text) <> "DEVOLUCAO" And UCase(Natureza.Text) <> "TRANSFERENCIA" And UCase(Natureza.Text) <> "VENDAS A VISTA" And UCase(Natureza.Text) <> "RESSARCIMENTO DO ICMS S.T" Then
    'LcCriterioPes = "codigo='" & txt(8).Text & "'"
    'RsCliente.FindFirst LcCriterioPes
    If Not rsCliente.EOF Then
       rsCliente.Edit
       rsCliente("ULTCOMPRA") = CDate(Txt(12).Text)
       'If Natureza.Text <> "DEVOLUCAO" And Natureza.Text <> "TRANSFERENCIA" Then
          rsCliente("CreditoUtilizado") = rsCliente("CreditoUtilizado") + CCur(Txt(16).Text)
       'End If
       rsCliente.Update
    End If
End If
BaixaVale
LancaVendaEstado
GravaNumeroNota
SalvaNota = True
Saida:
rsCliente.Close
Set RsComissao = Nothing
Set rsCliente = Nothing
Exit Function
ErrSalva:
'conexaoAdo.RollbackTrans
SalvaNota = False
MsgBox err.Description & err.Number
Resume 0
LcComentario = "Função Salva Nota " & Txt(0).Text & " " & LcComentario
'LcNune = CStr(err.Number)
Call logErro(err.Number, err.Description, LcComentario)
LcResp = MsgBox("Ocorreu o Seguinte Erro Salvando a nota:" & Chr(13) & err.Description & Chr(13) & "Deseja Tentar Novamente ?", vbExclamation + vbYesNo, "Nº do Erro :" & err.Description)
If LcResp = 6 Then Resume 0
If LcResp = 7 Then
   'conexaoAdo.RollbackTrans
'   conexaoAdo.RollbackTrans
   'excluilancamentos
   SalvaNota = False
   'MsgBox "Todos os Lançamentos para esta Nota Foram Cancelados.", vbInformation, "Operação Cancelada"
   'On Error Resume Next
   Exit Function
End If

End Function
Function excluilancamentos()
On Error Resume Next
Dim RsProposta As Recordset
Dim RsLogNota As Recordset
If Len(proposta.Text) > 0 Then
        AbreBase
        LcSql6 = "Select * From proposta where NUMNF='" & proposta.Text & "'"
        Set RsProposta = Dbbase.OpenRecordset(LcSql6, dbOpenDynaset, dbSeeChanges, dbOptimistic)
        If Not RsProposta.EOF Then
           RsProposta.Edit
           RsProposta("faturado") = False
           RsProposta.Update
        End If
        RsProposta.Close
End If
Set RsLogNota = Dbbase.OpenRecordset("select * from LogImpressaoNota where nf='" & Txt(0).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

err.Number = 0
Do Until RsLogNota.EOF
   If err.Number > 0 Then Exit Do
   RsLogNota.Delete
   RsLogNota.MoveNext
Loop
RsLogNota.Close
RsProposta.Close

End Function
Function RefazEstoque()
Dim Rs      As Recordset
Dim Rsp     As ADODB.Recordset
Dim RsPro   As Recordset
Exit Function
Dim LcTotalVendido As Double
Dim LcTotalEstoque As Double
AbreBase
For a = 0 To LcTam - 1
   Set Rs = Dbbase.OpenRecordset("select * from subproposta where codProd='" & LcMat(a).CodPro & "' and faturado=false", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Set Rsp = AbreRecordset("select * from produtos where codigo=" & LcMat(a).CodPro) '& "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Set RsPro = Dbbase.OpenRecordset("Select * from proposta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Do Until Rs.EOF
      If Len(Trim(LcMat(a).CodPro)) > 0 Then
         If Not Rs!faturado Then
            LcTotalVendido = Rs!Qtde * Rs!QTDUM
            LcTotalEstoque = (Rsp!QtdMedida * Rsp!QuantEstoque) + Rsp!QuantUnidade
            If LcTotalVendido > LcTotalEstoque Then
               LcPesq = "NUMNF='" & Rs!numnf & "'"
               RsPro.FindFirst LcPesq
               If Not RsPro.NoMatch Then
                  RsPro.Edit
                  RsPro!Pendente = True
                  RsPro.Update
               End If
               Rs.Edit
               Rs!bloqueado = True
               If Rs!tipoliberacao <> 2 Then
                  If Rs!tipoliberacao <> 3 Then
                     
                     Rs!tipoliberacao = Rs!tipoliberacao + 2
                     
                     
                  End If
               End If
               Rs.Update
            End If
         End If
      End If
      Rs.MoveNext
   Loop
   Rs.Close
   Rsp.Close
Next
Dbbase.Close
Set Rs = Nothing
Set Rsp = Nothing
Set Dbbase = Nothing
      
End Function
Function AtualizaEstoque(LcUnidade, LcGalpao, LcProduto, LcUnidadeProduto As String, LcQuanti, Lcemb, LcEmbProduto As Double)
On Error Resume Next
Dim db As Database
Dim RsUnidade As Recordset, RsGalpao As Recordset
Dim LcCrigal, LcCriUnid As String
Dim LcSaldoCaixaAtual, LcsaldounitarioAtual As Double

LcCrigal = "Select * from alid013 where ITEM='" & LcProduto & "' and ALMOX='" & LcGalpao & "'"
LcCriUnid = "Select * from alid004 where simbolo='" & LcUnidade & "'"
Set db = OpenDatabase(GLBase)
Set RsUnidade = db.OpenRecordset(LcCriUnid, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsGalpao = db.OpenRecordset(LcCrigal, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'========> 1- Verifica se a unidade escolhida é a mesma do cadastro.
If RsUnidade!cod = LcUnidadeProduto Then
   If RsGalpao!Estoque >= LcQuanti Then
      RsGalpao.Edit
      RsGalpao("estoque") = RsGalpao("estoque") - LcSaldoCaixa
      TotalUnitario = 0
      RsGalpao.Update
      LcSaldoCaixa = 0
      LcSaldoUnit = 0
      LcProximo = False
   Else
      If RsGalpao("Estoque") > 0 Then
         LcSaldoCaixa = LcSaldoCaixa - RsGalpao("Estoque")
         RsGalpao.Edit
         RsGalpao("estoque") = 0
         RsGalpao.Update
         LcProximo = True
      Else
         LcProximo = True
      End If
   End If
Else
  '=======> A Quantidade vendida não é a Basica.
  '=======> Calcular a quantidade a dar baixa
  LcsaldounitarioAtual = LcQuanti * Lcemb
  'Verifica se o galpao Possui Quantidade disponivel, mesmo abrindo caixas
  If ((LcEmbProduto * RsGalpao("Estoque")) + RsGalpao("quantUnidade")) >= LcsaldounitarioAtual Then
  
  '===> Verifica se a quantidade  ultrapassa a Quantidade da unidade
  LcValorNovo = LcsaldounitarioAtual / LcEmbProduto
  '===> Se For Maior que Zero, Ultrapassou o Valor da Unidade
  If LcValorNovo >= 1 Then
     LcNovoInteiro = Int(LcValorNovo)
     LcResto = LcsaldounitarioAtual - (LcNovoInteiro * LcEmbProduto)
     RsGalpao.Edit
     RsGalpao("estoque") = RsGalpao("estoque") - LcNovoInteiro
     RsGalpao.Update
     LcsaldounitarioAtual = LcResto
  End If
  '===> Verifica se a unidade do galpao é suficiente
      If RsGalpao("quantUnidade") >= LcsaldounitarioAtual Then
            RsGalpao.Edit
            RsGalpao("quantUnidade") = RsGalpao("quantUnidade") - LcsaldounitarioAtual
            RsGalpao.Update
            LcSaldoCaixa = 0
            LcSaldoUnit = 0
            LcProximo = False
      Else
         '====> Vai abrir uma caixa
           RsGalpao.Edit
           LcTotalcaixa = RsGalpao!Estoque - 1
           RsGalpao("estoque") = LcTotalcaixa
           RsGalpao("quantUnidade") = (RsGalpao("quantUnidade") + LcEmbProduto) - LcsaldounitarioAtual
           RsGalpao.Update
           LcProximo = False
      End If
  Else
      LcProximo = True
      'LcSaldoCaixa = 0
      If (LcEmbProduto * RsGalpao("Estoque")) + RsGalpao("quantUnidade") > 0 Then
         LcSaldoUnit = Lcemb - (LcEmbProduto * RsGalpao("Estoque")) + RsGalpao("quantUnidade")
      Else
        LcSaldoUnit = LcQuanti
      End If
      LcSaldoCaixa = LcSaldoUnit
        
      RsGalpao.Edit
      RsGalpao("estoque") = 0
      RsGalpao("quantUnidade") = 0
      RsGalpao.Update
  End If
  
End If
RsGalpao.Close
RsUnidade.Close
db.Close
Set RsGalpao = Nothing
Set RsUnidade = Nothing
Set db = Nothing
End Function
Function Atualizacaixa(LcNumeroContas As Integer) As Boolean
Dim RsContasReceber As ADODB.Recordset, RsCaixa As Recordset
Dim RsTipoMonetario As Recordset
Dim LcSql1 As String
Dim LcSql As String
Dim LcNovo As Boolean
Dim a As Integer
Dim LcValor As Double
Dim LcValorPago As Double
On Error GoTo erroatualiza
LcSql1 = "Select * from Alid015"
LcSql2 = "Select * from Alid016"
LcSql3 = "Select * from Alid008"

AbreBase
''abreconexao

'Set RsContasReceber = AbreRecordset(LcSql1, RsContasReceber)
Set RsCaixa = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsTipoMonetario = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)

'===> Excluindo registros da conta a receber
If Len(Txt(0).Text) > 0 Then
    LcSql = "select * from Alid015 where NF like '" & Txt(0).Text & "%'"
    Set RsContasReceber = AbreRecordset(LcSql)
    LcNovo = True
Else
    LcNovo = True
End If

Select Case Natureza.Text
    Case Is = "VENDAS A VISTA"
         If GlVistaSaida Then
            LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMonetario = RsTipoMonetario("TPMONET")
            Else
               LcTipoMonetario = "03"
            End If
            LcValor = CCur(Txt(16).Text)
            'LcValor = Replace(LcValor, ",", ".")
            
            LcValorPago = CCur(Txt(16).Text)
           ' LcValorPago = Replace(LcValorPago, ",", ".")
            If LcNovo Then RsContasReceber.AddNew
            
            RsContasReceber!NF = Txt(0).Text
            RsContasReceber!Cliente = Txt(8).Text
            RsContasReceber!TPMONET = LcTipoMonetario
            RsContasReceber!valor = LcValor
            RsContasReceber!Data = Format(Txt(12).Text, "dd/mm/yy")
            RsContasReceber!DTVENC = Format(Txt(12).Text, "dd/mm/yy")
            'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
            RsContasReceber!VALPAGO = 0
            RsContasReceber!tipord = "R"
            RsContasReceber!acrescimo = 0
            RsContasReceber.Update
            

          End If
          
         If GlCaixaSaida Then
           
            LcCriterio = "NF='" & Txt(0).Text & "'"
            RsCaixa.FindFirst LcCriterio
            If Not RsCaixa.NoMatch Then
               RsCaixa.Edit
            Else
               RsCaixa.AddNew
            End If
            RsCaixa("NF") = Txt(0).Text
            RsCaixa("RECDESP") = "R"
            RsCaixa("CLICRED") = Txt(8).Text
            LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
            If Not RsTipoMonetario.NoMatch Then
               RsCaixa("TPMONET") = RsTipoMonetario("TPMONET")
            End If
            RsCaixa("VALOR") = CCur(Txt(16).Text)
            RsCaixa("DATA") = CDate(Txt(12).Text)
            RsCaixa.Update
          End If
    Case Is = "VENDAS A PRAZO"
         If GlFaturaSaida Then
            For a = 1 To LcNumeroContas
                LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
                RsTipoMonetario.FindFirst LcCriterioPes
                If Not RsTipoMonetario.NoMatch Then
                   LcTipoMonetario = RsTipoMonetario("TPMONET")
                Else
                   LcTipoMonetario = "03"
                End If
                LcValor = CCur(DadosTransp.valor.Text)
                'LcValor = Replace(LcValor, ",", ".")
                
                LcValorPago = CCur(Txt(16).Text)
                'LcValorPago = Replace(LcValorPago, ",", ".")
                If LcNovo Then RsContasReceber.AddNew
            
                RsContasReceber!NF = Txt(0).Text & "/" & Right("00" & a, 2)
                RsContasReceber!Cliente = Txt(8).Text
                RsContasReceber!TPMONET = LcTipoMonetario
                RsContasReceber!valor = LcValor
                RsContasReceber!Data = Format(Txt(12).Text, "dd/mm/yy")
                Select Case a
                    Case Is = 1
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(0).Text)
                    Case Is = 2
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(1).Text)
                    Case Is = 3
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(2).Text)
                End Select
                'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
                RsContasReceber!VALPAGO = 0
                RsContasReceber!tipord = "R"
                RsContasReceber!acrescimo = 0
                RsContasReceber.Update
            Next
          End If
    Case Is = "ORG PUBL. EST."
          If GlVistaSaida Then
         
            LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMonetario = RsTipoMonetario("TPMONET")
            Else
               LcTipoMonetario = "03"
            End If
            LcValor = CCur(Txt(16).Text)
            'LcValor = Replace(LcValor, ",", ".")
            
            LcValorPago = CCur(Txt(16).Text)
            'LcValorPago = Replace(LcValorPago, ",", ".")
            If LcNovo Then RsContasReceber.AddNew
            
            RsContasReceber!NF = Txt(0).Text
            RsContasReceber!Cliente = Txt(8).Text
            RsContasReceber!TPMONET = LcTipoMonetario
            RsContasReceber!valor = LcValor
            RsContasReceber!Data = Format(Txt(12).Text, "dd/mm/yy")
            RsContasReceber!DTVENC = Format(CDate(Txt(12).Text) + 30, "dd/mm/yy")
            'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
            RsContasReceber!VALPAGO = 0
            RsContasReceber!tipord = "R"
            RsContasReceber!acrescimo = 0
            RsContasReceber.Update
          End If
         
    Case Is = "EMPENHO"
         If GlVistaSaida Then
         
            LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMonetario = RsTipoMonetario("TPMONET")
            Else
               LcTipoMonetario = ""
            End If
            LcValor = CCur(Txt(16).Text)
            'LcValor = Replace(LcValor, ",", ".")
            
            LcValorPago = CCur(Txt(16).Text)
            'LcValorPago = Replace(LcValorPago, ",", ".")
            If LcNovo Then RsContasReceber.AddNew
            
            RsContasReceber!NF = Txt(0).Text
            RsContasReceber!Cliente = Txt(8).Text
            RsContasReceber!TPMONET = LcTipoMonetario
            RsContasReceber!valor = LcValor
            RsContasReceber!Data = Format(Txt(12).Text, "dd/mm/yy")
            RsContasReceber!DTVENC = Format(CDate(Txt(12).Text) + 30, "dd/mm/yy")
            'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
            RsContasReceber!VALPAGO = 0
            RsContasReceber!tipord = "R"
            RsContasReceber!acrescimo = 0
            RsContasReceber.Update
          End If
        Case Else
           If GlFaturaSaida Then
            For a = 1 To LcNumeroContas
                LcCriterioPes = "XTPMONET='" & DadosTransp.TipoMonetario.Text & "'"
                RsTipoMonetario.FindFirst LcCriterioPes
                If Not RsTipoMonetario.NoMatch Then
                   LcTipoMonetario = RsTipoMonetario("TPMONET")
                Else
                   LcTipoMonetario = "03"
                End If
                LcValor = CCur(DadosTransp.valor.Text)
                'LcValor = Replace(LcValor, ",", ".")
                
                LcValorPago = CCur(Txt(16).Text)
                'LcValorPago = Replace(LcValorPago, ",", ".")
                If LcNovo Then RsContasReceber.AddNew
            
                RsContasReceber!NF = Txt(0).Text & "/" & Right("00" & a, 2)
                RsContasReceber!Cliente = Txt(8).Text
                RsContasReceber!TPMONET = LcTipoMonetario
                RsContasReceber!valor = LcValor
                RsContasReceber!Data = Format(Txt(12).Text, "dd/mm/yy")
                Select Case a
                    Case Is = 1
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(0).Text)
                    Case Is = 2
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(1).Text)
                    Case Is = 3
                         RsContasReceber("DTVENC") = CDate(DadosTransp.Vencimento(2).Text)
                End Select
                'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
                If RsContasReceber("DTVENC") = Date Then
                    RsContasReceber!VALPAGO = LcValor
                    RsContasReceber!tipord = "R"
                    RsContasReceber!acrescimo = 0
                    RsContasReceber!DTPAGTO = Format(Txt(12).Text, "dd/mm/yy")
                Else
                    RsContasReceber!VALPAGO = 0
                    RsContasReceber!tipord = "R"
                    RsContasReceber!acrescimo = 0
                End If
                RsContasReceber.Update
            Next
          End If
End Select

Atualizacaixa = True
RsContasReceber.Close
RsCaixa.Close
RsTipoMonetario.Close

Set RsContasReceber = Nothing
Set RsCaixa = Nothing
Set RsTipoMonetario = Nothing

Exit Function

erroatualiza:
Resume Next
Atualizacaixa = False
Exit Function

End Function
Function ImprimeNotaDhel()
On Error Resume Next

Dim Item, Descricao, cst, icms, Unidade As String
Dim quant, a As Long
Dim Unitario, total As Currency
CalculaNumeroNota
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1
LcMargem = ""
For a = 1 To Glmargemnota
    LcMargem = LcMargem & " "
Next
LcSalto = Val(GLSaltoLinhaNota)
LcNota = GlPortaNota
LcBoleto = GlPortaBoleto

If IsNull(LcNota) Then LcNota = "LPT1"
If GlboletoA4 = 0 Then If IsNull(LcBoleto) Then LcBoleto = "LPT2"

LcImpressoes = 0
LcLinha = ""

LcLinha = ""
LcQuantiImpressao = 0
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(27) + Chr(48)
    LcQuantiImpressao = LcQuantiImpressao + 1
CabecalhoNotaDhel
If Natureza.Text <> "ORG PUBL. EST." Or Natureza.Text <> "EMPENHO" Or Natureza.Text <> "TRANSFERENCIA" Then
  If DadosTransp.TipoMonetario.Text = "BOLETO" Then
    If GlboletoA4 = 0 Then
        If GlBoletoCEF Then
            ImprimeBoletoCaixa (Val(DadosTransp.Quantidade.Text))
        Else
            ImprimeBoleto (Val(DadosTransp.Quantidade.Text))
        End If
    Else
       GeraBoletoA4 (Val(DadosTransp.Quantidade.Text))
    End If
  End If
End If
If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item = LcMat(a).CodPro
       If Natureza.Text <> "EMPENHO" And Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
          Descricao = LcMat(a).Produto & " " & LcMat(a).Und & " C/" & LcMat(a).com
       Else
          Descricao = LcMat(a).Produto
       End If
       cst = LcMat(a).cst
       If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
            If Val(cst) = 60 Or Val(cst) = 160 Or Val(cst) = 260 Then
               Descricao = Descricao
            Else
               Descricao = Descricao
            End If
       End If
       icms = LcMat(a).icms
       Unidade = LcMat(a).Und
       quant = LcMat(a).Qut
       Unitario = LcMat(a).VUnit
       total = LcMat(a).Vtotal
       Call ImprimeItemDhel(Item, Descricao, cst, icms, Unidade, quant, Unitario, total)
      
       LcImpressoes = LcImpressoes + 1
       If LcImpressoes > 28 Then
          FechaImpressaoDhel (LcImpressoes)
          CabecalhoNotaDhel
          Txt(0).Text = Right("000000" & CStr(Val(Txt(0).Text) + 1), 6)
          SalvaNotaNumeroAlterado
          LcImpressoes = 0
       End If
    End If
Next
FechaImpressaoDhel (LcImpressoes)
'Close #FnunNota
GeraSpool
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
  If GlboletoA4 = 0 Then
     GeraSpoolBoleto
  End If
End If
'LogNotaFiscal
ImprimeGalpao
End Function

Function ImprimeNotaBel()
On Error Resume Next

Dim Item, Descricao, cst, icms, Unidade As String
Dim quant, a As Long
Dim Unitario, total As Currency
CalculaNumeroNota
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1
LcMargem = ""
For a = 1 To Glmargemnota
    LcMargem = LcMargem & " "
Next
LcQuantiImpressao = 0
LcSalto = Val(GLSaltoLinhaNota)
LcNota = GlPortaNota
LcBoleto = GlPortaBoleto

If IsNull(LcNota) Then LcNota = "LPT1"
If GlboletoA4 = 0 Then If IsNull(LcBoleto) Then LcBoleto = "LPT2"

LcImpressoes = 0
LcLinha = ""

LcLinha = ""
CabecalhoNotabel
If Natureza.Text <> "ORG PUBL. EST." Or Natureza.Text <> "EMPENHO" Or Natureza.Text <> "TRANSFERENCIA" Then
  If DadosTransp.TipoMonetario.Text = "BOLETO" Then
    If GlboletoA4 = 0 Then
        If GlBoletoCEF Then
            ImprimeBoletoCaixa (Val(DadosTransp.Quantidade.Text))
        Else
            ImprimeBoleto (Val(DadosTransp.Quantidade.Text))
        End If
    Else
       GeraBoletoA4 (Val(DadosTransp.Quantidade.Text))
    End If
  End If
End If
If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item = LcMat(a).CodPro
       If Natureza.Text <> "EMPENHO" And Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
          Descricao = LcMat(a).Produto & " " & LcMat(a).Und & " C/" & LcMat(a).com
       Else
          Descricao = LcMat(a).Produto
       End If
       cst = LcMat(a).cst
       If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
            If Val(cst) = 60 Or Val(cst) = 160 Or Val(cst) = 260 Then
               Descricao = Descricao
            Else
               Descricao = Descricao
            End If
       End If
       icms = LcMat(a).icms
       Unidade = LcMat(a).Und
       quant = LcMat(a).Qut
       Unitario = LcMat(a).VUnit
       total = LcMat(a).Vtotal
       Call ImprimeItemBel(Item, Descricao, cst, icms, Unidade, quant, Unitario, total)
      
       LcImpressoes = LcImpressoes + 1
       If LcImpressoes > 28 Then
          FechaImpressaoBel (LcImpressoes)
          CabecalhoNotabel
          Txt(0).Text = Right("000000" & CStr(Val(Txt(0).Text) + 1), 6)
          SalvaNotaNumeroAlterado
          LcImpressoes = 0
       End If
    End If
Next
FechaImpressaoBel (LcImpressoes)
'Close #FnunNota
GeraSpool
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
  If GlboletoA4 = 0 Then
     GeraSpoolBoleto
  End If
End If
'LogNotaFiscal
ImprimeGalpao
End Function
Function ImprimeNota()
On Error Resume Next

Dim Item, Descricao, cst, icms, Unidade As String
Dim quant, a As Long
Dim Unitario, total As Currency
CalculaNumeroNota
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcEspaco = ""
FnunNota = FreeFile
FnunBoleto = FreeFile + 1
LcMargem = ""
For a = 1 To Glmargemnota
    LcMargem = LcMargem & " "
Next
LcSalto = Val(GLSaltoLinhaNota)
LcNota = GlPortaNota
LcBoleto = GlPortaBoleto

If IsNull(LcNota) Then LcNota = "LPT1"
If GlboletoA4 = 0 Then If IsNull(LcBoleto) Then LcBoleto = "LPT2"

LcImpressoes = 0
'Open LcNota For Output Access Write As #FnunNota 'Abre Porta Nf
'Salta linhas no inicio da nota
LcLinha = ""

LcLinha = ""
cabecalhonota
If Natureza.Text <> "ORG PUBL. EST." Or Natureza.Text <> "EMPENHO" Or Natureza.Text <> "TRANSFERENCIA" Then
  If DadosTransp.TipoMonetario.Text = "BOLETO" Then
    If GlboletoA4 = 0 Then
        If GlBoletoCEF Then
            ImprimeBoletoCaixa (Val(DadosTransp.Quantidade.Text))
        Else
            ImprimeBoleto (Val(DadosTransp.Quantidade.Text))
        End If
    Else
       GeraBoletoA4 (Val(DadosTransp.Quantidade.Text))
    End If
  End If
End If
If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item = LcMat(a).Item
       If Natureza.Text <> "EMPENHO" And Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
          Descricao = LcMat(a).Produto & " " & LcMat(a).Und & " C/" & LcMat(a).com
       Else
          Descricao = LcMat(a).Produto
       End If
       cst = LcMat(a).cst
       If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
            If Val(cst) = 60 Or Val(cst) = 160 Or Val(cst) = 260 Then
               Descricao = Descricao & " - 5.403"
            Else
               Descricao = Descricao & " - 5.102"
            End If
       End If
       icms = LcMat(a).icms
       Unidade = LcMat(a).Und
       quant = LcMat(a).Qut
       Unitario = LcMat(a).VUnit
       total = LcMat(a).Vtotal
       Call imprimeitem(Item, Descricao, cst, icms, Unidade, quant, Unitario, total)
      
       LcImpressoes = LcImpressoes + 1
       If LcImpressoes > 20 Then
          FechaImpressao (LcImpressoes)
          cabecalhonota
          Txt(0).Text = Right("000000" & CStr(Val(Txt(0).Text) + 1), 6)
          SalvaNotaNumeroAlterado
          LcImpressoes = 0
       End If
    End If
Next
FechaImpressao (LcImpressoes)
'Close #FnunNota
GeraSpool
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
  If GlboletoA4 = 0 Then
     GeraSpoolBoleto
  End If
End If
'LogNotaFiscal
ImprimeGalpao

End Function
Function GeraSpool()
Dim RsLogNota As Recordset, RsImpressora As Recordset
Dim a As Integer
On Error GoTo Errimpr
AbreBase
Set RsLogNota = Dbbase.OpenRecordset("select * from LogImpressaoNota", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsImpressora = Dbbase.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'")
Set RsImpressora = Dbbase.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

For a = 0 To LcQuantiImpressao
   RsLogNota.AddNew
   RsLogNota!impressora = GlPortaNota
   If Len(Trim(GlNomeMaquina)) > 0 Then
      RsLogNota!Maquina = RsImpressora!Maquina
   Else
      RsLogNota!Maquina = "Maquina Local"
   End If
   RsLogNota!NF = Txt(0).Text
   RsLogNota!dados = MtImpressao(a)
'   MsgBox MtImpressao(a)
   RsLogNota.Update
Next
LcQuantiImpressao = 0
ReDim Preserve MtImpressao(0)
RsLogNota.Close
Set RsLogNota = Nothing
LogNotaFiscal
Exit Function
Errimpr:
'MsgBox err.Number & err.Description
End Function
Function GeraBoletoA4(LcQuantidade As Integer)
Dim ClBoleto As ControlaBoleto
Dim ClAcesso As New AcessoAdo.Acessos
Dim RsCedente As ADODB.Recordset
Dim StrSql As String
Dim a As Integer
Dim LcMargemBo As String
Dim Protesto    As String
Dim protesto1   As String
Dim LocalAr As String
Set ClBoleto = New ControlaBoleto

If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Natureza.Text = "EMPENHO" Then Exit Function
If Natureza.Text = "ORG PUBL. EST." Then Exit Function
LocalAr = App.EXEName & ".ini"
ClBoleto.NomeProjeto = LocalAr
StrSql = "Select * from contasacado"

Set RsCedente = AbreRecordset(StrSql, True)
Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsCedente.EOF Then
   MsgBox "A conta do cedente não foi cadastrada.", 64, "Aviso"
   Exit Function
End If

LcJuros = "             JUROS DE 5 % AO MES"
LcPag = "             ATE A DATA DO VENCIMENTO PAGAR EM QUALQUER BANCO / QUALQUER AGENCIA"
Protesto = "             NAO RECEBER APOS 4 (QUATRO) DIAS UTEIS DO VENCIMENTO."
protesto1 = "             SUJEITO A PROTESTO"

ClBoleto.BoletodoBanco = "Itau"
ClBoleto.Agencia = RsCedente!Agencia
ClBoleto.Cedente = RsCedente!NomeSacado
ClBoleto.Conta = RsCedente!Conta
ClBoleto.ContaDigito = RsCedente!DigitoConta
ClBoleto.Carteira = RsCedente!Carteira
ClBoleto.BairroSacado = RsClientes!Bairro & ""
ClBoleto.CepSacado = RsClientes!Cep & ""

If Not RsCidade.EOF Then
   ClBoleto.CidadeSacado = RsCidade!Nome & ""
Else
  ClBoleto.CidadeSacado = ""
End If
ClBoleto.CnpjSacado = "CNPJ/CPF:" & RsClientes!CGC & ""
ClBoleto.DataDocumento = Format(Date, "dd/mm/yy")
ClBoleto.EnderecoSacado = RsClientes!End & ""
ClBoleto.EspecieDoc = "DP"
ClBoleto.EstadoSacado = RsClientes!Estado & ""
ClBoleto.IncricaoSacado = RsClientes!INSCEST & ""
ClBoleto.Instrucao1 = LcJuros
ClBoleto.Instrucao2 = Protesto
ClBoleto.Instrucao3 = protesto1
ClBoleto.Instrucao4 = ""
ClBoleto.NomeSacado = RsClientes!razaosoc & ""
ClBoleto.RgSacado = ""
ClBoleto.Especie = "R$"
ClBoleto.Aceite = "N"

For a = 1 To LcQuantidade
    ClBoleto.NumeroDocumento = Txt(0).Text & "/" & Right("00" & a, 2)
    ClBoleto.Vencimento = DadosTransp.Vencimento(a - 1).Text
    Select Case a
        Case Is = 1
            ClBoleto.valor = Format(LcValor1, "Standard")
        Case Is = 2
            ClBoleto.valor = Format(LcValor2, "Standard")
        Case Is = 3
            ClBoleto.valor = Format(LcValor3, "Standard")
    End Select
    ClBoleto.GeraBoleto imprimir
    
Next

Set ClBoleto = Nothing





End Function
Function GeraSpoolBoleto()
Dim RsLogBoleto As Recordset
Dim a, b As Integer

If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "EMPENHO" And Natureza.Text <> "TRANSFERENCIA" And Natureza.Text <> "DEVOLUCAO" Then
  If DadosTransp.TipoMonetario.Text <> "BOLETO" Then Exit Function
Else
  Exit Function
End If
AbreBase
Set RsLogBoleto = Dbbase.OpenRecordset("select * from logboleto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsImpressora = Dbbase.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

'Set RsImpressora = Dbbase.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'")
LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto - 1

For b = 0 To LcQuantiImpressaoBoleto
   RsLogBoleto.AddNew
   RsLogBoleto!impressora = GlPortaBoleto
   If Len(Trim(GlNomeMaquina)) > 0 Then
      RsLogBoleto!Maquina = RsImpressora!Maquina
   Else
      RsLogBoleto!Maquina = "Maquina Local"
   End If
   RsLogBoleto!NF = Txt(0).Text
   RsLogBoleto!dados = MtBoleto(b)
   RsLogBoleto.Update
Next
LcQuantiImpressao = 0
ReDim Preserve MtBoleto(0)
RsLogBoleto.Close
Set RsLogBoleto = Nothing
End Function
Function SalvaNotaNumeroAlterado()
On Error Resume Next
Dim RsNotaFiscal As Recordset, RsItens As Recordset
Dim rsCliente As Recordset, RsProduto As Recordset
Dim RsEstoque As Recordset
Dim LcCom As Long
LcSql1 = "Select * from saidas"

AbreBase
Set RsNotaFiscal = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)

'==== Grava Os dados da Nota Fiscal
Lccr = "NUMNF='" & Txt(0).Text & "'"
RsNotaFiscal.FindFirst Lccr
If Not RsNotaFiscal.NoMatch Then
   RsNotaFiscal.Edit
Else
   RsNotaFiscal.AddNew
End If
RsNotaFiscal("NUMNF") = Txt(0).Text
RsNotaFiscal("DTEMIS") = CDate(Txt(12).Text)
LcCap1 = DadosTransp.Caption

Select Case Natureza.Text
    Case Is = "VENDAS A VISTA"
         RsNotaFiscal("NATUREZA") = "VV"
    Case Is = "VENDAS A PRAZO"
         RsNotaFiscal("NATUREZA") = "VP"
    Case Is = "ORG PUBL. EST."
         RsNotaFiscal("NATUREZA") = "OR"
    Case Is = "EMPENHO"
         RsNotaFiscal("NATUREZA") = "EM"
     Case Is = "TRANSFERENCIA"
         RsNotaFiscal("NATUREZA") = "TR"
     Case Is = "DEVOLUCAO"
        RsNotaFiscal("NATUREZA") = "DE"
End Select
RsNotaFiscal("CFOP") = CFOP.Text
RsNotaFiscal("CLIENTE") = Txt(8).Text
RsNotaFiscal("TRANSP") = DadosTransp.Txt(0).Text
RsNotaFiscal("TIPOTRANS") = Mid(DadosTransp.Tipo.Text, 1, 1)
RsNotaFiscal("PLACATRANS") = DadosTransp.Placa.Text
RsNotaFiscal("UFTRANS") = DadosTransp.Txt(1).Text
RsNotaFiscal("CGCCPFTRAN") = DadosTransp.Txt(2).Text
RsNotaFiscal("ENDTRANS") = DadosTransp.Txt(3).Text
RsNotaFiscal("MUNICTRANS") = DadosTransp.Txt(4).Text
RsNotaFiscal("UFMUNIC") = DadosTransp.Txt(5).Text
RsNotaFiscal("INSCEST") = DadosTransp.Txt(6).Text
RsNotaFiscal("OBS02") = DadosTransp.Txt(7).Text
RsNotaFiscal("OBS03") = DadosTransp.Txt(8).Text
RsNotaFiscal("OBS04") = DadosTransp.Txt(9).Text
RsNotaFiscal("valorproduto") = CCur(Txt(15).Text)
RsNotaFiscal("ValorNota") = CCur(Txt(16).Text)
RsNotaFiscal("vencimento1") = DadosTransp.Vencimento(0).Text
RsNotaFiscal("vencimento2") = DadosTransp.Vencimento(1).Text
RsNotaFiscal("vencimento3") = DadosTransp.Vencimento(2).Text
RsNotaFiscal("Vendedor") = Right("00000" & Txt(10).Text, 5)
RsNotaFiscal.Update
RsNotaFiscal.Close
Set RsNotaFiscal = Nothing
End Function

Function FechaImpressaoBel(Linhas As Integer)
On Error Resume Next

Dim LcCompl, a, ax As Integer
Dim lcLinhasSalto As Integer
Dim LcDesc, LCLEtra As String
'==== Imprime Desconto Na Nota
'MsgBox CInt(txt(17).Text)
LcCal = Format(AcertaNumero(CCur(Txt(15).Text) * 0.18, 2), "Currency")
If Natureza.Text = "ORG PUBL. EST." Then Txt(14).Text = "ICMS isento conforme item 136 do anexo 1 do RICMS" & Chr(13) & "Valor da mercadoria com imposto:" & Txt(16).Text & Chr(13) & "Valor do desconto/imposto:" & Txt(11).Text
'Txt(14).Text = "ICMS Recolhido Conforme Decreto 43349/03 de 30/05/03" & Chr(13) & "Valor Total sem Desc. ICMS:" & DadosTransp.VlTotalSdesconto.Text & Chr(13) & " Valor da Isenção do ICMS: " & DadosTransp.DescontoEst.Text
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Txt(14).Text = "Ressarcimento do ICMS S.T. Conforme artigo 27 do anexo XV do decreto 43080-02" & Chr(13) & "do RICMS/MG. Referente a(s) NF(s):" & DadosTransp.Txt(7).Text
   If Len(DadosTransp.Txt(8).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(8).Text
   If Len(DadosTransp.Txt(9).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(9).Text
End If
'MsgBox Txt(14).Text
If Len(Trim(Txt(17).Text)) > 0 And CInt(Txt(17).Text) > 0 Then
   If Linhas < 26 Then
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
'      Print #FnunNota, Chr(13)
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
     ' Print #FnunNota, Chr(13)
      LcLinha = ""
      For a = 1 To 130
         LcLinha = LcLinha + " "
      Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 128) + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
 '    Print #FnunNota, Chr(15) + Right(LcLinha, 128) + Chr(18)
     Linhas = Linhas + 3
   Else
     LcLinha = ""
     For a = 1 To 130
         LcLinha = LcLinha + " "
     Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 128) + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     
  '   Print #FnunNota, Chr(15) + Right(LcLinha, 128) + Chr(18)
     Linhas = Linhas + 1
   End If
End If
For a = Linhas To 29 '22
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
'    Print #FnunNota, Chr(13)
Next
LcLinha = ""
For a = 1 To 1
  LcLinha = LcLinha + " "
Next
'=== Imprime a Base de calculo de ICMS
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
   'LcLinha = LcLinha + Left(txt(13).Text & LcEspC, 13)
   If GLNImprimeBaseC Then
      LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
   Else
        LcLinha = LcLinha + Left(Format(Txt(13).Text, "Currency") & LcEspC, 13)
   End If
Else
   LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
End If
For a = 15 To 19
  LcLinha = LcLinha + " "
Next
'=== Imprime O Valor do Icms
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If GLNImprimeBaseC Then
        LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
    Else
        LcLinha = LcLinha + Left(Format(Txt(11).Text, "Currency") & LcEspC, 13)
    End If
Else
    LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
End If
For a = 31 To 63
  LcLinha = LcLinha + " "
Next
'=== Imprime O Valor do TOTAL DE PRODUTOS
'LcLinha = LcLinha + Left("0" & LcEspC, 13)
LcLinha = LcLinha + Left(Txt(15).Text & LcEspC, 13)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
'    Print #FnunNota, Chr(13)
Next
LcLinha = ""
For a = 1 To 65
  LcLinha = LcLinha + " "
Next
'=== Imprime o total da NOTA
LcLinha = LcLinha + Left(Txt(16).Text & LcEspC, 13)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

LcLinha = ""
For a = 1 To 2
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1

    'Print #FnunNota, Chr(13)
Next
For a = 1 To 3
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Left(DadosTransp.Txt(0).Text & LcEspC, 20)
For a = 22 To 43
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Mid(DadosTransp.Tipo.Text, 1, 1)
For a = 48 To 49
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + DadosTransp.Placa.Text
LcLinha = LcLinha + " " + DadosTransp.Txt(1).Text
LcLinha = LcLinha + " " + DadosTransp.Txt(2).Text
    
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 5
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    'Print #FnunNota, Chr(13)
Next
If Len(Txt(14).Text) = 0 Then
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
End If
'==== Imprime dados Complementares
LcCompl = 0
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcDesc = ""
cb = 1
LcFrase = ""
LcPrimeira = True
If Natureza.Text = "ORG PUBL. EST." Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Conpemsacao = 2
Else
  If Len(Txt(14).Text) > 0 Then Conpemsacao = 2 Else Conpemsacao = 2
End If
'For ax = 1 To Len(txt(14).Text)
    LcAdicionais = Replace(Txt(14).Text, Chr(13), "")
    LcAdicionais = Replace(LcAdicionais, Chr(10), " ")
    If Len(LcAdicionais) > 56 Then
        LcAdicionais1 = Mid(LcAdicionais, 1, 56)
        LcAdicionais2 = Mid(LcAdicionais, 57, Len(LcAdicionais))
        If Len(LcAdicionais2) > 56 Then
            LcAdicionais = LcAdicionais2
            LcAdicionais2 = Mid(LcAdicionais, 1, 56)
            LcAdicionais3 = Mid(LcAdicionais, 57, Len(LcAdicionais))
            If Len(LcAdicionais3) > 56 Then
                LcAdicionais = LcAdicionais3
                LcAdicionais3 = Mid(LcAdicionais, 1, 56)
            Else
                LcAdicionais3 = LcAdicionais3
            End If
        Else
            LcAdicionais2 = LcAdicionais2
        End If
    Else
        LcAdicionais1 = Replace(Txt(14).Text, Chr(13), "")
        LcAdicionais1 = Replace(LcAdicionais1, Chr(10), " ")
    End If
  '  LCLEtra = Mid(txt(14).Text, ax, 1)
 '   If LCLEtra = Chr(13) Then
    If Len(LcAdicionais1) > 0 Then
        LcQuantiImpressao = LcQuantiImpressao + 1
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais1
    End If
    If Len(LcAdicionais2) > 0 Then
        LcQuantiImpressao = LcQuantiImpressao + 1
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais2
    End If
    If Len(LcAdicionais3) > 0 Then
        LcQuantiImpressao = LcQuantiImpressao + 1
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais3
    End If
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Txt(14).Text & Chr(13)
If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Then
   'DadosTransp.txt(7).Text = "Nao Incidencia do ICMS" & Chr(13) & "Conf. Artigo V Inciso X" & Chr(13) & DadosTransp.txt(7).Text
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "Nao Incidencia do ICMS" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "Conf. Artigo V Inciso X" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
  
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If Len(DadosTransp.Txt(7).Text) > 0 Then
    
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(7).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       'Print #FnunNota, DadosTransp.Txt(7).Text & Chr(13)
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(8).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(8).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, DadosTransp.Txt(8).Text & Chr(13)
       
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(9).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(9).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
    
      ' Print #FnunNota, DadosTransp.Txt(9).Text & Chr(13)
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(10).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + "End.Entrega: " & DadosTransp.Txt(10).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(11).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + "O.C " & DadosTransp.Txt(11).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If GLAproveitamentoICMS Then
        LcAdicionais = ""
        LcAdicionais1 = ""
        LcAdicionais2 = ""
        LcAdicionais3 = ""
        LcAdicionais = "Permite o aproveitamento do crédito de ICMS no valor de " & Txt(11).Text
        LcAdicionais = LcAdicionais & " correspontente a aliquota de " & Txt(5).Text & "% Nos termos do artigo 2 da resoluçao do CGSN 53/2008"
        LcAdicionais1 = Mid(LcAdicionais, 1, 56)
        LcAdicionais2 = Mid(LcAdicionais, 57, 54)
        LcAdicionais3 = Mid(LcAdicionais, 111, 54)
        If Len(LcAdicionais1) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais1
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais2) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais2
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais3) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais3
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
    End If
    If Len(GLInformacaoNF) > 0 Then
        LcAdicionais = ""
        LcAdicionais1 = ""
        LcAdicionais2 = ""
        LcAdicionais3 = ""
        LcAdicionais = GLInformacaoNF
        LcAdicionais1 = Mid(LcAdicionais, 1, 56)
        LcAdicionais2 = Mid(LcAdicionais, 57, 54)
        LcAdicionais3 = Mid(LcAdicionais, 111, 54)
        If Len(LcAdicionais1) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais1
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais2) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais2
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais3) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais3
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
    End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(18)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(18)
If Natureza.Text = "ORG PUBL. EST." Then LcCompl = 2
x = LcQuantiImpressao
For a = x To 71
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(13)
Next
'=== Imprime O Numero da nota no canhoto
LcLinha = ""
For a = 0 To 50
    LcLinha = LcLinha + " "
Next
For a = 50 To 67
   LcLinha = LcLinha + " "
Next
LcLinha = LcLinha & Txt(0).Text
'===> Imprime a 1º Linha Gerada

ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

'Print #FnunNota, Chr(20)

LcLinha = ""
For a = 1 + Conpemsacao To GlPuloFim
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
 '  Print #FnunNota, Chr(13)
Next
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = " "
End Function
Function FechaImpressaoDhel(Linhas As Integer)
On Error Resume Next
Dim LcCompl, a, ax As Integer
Dim lcLinhasSalto As Integer
Dim LcDesc, LCLEtra As String
'==== Imprime Desconto Na Nota
LcCal = Format(CCur(Txt(11).Text), "Currency")
If Natureza.Text = "ORG PUBL. EST." Then Txt(14).Text = "ICMS isento conforme item 136 do anexo 1 do RICMS" & Chr(13) & "Valor da mercadoria com imposto:" & Txt(16).Text & Chr(13) & "Valor do desconto/imposto:" & Txt(11).Text
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Txt(14).Text = "Ressarcimento do ICMS S.T. Conforme artigo 27 do anexo XV do decreto 43080-02" & Chr(13) & "do RICMS/MG. Referente a(s) NF(s):" & DadosTransp.Txt(7).Text
   If Len(DadosTransp.Txt(8).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(8).Text
   If Len(DadosTransp.Txt(9).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(9).Text
End If

If Len(Trim(Txt(17).Text)) > 0 And CInt(Txt(17).Text) > 0 Then
   If Linhas < 26 Then
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
     LcLinha = ""
      For a = 1 To 130
         LcLinha = LcLinha + " "
      Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 90)
     LcQuantiImpressao = LcQuantiImpressao + 1
     Linhas = Linhas + 3
   Else
     LcLinha = ""
     For a = 1 To 130
         LcLinha = LcLinha + " "
     Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 90)
     LcQuantiImpressao = LcQuantiImpressao + 1
     Linhas = Linhas + 1
   End If
End If
For a = Linhas To 32 '22
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
Next
LcLinha = ""
For a = 1 To 2
  LcLinha = LcLinha + " "
Next
'=== Imprime a Base de calculo de ICMS
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If GLNImprimeBaseC Then
        LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 21)
    Else
        LcLinha = LcLinha + Left(Txt(13).Text & LcEspC, 21)
    End If
Else
   LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 21)
End If
'=== Imprime O Valor do Icms
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If GLNImprimeBaseC Then
        LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 16)
    Else
        LcLinha = LcLinha + Left(LcCal & LcEspC, 16)
    End If
Else
    LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 16)
End If
For a = 16 To 63
  LcLinha = LcLinha + " "
Next
'=== Imprime O Valor do TOTAL DE PRODUTOS
LcLinha = LcLinha + "   " & Left(Txt(15).Text & LcEspC, 13)
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15)
     LcQuantiImpressao = LcQuantiImpressao + 1
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
For a = 1 To 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
Next
LcLinha = ""
For a = 1 To 89
  LcLinha = LcLinha + " "
Next
'=== Imprime o total da NOTA
LcLinha = LcLinha + Left(Txt(16).Text & LcEspC, 13)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
LcLinha = ""
For a = 1 To 2
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
Next
LcLinha = ""
LcLinha = LcLinha + Left(DadosTransp.Txt(0).Text & LcEspC, 40)
For a = 22 To 44
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Mid(DadosTransp.Tipo.Text, 1, 1)
For a = 48 To 52
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + DadosTransp.Placa.Text
LcLinha = LcLinha + "     " + DadosTransp.Txt(1).Text
LcLinha = LcLinha + "    " + DadosTransp.Txt(2).Text
    
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
For a = 1 To 8
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
Next
If Len(Txt(14).Text) = 0 Then
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
End If
'==== Imprime dados Complementares
LcCompl = 0
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcDesc = ""
cb = 1
LcFrase = ""
LcPrimeira = True
If Natureza.Text = "ORG PUBL. EST." Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Conpemsacao = 2
Else
  If Len(Txt(14).Text) > 0 Then Conpemsacao = 2 Else Conpemsacao = 2
End If
    LcAdicionais = Replace(Txt(14).Text, Chr(13), "")
    LcAdicionais = Replace(LcAdicionais, Chr(10), " ")
    If Len(LcAdicionais) > 50 Then
        LcAdicionais1 = Mid(LcAdicionais, 1, 50)
        LcAdicionais2 = Mid(LcAdicionais, 51, Len(LcAdicionais))
        If Len(LcAdicionais2) > 50 Then
            LcAdicionais = LcAdicionais2
            LcAdicionais2 = Mid(LcAdicionais, 1, 50)
            LcAdicionais3 = Mid(LcAdicionais, 51, Len(LcAdicionais))
            If Len(LcAdicionais3) > 50 Then
                LcAdicionais = LcAdicionais3
                LcAdicionais3 = Mid(LcAdicionais, 1, 50)
            Else
                LcAdicionais3 = LcAdicionais3
            End If
        Else
            LcAdicionais2 = LcAdicionais2
        End If
    Else
        LcAdicionais1 = Replace(Txt(14).Text, Chr(13), "")
        LcAdicionais1 = Replace(LcAdicionais1, Chr(10), " ")
    End If
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + LcAdicionais1
        LcQuantiImpressao = LcQuantiImpressao + 1

        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + LcAdicionais2
        LcQuantiImpressao = LcQuantiImpressao + 1

        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + LcAdicionais3
        LcQuantiImpressao = LcQuantiImpressao + 1
If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Then
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "Nao Incidencia do ICMS" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + "Conf. Artigo V Inciso X" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
  
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If Len(DadosTransp.Txt(7).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + DadosTransp.Txt(7).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(8).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + DadosTransp.Txt(8).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(9).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + DadosTransp.Txt(9).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(10).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + "End.Entrega: " & DadosTransp.Txt(10).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(11).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + "O.C " & DadosTransp.Txt(11).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       LcCompl = LcCompl + 1
    End If
    If GLAproveitamentoICMS Then
        LcAdicionais = "Permite o aproveitamento do crédito de ICMS no valor de " & Txt(11).Text
        LcAdicionais = LcAdicionais & " correspontente a aliquota de " & Txt(5).Text & "% Nos termos do artigo 2 da resoluçao do CGSN 53/2008"
        LcAdicionais1 = Mid(LcAdicionais, 1, 46)
        LcAdicionais2 = Mid(LcAdicionais, 48, 50)
        LcAdicionais3 = Mid(LcAdicionais, 98, 54)
        If Len(LcAdicionais1) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais1
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais2) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais2
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais3) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais3
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
    End If
    If Len(GLInformacaoNF) > 0 Then
        LcAdicionais = ""
        LcAdicionais1 = ""
        LcAdicionais2 = ""
        LcAdicionais3 = ""
        LcAdicionais = GLInformacaoNF
        LcAdicionais1 = Mid(LcAdicionais, 1, 50)
        LcAdicionais2 = Mid(LcAdicionais, 51, 50)
        LcAdicionais3 = Mid(LcAdicionais, 102, 50)
        If Len(LcAdicionais1) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais1
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais2) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais2
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
        If Len(LcAdicionais3) > 0 Then
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcAdicionais3
            LcQuantiImpressao = LcQuantiImpressao + 1
        End If
    End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(18)
LcQuantiImpressao = LcQuantiImpressao + 1
If Natureza.Text = "ORG PUBL. EST." Then LcCompl = 2
x = LcQuantiImpressao
For a = x To 88
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
Next
'=== Imprime O Numero da nota no canhoto
LcLinha = ""
For a = 0 To 47
    LcLinha = LcLinha + " "
Next
LcLinha = LcLinha & Txt(0).Text
'===> Imprime a 1º Linha Gerada
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
LcLinha = ""
For a = 1 + Conpemsacao To GlPuloFim
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
Next
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = " "
End Function

Function FechaImpressao(Linhas As Integer)
On Error Resume Next

Dim LcCompl, a, ax As Integer
Dim lcLinhasSalto As Integer
Dim LcDesc, LCLEtra As String
'==== Imprime Desconto Na Nota
'MsgBox CInt(txt(17).Text)
LcCal = Format(AcertaNumero(CCur(Txt(15).Text) * 0.18, 2), "Currency")

If Natureza.Text = "ORG PUBL. EST." Then Txt(14).Text = Txt(14).Text = "ICMS isento conforme item 136 do anexo 1 do RICMS" & Chr(13) & "Valor da mercadoria com imposto:" & Txt(16).Text & Chr(13) & "Valor do desconto/imposto:" & Txt(11).Text '"ICMS Recolhido Conforme Decreto 43349/03 de 30/05/03" & Chr(13) & "Valor Total sem Desc. ICMS:" & DadosTransp.VlTotalSdesconto.Text & Chr(13) & " Valor da Isenção do ICMS: " & DadosTransp.DescontoEst.Text
If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Txt(14).Text = "Ressarcimento do ICMS S.T. Conforme artigo 27 do anexo XV do decreto 43080-02" & Chr(13) & "do RICMS/MG. Referente a(s) NF(s):" & DadosTransp.Txt(7).Text
   If Len(DadosTransp.Txt(8).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(8).Text
   If Len(DadosTransp.Txt(9).Text) > 0 Then Txt(14).Text = Txt(14).Text & Chr(13) & DadosTransp.Txt(9).Text
End If
'MsgBox Txt(14).Text
If Len(Trim(Txt(17).Text)) > 0 And CInt(Txt(17).Text) > 0 Then
   If Linhas < 20 Then
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
'      Print #FnunNota, Chr(13)
      ReDim Preserve MtImpressao(LcQuantiImpressao)
      MtImpressao(LcQuantiImpressao) = Chr(13)
      LcQuantiImpressao = LcQuantiImpressao + 1
     ' Print #FnunNota, Chr(13)
      LcLinha = ""
      For a = 1 To 130
         LcLinha = LcLinha + " "
      Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 128) + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
 '    Print #FnunNota, Chr(15) + Right(LcLinha, 128) + Chr(18)
     Linhas = Linhas + 3
   Else
     LcLinha = ""
     For a = 1 To 130
         LcLinha = LcLinha + " "
     Next
     LcLinha = LcLinha & "DESCONTO DE " & Format(AcertaNumero(Txt(17).Text, 2), "CURRENCY")
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + Right(LcLinha, 128) + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     
  '   Print #FnunNota, Chr(15) + Right(LcLinha, 128) + Chr(18)
     Linhas = Linhas + 1
   End If
End If
For a = Linhas To 22
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
'    Print #FnunNota, Chr(13)
Next
LcLinha = ""
For a = 1 To 1
  LcLinha = LcLinha + " "
Next
'=== Imprime a Base de calculo de ICMS
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
   LcLinha = LcLinha + Left(Txt(13).Text & LcEspC, 13)
Else
   LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
End If
For a = 15 To 19
  LcLinha = LcLinha + " "
Next
'=== Imprime O Valor do Icms
If Natureza.Text <> "ORG PUBL. EST." And Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    LcLinha = LcLinha + Left(Txt(11).Text & LcEspC, 13)
Else
    LcLinha = LcLinha + Left(Format("0", "Currency") & LcEspC, 13)
End If
For a = 31 To 63
  LcLinha = LcLinha + " "
Next
'=== Imprime O Valor do TOTAL DE PRODUTOS
LcLinha = LcLinha + Left(Txt(15).Text & LcEspC, 13)

ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
'    Print #FnunNota, Chr(13)
Next
LcLinha = ""
For a = 1 To 65
  LcLinha = LcLinha + " "
Next
'=== Imprime o total da NOTA
LcLinha = LcLinha + Left(Txt(16).Text & LcEspC, 13)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

LcLinha = ""
For a = 1 To 2
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1

    'Print #FnunNota, Chr(13)
Next
For a = 1 To 3
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Left(DadosTransp.Txt(0).Text & LcEspC, 20)
For a = 22 To 45
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Mid(DadosTransp.Tipo.Text, 1, 1)
For a = 48 To 49
  LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + DadosTransp.Placa.Text
LcLinha = LcLinha + " " + DadosTransp.Txt(1).Text
LcLinha = LcLinha + " " + DadosTransp.Txt(2).Text
    
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 5
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    'Print #FnunNota, Chr(13)
Next
If Len(Txt(14).Text) = 0 Then
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
End If
'==== Imprime dados Complementares
LcCompl = 0
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
'LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15)
'ReDim Preserve MtImpressao(LcQuantiImpressao)
LcDesc = ""
'If Len(txt(14).Text) = 0 Then
'   MtImpressao(LcQuantiImpressao) = Chr(13)
 '  LcQuantiImpressao = LcQuantiImpressao + 1
'   ReDim Preserve MtImpressao(LcQuantiImpressao)
'   MtImpressao(LcQuantiImpressao) = Chr(13)
'End If
cb = 1
LcFrase = ""
LcPrimeira = True
If Natureza.Text = "ORG PUBL. EST." Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
   Conpemsacao = 0
Else
  If Len(Txt(14).Text) > 0 Then Conpemsacao = 2 Else Conpemsacao = 0
End If
For ax = 1 To Len(Txt(14).Text)
    LCLEtra = Mid(Txt(14).Text, ax, 1)
    If LCLEtra = Chr(13) Then
       ' LcQuantiImpressao = LcQuantiImpressao + 1
       ' ReDim Preserve MtImpressao(LcQuantiImpressao)
       ' MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcFrase
        If LcPrimeira Then
            LcPrimeira = False
            LcQuantiImpressao = LcQuantiImpressao + 1
            ReDim Preserve MtImpressao(LcQuantiImpressao)
            MtImpressao(LcQuantiImpressao) = Chr(13)
            Conpemsacao = Conpemsacao + 3
        End If
        Conpemsacao = Conpemsacao + 1
       ' lcdesc = lcdesc & Chr(13)
       LcFrase = ""
    Else
       LcFrase = LcFrase & LCLEtra
    End If
Next
If Len(LcFrase) > 0 Then
   LcQuantiImpressao = LcQuantiImpressao + 1
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcFrase
End If
'MtImpressao(LcQuantiImpressao) = lcdesc & Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Txt(14).Text & Chr(13)
If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Then
   'DadosTransp.txt(7).Text = "Nao Incidencia do ICMS" & Chr(13) & "Conf. Artigo V Inciso X" & Chr(13) & DadosTransp.txt(7).Text
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "Nao Incidencia do ICMS" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "Conf. Artigo V Inciso X" & Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
  
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    If Len(DadosTransp.Txt(7).Text) > 0 Then
    
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(7).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       'Print #FnunNota, DadosTransp.Txt(7).Text & Chr(13)
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(8).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(8).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, DadosTransp.Txt(8).Text & Chr(13)
       
       LcCompl = LcCompl + 1
    End If
    If Len(DadosTransp.Txt(9).Text) > 0 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + DadosTransp.Txt(9).Text & Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
    
      ' Print #FnunNota, DadosTransp.Txt(9).Text & Chr(13)
       LcCompl = LcCompl + 1
    End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(18)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(18)
If Natureza.Text = "ORG PUBL. EST." Then LcCompl = 2
x = LcQuantiImpressao
For a = x To 65
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(13)
Next
'=== Imprime O Numero da nota no canhoto
LcLinha = ""
For a = 0 To 50
    LcLinha = LcLinha + " "
Next
For a = 50 To 67
   LcLinha = LcLinha + " "
Next
LcLinha = LcLinha & Txt(0).Text
'===> Imprime a 1º Linha Gerada

ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem + LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

'Print #FnunNota, Chr(20)

LcLinha = ""
For a = 1 + Conpemsacao To GlPuloFim
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
 '  Print #FnunNota, Chr(13)
Next
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = " "
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, " "

End Function
Function imprimeitem(Item, Descricao, cst, icms, Unidade As String, quant, Unitario, total As Currency) As Integer
On Error Resume Next
Dim a, b As Integer
Dim descricao1, DESCRICAO2, DESCRICAO3, DESCRICAO4, DESCRICAO5 As String
'Print #FnunNota, Chr(15)
LcLinha = ""
For a = 1 To 1
   LcLinha = LcLinha + " "
Next
LcLinha = LcLinha + Right("00" & Item, 2)
For a = 4 To 5
   LcLinha = LcLinha + " "
Next
b = 1
C = 1
If Len(Descricao) > 74 Then
   For a = 1 To Len(Descricao)
       LCLEtra = Mid(Descricao, a, 1)
       Select Case b
              Case Is = 1
                  descricao1 = descricao1 & LCLEtra
               Case Is = 2
                  DESCRICAO2 = DESCRICAO2 & LCLEtra
               Case Is = 3
                  DESCRICAO3 = DESCRICAO3 & LCLEtra
               Case Is = 4
                  DESCRICAO4 = DESCRICAO4 & LCLEtra
               Case Is = 5
                  DESCRICAO5 = DESCRICAO5 & LCLEtra
        End Select
       If C = 74 Then
          b = b + 1
          C = 1
       End If
       C = C + 1
   Next
   LcVarios = True
   LcLinha = LcLinha + Left(" " & descricao1 & LcEspC, 73)
Else
   LcLinha = LcLinha + Left(" " & Descricao & LcEspC, 73)
   LcVarios = False
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    For a = 58 To 59
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Left(cst & LcEspC, 3)
    For a = 83 To 85
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Left(Unidade & "   ", 3)
    
    For a = 90 To 90
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(quant), 6)
    
    For a = 99 To 103
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(Unitario, "currency")), 10)
    
    For a = 114 To 119
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(total, "currency")), 10)
    
    For a = 130 To 134
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & icms, 2)
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcLinha + Chr(18)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15) + LcLinha + Chr(18)
If LcVarios Then
   If Len(DESCRICAO2) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO2 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     ' Print #FnunNota, Chr(15) + "      " & DESCRICAO2 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
   If Len(DESCRICAO3) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO3 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO4) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO4 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO5) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO5 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
End If
LcVarios = False
End Function
Function ImprimeItemBel(Item, Descricao, cst, icms, Unidade As String, quant, Unitario, total As Currency) As Integer
On Error Resume Next
Dim a, b As Integer
Dim descricao1, DESCRICAO2, DESCRICAO3, DESCRICAO4, DESCRICAO5 As String
'Print #FnunNota, Chr(15)
LcLinha = ""
For a = 1 To 1
   LcLinha = LcLinha + ""
Next
LcLinha = LcLinha + Right("      " & Item, 6)
For a = 7 To 7
   LcLinha = LcLinha + ""
Next
b = 1
C = 1
'MsgBox Mid(Descricao, 1, 62)
descricao1 = ""
DESCRICAO2 = ""
DESCRICAO3 = ""
DESCRICAO4 = ""
DESCRICAO5 = ""
If Len(Descricao) > 62 Then
   For a = 1 To Len(Descricao)
       LCLEtra = Mid(Descricao, a, 1)
       Select Case b
              Case Is = 1
                  descricao1 = descricao1 & LCLEtra
               Case Is = 2
                  DESCRICAO2 = DESCRICAO2 & LCLEtra
               Case Is = 3
                  DESCRICAO3 = DESCRICAO3 & LCLEtra
               Case Is = 4
                  DESCRICAO4 = DESCRICAO4 & LCLEtra
               Case Is = 5
                  DESCRICAO5 = DESCRICAO5 & LCLEtra
        End Select
       If C = 62 Then
          b = b + 1
          C = 1
       End If
       C = C + 1
   Next
   LcVarios = True
   LcLinha = LcLinha + Left(" " & descricao1 & LcEspC, 63)
Else
   LcLinha = LcLinha + Left("  " & Descricao & LcEspC, 63)
   LcVarios = False
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    For a = 58 To 59
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Left(cst & LcEspC, 3)
    For a = 83 To 85
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Left(Unidade & "   ", 3)
    
    For a = 90 To 90
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(quant), 6)
    
    For a = 99 To 103
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(Unitario, "currency")), 12)
    
    For a = 114 To 121
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(total, "currency")), 10)
    
    For a = 130 To 135
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & icms, 5)
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + LcLinha + Chr(18)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15) + LcLinha + Chr(18)
If LcVarios Then
   If Len(DESCRICAO2) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO2 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     ' Print #FnunNota, Chr(15) + "      " & DESCRICAO2 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
   If Len(DESCRICAO3) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO3 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO4) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO4 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO5) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO5 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
End If
LcVarios = False
End Function
Function ImprimeItemDhel(Item, Descricao, cst, icms, Unidade As String, quant, Unitario, total As Currency) As Integer
On Error Resume Next
Dim a, b As Integer
Dim descricao1, DESCRICAO2, DESCRICAO3, DESCRICAO4, DESCRICAO5 As String
'Print #FnunNota, Chr(15)
LcLinha = ""
For a = 1 To 1
   LcLinha = LcLinha + ""
Next
LcLinha = LcLinha + Right("      " & Item, 6)
For a = 7 To 7
   LcLinha = LcLinha + ""
Next
b = 1
C = 1
If Len(Descricao) > 74 Then
   For a = 1 To Len(Descricao)
       LCLEtra = Mid(Descricao, a, 1)
       Select Case b
              Case Is = 1
                  descricao1 = descricao1 & LCLEtra
               Case Is = 2
                  DESCRICAO2 = DESCRICAO2 & LCLEtra
               Case Is = 3
                  DESCRICAO3 = DESCRICAO3 & LCLEtra
               Case Is = 4
                  DESCRICAO4 = DESCRICAO4 & LCLEtra
               Case Is = 5
                  DESCRICAO5 = DESCRICAO5 & LCLEtra
        End Select
       If C = 74 Then
          b = b + 1
          C = 1
       End If
       C = C + 1
   Next
   LcVarios = True
   LcLinha = LcLinha + Left(" " & descricao1 & LcEspC, 56)
Else
   LcLinha = LcLinha + Left("  " & Descricao & LcEspC, 56)
   LcVarios = False
End If
If Natureza.Text <> "RESSARCIMENTO DO ICMS S.T" Then
    For a = 58 To 59
       LcLinha = LcLinha + ""
    Next
    LcLinha = LcLinha + Left(cst & LcEspC, 3)
    For a = 83 To 85
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Left(Unidade & "   ", 3)
    
    For a = 90 To 90
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(quant), 6)
    
    For a = 99 To 100
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(Unitario, "currency")), 10)
    
    For a = 114 To 115
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & CStr(Format(total, "currency")), 14)
    
    For a = 130 To 130
       LcLinha = LcLinha + " "
    Next
    LcLinha = LcLinha + Right(LcEspC & icms, 5)
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15) + LcLinha + Chr(18)
If LcVarios Then
   If Len(DESCRICAO2) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO2 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     ' Print #FnunNota, Chr(15) + "      " & DESCRICAO2 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
   If Len(DESCRICAO3) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO3 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO4) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO4 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
    If Len(DESCRICAO5) > 0 Then
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = Chr(15) + LcMargem + Left(LcMargem, Len(LcMargem) / 1.5) + "      " & DESCRICAO5 + Chr(18)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, Chr(15) + "      " & DESCRICAO3 + Chr(18)
      LcImpressoes = LcImpressoes + 1
   End If
End If
LcVarios = False
End Function

Function cabecalhonota()
On Error Resume Next
Dim LcExtenso, LcExtenso1, LcExtenso2, LcExtenso3 As String
Dim LcPesq  As String
Dim LcCgc1 As String
Dim LcCpf1 As String
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Dim a, b As Integer
On Error Resume Next
LcQuantiImpressao = 0

For a = 1 To LcSalto
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
   ' Print #FnunNota, Chr(13)
Next
'===> Gera a Primeira Linha
LcLinha = ""
For a = 0 To 49
    LcLinha = LcLinha + " "
Next

LcLinha = LcLinha + "X"
For a = 50 To 69
   LcLinha = LcLinha + " "
Next
LcUltimo = CCur(Txt(16).Text) - (CCur(AcertaNumero(DadosTransp.valor.Text, 1.5)) * CCur(DadosTransp.Quantidade.Text))

LcValor1 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor2 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor3 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))

Select Case Val(DadosTransp.Quantidade.Text)
 Case Is = 1
      LcValor1 = LcValor1 + LcUltimo
 Case Is = 2
      LcValor2 = LcValor2 + LcUltimo
 Case Is = 3
      LcValor3 = LcValor3 + LcUltimo
End Select

LcLinha = LcLinha + FrmSaidaProduto.Txt(0).Text
'===> Imprime a 1º Linha Gerada
'Print #FnunNota, Chr(14)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)


For a = 1 To 4 'Salto de Linhas
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   ' Print #FnunNota, Chr(13)
Next
LcLinha = " "
If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
   LcLinha = LcLinha + "VENDAS A PRAZO "
Else
   If Natureza.Text = "TRANSFERENCIA" Then
      LcLinha = LcLinha + "REM. P/ DEP. FECH."
   Else
      If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
         LcLinha = LcLinha + "RESSARC. ICMS S.T"
      Else
         LcLinha = LcLinha + FrmSaidaProduto.Natureza.Text
      End If
   End If
End If
If FrmSaidaProduto.Natureza.Text = "VENDAS A PRAZO" Then LcLinha = LcLinha + " "

For a = 12 To 22
   LcLinha = LcLinha + " "
Next

LcLinha = LcLinha + CFOP.Text
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 2 'Salto de Linhas
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   ' Print #FnunNota, Chr(13)
Next
LcLinha = " "
For a = 1 To 80
   LcEspC = LcEspC + " "
Next
'If Mask(1).Text = "  .   .   /    -  " And Mask(4).Text = "   .   .   -  " Then Exit Sub
LcCgc = ""
LcCpf = ""
If Not RsClientes.EOF Then
   LcLinha = Left(RsClientes!razaosoc & LcEspC, 48) '==== Diminui 2 carac
   LcCgc1 = RsClientes!CGC & ""
   LcCgc1 = Replace(LcCgc1, ".", "")
   LcCgc1 = Replace(LcCgc1, "/", "")
   LcCgc1 = Replace(LcCgc1, "-", "")
   LcCgc1 = Trim(LcCgc1)
   'If RsClientes!cgc <> "  .   .   /    -  " Then
   If Len(LcCgc1) > 0 Then
      
   '   For a = 1 To Len(RsClientes!cgc)
   '       If IsNumeric(Mid(RsClientes!cgc, a, 1)) Then
   '          LcCgc = LcCgc & Mid(RsClientes!cgc, a, 1)
   '       End If
    ' Next
     LcCgc = ExibeCgc(LcCgc1) 'Mid(LcCgc, 1, 2) & "." & Mid(LcCgc, 3, 3) & "." & Mid(LcCgc, 6, 3) & "/" & Mid(LcCgc, 9, 4) & "-" & Mid(LcCgc, 13)
     LcLinha = LcLinha + " " + LcCgc & ""
 
   Else
        LcCpf1 = RsClientes!cpf & ""
        LcCpf1 = Replace(LcCpf1, ".", "")
        LcCpf1 = Replace(LcCpf1, "-", "")
        LcCpf1 = Trim(LcCpf1)

        If Len(LcCpf1) > 0 Then
      ' If RsClientes!cpf <> "   .   .   -  " Then
        '==> É CPF
         '==> Tira Formatacao
         'For a = 1 To Len(RsClientes!cpf)
         '   If IsNumeric(Mid(RsClientes!cpf, a, 1)) Then
         '      LcCpf = LcCpf & Mid(RsClientes!cpf, a, 1)
         '   End If
         'Next
         LcCpf = ExibeCpf(LcCpf1) 'Mid(LcCpf, 1, 3) & "." & Mid(LcCpf, 4, 3) & "." & Mid(LcCpf, 7, 3) & "-" & Mid(LcCpf, 10)
         
         'LcCpf = RsClientes!cpf
             '===> Formata De novo
         'LcCpf = Mid(LcNovoCpf, 1, 3) & "." & Mid(LcNovoCpf, 4, 3) & "." & Mid(LcNovoCpf, 7, 3) & "-" & Mid(LcNovoCpf, 10)
         LcLinha = LcLinha + " " + Left(LcCpf & "                  ", 18) & ""
       Else
         LcLinha = LcLinha + " " + "              " & ""
       End If
      
   End If
   LcLinha = LcLinha + "  " + Format(Date, "dd/mm/yyyy")
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   '  Print #FnunNota, LcLinha + Chr(13)
   LcLinha = "  "
   
   If Not IsNull(RsClientes!End) Then
      LcLinha = Left(RsClientes!End & LcEspC, 40)
   Else
     LcLinha = Left(" " & LcEspC, 40)
   End If
   
   For a = 42 To 70
      LcLinha = LcLinha + " "
   Next
   If Not IsNull(RsClientes!Bairro) Then
      LcLinha = LcLinha + Left(RsClientes!Bairro & LcEspC, 20)
   Else
      LcLinha = LcLinha + Left(" " & LcEspC, 20)
   End If
   For a = 71 To 81
      LcLinha = LcLinha + " "
   Next
   If Not IsNull(RsClientes!Cep) Then
      LcLinha = LcLinha + RsClientes!Cep & ""
   End If
Else
   LcLinha = Left("Não Cadastrado" & LcEspC, 50)
   LcLinha = LcLinha + " " + "               "
   LcLinha = LcLinha + "    " + Format(Date, "dd/mm/yyyy")
   
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1

   'Print #FnunNota, LcLinha + Chr(13)
   LcLinha = "  "
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(13)
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 1 'Salto de Linhas
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    'Print #FnunNota, Chr(13)
Next
LcLinha = " "

Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Not RsCidade.EOF Then
   LcLinha = LcLinha + Left(RsCidade!Nome & LcEspC, 30)
   For a = 32 To 57
       LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   For a = 69 To 80
      LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + RsClientes!Estado
   For a = 83 To 89
       LcLinha = LcLinha + " "
   Next
   
   If Len(LcCgc1) > 0 Then
      If Len(RsClientes!INSCEST) > 0 Then
           LcLinha = LcLinha + RsClientes!INSCEST
      Else
           LcLinha = LcLinha + "ISENTO"
      End If
   Else
      If Len(RsClientes!rg) > 0 Then
         LcLinha = LcLinha + RsClientes!rg
      End If
   End If
 
 Else
   LcLinha = LcLinha + Left("NÃO CADASTRADO" & LcEspC, 30)
   For a = 32 To 57
       LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   For a = 69 To 80
      LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + "**"
   For a = 83 To 89
       LcLinha = LcLinha + " "
   Next
   If Len(RsClientes!INSCEST) > 0 Then
      LcLinha = LcLinha + RsClientes!INSCEST
   Else
      LcLinha = LcLinha + "ISENTO"
   End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)
'For a = 1 To 2 'Salto de Linhas
 '   Print #FnunNota, Chr(13)
'Next
LcLinha = ""
'===== Gerar extenso
LcExtenso = GeraExtenso(CCur(Txt(16).Text))
LcTamanhoExt = Len(LcExtenso)
LcExtenso1 = "      " & Mid(LcExtenso, 1, 76)
LcExtenso2 = "      " & Mid(LcExtenso, 81, 76)
LcExtenso3 = "      " & Mid(LcExtenso, 161, 76)

For a = 1 To 4 'Salto de Linhas
    LcLinha = LcLinha & " "
Next
'== usar tamanho de 80

If (FrmSaidaProduto.Natureza.Text = "VENDAS A PRAZO") Then
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(15)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(15)
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
  ' Print #FnunNota, Chr(13)
      For a = 1 To Val(DadosTransp.Quantidade.Text)
       Select Case a
           Case Is = 1
                LcExtenso1 = Left(LcExtenso1 & "                                                                                ", 80)
                LcLinha = LcLinha & LcExtenso1
                LcLinha = Left(LcLinha & "                                             ", 90) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
                   LcLinha = LcLinha + "C/APRES.      "
                Else
                   If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                      LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(0).Text
                   Else
                       LcLinha = LcLinha + DadosTransp.Vencimento(0).Text
                   End If
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor1, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
               ' Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 2
                LcLinha = ""
                If Len(Trim(LcExtenso2)) > 0 Then
                    LcExtenso1 = Left(LcExtenso1 & "                                                                                ", 80)
                    LcLinha = LcLinha & LcExtenso2
                Else
                    LcLinha = ""
                    For ron = 1 To 84
                        LcLinha = LcLinha & " "
                    Next
                End If
                LcLinha = Left(LcLinha & "                                                      ", 90) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(1).Text
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(1).Text
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor2, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 3
                LcLinha = ""
                If Len(Trim(LcExtenso3)) > 0 Then
                    LcExtenso3 = Left(LcExtenso3 & "                                                                                ", 80)
                    LcLinha = LcLinha & LcExtenso3
                Else
                    LcLinha = ""
                    For ron = 1 To 84
                        LcLinha = LcLinha & " "
                    Next
                End If
                LcLinha = Left(LcLinha & "                                             ", 90) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(2).Text
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(2).Text
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor3, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
        End Select
     Next
     For wq = 1 To Val(DadosTransp.Quantidade.Text)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       Print #FnunNota, Chr(13)
     Next
     If Val(DadosTransp.Quantidade.Text) = 2 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, Chr(13)
     End If
     If Val(DadosTransp.Quantidade.Text) = 1 Then
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       'Print #FnunNota, Chr(13)
     End If
     
Else
   For a = 1 To 2
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, Chr(13)
   Next
     LcExtenso1 = Left(LcExtenso1 & "                                             ", 80)
     LcLinha = LcLinha & LcExtenso1

     'For b = 80 To 119
    '     LcLinha = LcLinha + " "
    ' Next
     LcLinha = Left(LcLinha & "                                             ", 90) + Txt(0).Text & "/" & Right("00" & 1, 2)
     For b = 100 To 107
         LcLinha = LcLinha + " "
     Next
     If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
        LcLinha = LcLinha + "C/APRESENTAÇÃO  "
        DadosTransp.valor.Text = Txt(16).Text
        
     Else
        If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
           LcLinha = " "
           DadosTransp.valor.Text = 0
        Else
           LcLinha = LcLinha + "A VISTA "
           For b = 1 To 8
              LcLinha = LcLinha + " "
           Next
        End If
     End If
     If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
        LcLinha = ""
     Else
        LcLinha = LcLinha + Format(DadosTransp.valor.Text, "currency")
     End If
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, LcLinha + Chr(13)
     If Len(LcExtenso2) > 0 Then
        LcLinha = LcLinha & LcExtenso2
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     If Len(LcExtenso3) > 0 Then
        LcLinha = LcLinha & LcExtenso3
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     For a = 1 To 3
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
     Next
End If

End Function
Function CabecalhoNotaDhel()
On Error Resume Next
Dim LcExtenso, LcExtenso1, LcExtenso2, LcExtenso3 As String
Dim LcPesq  As String
Dim LcCgc1 As String
Dim LcCpf1 As String
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Dim a, b As Integer
On Error Resume Next
'LcQuantiImpressao = 0
For a = 1 To LcSalto

    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
Next
'===> Gera a Primeira Linha
LcUltimo = CCur(Txt(16).Text) - (CCur(AcertaNumero(DadosTransp.valor.Text, 1.5)) * CCur(DadosTransp.Quantidade.Text))

LcValor1 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor2 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor3 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))

Select Case Val(DadosTransp.Quantidade.Text)
 Case Is = 1
      LcValor1 = LcValor1 + LcUltimo
 Case Is = 2
      LcValor2 = LcValor2 + LcUltimo
 Case Is = 3
      LcValor3 = LcValor3 + LcUltimo
End Select
'===> Imprime a 1º Linha Gerada
LcLinha = ""

LcLinha = ""
LcLinha = String(36, " ")
LcLinha = LcLinha + "X"
LcLinha = LcLinha & String(14, " ")
LcLinha = LcLinha + FrmSaidaProduto.Txt(0).Text
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1

For a = 1 To 7
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
Next
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcQuantiImpressao = LcQuantiImpressao + 1
LcLinha = ""
If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
   LcLinha = LcLinha + "VENDAS A PRAZO "
Else
   If Natureza.Text = "TRANSFERENCIA" Then
      LcLinha = LcLinha + "REM. P/ DEP. FECH."
   Else
      If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
         LcLinha = LcLinha + "RESSARC. ICMS S.T"
      Else
         LcLinha = LcLinha & FrmSaidaProduto.Natureza.Text
      End If
   End If
End If

LcLinha = LcLinha & String(21, " ")
LcLinha = LcLinha + CFOP.Text

   'Print #FnunNota, Chr(15)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcQuantiImpressao = LcQuantiImpressao + 1


ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1

LcLinha = ""
LcEspC = String(79, " ")
LcCgc = ""
LcCpf = ""
If Not RsClientes.EOF Then
   LcLinha = LcLinha & Left(RsClientes!razaosoc & LcEspC, 70) '==== Diminui 2 carac
   LcCgc1 = RsClientes!CGC & ""
   LcCgc1 = Replace(LcCgc1, ".", "")
   LcCgc1 = Replace(LcCgc1, "/", "")
   LcCgc1 = Replace(LcCgc1, "-", "")
   LcCgc1 = Trim(LcCgc1)

   If Len(LcCgc1) > 0 Then
     LcCgc = ExibeCgc(LcCgc1) 'Mid(LcCgc, 1, 2) & "." & Mid(LcCgc, 3, 3) & "." & Mid(LcCgc, 6, 3) & "/" & Mid(LcCgc, 9, 4) & "-" & Mid(LcCgc, 13)
     LcLinha = LcLinha + "    " + LcCgc & ""
 
   Else
        LcCpf1 = RsClientes!cpf & ""
        LcCpf1 = Replace(LcCpf1, ".", "")
        LcCpf1 = Replace(LcCpf1, "-", "")
        LcCpf1 = Trim(LcCpf1)

        If Len(LcCpf1) > 0 Then
         LcCpf = ExibeCpf(LcCpf1) 'Mid(LcCpf, 1, 3) & "." & Mid(LcCpf, 4, 3) & "." & Mid(LcCpf, 7, 3) & "-" & Mid(LcCpf, 10)
           LcLinha = LcLinha + " " + Left(LcCpf & "                  ", 18) & ""
       Else
         LcLinha = LcLinha + " " + "              " & ""
       End If
      
   End If
   LcLinha = LcLinha + "           " + Format(Date, "dd/mm/yy")
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha
   LcQuantiImpressao = LcQuantiImpressao + 1
   LcLinha = ""
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   LcLinha = ""
   If Not IsNull(RsClientes!End) Then
      LcLinha = Left(RsClientes!End & LcEspC, 50)
   Else
     LcLinha = Left(" " & LcEspC, 50)
   End If
   LcLinha = LcLinha & String(9, " ")
   If Not IsNull(RsClientes!Bairro) Then
      LcLinha = LcLinha + Left(RsClientes!Bairro & LcEspC, 20)
   Else
      LcLinha = LcLinha + Left(" " & LcEspC, 20)
   End If
   LcLinha = LcLinha & String(7, " ")
   If Not IsNull(RsClientes!Cep) Then
      LcLinha = LcLinha + RsClientes!Cep & ""
   End If
Else
   LcLinha = Left("Não Cadastrado" & LcEspC, 50)
   LcLinha = LcLinha + " " + "               "
   LcLinha = LcLinha + "    " + Format(Date, "dd/mm/yyyy")
   
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1

   'Print #FnunNota, LcLinha + Chr(13)
   LcLinha = "  "
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(13)
End If
'Print #FnunNota, Chr(15)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
LcLinha = "       "
Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Not RsCidade.EOF Then
   LcLinha = LcLinha + Left(RsCidade!Nome & LcEspC, 30)
   LcLinha = LcLinha & String(8, " ")
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   LcLinha = LcLinha & String(11, " ")
   LcLinha = LcLinha + RsClientes!Estado
   LcLinha = LcLinha & String(5, " ")
   If Len(LcCgc1) > 0 Then
      If Len(RsClientes!INSCEST) > 0 Then
           LcLinha = LcLinha + RsClientes!INSCEST
      Else
           LcLinha = LcLinha + "ISENTO"
      End If
   Else
      If Len(RsClientes!rg) > 0 Then
         LcLinha = LcLinha + RsClientes!rg
      End If
   End If
 
 Else
   LcLinha = LcLinha + Left("NÃO CADASTRADO" & LcEspC, 30)
   For a = 32 To 57
       LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   For a = 69 To 80
      LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + "**"
   For a = 83 To 89
       LcLinha = LcLinha + " "
   Next
   If Len(RsClientes!INSCEST) > 0 Then
      LcLinha = LcLinha + RsClientes!INSCEST
   Else
      LcLinha = LcLinha + "ISENTO"
   End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1

LcLinha = ""
'===== Gerar extenso
LcExtenso = GeraExtenso(CCur(Txt(16).Text))
LcTamanhoExt = Len(LcExtenso)
LcExtenso1 = "           " & Mid(LcExtenso, 1, 40)
LcExtenso2 = "           " & Mid(LcExtenso, 41, 81)
LcExtenso3 = "           " & Mid(LcExtenso, 82, 123)

LcLinha = LcLinha & String(1, " ")
'== usar tamanho de 80
If (FrmSaidaProduto.Natureza.Text = "VENDAS A PRAZO") Then
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1

      For a = 1 To Val(DadosTransp.Quantidade.Text)
       Select Case a
           Case Is = 1
                LcExtenso1 = Left(LcExtenso1 & "                                             ", 63)
                LcLinha = LcLinha & LcExtenso1
                LcLinha = Left(LcLinha & "                                             ", 73) + Txt(0).Text & "/" & Right("00" & 1, 2)
                For b = 100 To 105
                    LcLinha = LcLinha + " "
                Next
                If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
                   LcLinha = LcLinha + "C/APRES.       "
                Else
                   If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                      LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(0).Text
                      LcLinha = LcLinha & "    "
                   Else
                       LcLinha = LcLinha + DadosTransp.Vencimento(0).Text
                       LcLinha = LcLinha & "       "
                   End If
                End If
                LcLinha = LcLinha & Format(LcValor1, "currency")
                'MsgBox LcLinha
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
               ' Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 2
                LcLinha = ""
                LcExtenso2 = Left(LcExtenso2 & "                                             ", 63)
                LcLinha = LcLinha & LcExtenso2
                LcLinha = Left(LcLinha & "                                             ", 73) + Txt(0).Text & "/" & Right("00" & 2, 2)
                For b = 100 To 105
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(1).Text
                   LcLinha = LcLinha & "    "
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(1).Text
                   LcLinha = LcLinha & "       "
                End If
                LcLinha = LcLinha & Format(LcValor2, "currency")
                'MsgBox LcLinha
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 3
                LcLinha = ""
                LcExtenso3 = Left(LcExtenso3 & "                                             ", 63)
                LcLinha = LcLinha & LcExtenso3
                LcLinha = Left(LcLinha & "                                             ", 73) + Txt(0).Text & "/" & Right("00" & 3, 2)
                For b = 100 To 105
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(2).Text
                   LcLinha = LcLinha & "    "
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(2).Text
                   LcLinha = LcLinha & "       "
                End If
                LcLinha = LcLinha & Format(LcValor3, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
        End Select
     Next
     For wq = Val(DadosTransp.Quantidade.Text) To 5
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
     Next
Else
   For a = 1 To 1
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, Chr(13)
   Next
     LcExtenso1 = Left(LcExtenso1 & "                                             ", 63)
     LcLinha = LcLinha & LcExtenso1

     'For b = 80 To 119
    '     LcLinha = LcLinha + " "
    ' Next
     LcLinha = Left(LcLinha & "                                             ", 73) + Txt(0).Text & "/" & Right("00" & 1, 2)
     For b = 100 To 105
         LcLinha = LcLinha + " "
     Next
     If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
        LcLinha = LcLinha + "C/APRESENTAÇÃO  "
        DadosTransp.valor.Text = Txt(16).Text
        
     Else
        If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
           LcLinha = " "
           DadosTransp.valor.Text = 0
        Else
           LcLinha = LcLinha + "A VISTA "
           For b = 1 To 7
              LcLinha = LcLinha + " "
           Next
        End If
     End If
     If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
        LcLinha = ""
     Else
        LcLinha = LcLinha + Format(DadosTransp.valor.Text, "currency")
     End If
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, LcLinha + Chr(13)
     If Len(LcExtenso2) <> 11 Then
        LcLinha = LcLinha & LcExtenso2
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     If Len(LcExtenso3) <> 11 Then
        LcLinha = LcLinha & LcExtenso3
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     For a = 1 To 5
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
     Next
End If


End Function
Function CabecalhoNotabel()
On Error Resume Next
Dim LcExtenso, LcExtenso1, LcExtenso2, LcExtenso3 As String
Dim LcPesq  As String
Dim LcCgc1 As String
Dim LcCpf1 As String
AbreBase
Set RsClientes = Dbbase.OpenRecordset("select * from alid001 where codigo='" & Txt(8).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

Dim a, b As Integer
On Error Resume Next
LcQuantiImpressao = 0
For a = 1 To LcSalto
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
Next
'===> Gera a Primeira Linha

LcUltimo = CCur(Txt(16).Text) - (CCur(AcertaNumero(DadosTransp.valor.Text, 1.5)) * CCur(DadosTransp.Quantidade.Text))

LcValor1 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor2 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))
LcValor3 = CCur(AcertaNumero(DadosTransp.valor.Text, 2))

Select Case Val(DadosTransp.Quantidade.Text)
 Case Is = 1
      LcValor1 = LcValor1 + LcUltimo
 Case Is = 2
      LcValor2 = LcValor2 + LcUltimo
 Case Is = 3
      LcValor3 = LcValor3 + LcUltimo
End Select
'===> Imprime a 1º Linha Gerada
LcLinha = " "
LcLinha = ""
LcLinha = String(49, " ")
LcLinha = LcLinha + "X"
LcLinha = LcLinha & String(19, " ")
LcLinha = LcLinha + FrmSaidaProduto.Txt(0).Text
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
For a = 1 To 3 'Salto de Linhas
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
Next
LcLinha = " "
If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
   LcLinha = LcLinha + "VENDAS A PRAZO "
Else
   If Natureza.Text = "TRANSFERENCIA" Then
      LcLinha = LcLinha + "REM. P/ DEP. FECH."
   Else
      If Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
         LcLinha = LcLinha + "RESSARC. ICMS S.T"
      Else
         LcLinha = LcLinha + FrmSaidaProduto.Natureza.Text
      End If
   End If
End If
If FrmSaidaProduto.Natureza.Text = "VENDAS A PRAZO" Then LcLinha = LcLinha + " "

LcLinha = LcLinha & String(7, " ")
LcLinha = LcLinha + CFOP.Text
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1

For a = 1 To 2 'Salto de Linhas
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
Next
LcLinha = " "
LcEspC = String(79, " ")
LcCgc = ""
LcCpf = ""
If Not RsClientes.EOF Then
   LcLinha = Left(RsClientes!razaosoc & LcEspC, 48) '==== Diminui 2 carac
   LcCgc1 = RsClientes!CGC & ""
   LcCgc1 = Replace(LcCgc1, ".", "")
   LcCgc1 = Replace(LcCgc1, "/", "")
   LcCgc1 = Replace(LcCgc1, "-", "")
   LcCgc1 = Trim(LcCgc1)
   If Len(LcCgc1) > 0 Then
     LcCgc = ExibeCgc(LcCgc1) 'Mid(LcCgc, 1, 2) & "." & Mid(LcCgc, 3, 3) & "." & Mid(LcCgc, 6, 3) & "/" & Mid(LcCgc, 9, 4) & "-" & Mid(LcCgc, 13)
     LcLinha = LcLinha + "" + LcCgc & ""
 
   Else
        LcCpf1 = RsClientes!cpf & ""
        LcCpf1 = Replace(LcCpf1, ".", "")
        LcCpf1 = Replace(LcCpf1, "-", "")
        LcCpf1 = Trim(LcCpf1)

        If Len(LcCpf1) > 0 Then
        LcCpf = ExibeCpf(LcCpf1) 'Mid(LcCpf, 1, 3) & "." & Mid(LcCpf, 4, 3) & "." & Mid(LcCpf, 7, 3) & "-" & Mid(LcCpf, 10)
         LcLinha = LcLinha + " " + Left(LcCpf & "                  ", 18) & ""
       Else
         LcLinha = LcLinha + " " + "              " & ""
       End If
      
   End If
   LcLinha = LcLinha + "   " + Format(Date, "dd/mm/yy")
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   '  Print #FnunNota, LcLinha + Chr(13)
   LcLinha = "  "
   
   If Not IsNull(RsClientes!End) Then
      LcLinha = Left(RsClientes!End & LcEspC, 40)
   Else
     LcLinha = Left(" " & LcEspC, 40)
   End If
   LcLinha = LcLinha & String(28, " ")
   If Not IsNull(RsClientes!Bairro) Then
      LcLinha = LcLinha + Left(RsClientes!Bairro & LcEspC, 20)
   Else
      LcLinha = LcLinha + Left(" " & LcEspC, 20)
   End If
   LcLinha = LcLinha & String(10, " ")
   If Not IsNull(RsClientes!Cep) Then
      LcLinha = LcLinha + RsClientes!Cep & ""
   End If
Else
   LcLinha = Left("Não Cadastrado" & LcEspC, 50)
   LcLinha = LcLinha + " " + "               "
   LcLinha = LcLinha + "    " + Format(Date, "dd/mm/yyyy")
   
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = LcMargem & LcLinha + Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1

   'Print #FnunNota, LcLinha + Chr(13)
   LcLinha = "  "
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(13)
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = Chr(15)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, Chr(15)
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
'Print #FnunNota, LcLinha + Chr(13)

For a = 1 To 1 'Salto de Linhas
    ReDim Preserve MtImpressao(LcQuantiImpressao)
    MtImpressao(LcQuantiImpressao) = Chr(13)
    LcQuantiImpressao = LcQuantiImpressao + 1
    'Print #FnunNota, Chr(13)
Next
LcLinha = " "

Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Not RsCidade.EOF Then
   LcLinha = LcLinha + Left(RsCidade!Nome & LcEspC, 30)
   LcLinha = LcLinha & String(25, " ")
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   LcLinha = LcLinha & String(11, " ")
   LcLinha = LcLinha + RsClientes!Estado
   LcLinha = LcLinha & String(6, " ")
   If Len(LcCgc1) > 0 Then
      If Len(RsClientes!INSCEST) > 0 Then
           LcLinha = LcLinha + RsClientes!INSCEST
      Else
           LcLinha = LcLinha + "ISENTO"
      End If
   Else
      If Len(RsClientes!rg) > 0 Then
         LcLinha = LcLinha + RsClientes!rg
      End If
   End If
 
 Else
   LcLinha = LcLinha + Left("NÃO CADASTRADO" & LcEspC, 30)
   For a = 32 To 57
       LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + Left(RsClientes!Fone1 & LcEspC, 11)
   For a = 69 To 80
      LcLinha = LcLinha + " "
   Next
   LcLinha = LcLinha + "**"
   For a = 83 To 89
       LcLinha = LcLinha + " "
   Next
   If Len(RsClientes!INSCEST) > 0 Then
      LcLinha = LcLinha + RsClientes!INSCEST
   Else
      LcLinha = LcLinha + "ISENTO"
   End If
End If
ReDim Preserve MtImpressao(LcQuantiImpressao)
MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
LcQuantiImpressao = LcQuantiImpressao + 1
LcLinha = ""
'===== Gerar extenso
LcExtenso = GeraExtenso(CCur(Txt(16).Text))
LcTamanhoExt = Len(LcExtenso)
LcExtenso1 = "      " & Mid(LcExtenso, 1, 76)
LcExtenso2 = "      " & Mid(LcExtenso, 81, 76)
LcExtenso3 = "      " & Mid(LcExtenso, 161, 76)

LcLinha = LcLinha & String(1, " ")
'== usar tamanho de 80

If (FrmSaidaProduto.Natureza.Text = "VENDAS A PRAZO") Then
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(15)
   LcQuantiImpressao = LcQuantiImpressao + 1
   'Print #FnunNota, Chr(15)
   ReDim Preserve MtImpressao(LcQuantiImpressao)
   MtImpressao(LcQuantiImpressao) = Chr(13)
   LcQuantiImpressao = LcQuantiImpressao + 1
  ' Print #FnunNota, Chr(13)
      For a = 1 To Val(DadosTransp.Quantidade.Text)
      
       Select Case a
           Case Is = 1
                LcExtenso1 = Left(LcExtenso1 & "                                                                                ", 80)
                LcLinha = LcLinha & LcExtenso1
                LcLinha = Left(LcLinha & "                                        ", 85) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
                   LcLinha = LcLinha + "C/APRES.      "
                Else
                   If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                      LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(0).Text
                   Else
                       LcLinha = LcLinha + DadosTransp.Vencimento(0).Text
                   End If
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor1, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
               ' Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 2
                LcLinha = ""
                If Len(Trim(LcExtenso2)) > 0 Then
                    LcExtenso1 = Left(LcExtenso1 & "                                                                                ", 80)
                    LcLinha = LcLinha & LcExtenso2
                Else
                    LcLinha = ""
                    For ron = 1 To 84
                        LcLinha = LcLinha & " "
                    Next
                End If
                LcLinha = Left(LcLinha & "                                                      ", 90) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(1).Text
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(1).Text
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor2, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
            Case Is = 3
                LcLinha = ""
                If Len(Trim(LcExtenso3)) > 0 Then
                    LcExtenso3 = Left(LcExtenso3 & "                                                                                ", 80)
                    LcLinha = LcLinha & LcExtenso3
                Else
                    LcLinha = ""
                    For ron = 1 To 84
                        LcLinha = LcLinha & " "
                    Next
                End If
                LcLinha = Left(LcLinha & "                                             ", 90) + Txt(0).Text & "/" & Right("00" & a, 2)
                For b = 100 To 107
                    LcLinha = LcLinha + " "
                Next
                If DadosTransp.TipoMonetario.Text = "CHEQUE" Then
                   LcLinha = LcLinha + "CH " + DadosTransp.Vencimento(2).Text
                Else
                   LcLinha = LcLinha + DadosTransp.Vencimento(2).Text
                End If
                For b = 1 To 8
                    LcLinha = LcLinha + " "
                Next
                LcLinha = LcLinha & Format(LcValor3, "currency")
                ReDim Preserve MtImpressao(LcQuantiImpressao)
                MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
                LcQuantiImpressao = LcQuantiImpressao + 1
                'Print #FnunNota, LcLinha + Chr(13)
               
        End Select
     Next
     For wq = 1 To Val(DadosTransp.Quantidade.Text)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       Print #FnunNota, Chr(13)
     Next
     If Val(DadosTransp.Quantidade.Text) = 2 Then
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, Chr(13)
     End If
     If Val(DadosTransp.Quantidade.Text) = 1 Then
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
       'Print #FnunNota, Chr(13)
     End If
     
Else
   For a = 1 To 2
       ReDim Preserve MtImpressao(LcQuantiImpressao)
       MtImpressao(LcQuantiImpressao) = Chr(13)
       LcQuantiImpressao = LcQuantiImpressao + 1
      ' Print #FnunNota, Chr(13)
   Next
     LcExtenso1 = Left(LcExtenso1 & "                                             ", 73)
     LcLinha = LcLinha & LcExtenso1
     LcLinha = Left(LcLinha & "                                             ", 83) + Txt(0).Text & "/" & Right("00" & 1, 2)
     For b = 100 To 107
         LcLinha = LcLinha + " "
     Next
     If Natureza.Text = "EMPENHO" Or Natureza.Text = "ORG PUBL. EST." Then
        LcLinha = LcLinha + "C/APRESENTAÇÃO  "
        DadosTransp.valor.Text = Txt(16).Text
        
     Else
        If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
           LcLinha = " "
           DadosTransp.valor.Text = 0
        Else
           LcLinha = LcLinha + "A VISTA "
           For b = 1 To 11
              LcLinha = LcLinha + " "
           Next
        End If
     End If
     If Natureza.Text = "TRANSFERENCIA" Or Natureza.Text = "DEVOLUCAO" Or Natureza.Text = "RESSARCIMENTO DO ICMS S.T" Then
        LcLinha = ""
     Else
        LcLinha = LcLinha + Format(DadosTransp.valor.Text, "currency")
     End If
     ReDim Preserve MtImpressao(LcQuantiImpressao)
     MtImpressao(LcQuantiImpressao) = LcMargem & Left(LcMargem, Len(LcMargem) / 1.5) & LcLinha + Chr(13)
     LcQuantiImpressao = LcQuantiImpressao + 1
     'Print #FnunNota, LcLinha + Chr(13)
     If Len(LcExtenso2) > 0 Then
        LcLinha = LcLinha & LcExtenso2
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     If Len(LcExtenso3) > 0 Then
        LcLinha = LcLinha & LcExtenso3
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
        'Print #FnunNota, Chr(13)
     End If
     For a = 1 To 3
        ReDim Preserve MtImpressao(LcQuantiImpressao)
        MtImpressao(LcQuantiImpressao) = Chr(13)
        LcQuantiImpressao = LcQuantiImpressao + 1
       ' Print #FnunNota, Chr(13)
     Next
End If

End Function

Function ImprimeBoleto(LcQuantidade As Integer)
On Error Resume Next
Dim a As Integer
Dim LcMargemBo As String
Dim Protesto    As String
Dim protesto1   As String

If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Natureza.Text = "EMPENHO" Then Exit Function
If Natureza.Text = "ORG PUBL. EST." Then Exit Function
LcQuantiImpressaoBoleto = 0
'Open LcBoleto For Output Access Write As #FnunBoleto 'Abre Porta Boleto
Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcMargemBo = LcMargemBo & String(9, " ")

LcEsp = LcEsp & String(81, " ")
LcJuros = "             JUROS DE 5 % AO MES"
LcPag = "             ATE A DATA DO VENCIMENTO PAGAR EM QUALQUER BANCO / QUALQUER AGENCIA"
Protesto = "             NAO RECEBER APOS 4 (QUATRO) DIAS UTEIS DO VENCIMENTO."
protesto1 = "             SUJEITO A PROTESTO"
For a = 1 To LcQuantidade
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = Chr(27) + Chr(48) + Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
   'Print #FnunBoleto, Chr(13)
    LcLinha = Left(LcPag & LcEsp, 90) & DadosTransp.Vencimento(a - 1).Text
    
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = Chr(15) & LcMargemBo & LcLinha & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    
   ' Print #FnunBoleto, Chr(15) & LcMargemBo & LcLinha & Chr(13)
    LcLinha = ""
    For az = 1 To 2
       ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
       MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
       LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
      ' Print #FnunBoleto, Chr(13)
    Next
    
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    
    'Print #FnunBoleto, Chr(13)
    LcLinha = LcLinha & Left(Date & LcEsp, 30) & Txt(0).Text & "/" & Right("00" & a, 2)
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = LcMargemBo & LcLinha & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
   ' Print #FnunBoleto, LcMargemBo & LcLinha & Chr(13)
    Select Case a
        Case Is = 1
            LcLinha = "   " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
        Case Is = 2
            LcLinha = "   " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
        Case Is = 3
            LcLinha = "   " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
    End Select
    'MsgBox LcLinha
    'LcLinha = "   " & Right(LcEsp & Format(DadosTransp.valor.Text, "currency"), 95)
    
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Print #FnunBoleto, Chr(13)
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = LcMargemBo & LcLinha & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    'Print #FnunBoleto, LcMargemBo & LcLinha & Chr(13)
  
    For z = 1 To 3
        ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
        LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
       'Print #FnunBoleto, Chr(13)
    Next
    For z = 1 To 2
       LcLinha = LcLinha & "  "
    Next
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcJuros & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = LigaNegrito
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = "       " & Protesto & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = "       " & protesto1 & Chr(13)
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    MtBoleto(LcQuantiImpressaoBoleto) = DesligaNegrito
    LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
'    Print #FnunBoleto, "       " & LcJuros & Chr(13)
    For z = 1 To 6
       LcLinha = LcMargemBo & LcLinha & "  "
    Next
  '===== Busca Dados Cliente
   If Not RsClientes.EOF Then
      LcLinha = "C.G.C : " & RsClientes!CGC & ""
      
      ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
     ' Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      
      'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
      'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
     'Print #FnunBoleto, Chr(13)
      LcLinha = Left(RsClientes!razaosoc & LcEspC, 50)
      
      ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
      'Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      LcLinha = Left(RsClientes!End & LcEspC, 40)
      
      'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      'MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
      ' Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      
      LcLinha = Trim(LcLinha) & "  " & Left(RsClientes!Bairro & LcEspC, 23)
      If Not RsCidade.EOF Then
         LcLinha = Trim(LcLinha) & "  " & Left(RsCidade!Nome & LcEspC, 30)
      End If
      LcLinha = Trim(LcLinha) & "  " & RsClientes!Estado
      
      ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1

   '   Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
    End If
    For aq = 1 To 6
         ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
        LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1

     '   Print #FnunBoleto, Chr(13)
    Next
Next
'Close #FnunBoleto

End Function

Private Sub Unidade_LostFocus()
On Error Resume Next
Dim a As Long
For a = 0 To LcQUn
    If MtUnidade(a).Simbolo = Unidade.Text Then
       If MtUnidade(a).Quantidade <> 0 Then
          Txt(4).Text = MtUnidade(a).Quantidade
       End If
       Exit For
    End If
Next
End Sub

Private Sub valor_Change(Index As Integer)
On Error Resume Next
If Not LcLimpaValor Then CalculaValores

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
If KeyCode = 122 Then Txt(17).SetFocus
If KeyCode = 13 Then
    SendKeys "{TAB}"
Else
    Teclas (KeyCode)
    LcCalculado = False
    If KeyCode = 123 Then UltimasComprasCliente.Show , Me
    If KeyCode = 113 Then SendKeys "%+{B}"
    If KeyCode = 114 Then SendKeys "%+{F}"
    If KeyCode = 115 Then SendKeys "%+{E}"
    If KeyCode = 118 Then Call Command4_Click
    If KeyCode = 121 Then SendKeys "%+{C}"
End If

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
'===> Verifica se já foi lançado no vale

If Index = 1 Then
    For a = 0 To UBound(LcMat)
        If err.Number > 0 Then Exit For
        If LcMat(a).CodPro = Txt(1).Text Then
           If (UCase(Unidade.Text) <> UCase(LcMat(a).Und)) Or (CLng(Txt(4).Text) <> CLng(LcMat(a).com)) Then
              MsgBox "Para o lançamento de mais itens deste produto na nota fiscal, deverá ser feito na mesma unidade da que já esta lançado.", 64, "Aviso"
              Unidade.SetFocus
              Exit Sub
           End If
        End If
    Next
End If
If Index = 0 Then ConferePreco
If Index = 1 And GlLibera Then
 
      montagrid
 End If
End Sub
Function LiberacaoICms() As Boolean
GlSaidaIcms = False

Dim Rs As Recordset
Dim StrSql As String
StrSql = "Select * from alid001 where codigo='" & Txt(8).Text & "'"
AbreBase
Set Rs = Dbbase.OpenRecordset(StrSql)
If Not Rs.EOF Then
    If Rs!Estado <> "MG" Then
        GlLiberaIcms = False
        LiberaIcms.Show , Me
        Do While Not GlSaidaIcms
            DoEvents
        Loop
        LiberacaoICms = GlLiberaIcms
     Else
        LiberacaoICms = True
     End If
Else
    MsgBox "Cliente não encontrado.", 64, "Cliente não encontrado"
    GlLiberaIcms = False
End If


End Function
Sub ProcessaSintegra()
Dim StrSql As String
Dim Mt() As Tipo50
Dim a As Integer
Dim b As Integer
Dim C As Integer
Dim Achou As Boolean
Dim Rs As Recordset
Dim db As Database
Dim Cnpj As String
Dim Valor_Desconto As Double

Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid001 where codigo='" & Txt(8).Text & "'")
'MsgBox Txt(8).Text
If Not Rs.EOF Then
   If Not IsNull(Rs!CGC) Then
      Cnpj = Replace(Rs!CGC, ".", "")
      Cnpj = Replace(Cnpj, ",", "")
      Cnpj = Replace(Cnpj, "-", "")
      Cnpj = Replace(Cnpj, "/", "")
      Cnpj = Replace(Cnpj, "\", "")
      Cnpj = Trim(Cnpj)
      If Len(Cnpj) = 0 Then
        Cnpj = Replace(Rs!cpf, ".", "")
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Trim(Cnpj)
      End If
   Else
     Cnpj = Replace(Rs!cpf, ".", "")
     Cnpj = Replace(Cnpj, ",", "")
     Cnpj = Replace(Cnpj, "-", "")
     Cnpj = Replace(Cnpj, "/", "")
     Cnpj = Replace(Cnpj, "\", "")
     Cnpj = Trim(Cnpj)
   End If
   Inscricao = Rs!INSCEST
   Inscricao = Replace(Inscricao, ".", "")
   Inscricao = Replace(Inscricao, ",", "")
   Inscricao = Replace(Inscricao, "-", "")
   Inscricao = Replace(Inscricao, "/", "")
   Inscricao = Replace(Inscricao, "\", "")
   Inscricao = Trim(Inscricao)
   Estado = Rs!Estado
   
End If
'==> Insere o cabecalho do Sintegra
StrSql = "Insert into sintegra (data,nf,cfop,valor,Cliente_Forn,Origem) Values ('" & _
       Format(Txt(12).Text, "yyyy-mm-dd") & "','" & _
       Right("000000" & Txt(0).Text, 6) & "','" & _
       Replace(Replace(CFOP.Text, ",", ""), ".", "") & "'," & _
       Replace(Replace(Replace(Replace(Replace(Txt(16).Text, ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
       Txt(8).Text & ",'" & _
       "S" & "')"
'MsgBox strSql
ExecutaSql StrSql

'==> Processa os dados para o 50
a = 0
b = 0
C = 0
'==> Calcula o valor do desconto
If Len(Txt(17).Text) = 0 Then Txt(17).Text = 0

If CDbl(Txt(17).Text) > 0 Then
   Valor_Desconto = CDbl(Txt(17).Text) / (Item.Rows - 1)
   Valor_Desconto = CDbl(AcertaNumero(CStr(Valor_Desconto), 2))
   
End If
For a = 1 To Item.Rows - 1
    Achou = False
    If b = 0 Then
       ReDim Preserve Mt(b)
       Mt(b).icms = Item.TextMatrix(a, 8)
       Mt(b).valor = CDbl(Item.TextMatrix(a, 7))
       b = b + 1
    Else
       For C = 0 To UBound(Mt)
          If Mt(C).icms = Item.TextMatrix(a, 8) Then
             Achou = True
             Exit For
          End If
       Next
       If Achou Then
            Mt(C).valor = CDbl(Item.TextMatrix(a, 7)) + Mt(C).valor
       Else
            ReDim Preserve Mt(b)
            Mt(b).icms = Item.TextMatrix(a, 8)
            Mt(b).valor = CDbl(Item.TextMatrix(a, 7))
            b = b + 1
       End If
    End If
    StrSql = "insert into sintegra_54 (cnpj,modelo,serie,nf,cfop,cst,item," & _
             "codproduto,quantidade,valor_total_bruto,valor_desconto," & _
             "Base_calculo,base_calculo_subst,ipi,Aliquota_icms,data) Values('" & _
             Cnpj & "','" & _
             "01" & "','" & _
             "1  " & "','" & _
             Right("000000" & Txt(0).Text, 6) & "','" & _
             Replace(Replace(CFOP.Text, ",", ""), ".", "") & "','" & _
             Item.TextMatrix(a, 3) & "','" & _
             a & "','" & _
             Item.TextMatrix(a, 1) & "'," & _
             Replace(Item.TextMatrix(a, 5), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(Replace(Item.TextMatrix(a, 7), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
             Replace(Replace(Replace(Replace(Replace(CStr(Valor_Desconto), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
             IIf(CDbl(Item.TextMatrix(a, 8)) > 0, Replace(Replace(Replace(Replace(Replace(CDbl(Item.TextMatrix(a, 7)) - Valor_Desconto, ".", ""), ",", "."), "R", ""), "$", ""), " ", ""), 0) & "," & _
             "0" & "," & _
             "0" & "," & _
             Replace(Item.TextMatrix(a, 8), ",", ".") & ",'" & _
             Format(Txt(12).Text, "yyyy-mm-dd") & "')"
             'MsgBox StrSql
    ExecutaSql StrSql
             
Next
'==> Grava o reguistro 50
If Natureza.Text <> "COMPLEMENTO ICMS" Then
    For a = 0 To UBound(Mt)
        StrSql = "Insert into sintegra_50 (Cnpj,inscricao,data,uf,modelo,serie," & _
                 "nf,cfop,emitente,valortotal,base_calculo_icms,Valor_icms," & _
                 "isenta,outra,aliquota,situacao) values ('" & _
                 Cnpj & "','" & _
                 Inscricao & "','" & _
                 Format(Txt(12).Text, "yyyy-mm-dd") & "','" & _
                 Estado & "','" & _
                 "01" & "','" & _
                 "1   " & "','" & _
                 Right("000000" & Txt(0).Text, 6) & "','" & _
                 Replace(Replace(CFOP.Text, ",", ""), ".", "") & "','" & _
                 "P" & "'," & _
                 Replace(Mt(a).valor, ",", ".") & "," & _
                 IIf(CDbl(Mt(a).icms) > 0, Replace(Mt(a).valor, ",", "."), 0) & "," & _
                 Replace(AcertaNumero(CStr((CDbl(Mt(a).icms) / 100) * Mt(a).valor), 2), ",", ".") & "," & _
                 "0" & "," & _
                 "0" & "," & _
                 Replace(Mt(a).icms, ",", ".") & ",'" & _
                 "N" & "')"
                 'MsgBox StrSql
         ExecutaSql StrSql
                 
    Next
End If
End Sub
Function ValidaEntradaSintegra() As Boolean
On Error GoTo errorVali
Dim Rs As Recordset
Dim db As Database
Dim Estado As String
Dim Cnpj As String
Dim Inscricao As String

Set db = OpenDatabase(GLBase)
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
'   DadosTransp.Txt(9).Text = Rs!dadosnota & ""
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
     Cnpj = Rs!CGC & ""
     Cnpj = Replace(Cnpj, ",", "")
     Cnpj = Replace(Cnpj, ".", "")
     Cnpj = Replace(Cnpj, "-", "")
     Cnpj = Replace(Cnpj, "/", "")
     Cnpj = Replace(Cnpj, "\", "")
     Cnpj = Replace(Cnpj, " ", "")
     Cnpj = Trim(Cnpj)
     If Len(Cnpj) = 0 Then
        Cnpj = Rs!cpf & ""
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, ".", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Replace(Cnpj, " ", "")
        Cnpj = Trim(Cnpj)
     End If
     Inscricao = Rs!INSCEST & ""
     Inscricao = Replace(Inscricao, ",", "")
     Inscricao = Replace(Inscricao, ".", "")
     Inscricao = Replace(Inscricao, "-", "")
     Inscricao = Replace(Inscricao, "/", "")
     Inscricao = Replace(Inscricao, "\", "")
     Inscricao = Replace(Inscricao, " ", "")
     Inscricao = Trim(Inscricao)
     If Len(Cnpj) = 0 Then
        MsgBox "O CNPJ / CPF do cliente não foi cadastrado.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
     If Len(Cnpj) > 11 Then
        If Not Calc_CNPJ(Cnpj) Then
           MsgBox "O CNPJ do cliente é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     Else
        If Not Calc_CPF(Cnpj) Then
           MsgBox "O CPF do cliente é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     End If
     '==> Verifica a Inscricao estadual
     If Len(Inscricao) = 0 Then
        MsgBox "A inscrição Estadual do cliente não foi cadastrada." & Chr(13) & "Caso ele não possua inscrição estadual ou seje pessoa física, casatre como ISENTO.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
     If Consiste(Inscricao, Estado) <> 0 Then
        MsgBox "A Inscrição Estadual do cliente é invalida.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
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
Function ImprimeBoletoCaixa(LcQuantidade As Integer)
On Error Resume Next
Dim a As Integer
Dim LcMargemBo As String
Dim Protesto    As String
Dim protesto1   As String
Dim RsImpressora As Recordset
Dim FnunBoleto As Integer
Dim LcPorta As String
Dim DbbaseLogb As Database
FnunBoleto = FreeFile + 2
Set DbbaseLogb = OpenDatabase(GLBase)
Set RsImpressora = DbbaseLogb.OpenRecordset("select * from impressoras where Impressora='" & GlPortaBoleto & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsImpressora.EOF Then
    LcPorta = RsImpressora!EnderecoLocal
Else
    LcPorta = "Lpt1"
End If
RsImpressora.Close
Set RsImpressora = Nothing
Open LcPorta For Output Access Write As #FnunBoleto
If Natureza.Text = "TRANSFERENCIA" Then Exit Function
If Natureza.Text = "EMPENHO" Then Exit Function
If Natureza.Text = "ORG PUBL. EST." Then Exit Function
LcQuantiImpressaoBoleto = 0
Set RsCidade = Dbbase.OpenRecordset("select * from alid005 where Cod='" & RsClientes!Cidade & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcMargemBo = LcMargemBo & String(3, " ")
LcEsp = LcEsp & String(81, " ")
LcJuros = GlIntrucao1
Protesto = GlIntrucao2
protesto1 = GlIntrucao3
For a = 1 To LcQuantidade
    LcLinha = ""
    For az = 1 To GlLinhasSaltarInicio
       'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
       'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
       Print #FnunBoleto, Chr(13)
       'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Next
    LcLinha = Left(LcEsp, 90) & DadosTransp.Vencimento(a - 1).Text
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = Chr(15) + Chr(27) + Chr(48) & LcMargemBo & LcLinha & Chr(13)
    Print #FnunBoleto, Chr(15) + Chr(27) + Chr(48) & LcMargemBo & LcLinha & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    
    LcLinha = ""
    For az = 1 To 3
        'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
        Print #FnunBoleto, Chr(13)
        'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Next
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
    Print #FnunBoleto, Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    LcLinha = LcLinha & Left(Format(CDate(Date), "dd/mm/yy") & LcEsp, 13) & Txt(0).Text & "/" & Right("00" & a, 2)
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = LcMargemBo & LcLinha & Chr(13)
    Print #FnunBoleto, LcMargemBo & LcLinha & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Select Case a
        Case Is = 1
            LcLinha = "          " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
        Case Is = 2
            LcLinha = "          " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
        Case Is = 3
            LcLinha = "          " & Right(LcEsp & Format(LcValor1, "Standard"), 95)
    End Select
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
    Print #FnunBoleto, Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = LcLinha & Chr(13)
    Print #FnunBoleto, LcLinha & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    For z = 1 To 2
        'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
        Print #FnunBoleto, Chr(13)
        'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Next
    For z = 1 To 2
       LcLinha = LcLinha & "  "
    Next
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = LcJuros & Chr(13)
    Print #FnunBoleto, LcJuros & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = LigaNegrito
    'Print #FnunBoleto, LcJuros & LigaNegrito
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = Protesto & Chr(13)
    Print #FnunBoleto, Protesto & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
    'MtBoleto(LcQuantiImpressaoBoleto) = protesto1 & Chr(13)
    Print #FnunBoleto, protesto1 & Chr(13)
    'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    For z = 1 To 6
        'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        'MtBoleto(LcQuantiImpressaoBoleto) = " "
        Print #FnunBoleto, " "
        'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Next
    For z = 1 To 6
       LcLinha = LcMargemBo & LcLinha & "  "
    Next
  '===== Busca Dados Cliente
   If Not RsClientes.EOF Then
      LcLinha = "C.G.C : " & RsClientes!CGC & ""
      
      'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      'MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
      LcLinha = Left(RsClientes!razaosoc & LcEspC, 50)
      
      'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      'MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
      LcLinha = Left(RsClientes!End & LcEspC, 40)
      LcLinha = Trim(LcLinha) & "  " & Left(RsClientes!Bairro & LcEspC, 23)
      If Not RsCidade.EOF Then
         LcLinha = Trim(LcLinha) & "  " & Left(RsCidade!Nome & LcEspC, 30)
      End If
      LcLinha = Trim(LcLinha) & "  " & RsClientes!Estado
      
      'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
      'MtBoleto(LcQuantiImpressaoBoleto) = "       " & LcMargemBo & LcLinha & Chr(13)
      Print #FnunBoleto, "       " & LcMargemBo & LcLinha & Chr(13)
      'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    End If
    For aq = 1 To GlLinhasSaltarBFim
        'ReDim Preserve MtBoleto(LcQuantiImpressaoBoleto)
        'MtBoleto(LcQuantiImpressaoBoleto) = Chr(13)
        Print #FnunBoleto, Chr(13)
        'LcQuantiImpressaoBoleto = LcQuantiImpressaoBoleto + 1
    Next
Next
Close #FnunBoleto
End Function

