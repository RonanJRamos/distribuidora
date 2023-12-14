VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmEntradaProduto 
   BackColor       =   &H00C7CBBA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Estoque"
   ClientHeight    =   8805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NCM 
      Height          =   285
      Left            =   4080
      TabIndex        =   21
      Top             =   4185
      Width           =   1095
   End
   Begin VB.CommandButton CmdCte 
      Caption         =   "CT-e Referênciado"
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
      TabIndex        =   103
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Modelo 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Protocolo 
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
      Left            =   7320
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Chave 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
   End
   Begin VB.CheckBox EstoqueSeguranca 
      BackColor       =   &H00C7CBBA&
      Caption         =   "Compra Fora do Estado"
      Height          =   495
      Left            =   8400
      TabIndex        =   97
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Desconto 
      Height          =   285
      Left            =   7920
      TabIndex        =   19
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   3435
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Emissao 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Frete 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Seguro 
      Height          =   285
      Left            =   5280
      TabIndex        =   12
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox PIS_COFINS 
      Height          =   285
      Left            =   7440
      TabIndex        =   13
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Complemento 
      Height          =   285
      Left            =   8760
      TabIndex        =   14
      ToolTipText     =   "Não acescentar no total da nota."
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox NaoTributado 
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   3435
      Width           =   1095
   End
   Begin VB.TextBox Custos 
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "Acrescentar no total da nota."
      Top             =   3435
      Width           =   1215
   End
   Begin VB.ComboBox TipoFrete 
      Height          =   315
      ItemData        =   "FrmEntradaProduto.frx":0000
      Left            =   3240
      List            =   "FrmEntradaProduto.frx":000D
      TabIndex        =   11
      Text            =   "2 - FOB"
      Top             =   2625
      Width           =   1215
   End
   Begin VB.TextBox CodigoDaNota 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   8640
      TabIndex        =   87
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdNova 
      Caption         =   "&Nova Nota"
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
      TabIndex        =   86
      Top             =   2250
      Width           =   1575
   End
   Begin VB.CommandButton CmdExcluirNota 
      Caption         =   "Excluir a Nota"
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
      TabIndex        =   85
      Top             =   1875
      Width           =   1575
   End
   Begin VB.TextBox IcmsSubst 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox BaseIcmsSubs 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   3105
      Width           =   1095
   End
   Begin VB.TextBox EPesquisa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   82
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdListaNota 
      Caption         =   "Listar Notas"
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
      TabIndex        =   81
      Top             =   750
      Width           =   1575
   End
   Begin VB.CommandButton CmdPesqisar 
      Caption         =   "&Pesquisar Nota"
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
      TabIndex        =   80
      Top             =   1500
      Width           =   1575
   End
   Begin VB.CheckBox PagEntrega 
      BackColor       =   &H00C7CBBA&
      Caption         =   "Pag. na entrega."
      Height          =   495
      Left            =   5040
      TabIndex        =   33
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox Serie 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Text            =   "1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Cfop 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox Baixa 
      BackColor       =   &H00C7CBBA&
      Caption         =   "Não Entra no Estoque"
      Height          =   495
      Left            =   6840
      TabIndex        =   34
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox SantaMaria1 
      Height          =   375
      Left            =   6960
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox SantaMaria 
      Height          =   375
      Left            =   4800
      TabIndex        =   76
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox California 
      Height          =   375
      Left            =   6000
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox comissao 
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
      Left            =   6480
      TabIndex        =   74
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Index           =   5
      Left            =   3480
      TabIndex        =   72
      Top             =   6240
      Visible         =   0   'False
      Width           =   6135
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
      Left            =   3960
      TabIndex        =   68
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   25
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox tam 
      Height          =   375
      Left            =   4560
      TabIndex        =   66
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
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
      Left            =   1920
      TabIndex        =   63
      Top             =   7920
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
      Left            =   120
      TabIndex        =   62
      Top             =   7920
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
      Height          =   405
      Index           =   14
      Left            =   7080
      ScrollBars      =   2  'Vertical
      TabIndex        =   60
      Top             =   7920
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
      Index           =   13
      Left            =   3960
      TabIndex        =   57
      Top             =   7920
      Width           =   1335
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
      Left            =   5400
      TabIndex        =   56
      Top             =   7920
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
      Left            =   5160
      TabIndex        =   55
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   30
      Top             =   4560
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
      Left            =   2160
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4200
      Width           =   810
   End
   Begin VB.ComboBox Unidade 
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4170
      Width           =   975
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
      TabIndex        =   53
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
      ItemData        =   "FrmEntradaProduto.frx":002F
      Left            =   8040
      List            =   "FrmEntradaProduto.frx":0039
      TabIndex        =   3
      Text            =   "A VISTA"
      Top             =   570
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
      TabIndex        =   50
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
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   2535
      Left            =   120
      TabIndex        =   47
      Top             =   5040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      BackColor       =   -2147483624
      BackColorBkg    =   16777215
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
      TabIndex        =   45
      Top             =   375
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
      TabIndex        =   44
      Top             =   1125
      Width           =   1575
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Cancelar F10"
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
      TabIndex        =   43
      Top             =   3240
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
      TabIndex        =   6
      Top             =   1560
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
      TabIndex        =   32
      Top             =   4560
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
      Left            =   6960
      TabIndex        =   24
      Top             =   4200
      Width           =   690
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
      TabIndex        =   20
      Top             =   4185
      Width           =   3975
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
      Left            =   7080
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   1
      Left            =   10440
      TabIndex        =   29
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   2
      Left            =   9840
      TabIndex        =   28
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox valor 
      Height          =   285
      Index           =   3
      Left            =   9240
      TabIndex        =   27
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Entrada 
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CFOPItem 
      Height          =   285
      Left            =   8640
      TabIndex        =   26
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "NCM"
      Height          =   255
      Left            =   4080
      TabIndex        =   104
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RRRRR"
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
      Left            =   960
      TabIndex        =   102
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
      Height          =   195
      Index           =   23
      Left            =   8640
      TabIndex        =   101
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      Height          =   195
      Index           =   22
      Left            =   3120
      TabIndex        =   100
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protocolo"
      Height          =   195
      Left            =   6240
      TabIndex        =   99
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chave"
      Height          =   195
      Left            =   120
      TabIndex        =   98
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto"
      Height          =   195
      Left            =   6960
      TabIndex        =   96
      Top             =   3480
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emissao"
      Height          =   195
      Index           =   21
      Left            =   2640
      TabIndex        =   95
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frete"
      Height          =   195
      Left            =   120
      TabIndex        =   94
      Top             =   2685
      Width           =   450
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
      Height          =   195
      Left            =   4560
      TabIndex        =   93
      Top             =   2685
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PIS/COFINS"
      Height          =   195
      Left            =   6360
      TabIndex        =   92
      Top             =   2685
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compl. "
      Height          =   195
      Left            =   8160
      TabIndex        =   91
      Top             =   2685
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serviço não Tributado"
      Height          =   195
      Left            =   120
      TabIndex        =   90
      Top             =   3480
      Width           =   2025
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desp. Acessórias"
      Height          =   195
      Left            =   3480
      TabIndex        =   89
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo frete"
      Height          =   195
      Left            =   2160
      TabIndex        =   88
      Top             =   2685
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS Subst"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   3480
      TabIndex        =   84
      Top             =   3150
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Base.Calc.ICMS.Subs"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   83
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Série"
      Height          =   195
      Index           =   18
      Left            =   2040
      TabIndex        =   79
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   78
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inf. Compl."
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
      Index           =   16
      Left            =   3240
      TabIndex        =   73
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% ICMS"
      Height          =   195
      Index           =   15
      Left            =   9240
      TabIndex        =   71
      Top             =   3960
      Width           =   660
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   7080
      Y2              =   8400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Produto Digite Seu Código,  Nome ou pressione F5"
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
      TabIndex        =   70
      Top             =   4560
      Width           =   5025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% IPI"
      Height          =   195
      Index           =   14
      Left            =   9960
      TabIndex        =   69
      Top             =   3960
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Nota"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   1920
      TabIndex        =   65
      Top             =   7680
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Produtos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   64
      Top             =   7680
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI"
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
      Left            =   7080
      TabIndex        =   61
      Top             =   7680
      Width           =   600
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
      Left            =   3960
      TabIndex        =   59
      Top             =   7680
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
      Left            =   5400
      TabIndex        =   58
      Top             =   7680
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Vend"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Selecionar um Fornecedor pressione F5"
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
      TabIndex        =   52
      Top             =   1920
      Width           =   3180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de  Pag"
      Height          =   195
      Index           =   8
      Left            =   6720
      TabIndex        =   51
      Top             =   600
      Width           =   1245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   9960
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   9960
      X2              =   9960
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos Já Lançados"
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   4800
      Width           =   1905
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc."
      Height          =   195
      Left            =   120
      TabIndex        =   46
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   42
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Entrada"
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   41
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. Total"
      Height          =   195
      Index           =   5
      Left            =   10560
      TabIndex        =   67
      Top             =   3960
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V. Unit."
      Height          =   195
      Index           =   4
      Left            =   7680
      TabIndex        =   40
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unid. / Com"
      Height          =   255
      Index           =   3
      Left            =   5505
      TabIndex        =   39
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
      Height          =   195
      Index           =   2
      Left            =   6960
      TabIndex        =   38
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   3990
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notas Fiscais de Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   36
      Top             =   240
      Width           =   4425
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Configurar"
      Begin VB.Menu MnSomas 
         Caption         =   "Somas"
      End
   End
End
Attribute VB_Name = "FrmEntradaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Tipo50
    icms As String
    Valor As Double
End Type

Private LcItem As Long, LcTam As Long
Private FnunNota, FnunBoleto
Private LcNota, LcBoleto, LcEspC As String
Private LcFocus, LcCalculado, LcSalto As Integer
Private LcPrecoVelho As Currency
Private ComNormal, ComAlterado As Long
Private LcLinha As String
Private RsOpcoes As Recordset, RsClientes As Recordset
Private RsCidade As Recordset
Private LcAlteradoCliente, LcAlteradoProduto, LcAlteradoFuncionario As Integer
Private LcMat() As DadosEntrada
Private Liberado, a As Integer
Private PerguntaNNota As Boolean
Private Msgerro As String
Private LcSq As String


Private Sub BaseIcmsSubs_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Cfop_GotFocus()
On Error Resume Next
If NfDuplicada Then
   MsgBox "Nota fiscal já lançada.", 64, "Aviso"
   txt(0).SetFocus
   SendKeys "{End}+{Home}"
End If
End Sub

Private Sub CFOPItem_GotFocus()
On Error Resume Next
SendKeys ("{home}+{End}")
End Sub

Private Sub Chave_Change()
On Error Resume Next
If Len(Chave.Text) > 0 Then
   Modelo.Text = "55"
Else
  Modelo.Text = ""
End If

End Sub

Private Sub CmdCte_Click()
On Error Resume Next
If Len(CodigoDaNota.Text) = 0 Then Exit Sub
FrmCte.Show (1), Me
End Sub

Private Sub CmdExcluir_Click()
FrmExcluiItem.Show
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode <> 116 Then Teclas (KeyCode)
  txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdExcluirNota_Click()
On Error GoTo erroExc
Dim StrSql As String
If MsgBox("Confirma a Exclusão da Nota " & txt(0).Text & "?", vbYesNo, "Aviso") = vbNo Then Exit Sub
Estonar
StrSql = "Delete from entradanf where nf='" & txt(0).Text & "' and data='" & Format(Entrada.Text, "yyyy-mm-dd") & "'"
afetados = ExecutaSql(StrSql)
StrSql = "Delete from itensentradanf where numnf='" & txt(0).Text & "' and data='" & Format(Entrada.Text, "yyyy-mm-dd") & "'"
afetados = ExecutaSql(StrSql)



MsgBox "A Nota foi excluida.", 64, "Aviso"

limpanota
Exit Sub
erroExc:
MsgBox "Ocorreu o seguinte erro excluido a nota :" & Chr(13) & err.Description & " N:" & err.Number, 64, "Aviso"

End Sub

Private Sub CmdFechar_Click()
Unload Me
ReDim LcMat(0)
LcTam = 0
LcItem = 0
End Sub



Function CalculaValores()
Dim LcTotal As Single, LcQuant As Single, LcUnit As Single
On Error Resume Next
If LcCalculado Then Exit Function
LcCalculado = True
'=== Converte os Valores
LcQuant = CCur(txt(3).Text)
LcUnit = CCur(Valor(0).Text)
LcTotal = LcQuant * LcUnit
Valor(1).Text = LcTotal

End Function
Function GeraGrid()
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
Item.ColWidth(0) = 500
Item.ColWidth(1) = 700
Item.ColWidth(2) = 4600
Item.ColWidth(3) = 500
Item.ColWidth(4) = 1000
Item.ColWidth(5) = 850
Item.ColWidth(6) = 1100
Item.ColWidth(7) = 1100
Item.ColWidth(8) = 550
Item.ColWidth(9) = 550
Item.ColWidth(10) = 0
Item.ColWidth(11) = 550
Item.ColWidth(12) = 550

Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Código"
Item.TextMatrix(0, 2) = "Descrição"
Item.TextMatrix(0, 3) = "CST"
Item.TextMatrix(0, 4) = "Unidade"
Item.TextMatrix(0, 5) = "Quant"
Item.TextMatrix(0, 6) = "Unitário"
Item.TextMatrix(0, 7) = "Total"
Item.TextMatrix(0, 8) = "ICMS"
Item.TextMatrix(0, 9) = "IPI"
Item.TextMatrix(0, 11) = "CFOP"
Item.TextMatrix(0, 12) = "NCM"

LcTamanhoGrid = 1
End Function
Function montagrid()
Dim LcAchou, a As Integer
On Error Resume Next
'==== Verifica se Foi digitados todos os campos
If Len(Trim(txt(1).Text)) = 0 Then
   MsgBox "Necessário Informar o Produto.", 48, "Aviso"
   txt(1).SetFocus
   Exit Function
End If
If Len(Trim(txt(3).Text)) = 0 Or (txt(3).Text = "0") Then
   MsgBox "Necessário Informar a Quantidade de Saída.", 48, "Aviso"
   txt(3).SetFocus
   Exit Function
End If
If Len(Trim(Valor(0).Text)) = 0 Or Valor(0).Text = "0" Then
   MsgBox "Necessário Informar o Valor Unitario do Item.", 48, "Aviso"
   Valor(0).SetFocus
   Exit Function
End If
GlLibera = False
If GlArmazenaGalpao Then
  DistribuiMerc.Show , Me
  DistribuiMerc.SetFocus
Else
  GlLibera = True
End If
While Not GlLibera
   DoEvents
Wend
LcItem = LcItem + 1
ReDim Preserve LcMat(LcTam)
LcMat(LcTam).Item = LcItem
LcMat(LcTam).CodPro = txt(1).Text
LcMat(LcTam).produto = txt(2).Text
LcMat(LcTam).Qut = CSng(txt(3).Text)
LcMat(LcTam).Und = Unidade.Text
LcMat(LcTam).CFOP = CFOP.Text
LcMat(LcTam).NCM = NCM.Text
If Len(txt(4).Text) > 0 Then
   LcMat(LcTam).Com = txt(4).Text
Else
   LcMat(LcTam).Com = 1
End If
LcMat(LcTam).VUnit = CCur(Valor(0).Text)
LcMat(LcTam).Vtotal = CCur(Valor(1).Text)
'LcMat(LcTam).cst = cst.Text
LcMat(LcTam).california = CSng(california.Text)
LcMat(LcTam).santamaria = CSng(santamaria.Text)
LcMat(LcTam).santamaria1 = CSng(santamaria1.Text)
If Len(Valor(3).Text) > 0 Then LcMat(LcTam).icms = Valor(3).Text Else LcMat(LcTam).icms = 0
If Len(Valor(2).Text) > 0 Then LcMat(LcTam).ipi = Valor(2).Text Else LcMat(LcTam).ipi = 0

LcTam = LcTam + 1

   
EscreveGrid

For a = 1 To 6
   txt(a).Text = ""
   Valor(a).Text = ""
Next
txt(3).Text = " "
Valor(0).Text = " "
Valor(0).Text = " "
'Custo.Text = "0"
icms.Text = "0"
cst.Text = "0"
minimo.Text = "0"
santamaria.Text = ""
santamaria1.Text = ""
california.Text = ""
CFOPItem.Text = ""
Label11.Caption = ""
NCM.Text = ""
txt(2).SetFocus
End Function
Function limpanota()
On Error Resume Next
Dim a As Integer
Liberado = False
LcTam = 0
ReDim LcMat(0)
Item.Rows = 1
For a = 0 To 16
   txt(a).Text = ""
   Valor(a).Text = ""
Next
Complemento.Text = ""
CodigoDaNota.Text = ""
EPesquisa.Text = ""
LcItem = 0
limite.Text = 0
utilizado.Text = 0
BaseIcmsSubs.Text = 0
IcmsSubst.Text = 0
Frete.Text = ""
Seguro.Text = ""
PIS_COFINS.Text = ""
NaoTributado.Text = ""
Custos.Text = ""
CFOP.Text = ""
Chave.Text = ""
Protocolo.Text = ""
Modelo.Text = ""
Serie.Text = "1"
Entrada.Text = "  /  /  "
emissao.Text = "  /  /  "
'CalculaNumeroNota
'txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
Command3.Enabled = False
CmdSalvar.Enabled = False
CmdExcluir.Enabled = False
CmdExcluirNota.Enabled = False
Entrada.SetFocus

End Function
Function EscreveGrid()
Dim b, a As Integer
b = 1
Item.Rows = 1
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
       Item.Rows = b + 1
       Item.TextMatrix(b, 0) = LcMat(a).Item
       Item.TextMatrix(b, 1) = LcMat(a).CodPro
       Item.TextMatrix(b, 2) = LcMat(a).produto
       Item.TextMatrix(b, 3) = LcMat(a).cst
       Item.TextMatrix(b, 4) = LcMat(a).Und & " C/" & LcMat(a).Com
       Item.TextMatrix(b, 5) = LcMat(a).Qut
       Item.TextMatrix(b, 6) = Format(LcMat(a).VUnit, "Currency")
       Item.TextMatrix(b, 7) = Format(LcMat(a).Vtotal, "Currency")
       Item.TextMatrix(b, 8) = LcMat(a).icms
       Item.TextMatrix(b, 9) = LcMat(a).ipi
       Item.TextMatrix(b, 11) = LcMat(a).CFOP
        Item.TextMatrix(b, 12) = LcMat(a).NCM
       b = b + 1
    End If
Next
CalculaIcms
CalculaTotalSubst
Command3.Enabled = True
CmdSalvar.Enabled = True
CmdExcluir.Enabled = True



End Function
Function CalculaIcms()
Dim LcBaseCalculo, LcIcms, LcPRodutos, LcNota, LcIpi As Currency
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
      If LcMat(a).ipi > 0 Then
         LcIpi = LcIpi + ((LcMat(a).ipi / 100) * LcMat(a).Vtotal)
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
   txt(5).Text = LcComp
End If
LcNota = CCur(LcNota) + CCur(LcIpi)
txt(13).Text = Format(LcBaseCalculo, "Currency")
txt(11).Text = Format(LcIcms, "Currency")
txt(15).Text = Format(LcPRodutos, "Currency")
txt(16).Text = Format(LcNota, "Currency")
txt(14).Text = Format(LcIpi, "currency")

End Function
Function RemontaIndice()
Dim a As Integer
LcItem = 0
For a = 0 To LcTam - 1
   If Len(Trim(LcMat(a).CodPro)) > 0 Then
      LcItem = LcItem + 1
      LcMat(a).Item = LcItem
   End If
Next


End Function
Function CarregaCboUnidade()
On Error Resume Next
Dim LcAchou As Integer
Dim RsUnidade As Recordset
Dim LcPrimeiro As String
AbreBase
Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 order By SIMBOLO", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsUnidade.EOF

   Unidade.AddItem RsUnidade!Simbolo
   RsUnidade.MoveNext
Loop
RsUnidade.Close
Dbbase.Close
Set RsUnidade = Nothing
Set Dbbase = Nothing


End Function
Function calculaunitario()
On Error Resume Next
Valor(0).Text = CSng(txt(4).Text) * PrecoVendaNormal
minimo.Text = CSng(txt(4).Text) * PrecoMimimodeVendaAlterado
End Function
Function BuscaProduto(LcTipo As Integer)
On Error Resume Next

On Error GoTo errBuscaFor
Dim RsProduto As ADODB.Recordset

Dim LcValorDigitado
Dim LcCodigo As String
'If Not LcAlteradoProduto Then Exit Function
AbreBase
Set RsProduto = AbreRecordset("select * from produtos", True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcCalculado = True
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "Codigo=" & txt(2).Text
         RsProduto.Find LcCriterioCli
         If Not RsProduto.EOF Then
            If RsProduto!Desativado Then
               Label11.Caption = "PRODUTO DESATIVADO"
            Else
               Label11.Caption = ""
            End If
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            txt(1).Text = RsProduto!codigo
            txt(2).Text = RsProduto!Nome
            Unidade.Text = LcUnidade
            txt(4).Text = RsProduto!QtdMedida
            cst.Text = RsProduto!cst
            NCM.Text = RsProduto!ClassificacaoFiscal & ""
               
            'Custo.Text = RsProduto!Custo
            LcAchou = True
            'SendKeys "{TAB}"
       Else
            txt(2).Text = ""
      End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = txt(2).Text
        If Len(txt(2).Text) = 0 Then Exit Function
        lcchave = txt(2).Text
        If GLCalculacodigoProduto Then
            
            If IsNumeric(lcchave) Then
                LcCriterioCli = "Codigo=" & lcchave
            Else
                LcCriterioCli = "nome='" & lcchave & "'"
            End If
        Else
           If IsNumeric(lcchave) Then
             LcCriterioCli = "Codigo=" & lcchave
           Else
             LcCriterioCli = "nome='" & lcchave & "'"
           End If
        End If
        RsProduto.Find LcCriterioCli
        If Not RsProduto.EOF Then
             If RsProduto!Desativado Then
               Label11.Caption = "PRODUTO DESATIVADO"
            Else
               Label11.Caption = ""
            End If
            Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!UnidMedida & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)

            'RsUnidade.FindFirst LcCriterio
            If Not RsUnidade.EOF Then
                LcUnidade = RsUnidade!Simbolo
            End If
            txt(1).Text = RsProduto!codigo & ""
            txt(2).Text = RsProduto!Nome & ""
            Unidade.Text = LcUnidade & ""
            txt(4).Text = RsProduto!QtdMedida & ""
            cst.Text = RsProduto!cst & ""
            NCM.Text = RsProduto!ClassificacaoFiscal & ""
            'SendKeys "{TAB}"
        Else
            txt(2).Text = LcValorDigitado
            GlCriterioSql = "select * From produtos where nome like '" & UCase(txt(2).Text) & "%'  order by nome"
            'FrmBuscaProduto.Tag = Txt(2).Text
            FrmPesquisaProdutos.txt.Text = txt(2).Text
            If LcAlteradoProduto Then
               PerguntaNNota = True
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
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   
   Exit Function
End If




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
         LcCriterioCli = "CODIGO='" & txt(10).Text & "'"
         RsVendedor.FindFirst LcCriterioCli
         If Not RsVendedor.NoMatch Then
            txt(7).Text = RsVendedor!Nome
            If CLng(Comissao.Text) <> 1 Then
                Comissao.Text = RsVendedor!Comissao
            End If
            SendKeys "{TAB}"
         Else
            txt(7).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = txt(7).Text
        If Len(txt(7).Text) = 0 Then Exit Function
        
        lcchave = Right("00000" & txt(7).Text, 5)
        LcCriterioCli = "CODIGO='" & lcchave & "'"
        RsVendedor.FindFirst LcCriterioCli
        If Not RsVendedor.NoMatch Then
            txt(7).Text = RsVendedor!Nome
            txt(10).Text = RsVendedor!codigo
            
            If Len(Comissao.Text) > 0 Then
               If CSng(Comissao.Text) <> 1 Then
                  Comissao.Text = RsVendedor!Comissao
               End If
            Else
              Comissao.Text = RsVendedor!Comissao
            End If
            'SendKeys "{TAB}"
        Else
            txt(7).Text = LcValorDigitado
            FrmPesquisaFuncionarios.txt.Text = txt(7).Text
            GlCriterioSql = "select * From alid200 where nome like '" & UCase(txt(7).Text) & "*'  order by nome"
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
Function BuscaCliente(LcTipo As Integer)
On Error GoTo errBuscaFor
Dim rsCliente As Recordset
Dim LcValorDigitado As String
Dim LcCodigo As String
Dim LcCredito, LcUtilizado As Currency

If Not LcAlteradoCliente Then Exit Function
AbreBase
GlLibera = False
Set rsCliente = Dbbase.OpenRecordset("select * from alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & txt(8).Text & "'"
         rsCliente.FindFirst LcCriterioCli
         If Not rsCliente.NoMatch Then
            txt(9).Text = rsCliente!razaosoc
            LcDesCidade = rsCliente!razaosoc
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
                   txt(9).Text = ""
                   txt(9).SetFocus
                Else
                   Liberado = True
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
            End If
            SendKeys "{TAB}"
         Else
            txt(9).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = txt(9).Text
        If Len(txt(9).Text) = 0 Then Exit Function
      
        lcchave = Right("00000" & txt(9).Text, 5)
        LcCriterioCli = "CODIGO='" & lcchave & "'"
        rsCliente.FindFirst LcCriterioCli
        If Not rsCliente.NoMatch Then
            txt(9).Text = rsCliente!razaosoc
            txt(8).Text = rsCliente!codigo
            LcDesCidade = rsCliente!razaosoc
            
            'SendKeys "{TAB}"
        Else
            txt(9).Text = LcValorDigitado
            FrmPesquisaCliente.txt.Text = txt(9).Text
            GlCriterioSql = "select * From alid002 where RAZAOSOC like '" & UCase(txt(9).Text) & "*'  order by RAZAOSOC"
            If LcAlteradoCliente Then
                PesquisandoNota = True
                PerguntaNNota = True
               FrmPesquisaFornecedores.Show , Me
               Do While PesquisandoNota
                  DoEvents
               Loop
               PerguntaNNota = False
               LcAlteradoCliente = False
            End If
            'Data(1).SetFocus
        End If
  
End Select

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
   Resume Next
End If

End Function

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 116 Then Teclas (KeyCode)
 txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdListaNota_Click()
On Error Resume Next
ListaNotaEntrada.Show , Me
End Sub

Private Sub CmdNova_Click()
On Error Resume Next
limpanota
End Sub

Private Sub CmdPesqisar_Click()
On Error Resume Next
'Dim StrNota As String
'StrNota = InputBox("Entre com o numero da nota fiscal.", "Localizar nota fiscal.")
'BuscaNota StrNota
PesquisaNfEntrada.Show , Me
End Sub

Private Sub CmdSalvar_Click()
On Error GoTo errosalva

conexaoAdo.BeginTrans
If SalvaNota Then
   conexaoAdo.CommitTrans
   MsgBox "nota Salva com sucesso.", 64, "Aviso"
Else
   conexaoAdo.RollbackTrans
End If
Exit Sub

errosalva:
MsgBox err.Description & err.Number
'Resume 0
conexaoAdo.RollbackTrans
MsgBox "Ocorreu um Erro Lançando a nota." & Chr(13) & "Todos os lançamentos foram cancelados para manter a integridade do sistema.", 64, "Aviso"

End Sub

Private Sub Command1_Click()
FrmPesquisaCliente.Show , Me
End Sub

Private Sub Command2_Click()
FrmPesquisaProdutos.Show , Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If PagEntrega.Value = 1 Then
    If Len(txt(0).Text) = 0 Then
        MsgBox "Informe o numero da nota fiscal.", 64, "Aviso"
        txt(0).SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txt(0).Text) Then
        MsgBox "O numero da nota fiscal deve ser numerico.", 64, "Aviso"
        txt(0).SetFocus
        Exit Sub
    End If
    If Not IsDate(Entrada.Text) Then
        MsgBox "Informe uma data Valida para a entrada.", 64, "Aviso"
        Entrada.SetFocus
        SendKeys "+{Home}+{End}"
        Exit Sub
    End If
        If Not IsDate(emissao.Text) Then
        MsgBox "Informe uma data Valida para a emissão.", 64, "Aviso"
        emissao.SetFocus
        SendKeys "+{Home}+{End}"
        Exit Sub
    End If

'    If Len(txt(7).Text) = 0 Then
'        MsgBox "Informe o vendedor da entrada.", 64, "Aviso"
'        txt(7).SetFocus
'        SendKeys "+{Home}+{End}"
'        Exit Sub
'    End If
    If Len(txt(9).Text) = 0 Then
        MsgBox "Informe o Fornecedor da entrada.", 64, "Aviso"
        txt(9).SetFocus
        SendKeys "+{Home}+{End}"
        Exit Sub
    End If
    If Len(CFOP.Text) = 0 Then
        MsgBox "Informe CFOP da entrada.", 64, "Aviso"
        CFOP.SetFocus
        SendKeys "+{Home}+{End}"
        Exit Sub
    End If
    If Len(Serie.Text) = 0 Then
        MsgBox "Informe Seride da nota de entrada.", 64, "Aviso"
        Serie.SetFocus
        SendKeys "+{Home}+{End}"
        Exit Sub
    End If
    If Item.Rows = 1 Then
       MsgBox "É nescessario lançar no minimo um item para confirmar a nota fiscal.", 64, "Aviso"
       txt(2).SetFocus
       SendKeys "+{Home}+{End}"
       Exit Sub
    End If
    If Not ValidaEntradaSintegra Then
       'Txt(0).SetFocus
       Exit Sub
    End If
End If
If Len(txt(0).Text) = 0 Then
   LcResp = MsgBox("O numero da Nota fiscal não foi informado." & Chr(13) & "Deseja que o Sistema Calcule um Número para esta Nota?", vbCritical + vbYesNo, "Aviso")
   If LcResp = vbNo Then
      MsgBox "Operação Cancelada.", 64, "Aviso"
      Exit Sub
   Else
      CalculaNumeroNota
   End If
End If
If NfDuplicada Then
   MsgBox "O Número da nota fiscal informado já foi lançada para este fornecedor.", 64, "Aviso"
   txt(0).SetFocus
   Exit Sub
End If
tam.Text = LcTam
DadosEntradaNota.Show , Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode <> 116 Then Teclas (KeyCode)
  txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Complemento_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Complemento_LostFocus()
CalculaTotalSubst
End Sub

Private Sub Custos_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Custos_LostFocus()
CalculaTotalSubst
End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Desconto_LostFocus()
CalculaTotalSubst
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
If Not GlCarregado Then
   Set GlFormA = Me
  ' txt(12).Text = Format(GlDataSistema, "dd/mm/yy")
   GlCarregado = True
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
Me.Height = 9585
Me.Width = 12000
GlEscolhe = 1
'CalculaNumeroNota
CarregaCboUnidade

End Sub

Private Sub Form_Unload(Cancel As Integer)
GlCarregado = False
End Sub

Private Sub Frete_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Frete_LostFocus()
CalculaTotalSubst
End Sub

Private Sub IcmsSubst_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub IcmsSubst_LostFocus()

CalculaTotalSubst
End Sub
Sub CalculaTotalSubst()
On Error Resume Next
Dim ValorNota As Double
Dim valorIcms As Double
Dim valorFrete As Double
Dim valorSeguro As Double
Dim valorPIS_COFINS As Double
Dim valorComplementar As Double
Dim valorNaoTributado As Double
Dim valorDespAce As Double
Dim ValorDesconto As Double

Dim ValorProduto As Double
Dim ValorIpi As Double
Dim LcCaminho As String
Dim StrFrete As Boolean
Dim StrSeguro As Boolean
Dim StrPIS_COFINS As Boolean
Dim StrComplementar As Boolean
Dim StrValorIcmsSubst As Boolean
Dim StrNaoTributado As Boolean
Dim StrDespAce As Boolean


LcCaminho = App.Path & "\configMeiaFolha.ini"

StrFrete = IIf(Len(LeIni("Soma", "Frete", LcCaminho)) = 0, 0, LeIni("Soma", "Frete", LcCaminho))
StrSeguro = IIf(Len(LeIni("Soma", "Seguro", LcCaminho)) = 0, 0, LeIni("Soma", "Seguro", LcCaminho))
StrPIS_COFINS = IIf(Len(LeIni("Soma", "PIS_COFINS", LcCaminho)) = 0, 0, LeIni("Soma", "PIS_COFINS", LcCaminho))
StrComplementar = IIf(Len(LeIni("Soma", "Complementar", LcCaminho)) = 0, 0, LeIni("Soma", "Complementar", LcCaminho))
StrValorIcmsSubst = IIf(Len(LeIni("Soma", "ValorIcmsSubst", LcCaminho)) = 0, 0, LeIni("Soma", "ValorIcmsSubst", LcCaminho))
StrNaoTributado = IIf(Len(LeIni("Soma", "NaoTributado", LcCaminho)) = 0, 0, LeIni("Soma", "NaoTributado", LcCaminho))
StrDespAce = IIf(Len(LeIni("Soma", "DespAce", LcCaminho)) = 0, 0, LeIni("Soma", "DespAce", LcCaminho))


If Len(txt(14).Text) = 0 Then txt(14).Text = 0

If Len(IcmsSubst.Text) = 0 Then IcmsSubst.Text = 0
If Not IsNumeric(IcmsSubst.Text) Then IcmsSubst.Text = 0

If Len(Frete.Text) = 0 Then Frete.Text = 0
If Not IsNumeric(Frete.Text) Then Frete.Text = 0

If Len(Seguro.Text) = 0 Then Seguro.Text = 0
If Not IsNumeric(Seguro.Text) Then Seguro.Text = 0

If Len(PIS_COFINS.Text) = 0 Then PIS_COFINS.Text = 0
If Not IsNumeric(PIS_COFINS.Text) Then PIS_COFINS.Text = 0

If Len(Complemento.Text) = 0 Then Complemento.Text = 0
If Not IsNumeric(Complemento.Text) Then Complemento.Text = 0

If Len(NaoTributado.Text) = 0 Then NaoTributado.Text = 0
If Not IsNumeric(NaoTributado.Text) Then NaoTributado.Text = 0

If Len(Custos.Text) = 0 Then Custos.Text = 0
If Not IsNumeric(Custos.Text) Then Custos.Text = 0

If Len(Desconto.Text) = 0 Then Desconto.Text = 0
If Not IsNumeric(Desconto.Text) Then Desconto.Text = 0

If StrValorIcmsSubst Then valorIcms = CDbl(IcmsSubst.Text) Else valorIcms = 0
If StrFrete Then valorFrete = CDbl(Frete.Text) Else valorFrete = 0
If StrSeguro Then valorSeguro = CDbl(Seguro.Text) Else valorSeguro = 0
If StrPIS_COFINS Then valorPIS_COFINS = CDbl(PIS_COFINS.Text) Else valorPIS_COFINS = 0
If StrComplementar Then valorComplementar = CDbl(Complemento.Text) Else valorComplementar = 0
If StrNaoTributado Then valorNaoTributado = CDbl(NaoTributado.Text) Else valorNaoTributado = 0
If StrDespAce Then valorDespAce = CDbl(Custos.Text) Else valorDespAce = 0

ValorDesconto = CDbl(Desconto.Text)
ValorProduto = CDbl(txt(15).Text)
ValorIpi = CDbl(txt(14).Text)

ValorNota = ValorProduto + valorIcms + ValorIpi + valorFrete + valorSeguro + valorPIS_COFINS + valorComplementar + valorNaoTributado + valorDespAce - ValorDesconto

txt(16).Text = AcertaNumero(CStr(ValorNota), 2)
End Sub

Private Sub MnSomas_Click()
On Error Resume Next
ConfiguraSomaEntrada.Show , Me
End Sub

Private Sub NaoTributado_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub NaoTributado_LostFocus()
CalculaTotalSubst
End Sub

Private Sub PIS_COFINS_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub PIS_COFINS_LostFocus()
CalculaTotalSubst
End Sub

Private Sub Seguro_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Seguro_LostFocus()
CalculaTotalSubst
End Sub

Private Sub txt_GotFocus(Index As Integer)
If Index = 9 Then
   If Len(txt(8).Text) > 0 Then
      txt(8).Text = Right("00000" & txt(8).Text, 5)
      If Len(Trim(txt(8).Text)) > 0 Then BuscaCliente (2)
   End If
End If
If Index = 8 Then
   If Len(Trim(txt(10).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Vendedor Responsável.", 64, "Aviso"
      txt(10).SetFocus
   End If
End If
If Index = 1 Then
   If Len(Trim(txt(8).Text)) = 0 Then
      MsgBox "É Necessário Escolher o Cliente para a Nota Fiscal.", 64, "Aviso"
      
   End If
End If
If Index = 9 Then LcAlteradoCliente = False
If Index = 2 Then
   LcAlteradoProduto = False
 If Not PerguntaNNota Then
   '===> Verifica se foi informado o numeor da nf
   If Len(txt(0).Text) = 0 Then
      LcResp = MsgBox("O número da NF não foi informado," & Chr(13) & "Deseja calcular um número para este Lançamento?", vbCritical + vbYesNo, "Aviso")
      If LcResp = vbNo Then
         txt(0).SetFocus
         Exit Sub
      Else
         CalculaNumeroNota
      End If
   End If
   '===> Verifica se e nota duplicata
   If NfDuplicada And Len(EPesquisa.Text) = 0 Then
      MsgBox "Já foi lançada uma nota com este número para este fornecedor.", 64, "Aviso"
      txt(0).SetFocus
      Exit Sub
   End If
 Else
'   txt(7).SetFocus
 End If
End If

If Index = 7 Then LcAlteradoFuncionario = False
End Sub

Function VoltaCampo(LcIndex As Integer)
Select Case LcIndex
   Case Is = 12
       txt(0).SetFocus
   Case Is = 8
       Natureza.SetFocus
   Case Is = 9
       txt(8).SetFocus
   Case Is = 1
       txt(8).SetFocus
   Case Is = 2
      txt(1).SetFocus
   Case Is = 4
     txt(2).SetFocus
   Case Is = 3
     txt(4).SetFocus
   Case Is = 5
     txt(3).SetFocus
   Case Is = 6
     txt(5).SetFocus
End Select

End Function

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
   VoltaCampo (KeyCode)
End If
If KeyCode = 117 Then FrmDescicaoProduto.Show , Me
If KeyCode = 116 Or KeyCode = 13 Then
   If Index = 8 Or Index = 9 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(txt(9).Text)) > 0 Then
            FrmPesquisaFornecedores.txt.Text = txt(9).Text
            GlCriterioSql = "select * From alid002 where RAZAOSOC like '" & UCase(txt(9).Text) & "*'  order by RAZAOSOC"
         Else
            GlCriterioSql = ""
         End If
         KeyCode = 116
      Teclas (KeyCode)
   Else
      If Index = 1 Or Index = 2 Then 'Exibe Produtos
         GlEscolhe = 2
         If Len(Trim(txt(2).Text)) > 0 Then
            FrmPesquisaProdutos.Tag = txt(2).Text
            LcAlteradoProduto = True
            GlCriterioSql = "select * From produtos where nome like '" & UCase(txt(2).Text) & "%'  order by nome"
         Else
            GlCriterioSql = ""
         End If
         FrmPesquisaProdutos.Show , Me
         Exit Sub
         'Teclas (KeyCode)
      End If
    End If
Else
  Teclas (KeyCode)
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
If Index = 6 And GlLibera Then montagrid
If Index = 0 Then
   If Not IsNumeric(txt(0).Text) Then
      If Len(txt(0).Text) = 0 Then Exit Sub
      MsgBox "O Código da Nota de Entrada Deve ser Numérico.", 64, "Aviso"
      txt(0).Text = ""
      'CalculaNumeroNota
      txt(0).SetFocus
      Exit Sub
   Else
      'txt(0).Text = Right("00000000" & txt(0).Text, 8)
   End If
End If
If Index = 1 Then
   If Len(Trim(txt(1).Text)) > 0 Then
      txt(1).Text = Right("00000" & txt(1).Text, 5)
      BuscaProduto (2)
   End If
End If
If Index = 2 Then
   If IsNumeric(txt(2).Text) Then
      LcAlteradoProduto = True
      BuscaProduto (1)
      txt(3).SetFocus
   Else
       BuscaProduto (2)
   End If
   
End If
If Index = 8 Then
   If Len(txt(8).Text) > 0 Then
      txt(8).Text = Right("00000" & txt(8).Text, 5)
      If Len(Trim(txt(8).Text)) > 0 Then BuscaCliente (2)
   End If
End If
'If Index = 4 Then calculaunitario

If Index = 5 Then
   ConferePreco
End If
If Index = 9 Then BuscaCliente (2)
If Index = 7 Then BuscaVendendor (2)
If Index = 2 Then If Len(txt(2).Text) = 0 Then BuscaProduto (2)
If Index = 10 And Len(Trim(txt(Index).Text)) <> 0 Then BuscaVendendor (2)
If Index = 14 Then
   On Error Resume Next
LcCalculado = False
If Index = 3 Or Index = 5 Then CalculaValores
If Index = 9 Then LcAlteradoCliente = True
If Index = 2 Then LcAlteradoProduto = True
If Index = 7 Then LcAlteradoFuncionario = True
If Index = 14 Then
    Dim ValorNota As Double
    Dim ValorIpi As Double
    
    If Len(IcmsSubst.Text) = 0 Then IcmsSubst.Text = 0
    If Len(txt(14).Text) = 0 Then txt(14).Text = 0
    ValorNota = CDbl(txt(15).Text) + CDbl(txt(14).Text) + CDbl(IcmsSubst.Text)
    
    txt(16).Text = AcertaNumero(CStr(ValorNota), 2)

End If
End If
End Sub
Function ConferePreco()
On Error Resume Next
End Function
Function ExcluiItem(LcNItem As Integer)
Dim a, b As Integer
On Error Resume Next
For a = 0 To LcTam - 1
    If LcMat(a).Item = LcNItem Then
       LcMat(a).CodPro = ""
       LcMat(a).Item = 0
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
Dim LcSql As String, LcNumeroNota As String
Dim RsNota As ADODB.Recordset
LcSql = "Select * from entradanf order by NF"
'AbreBase
Set RsNota = AbreRecordset(LcSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsNota.EOF Then
   LcNumeroNota = "000001"
Else
   RsNota.MoveLast
   LcNumeroNota = Right("000000" & CStr(Val(RsNota("NF")) + 1), 6)
End If
txt(0).Text = LcNumeroNota

RsNota.Close
'Dbbase.Close
Set RsNota = Nothing
'Set Dbbase = Nothing

End Function

Private Sub Unidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 117 Then FrmDescicaoProduto.Show , Me
If KeyCode <> 116 Then Teclas (KeyCode)
End Sub
Function DadosOk() As Boolean
Msgerro = ""
If Len(txt(0).Text) = 0 Then
   Msgerro = "É nescessário informar o numero da nota fiscal."
End If
If Not IsDate(emissao.Text) Then
   If Len(Msgerro) > 0 Then Msgerro = Msgerro & "chr(13)"
   Msgerro = "A data de emissão é invalida."
End If
If Not IsDate(Entrada.Text) Then
   If Len(Msgerro) > 0 Then Msgerro = Msgerro & "chr(13)"
   Msgerro = "A data de entrada é invalida."
End If
If Len(txt(9).Text) = 0 Then
    If Len(Msgerro) > 0 Then Msgerro = Msgerro & "chr(13)"
    Msgerro = "É nescessário informar o Fornecedor."
End If
If Len(CFOP.Text) = 0 Then
    If Len(Msgerro) > 0 Then Msgerro = Msgerro & "chr(13)"
    Msgerro = "É nescessário informar o CFOP."
End If
If Len(Serie.Text) = 0 Then
    Serie.Text = 1
End If
If Len(Frete.Text) = 0 Then
    Frete.Text = 0
End If
If Len(Seguro.Text) = 0 Then
    Seguro.Text = 0
End If
If Len(PIS_COFINS.Text) = 0 Then
    PIS_COFINS.Text = 0
End If
If Len(Complemento.Text) = 0 Then
    Complemento.Text = 0
End If
If Len(NaoTributado.Text) = 0 Then
    NaoTributado.Text = 0
End If
If Len(Custos.Text) = 0 Then
    Custos.Text = 0
End If
If Len(BaseIcmsSubs.Text) = 0 Then
    BaseIcmsSubs.Text = 0
End If
If Len(IcmsSubst.Text) = 0 Then
    IcmsSubst.Text = 0
End If
If Len(txt(13).Text) = 0 Then
    txt(13).Text = 0
End If
If Len(txt(11).Text) = 0 Then
    txt(11).Text = 0
End If
If Len(txt(14).Text) = 0 Then
    txt(14).Text = 0
End If
If Len(Msgerro) > 0 Then
  DadosOk = False
Else
  DadosOk = True
End If

End Function
Function Estonar() As Boolean
Dim StrSql As String
Dim RsDados As ADODB.Recordset
Dim RsHistorico As ADODB.Recordset
Dim LcSanta As Double
Dim LcSanta1 As Double
Dim LcCalifornia As Double
Dim LcQuantBaixa As Double
Dim LcEstoqueSeguranca As Double
StrSql = "Select * from itensentradanf where CodigoNota=" & CodigoDaNota.Text

Set RsDados = AbreRecordset(StrSql, True)

Do Until RsDados.EOF
  StrSql = "Select * from historicoproduto where produto=" & RsDados!Item & " And nf='" & txt(0).Text & "' and tipo='E' and clienteforn='" & txt(9).Text & "'"
  Set RsHistorico = AbreRecordset(StrSql, True)
  If Not RsHistorico.EOF Then
    If Not IsNull(RsHistorico!santa) Then
       LcSanta = RsHistorico!santa
    Else
       LcSanta = 0
    End If
     
     LcSanta1 = RsHistorico!Santa2
     LcCalifornia = RsHistorico!california
  End If
  LcQuantBaixa = LcSanta + LcSanta1 + LcCalifornia
  '==> Atualiza a tb de produtos.
  
  StrSql = "Update Produtos Set " & _
         "QuantEstoque=QuantEstoque-" & Replace(LcQuantBaixa, ",", ".") & _
         ",Santa1=Santa1-" & Replace(LcSanta, ",", ".") & _
          ",santa2=santa2-" & Replace(LcSanta1, ",", ".") & _
          ",California=California-" & Replace(LcCalifornia, ",", ".")
          If EstoqueSeguranca.Value = 1 Then
             StrSql = StrSql & ",EstoqueSeguranca=EstoqueSeguranca-" & Replace(LcQuantBaixa, ",", ".")
          End If
         StrSql = StrSql & " where codigo=" & RsDados!Item
         
  ' Debug.Print StrSql
  ExecutaSql StrSql
  StrSql = "delete from historicoproduto where produto=" & RsDados!Item & " And nf='" & txt(0).Text & "' and tipo='E' and clienteforn='" & txt(9).Text & "'"
  ExecutaSql StrSql
  RsDados.MoveNext
Loop

Set RsDados = Nothing


End Function
Function NotaJaLancada() As Boolean
On Error GoTo erroNotaLancada
Dim RsNota As ADODB.Recordset
Dim StrSql As String

StrSql = "Select * from entradanf where nf='" & txt(0).Text & "' and clicred='" & txt(8).Text & "'"
If Len(CodigoDaNota) > 0 Then
   '==>esta editando a nota fiscal.
   NotaJaLancada = False
   Exit Function
End If

Set RsNota = AbreRecordset(StrSql, True)

If Not RsNota.EOF Then
   NotaJaLancada = True
Else
  NotaJaLancada = False
End If

Set RsNota = Nothing

Exit Function
erroNotaLancada:
NotaJaLancada = False


End Function
Function SalvaNota() As Boolean
On Error GoTo ErrSalva
Dim RsNotaFiscal        As ADODB.Recordset
'Dim RsItens             As ADODB.Recordset
'Dim RsEstoque           As Recordset
Dim rsCliente           As Recordset
'Dim RsProduto           As ADODB.Recordset
Dim RsUnidade           As Recordset
Dim RsGalpao            As Recordset
Dim LcPerAtualizacao    As Double
Dim LcCom               As Double
Dim LcPerDe             As Double
Dim a                   As Double
Dim LcCusto             As Double
Dim LCLEtra             As String
Dim x                   As Double
Dim LcPerc              As String
Dim LcMesmaUnid         As Boolean
Dim LCLanca             As Double
Dim LcUni               As Double
Dim LcTemp              As Double
Dim Cestoque            As ControleDb
Dim LcUnimed            As String
Dim StrSql              As String
Dim SemNF               As Boolean
'Dim Lc
'LcSql1 = "Select * from entradanf"
'LcSql2 = "Select * from itensentradanf"
LcSql3 = "Select * from Alid002"
LcSql4 = "Select * from produtos"
LcSql5 = "Select * from Alid004"
LcSql6 = "Select * from Alid013"

'==== Grava Os dados da Nota Fiscal
'=> Efetua a Validação dos Dados doigitados.
If Not DadosOk Then
   MsgBox "Verifique os seguintes erros:" & Chr(13) & Chr(13) & Msgerro, 64, "Aviso"
   SalvaNota = False
   Exit Function
End If

'==> verifica se a nota ja foi lancada para o fornecedor.
If NotaJaLancada Then
   MsgBox "A nota fiscal já foi lançada para o fornecedor.", 64, "Aviso."
   SalvaNota = False
   Exit Function
End If
SemNF = EstoqueSeguranca.Value

AbreBase
'Set RsNotaFiscal = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsItens = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsFornecedor = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsProduto = Dbbase.OpenRecordset(LcSql4, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsUnidade = Dbbase.OpenRecordset(LcSql5, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsEstoque = Dbbase.OpenRecordset(LcSql6, dbOpenDynaset, dbSeeChanges, dbOptimistic)


If Natureza.Text = "A VISTA" Then LcNatureza = "V" Else LcNatureza = "P"
GlNomeMaquina = ""
NomeMaquina
'==> vamos verificar se é inclusao ou alteração
'txt(0).Text = Right("00000000" & txt(0).Text, 8)
If Len(CodigoDaNota.Text) = 0 Then
   '==> é inclusao
   StrSql = "insert into entradanf(Nf,data,Emissao,vp,CLICRED,valor,valorproduto,RECDESP," & _
            "BaseIcms,icms,ipi,Complementar,Sintegra,cfop,BaseIcmsSubst,IcmsSubst," & _
            "Frete,Seguro,PIS_COFINS,NaoTributado,DespesasAcessorias," & _
            "SubSerie,Maquina,TipoFrete,Desconto,CompraFora,chave,protocolo,modelo) Values ('" & _
            txt(0).Text & "','" & _
            Format(Entrada.Text, "yyyy-mm-dd") & "','" & _
            Format(emissao.Text, "yyyy-mm-dd") & "','" & _
            LcNatureza & "','" & _
            txt(8).Text & "'," & _
            Replace(Replace(Replace(Replace(txt(16).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(txt(15).Text, ".", ""), "R", ""), "$", ""), ",", ".") & ",'" & _
            "D" & "'," & _
            Replace(Replace(Replace(Replace(txt(13).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(txt(11).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(txt(14).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(Complemento.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            PagEntrega.Value & ",'" & _
            Replace(Replace(Replace(CFOP.Text, ".", ""), ",", ""), "-", "") & "'," & _
            Replace(Replace(Replace(Replace(BaseIcmsSubs.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(IcmsSubst.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(Frete.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(Seguro.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(PIS_COFINS.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            Replace(Replace(Replace(Replace(NaoTributado.Text, ".", ""), "R", ""), "$", ""), ",", ".") & ","
   
   StrSql = StrSql & Replace(Replace(Replace(Replace(Custos.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            "'1" & "','" & _
            GlNomeMaquina & "'," & _
            Mid(TipoFrete.Text, 1, 1) & "," & _
            Replace(Replace(Replace(Replace(Desconto.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            EstoqueSeguranca.Value & "," & _
            "'" & Chave.Text & "'," & _
            "'" & Protocolo.Text & "'," & _
            "'" & Modelo.Text & "')"
                
            
            
Else
   '==>é Alteração
             
   StrSql = "Update entradanf set " & _
           "nf='" & txt(0).Text & "'," & _
           "data='" & Format(Entrada.Text, "yyyy-mm-dd") & "'," & _
           "Emissao='" & Format(emissao.Text, "yyyy-mm-dd") & "'," & _
           "vp='" & LcNatureza & "'," & _
           "CLICRED='" & txt(8).Text & "'," & _
           "valor=" & Replace(Replace(Replace(Replace(txt(16).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "valorproduto=" & Replace(Replace(Replace(Replace(txt(15).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "RECDESP='D" & "'," & _
           "BaseIcms=" & Replace(Replace(Replace(Replace(txt(13).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "icms=" & Replace(Replace(Replace(Replace(txt(11).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "ipi=" & Replace(Replace(Replace(Replace(txt(14).Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "Complementar=" & Replace(Replace(Replace(Replace(Complemento.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "Sintegra=" & PagEntrega.Value & "," & _
           "cfop='" & Replace(Replace(Replace(CFOP.Text, ".", ""), ",", ""), "-", "") & "'," & _
           "BaseIcmsSubst=" & Replace(Replace(Replace(Replace(BaseIcmsSubs.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "IcmsSubst=" & Replace(Replace(Replace(Replace(IcmsSubst.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "Frete=" & Replace(Replace(Replace(Replace(Frete.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "Seguro=" & Replace(Replace(Replace(Replace(Seguro.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "PIS_COFINS=" & Replace(Replace(Replace(Replace(PIS_COFINS.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
           "NaoTributado=" & Replace(Replace(Replace(Replace(NaoTributado.Text, ".", ""), "R", ""), "$", ""), ",", ".") & ","
   
   StrSql = StrSql & _
            "DespesasAcessorias=" & Replace(Replace(Replace(Replace(Custos.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            "Desconto=" & Replace(Replace(Replace(Replace(Desconto.Text, ".", ""), "R", ""), "$", ""), ",", ".") & "," & _
            "SubSerie=" & "'1" & "'," & _
            "Maquina='" & GlNomeMaquina & "'," & _
             "Chave='" & Chave.Text & "'," & _
             "Protocolo='" & Protocolo.Text & "'," & _
             "modelo='" & Modelo.Text & "'," & _
             "TipoFrete=" & Mid(TipoFrete.Text, 1, 1) & _
             " Where Codigo=" & CodigoDaNota.Text
  
End If
Debug.Print StrSql
'MsgBox StrSql
LcRegistrosAfetados = ExecutaSql(StrSql)


If LcRegistrosAfetados < 1 Then
   err.Raise vbObjectError + 513, "Erro Lançando dados da Nota Fiscal. ", "Erro Lançando dados da Nota Fiscal."
   GoTo ErrSalva
Else
   '==> Recupera o Codigo caso seja inclusao.
   If Len(CodigoDaNota.Text) = 0 Then
      StrSql = "Select * from entradanf order by codigo desc limit 0,5"
      Debug.Print StrSql
      Set RsNotaFiscal = AbreRecordset(StrSql, True)
      
      If Not RsNotaFiscal.EOF Then
         CodigoDaNota.Text = RsNotaFiscal!codigo
      End If
      Set RsNotaFiscal = Nothing
   Else
     '==>estorna os itens anteriores.
     Estonar
     
   End If

End If
'==> Efetua a exclusao dos itens da nota
StrSql = "Delete from itensentradanf where CodigoNota=" & CodigoDaNota.Text
afetados = ExecutaSql(StrSql)

Set Cestoque = New ControleDb
For a = 0 To LcTam - 1
    If Len(Trim(LcMat(a).CodPro)) > 0 Then
         Cestoque.CodProduto = LcMat(a).CodPro
         Cestoque.CodClien_forn = txt(8).Text
         Cestoque.NF = txt(0).Text
         
         LcCritUnidade = "simbolo='" & LcMat(a).Und & "'"
         RsUnidade.FindFirst LcCritUnidade
         If Not RsUnidade.NoMatch Then
            LcUnimed = RsUnidade!cod
         End If
         LcVUnit = Replace(LcMat(a).VUnit, ".", "")
         LcQut = Replace(LcMat(a).Qut, ".", "")
         LcVtotal = Replace(LcMat(a).Vtotal, ".", "")
         LcIcms = Replace(LcMat(a).icms, ".", "")
         LcMatipi = Replace(LcMat(a).ipi, ".", "")

         LcVUnit = Replace(LcVUnit, ",", ".")
         LcQut = Replace(LcQut, ",", ".")
         LcVtotal = Replace(LcVtotal, ",", ".")
         LcIcms = Replace(LcIcms, ",", ".")
         LcMatipi = Replace(LcMatipi, ",", ".")
         
         LcSq = "insert into itensentradanf (numnf,item,QTDE,valunit,unimed,qtdum,descricao,valortotal,icms,ipi,fornecedor,data,cfop,serie,CodigoNota,Santa,California,NCM) values ('"
         LcSq = LcSq & txt(0).Text & "','" & LcMat(a).CodPro & "'," & LcQut & ","
         LcSq = LcSq & LcVUnit & ",'" & LcUnimed & "','" & CLng(LcMat(a).Com) & "','"
         LcSq = LcSq & Cestoque.RetiraCaracter(LcMat(a).produto) & "'," & LcVtotal & ","
         LcSq = LcSq & LcIcms & "," & LcMatipi & ",'" & txt(9).Text & "','" & Format(Entrada.Text, "yyyy-mm-dd") & "','"
         LcSq = LcSq & CFOPItem & "','"
         LcSq = LcSq & Serie.Text & "',"
         LcSq = LcSq & CodigoDaNota.Text & ","
         LcSq = LcSq & Replace(LcMat(a).santamaria, ",", ".") & ","
         LcSq = LcSq & Replace(LcMat(a).california, ",", ".") & ",'"
         LcSq = LcSq & LcMat(a).NCM & "')"
         Debug.Print LcSq
         LcRegistrosAfetados = ExecutaSql(LcSq)

        If LcRegistrosAfetados < 1 Then
           err.Raise vbObjectError + 513, "Erro Lançando o Item" & LcMat(a).CodPro, "Erro Lançando o Item" & LcMat(a).CodPro & " :" & DEscricaoErro
           GoTo ErrSalva
        Else
           Dim Str_Atualiza_prod As String
           Str_Atualiza_prod = "update produtos set classificacaofiscal='" & LcMat(a).NCM & "' where codigo=" & LcMat(a).CodPro
           LcRegistrosAfetados = ExecutaSql(Str_Atualiza_prod)
        End If

         
         'RsItens.AddNew
         LcValoripi = (LcMat(a).ipi / 100) * LcMat(a).VUnit
 
         If Baixa = 0 Then
            Dim Cl As New ControleEstoque
            Cl.EntradaEstoque CLng(LcMat(a).CodPro), CCur(LcMat(a).santamaria) * CCur(LcMat(a).Com), CCur(LcMat(a).california) * CCur(LcMat(a).Com), txt(0).Text, txt(9).Text, LcMat(a).Und, 0
            Dim RsEstoque As ADODB.Recordset
            StrSql = "Select Custo from Produtos where codigo=" & CLng(LcMat(a).CodPro)
            Set RsEstoque = AbreRecordset(StrSql, True)
            If Not RsEstoque.EOF Then
                If CDec(RsEstoque!Custo) < CDec(Replace(CStr(LcVUnit), ".", ",")) Then
                    LcSq = "Update Produtos set Custo=" & LcVUnit & ",CustoTotal=" & LcVUnit & " where codigo=" & CLng(LcMat(a).CodPro)
                    LcRegistrosAfetados = ExecutaSql(LcSq)
                End If
            Else
                LcSq = "Update Produtos set Custo=" & LcVUnit & ",CustoTotal=" & LcVUnit & " where codigo=" & CLng(LcMat(a).CodPro)
                LcRegistrosAfetados = ExecutaSql(LcSq)
            End If
            
           
        End If
    End If
Next
'==== Atualiza Dados Cliente
LcCriterioPes = "codigo='" & txt(8).Text & "'"
RsFornecedor.FindFirst LcCriterioPes
If Not RsFornecedor.NoMatch Then
   RsFornecedor.Edit
   RsFornecedor("ULTCOMPRA") = CDate(Entrada.Text)
   RsFornecedor.Update
End If
SalvaNota = True
Saida:
On Error Resume Next
'=== Fecha as Bases
RsNotaFiscal.Close
RsItens.Close
RsFornecedor.Close
RsProduto.Close
Dbbase.Close
Set RsNotaFiscal = Nothing
Set RsComissao = Nothing
Set RsFornecedor = Nothing
Set RsProduto = Nothing
Set Dbbase = Nothing
Exit Function

ErrSalva:
LcResp = MsgBox("Ocorreu o Seguinte erro salvando a nota:" & Chr(13) & Chr(13) & err.Description & Chr(13) & Chr(13) & "O que deseja fazer?", vbCritical + vbRetryCancel, "Erro nº:" & err.Number)
If LcResp = 4 Then
   Resume 0
Else
  SalvaNota = False
End If
MsgBox err.Description & err.Number

'Resume 0

End Function
Sub BuscaNota(CodigoNota As String, CodigoFornecedor As String, Optional NumerodaNotaFiscal As String)
On Error Resume Next
Dim RsFor As Recordset
Dim RsUnidade As Recordset
Dim db As Database
Dim RsNota As ADODB.Recordset
Dim RsItem As ADODB.Recordset
Dim StrSql As String

StrSql = "Select * from entradanf where codigo=" & CodigoNota

Set RsNota = AbreRecordset(StrSql)
Set db = OpenDatabase(GLBase)

If RsNota.EOF Then
   MsgBox "A nota fiscal " & NumerodaNotaFiscal & " não foi encontrada.", 64, "Aviso"
   Set RsNota = Nothing
   Exit Sub
End If
CodigoDaNota.Text = RsNota!codigo
txt(0).Text = RsNota!NF & ""
Entrada.Text = Format(RsNota!Data, "dd/mm/yy") & ""
emissao.Text = Format(RsNota!emissao, "dd/mm/yy") & ""
Natureza.Text = IIf(RsNota!vp = "V", "A VISTA", "A PRAZO")
Complemento.Text = RsNota!Complementar & ""
Modelo.Text = RsNota!Modelo & ""
Protocolo.Text = RsNota!Protocolo & ""
Chave.Text = RsNota!Chave & ""
EstoqueSeguranca.Value = RsNota!CompraFora
Set RsFor = db.OpenRecordset("Select * from alid002 where CODIGO='" & RsNota!clicred & "'")
txt(8).Text = RsNota!clicred & ""
If Not RsFor.EOF Then
   txt(9).Text = RsFor!razaosoc & ""
Else
   txt(9).Text = ""
End If
Set RsFor = Nothing
CFOP.Text = RsNota!CFOP & ""
Serie.Text = RsNota!Serie & ""
txt(5).Text = RsNota!Complementar & ""

StrSql = "Select * from itensentradanf where codigonota=" & CodigoNota & " order by codigo"
Set RsItem = AbreRecordset(StrSql)

Item.Rows = 1
LcItem = 0
LcTam = 0
Do Until RsItem.EOF
   Set RsUnidade = db.OpenRecordset("Select * from alid004 where cod='" & RsItem!UNIMED & "'")
   If Not RsUnidade.EOF Then
      Unidade = RsUnidade!Simbolo & ""
   End If
   LcItem = LcItem + 1
    ReDim Preserve LcMat(LcTam)
    LcMat(LcTam).Item = LcItem
    LcMat(LcTam).CodPro = RsItem!Item & ""
    LcMat(LcTam).produto = RsItem!Descricao & ""
    LcMat(LcTam).Qut = RsItem!Qtde & ""
    LcMat(LcTam).Und = Unidade
    LcMat(LcTam).Com = RsItem!QTDUM & ""
    LcMat(LcTam).VUnit = CCur(RsItem!VALUNIT)
    LcMat(LcTam).Vtotal = CCur(RsItem!ValorTotal)
    'LcMat(LcTam).cst = cst.Text
    LcMat(LcTam).california = 0
    LcMat(LcTam).santamaria = 0
    LcMat(LcTam).santamaria1 = 0
    LcMat(LcTam).icms = RsItem!icms & ""
    LcMat(LcTam).ipi = RsItem!ipi & ""
    LcMat(LcTam).CFOP = RsItem!CFOP & ""
    LcMat(LcTam).NCM = RsItem!NCM & ""
    If IsNumeric(RsItem!santa) Then
       LcMat(LcTam).santamaria1 = RsItem!santa & ""
    Else
       LcMat(LcTam).santamaria1 = 0
    End If
    If IsNumeric(RsItem!california) Then
       LcMat(LcTam).california = RsItem!california & ""
    Else
       LcMat(LcTam).california = 0
    End If
    LcTam = LcTam + 1

   RsItem.MoveNext
Loop
txt(14).Text = RsNota!ipi & ""
txt(16).Text = RsNota!Valor
BaseIcmsSubs.Text = IIf(Not IsNull(RsNota!BaseIcmsSubst), RsNota!BaseIcmsSubst, 0)
IcmsSubst.Text = IIf(Not IsNull(RsNota!IcmsSubst), RsNota!IcmsSubst, 0)
Frete.Text = RsNota!Frete & ""
Seguro.Text = RsNota!Seguro & ""
PIS_COFINS.Text = RsNota!PIS_COFINS & ""
NaoTributado.Text = RsNota!NaoTributado & ""
Custos.Text = RsNota!DespesasAcessorias & ""
Desconto.Text = RsNota!Desconto & ""

CmdSalvar.Enabled = True
CmdExcluirNota.Enabled = True
CmdExcluir.Enabled = True

Set RsNota = Nothing
Set RsItem = Nothing
EscreveGrid
CalculaTotalSubst
Command3.Enabled = False
End Sub
Function LancarFinanceiroCTE(LcNumeroContas As Integer, NumeroNF As String _
, CodFornecedor As String, CodTipoMonet As String, LcValor As Double) As Boolean
On Error GoTo errat
Dim RsContasPagar As Recordset, RsCaixa As Recordset
Dim RsTipoMonetario As Recordset
Dim LcTipoMOne As String

Dim a As Integer
LcSql1 = "Select * from Alid014"
LcSql2 = "Select * from Alid016"
LcSql3 = "Select * from Alid008"

AbreBase
Set RsContasPagar = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCaixa = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsTipoMonetario = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case Natureza.Text
    Case Is = "A VISTA"
         If GlVistaEntrada Then
            RsContasPagar.AddNew
            RsContasPagar("NF") = NumeroNF
            RsContasPagar("CREDOR") = CodFornecedor ' Txt(8).Text
            LcCriterioPes = "XTPMONET='" & CodTipoMonet & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               RsContasPagar("TPMONET") = RsTipoMonetario("TPMONET")
            End If
            RsContasPagar("VALOR") = CCur(LcValor)
            RsContasPagar("DATA") = CDate(Entrada.Text)
            RsContasPagar("DTVENC") = CDate(Entrada.Text)
            RsContasPagar("DTPAGTO") = CDate(Entrada.Text)
            RsContasPagar("VALPAGO") = CCur(LcValor)
            RsContasPagar("TIPORD") = "R"
            RsContasPagar("Acrescimo") = 0
            RsContasPagar("AutoNumericoNF") = CodigoDaNota.Text
            RsContasPagar.Update
          End If
          
         If GlCaixaEntrada Then
           ' RsCaixa.AddNew
          '  RsCaixa("NF") = Txt(0).Text
          '  RsCaixa("RECDESP") = "D"
          '  RsCaixa("CLICRED") = Txt(8).Text
            LcCriterioPes = "XTPMONET='" & CodTipoMonet & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMOne = RsTipoMonetario("TPMONET")
            End If
            LcValor = CCur(LcValor)
            If GlEntradaVista Then Call lancacaixa("Despesas", txt(0).Text, LcTipoMOne, LcValor)

           ' RsCaixa("VALOR") = CCur(Txt(16).Text)
           ' RsCaixa("DATA") = CDate(Txt(12).Text)
           ' RsCaixa.Update
          End If
    Case Is = "A PRAZO"
         If GlFaturaEntrada Then
            For a = 1 To LcNumeroContas
                RsContasPagar.AddNew
                RsContasPagar("NF") = NumeroNF & "/" & Right("00" & CStr(a), 2) ' Txt(0).Text & "/" & Right("00" & CStr(a), 2)
                RsContasPagar("credor") = CodFornecedor ' Txt(8).Text
                LcCriterioPes = "XTPMONET='" & CodTipoMonet & "'"
                RsTipoMonetario.FindFirst LcCriterioPes
                If Not RsTipoMonetario.NoMatch Then
                    RsContasPagar("TPMONET") = RsTipoMonetario("TPMONET")
                End If
                RsContasPagar("DATA") = CDate(Entrada.Text)
                RsContasPagar("VALOR") = CCur(LcValor)
                Select Case a
                    Case Is = 1
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(0).Text)
                    Case Is = 2
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(1).Text)
                    Case Is = 3
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(2).Text)
                    Case Is = 4
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(3).Text)
                    Case Is = 5
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(4).Text)
                    Case Is = 6
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(5).Text)
                   Case Is = 7
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(6).Text)
                    Case Is = 8
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(7).Text)
                   Case Is = 9
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(8).Text)
                   Case Is = 10
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(9).Text)
                   Case Is = 11
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(10).Text)
                   Case Is = 12
                         RsContasPagar("DTVENC") = CDate(FrmCte.Vencimento(11).Text)
                End Select
                If GlCaixaEntrada Then
                  ' RsCaixa.AddNew
                  '  RsCaixa("NF") = Txt(0).Text
                  '  RsCaixa("RECDESP") = "D"
                  '  RsCaixa("CLICRED") = Txt(8).Text
                    LcCriterioPes = "XTPMONET='" & FrmCte.TipoMonetario.Text & "'"
                    RsTipoMonetario.FindFirst LcCriterioPes
                    If Not RsTipoMonetario.NoMatch Then
                       LcTipoMOne = RsTipoMonetario("TPMONET")
                    End If
                    LcValor = CDbl(txt(16).Text)
                    If GlEntradaPrazo Then Call lancacaixa("Despesas", txt(0).Text & "/" & Right("00" & CStr(a), 2), LcTipoMOne, LcValor)

                   ' RsCaixa("VALOR") = CCur(Txt(16).Text)
                   ' RsCaixa("DATA") = CDate(Txt(12).Text)
                   ' RsCaixa.Update
                End If
                RsContasPagar("TIPORD") = "D"
                RsContasPagar("Acrescimo") = 0
                RsContasPagar("AutoNumericoNF") = CodigoDaNota.Text
                RsContasPagar.Update
            Next
          End If
        
         
End Select
'Atualizacaixa = True

Exit Function
errat:
LcResp = MsgBox("Ocorreu o Seguinte erro Lançando a(s) Conta(s):" & Chr(13) & Chr(13) & err.Description & Chr(13) & Chr(13) & "O que deseja fazer?", vbCritical + vbRetryCancel, "Erro nº:" & err.Number)
If LcResp = 4 Then
   Resume 0
Else
  'Atualizacaixa = False
End If

End Function

Function Atualizacaixa(LcNumeroContas As Integer) As Boolean
On Error GoTo errat
Dim RsContasPagar As Recordset, RsCaixa As Recordset
Dim RsTipoMonetario As Recordset
Dim LcTipoMOne As String
Dim LcValor As Double
Dim a As Integer
LcSql1 = "Select * from Alid014"
LcSql2 = "Select * from Alid016"
LcSql3 = "Select * from Alid008"

AbreBase
Set RsContasPagar = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsCaixa = Dbbase.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsTipoMonetario = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case Natureza.Text
    Case Is = "A VISTA"
         If GlVistaEntrada Then
            RsContasPagar.AddNew
            RsContasPagar("NF") = txt(0).Text
            RsContasPagar("CREDOR") = txt(8).Text
            LcCriterioPes = "XTPMONET='" & DadosEntradaNota.TipoMonetario.Text & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               RsContasPagar("TPMONET") = RsTipoMonetario("TPMONET")
            End If
            RsContasPagar("VALOR") = CCur(txt(16).Text)
            RsContasPagar("DATA") = CDate(Entrada.Text)
            RsContasPagar("DTVENC") = CDate(Entrada.Text)
            RsContasPagar("DTPAGTO") = CDate(Entrada.Text)
            RsContasPagar("VALPAGO") = CCur(txt(16).Text)
            RsContasPagar("TIPORD") = "R"
            RsContasPagar("Acrescimo") = 0
            RsContasPagar.Update
          End If
          
         If GlCaixaEntrada Then
           ' RsCaixa.AddNew
          '  RsCaixa("NF") = Txt(0).Text
          '  RsCaixa("RECDESP") = "D"
          '  RsCaixa("CLICRED") = Txt(8).Text
            LcCriterioPes = "XTPMONET='" & DadosEntradaNota.TipoMonetario.Text & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMOne = RsTipoMonetario("TPMONET")
            End If
            LcValor = CDbl(txt(16).Text)
            If GlEntradaVista Then Call lancacaixa("Despesas", txt(0).Text, LcTipoMOne, LcValor)

           ' RsCaixa("VALOR") = CCur(Txt(16).Text)
           ' RsCaixa("DATA") = CDate(Txt(12).Text)
           ' RsCaixa.Update
          End If
    Case Is = "A PRAZO"
         If GlFaturaEntrada Then
            For a = 1 To LcNumeroContas
                RsContasPagar.AddNew
                RsContasPagar("NF") = txt(0).Text & "/" & Right("00" & CStr(a), 2)
                RsContasPagar("credor") = txt(8).Text
                LcCriterioPes = "XTPMONET='" & DadosEntradaNota.TipoMonetario.Text & "'"
                RsTipoMonetario.FindFirst LcCriterioPes
                If Not RsTipoMonetario.NoMatch Then
                    RsContasPagar("TPMONET") = RsTipoMonetario("TPMONET")
                End If
                RsContasPagar("DATA") = CDate(Entrada.Text)
                RsContasPagar("VALOR") = CCur(DadosEntradaNota.Valor.Text)
                Select Case a
                    Case Is = 1
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(0).Text)
                    Case Is = 2
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(1).Text)
                    Case Is = 3
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(2).Text)
                    Case Is = 4
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(3).Text)
                    Case Is = 5
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(4).Text)
                    Case Is = 6
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(5).Text)
                   Case Is = 7
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(6).Text)
                    Case Is = 8
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(7).Text)
                   Case Is = 9
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(8).Text)
                   Case Is = 10
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(9).Text)
                   Case Is = 11
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(10).Text)
                   Case Is = 12
                         RsContasPagar("DTVENC") = CDate(DadosEntradaNota.Vencimento(11).Text)
                End Select
                If GlCaixaEntrada Then
                  ' RsCaixa.AddNew
                  '  RsCaixa("NF") = Txt(0).Text
                  '  RsCaixa("RECDESP") = "D"
                  '  RsCaixa("CLICRED") = Txt(8).Text
                    LcCriterioPes = "XTPMONET='" & DadosEntradaNota.TipoMonetario.Text & "'"
                    RsTipoMonetario.FindFirst LcCriterioPes
                    If Not RsTipoMonetario.NoMatch Then
                       LcTipoMOne = RsTipoMonetario("TPMONET")
                    End If
                    LcValor = CDbl(txt(16).Text)
                    If GlEntradaPrazo Then Call lancacaixa("Despesas", txt(0).Text & "/" & Right("00" & CStr(a), 2), LcTipoMOne, LcValor)

                   ' RsCaixa("VALOR") = CCur(Txt(16).Text)
                   ' RsCaixa("DATA") = CDate(Txt(12).Text)
                   ' RsCaixa.Update
                End If
                RsContasPagar("TIPORD") = "D"
                RsContasPagar("Acrescimo") = 0
                RsContasPagar.Update
            Next
          End If
        
         
End Select
Atualizacaixa = True

Exit Function
errat:
LcResp = MsgBox("Ocorreu o Seguinte erro Lançando a(s) Conta(s):" & Chr(13) & Chr(13) & err.Description & Chr(13) & Chr(13) & "O que deseja fazer?", vbCritical + vbRetryCancel, "Erro nº:" & err.Number)
If LcResp = 4 Then
   Resume 0
Else
  Atualizacaixa = False
End If

End Function


Private Sub valor_Change(Index As Integer)
CalculaValores
End Sub

Private Sub valor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   ' SendKeys "{TAB}"
Else
    Teclas (KeyCode)
    LcCalculado = False
End If

End Sub

Private Sub valor_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub
Function VerificaVenda()
On Error Resume Next
Dim Ce As ControleDb
Set Ce = New ControleDb

'Dim Rs As Recordset

If Len(Trim(Valor(0).Text)) = 0 Then Exit Function
'AbreBase
'Set Rs = AbreRecordsetLeitura("Select * from produtos where codigo=" & txt(1).Text) ' , dbOpenDynaset, dbSeeChanges, dbOptimistic)
'If Not Rs.EOF Then
Ce.CodProduto = txt(1).Text
 LcOldV = Valor(0).Text
 If CDbl(Ce.PrecoDeCusto) < CDbl(FrmEntradaProduto.Valor(0).Text) Then
      atualizapreco.Show , Me
End If
Set Ce = Nothing
'Rs.Close
'Dbbase.Close
End Function
Private Sub valor_LostFocus(Index As Integer)
If Index = 0 Then
   If Len(CFOP.Text) > 0 Then CFOPItem.Text = CFOP.Text
   VerificaVenda
   ConferePreco
End If
If Index = 1 Then montagrid
End Sub
Function processanota()
On Error GoTo errprocessa
Dim LcQuant As Integer
conexaoAdo.BeginTrans
'ProcessaSintegra
If Len(DadosEntradaNota.quantidade.Text) > 0 Then LcQuant = DadosEntradaNota.quantidade.Text Else LcQuant = 0
If SalvaNota Then
   If Atualizacaixa(LcQuant) Then
       'txt(0).SetFocus
       limpanota
   Else
      GoTo errprocessa
   End If
Else
   GoTo errprocessa
End If
conexaoAdo.CommitTrans

Exit Function
errprocessa:
MsgBox err.Description & err.Number
'Resume 0
conexaoAdo.RollbackTrans
MsgBox "Ocorreu um Erro Lançando a nota." & Chr(13) & "Todos os lançamentos foram cancelados para manter a integridade do sistema.", 64, "Aviso"

End Function
Function NfDuplicada() As Boolean
On Error Resume Next
Dim LcSql As String
Dim Rs As ADODB.Recordset
'AbreBase
'Exit Function
LcSql = "Select * from entradanf where NF='" & UCase(txt(0).Text) & "' and CLICRED='" & txt(8).Text & "'"
Set Rs = AbreRecordset(LcSql, True)
If Not Rs.EOF Then
    NfDuplicada = True
Else
    NfDuplicada = False
End If
Set Rs = Nothing
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
Dim CNPJ As String

Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid002 where codigo='" & txt(8).Text & "'")

If Not Rs.EOF Then
   If Not IsNull(Rs!CGC) Then
      CNPJ = Replace(Rs!CGC, ".", "")
      CNPJ = Replace(CNPJ, ",", "")
      CNPJ = Replace(CNPJ, "-", "")
      CNPJ = Replace(CNPJ, "/", "")
      CNPJ = Replace(CNPJ, "\", "")
      CNPJ = Replace(CNPJ, " ", "")
      CNPJ = Trim(CNPJ)
   Else
     CNPJ = Replace("", ".", "")
     CNPJ = Replace(CNPJ, ",", "")
     CNPJ = Replace(CNPJ, "-", "")
     CNPJ = Replace(CNPJ, "/", "")
     CNPJ = Replace(CNPJ, "\", "")
     CNPJ = Replace(CNPJ, " ", "")
     CNPJ = Trim(CNPJ)
   End If
   Inscricao = Rs!INSCEST & ""
   Inscricao = Replace(Inscricao, ".", "")
   Inscricao = Replace(Inscricao, ",", "")
   Inscricao = Replace(Inscricao, "-", "")
   Inscricao = Replace(Inscricao, "/", "")
   Inscricao = Replace(Inscricao, "\", "")
   Inscricao = Replace(Inscricao, " ", "")
   Inscricao = Trim(Inscricao)
   Estado = Rs!Estado & ""
   
End If
'==> Insere o cabecalho do Sintegra
StrSql = "Insert into sintegra (data,nf,cfop,valor,Cliente_Forn,origem) Values ('" & _
       Format(Entrada.Text, "yyyy-mm-dd") & "','" & _
       Right("000000" & txt(0).Text, 6) & "','" & _
       Replace(Replace(CFOP.Text, ",", ""), ".", "") & "'," & _
       Replace(Replace(Replace(Replace(Replace(txt(16).Text, ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
       txt(8).Text & ",'" & _
       "E" & "')"
'MsgBox strSql
ExecutaSql StrSql

'==> Processa os dados para o 50
a = 0
b = 0
C = 0
For a = 1 To Item.Rows - 1
    Achou = False
    If b = 0 Then
       ReDim Preserve Mt(b)
       Mt(b).icms = Item.TextMatrix(a, 8)
       Mt(b).Valor = CDbl(Item.TextMatrix(a, 7))
       b = b + 1
    Else
       For C = 0 To UBound(Mt)
          If Mt(C).icms = Item.TextMatrix(a, 8) Then
             Achou = True
             Exit For
          End If
       Next
       If Achou Then
            Mt(C).Valor = CDbl(Item.TextMatrix(a, 7)) + Mt(C).Valor
       Else
            ReDim Preserve Mt(b)
            Mt(b).icms = Item.TextMatrix(a, 8)
            Mt(b).Valor = CDbl(Item.TextMatrix(a, 7))
            b = b + 1
       End If
    End If
    StrSql = "insert into sintegra_54 (cnpj,modelo,serie,nf,cfop,cst,item," & _
             "codproduto,quantidade,valor_total_bruto,valor_desconto," & _
             "Base_calculo,base_calculo_subst,ipi,Aliquota_icms,data) Values('" & _
             CNPJ & "','" & _
             "01" & "','" & _
             Serie.Text & "','" & _
             Right("000000" & txt(0).Text, 6) & "','" & _
             Replace(Replace(CFOP.Text, ",", ""), ".", "") & "','" & _
             Item.TextMatrix(a, 3) & "'," & _
             a & ",'" & _
             Item.TextMatrix(a, 1) & "'," & _
             Replace(Item.TextMatrix(a, 5), ",", ".") & "," & _
             Replace(Replace(Replace(Replace(Replace(Item.TextMatrix(a, 7), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
             "0" & "," & _
             IIf(CDbl(Item.TextMatrix(a, 8)) > 0, Replace(Replace(Replace(Replace(Replace(Item.TextMatrix(a, 7), ".", ""), ",", "."), "R", ""), "$", ""), " ", ""), 0) & "," & _
             "0" & "," & _
             IIf(CDbl(Item.TextMatrix(a, 9)) > 0, Replace(Replace(Replace(Replace(Item.TextMatrix(a, 9), ",", "."), "R", ""), "$", ""), " ", ""), 0) & "," & _
             Replace(Item.TextMatrix(a, 8), ",", ".") & ",'" & _
             Format(Entrada.Text, "yyyy-mm-dd") & "')"
            ' MsgBox strSql
    ExecutaSql StrSql
             
Next

'==> Grava o reguistro 50
For a = 0 To UBound(Mt)
    StrSql = "Insert into sintegra_50 (Cnpj,inscricao,data,uf,modelo,serie," & _
             "nf,cfop,emitente,valortotal,base_calculo_icms,Valor_icms," & _
             "isenta,outra,aliquota,situacao) values ('" & _
             CNPJ & "','" & _
             Inscricao & "','" & _
             Format(Entrada.Text, "yyyy-mm-dd") & "','" & _
             Estado & "','" & _
             "01" & "','" & _
             Serie.Text & "','" & _
             Right("000000" & txt(0).Text, 6) & "','" & _
             Replace(Replace(CFOP.Text, ",", ""), ".", "") & "','" & _
             "P" & "'," & _
             Replace(Mt(a).Valor, ",", ".") & "," & _
             IIf(CDbl(Mt(a).icms) > 0, Replace(Mt(a).Valor, ",", "."), 0) & "," & _
             Replace(AcertaNumero(CStr((CDbl(Mt(a).icms) / 100) * Mt(a).Valor), 2), ",", ".") & "," & _
             "0" & "," & _
             "0" & "," & _
             Replace(Mt(a).icms, ",", ".") & ",'" & _
             "N" & "')"
            ' MsgBox strSql
     ExecutaSql StrSql
             
Next

End Sub

Function ValidaEntradaSintegra() As Boolean
On Error GoTo errorVali
Dim Rs As Recordset
Dim db As Database
Dim Estado As String
Dim CNPJ As String
Dim Inscricao As String

Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid002 where codigo='" & txt(8).Text & "'")

If Rs.EOF Then
   ValidaEntradaSintegra = False
   MsgBox "Fornecedor não encontrado.", 64, "Aviso"
Else
   '==> Verifica o Estado do Fornecedor
   If IsNull(Rs!Estado) Then
      Estado = ""
   Else
      Estado = UCase(Rs!Estado)
   End If
   If Len(Estado) = 0 Then
      ValidaEntradaSintegra = False
      MsgBox "O Estado do Fornecedor não foi cadastrado." & Chr(13) & "cadastre-o antes de entrar com a nota fiscal.", 64, "Aviso"
      GoTo Saida
   Else
     '==> Verifica se o Cfop é Valido
     If Estado = "MG" Then
        If Mid(CFOP.Text, 1, 1) <> "1" Then
           MsgBox "O CFOP é invalido para fornecedores do estado de MG.", 64, "Aviso"
           CFOP.SetFocus
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     Else
        If Mid(CFOP.Text, 1, 1) <> "2" Then
           MsgBox "O CFOP é invalido para fornecedores do fora do estado de MG.", 64, "Aviso"
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
     Inscricao = Trim(Inscricao)
     If Len(CNPJ) = 0 Then
        MsgBox "O CNPJ do fornecedor não foi cadastrado.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
     If Len(CNPJ) > 11 Then
        If Not Calc_CNPJ(CNPJ) Then
           MsgBox "O CNPJ do fornecedor é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     Else
        If Not Calc_CPF(CNPJ) Then
           MsgBox "O CPF do fornecedor é invalido.", 64, "Aviso"
           ValidaEntradaSintegra = False
           GoTo Saida
        End If
     
     End If
     '==> Verifica a Inscricao estadual
     If Len(Inscricao) = 0 Then
        MsgBox "A inscrição Estadual do fornecedor não foi cadastrada." & Chr(13) & "Caso ele não possua inscrição estadual, casatre como ISENTO.", 64, "Aviso"
        ValidaEntradaSintegra = False
        GoTo Saida
     End If
     If Consiste(Inscricao, Estado) <> 0 Then
        MsgBox "A Inscrição Estadual do fornecedor é invalida.", 64, "Aviso"
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
