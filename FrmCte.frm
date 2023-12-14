VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCte 
   BackColor       =   &H00E0D2C7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CT-e Referênciado"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "FrmCte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CodFornecedor 
      Height          =   285
      Left            =   6000
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox ValorParcela 
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox quantidade 
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox TipoMonetario 
      Height          =   315
      Left            =   3240
      TabIndex        =   12
      Top             =   2190
      Width           =   3975
   End
   Begin MSMask.MaskEdBox CNPJ 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99.999.999/9999-99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Nome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
   End
   Begin VB.TextBox ChaveAcesso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin VB.TextBox NumeroNF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Valor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "0,00"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Cmdfechar 
      BackColor       =   &H00CBB19C&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   5
      Left            =   5640
      TabIndex        =   20
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   21
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   7
      Left            =   3720
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   8
      Left            =   5640
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   9
      Left            =   1800
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   11
      Left            =   5640
      TabIndex        =   26
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdSalvar 
      BackColor       =   &H00CBB19C&
      Caption         =   "&Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "valor Parcela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4080
      TabIndex        =   30
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monetário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Transp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   11
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ Transp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chave de Acesso CT-e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº NF CT-e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   570
   End
End
Attribute VB_Name = "FrmCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Str_Sql As String
Private LcNatureza As String
Private LcNota, LcBoleta, LcEspaco, LcLinha, LcEspC As String
Private LcSalto, LcQuant, a As Integer
Private Sub CmdFechar_Click()
Unload Me
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
If Not IsNumeric(Valor.Text) Then Valor.Text = "0,00"
If Len(NumeroNF.Text) = 0 Then
    MsgBox "Entre com o Número da NF CT-e", vbInformation, "Aviso"
    NumeroNF.SetFocus
    Exit Sub
End If
'If Len(ChaveAcesso.Text) = 0 Then
'    MsgBox "Entre com a Chave de acesso da NF CT-e", vbInformation, "Aviso"
'    ChaveAcesso.SetFocus
'    Exit Sub
'End If
If Len(TipoMonetario.Text) = 0 Then
    MsgBox "Informe o tipo monetario para o Lancamento.", 64, "Aviso"
    TipoMonetario.SetFocus
    Exit Sub
Else
   If Not IsDate(Vencimento(0).Text) Then
       MsgBox "Informe pelo menos um vencimento para o lançamento.", 64, "Aviso"
       Vencimento(0).SetFocus
       Exit Sub
    End If
    
End If
Dim LcCNPJ As String
LcCNPJ = CNPJ.Text
LcCNPJ = Replace(LcCNPJ, ",", "")
LcCNPJ = Replace(LcCNPJ, ".", "")
LcCNPJ = Replace(LcCNPJ, "-", "")
LcCNPJ = Replace(LcCNPJ, "/", "")
LcCNPJ = Replace(LcCNPJ, "\", "")
LcCNPJ = Replace(LcCNPJ, " ", "")
If Len(LcCNPJ) <> 14 Then
    MsgBox "Entre com o CNPJ do(a) Transportador(a)", vbInformation, "Aviso"
    CNPJ.SetFocus
    Exit Sub
End If
If Len(Nome.Text) = 0 Then
    MsgBox "Entre com a Nome do(a) Transportador(a)", vbInformation, "Aviso"
    Nome.SetFocus
    Exit Sub
End If
If CDec(Valor.Text) <= 0 Then
    MsgBox "Entre com o Valor da NF CT-e", vbInformation, "Aviso"
    Valor.SetFocus
    Exit Sub
End If
If Not IsNumeric(quantidade.Text) Then quantidade.Text = 0
If Not IsNumeric(Valor.Text) Then Valor.Text = 0

Str_Sql = "Delete FROM NFentrada_Cte where CodNota=" & FrmEntradaProduto.CodigoDaNota.Text
afetados = ExecutaSql(Str_Sql)
Str_Sql = "Insert into NFentrada_Cte(CodNota,NumeroNFCte,ChaveAcesso,CNPJ,Nome,Valor,Emissao,Parcelas,FormaPag,NumeroNFe"
Str_Sql = Str_Sql & ")Values("
Str_Sql = Str_Sql & FrmEntradaProduto.CodigoDaNota.Text & ","
Str_Sql = Str_Sql & "'" & NumeroNF.Text & "',"
Str_Sql = Str_Sql & "'" & ChaveAcesso.Text & "',"
Str_Sql = Str_Sql & "'" & CNPJ.Text & "',"
Str_Sql = Str_Sql & "'" & Replace(Nome.Text, "'", "''") & "',"
Str_Sql = Str_Sql & Replace(CDec(Valor.Text), ",", ".") & ","
Str_Sql = Str_Sql & "'" & Format(FrmEntradaProduto.emissao.Text, "yyyy-mm-dd") & "',"
Str_Sql = Str_Sql & quantidade.Text & ","
Str_Sql = Str_Sql & "'" & Replace(TipoMonetario.Text, "'", "''") & "',"
Str_Sql = Str_Sql & "'" & FrmEntradaProduto.Txt(0).Text & "')"

afetados = ExecutaSql(Str_Sql)
'===> recupera o codigo gerado
Dim Rs As New ADODB.Recordset
Dim LcSql As String
Dim LcCodigo As Long
LcSql = "Select * from NFentrada_Cte where CodNota=" & FrmEntradaProduto.CodigoDaNota.Text
Set Rs = AbreRecordset(LcSql, True)
If Not Rs.EOF Then
    LcCodigo = Rs!Codigo
End If
If afetados > 0 Then GravaFinanceiro

If IsDate(Vencimento(0).Text) Then GravaVencimentos LcCodigo, Vencimento(0).Text, ValorParcela.Text
If IsDate(Vencimento(1).Text) Then GravaVencimentos LcCodigo, Vencimento(1).Text, ValorParcela.Text
If IsDate(Vencimento(2).Text) Then GravaVencimentos LcCodigo, Vencimento(2).Text, ValorParcela.Text
If IsDate(Vencimento(3).Text) Then GravaVencimentos LcCodigo, Vencimento(3).Text, ValorParcela.Text
If IsDate(Vencimento(4).Text) Then GravaVencimentos LcCodigo, Vencimento(4).Text, ValorParcela.Text
If IsDate(Vencimento(5).Text) Then GravaVencimentos LcCodigo, Vencimento(5).Text, ValorParcela.Text
If IsDate(Vencimento(6).Text) Then GravaVencimentos LcCodigo, Vencimento(6).Text, ValorParcela.Text
If IsDate(Vencimento(7).Text) Then GravaVencimentos LcCodigo, Vencimento(7).Text, ValorParcela.Text
If IsDate(Vencimento(8).Text) Then GravaVencimentos LcCodigo, Vencimento(8).Text, ValorParcela.Text
If IsDate(Vencimento(9).Text) Then GravaVencimentos LcCodigo, Vencimento(9).Text, ValorParcela.Text
If IsDate(Vencimento(10).Text) Then GravaVencimentos LcCodigo, Vencimento(10).Text, ValorParcela.Text
If IsDate(Vencimento(11).Text) Then GravaVencimentos LcCodigo, Vencimento(11).Text, ValorParcela.Text

'==> Grava os vencimentos


MsgBox "Salvo com sucesso!", vbInformation, "Aviso"
End Sub
Sub GravaVencimentos(CodigoCte As Long, Vencimento As String, Valor As Currency)
  Str_Sql = "Insert into nfentrada_cte_vencimentos(Cod_Cte,Vencimento,Valor_Parcela"
  Str_Sql = Str_Sql & ")Values("
  Str_Sql = Str_Sql & CodigoCte & ","
  Str_Sql = Str_Sql & "'" & Format(Vencimento, "yyyy-mm-dd") & "',"
  Str_Sql = Str_Sql & Replace(CDec(Valor), ",", ".") & ")"
  afetados = ExecutaSql(Str_Sql)


End Sub
Sub GravaFinanceiro()

FrmEntradaProduto.LancarFinanceiroCTE CInt(quantidade.Text), NumeroNF.Text, CodFornecedor.Text, TipoMonetario.Text, CCur(ValorParcela.Text)
'(LcNumeroContas As Integer, NumeroNF As String , CodFornecedor As String, CodTipoMonet As String, LcValor As Double)
End Sub
Private Sub CNPJ_Change()
On Error Resume Next
Dim LcCNPJ As String
LcCNPJ = CNPJ.Text
LcCNPJ = Replace(LcCNPJ, ",", "")
LcCNPJ = Replace(LcCNPJ, ".", "")
LcCNPJ = Replace(LcCNPJ, "-", "")
LcCNPJ = Replace(LcCNPJ, "/", "")
LcCNPJ = Replace(LcCNPJ, "\", "")
LcCNPJ = Replace(LcCNPJ, " ", "")

If Len(LcCNPJ) = 14 Then
   If Calc_CNPJ(LcCNPJ) Then
        Call Busca_Fornecedor(LcCNPJ)
   End If
End If
End Sub
Sub Busca_Fornecedor(LcCNPJ As String)
On Error Resume Next
Dim RsF As Recordset
AbreBase
GlLibera = False
Dim LcSql As String
LcSql = "select Codigo,RazaoSoc from alid002"
LcSql = LcSql & " Where cgc='" & CNPJ.Text & "' or cgc='" & LcCNPJ & "'"
Set RsF = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsF.EOF Then
    Nome.Text = RsF!RazaoSoc & ""
    CodFornecedor.Text = RsF!Codigo
End If
RsF.Close
Set RsF = Nothing
End Sub

Private Sub CNPJ_LostFocus()
On Error Resume Next
Dim LcCNPJ As String
LcCNPJ = CNPJ.Text
LcCNPJ = Replace(LcCNPJ, ",", "")
LcCNPJ = Replace(LcCNPJ, ".", "")
LcCNPJ = Replace(LcCNPJ, "-", "")
LcCNPJ = Replace(LcCNPJ, "/", "")
LcCNPJ = Replace(LcCNPJ, "\", "")
LcCNPJ = Replace(LcCNPJ, " ", "")

'If Len(LcCNPJ) = 14 Then
'   If Calc_CNPJ(LcCNPJ) Then
        Call Busca_Fornecedor(LcCNPJ)
 '  End If
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys ("{tab}")
End If

End Sub
Sub Busca_Dados()
On Error Resume Next
Dim Rs As New ADODB.Recordset
Dim RsVenc As New ADODB.Recordset
Dim LcSql As String
Dim LcCodigo As Long
LcSql = "Select * from NFentrada_Cte where CodNota=" & FrmEntradaProduto.CodigoDaNota.Text
Set Rs = AbreRecordset(LcSql, True)
If Not Rs.EOF Then
    LcCodigo = Rs!Codigo
    NumeroNF.Text = Rs!NumeroNFCte & ""
    ChaveAcesso.Text = Rs!ChaveAcesso & ""
    CNPJ.Text = Rs!CNPJ & ""
    Nome.Text = Rs!Nome & ""
    Valor.Text = FormatNumber(CDec(Rs!Valor), 2) & ""
    TipoMonetario.Text = Rs!FormaPag & ""
    quantidade.Text = Rs!Parcelas
    ValorParcela.Text = Rs!Valor
End If
Rs.Close
Set Rs = Nothing
If LcCodigo > 0 Then
    LcSql = "Select * from nfentrada_cte_vencimentos where Cod_Cte=" & LcCodigo
    Set RsVenc = AbreRecordset(LcSql, True)
    Dim i As Integer
    Do Until RsVenc.EOF
       Vencimento(i).Text = Format(RsVenc!Vencimento, "dd/mm/yy")
       i = i + 1
       RsVenc.MoveNext
    Loop
    RsVenc.Close
    Set RsVenc = Nothing
End If



End Sub
Function HabilitaPag()
Dim Exibe, ExibeMonetario As Integer


Select Case FrmEntradaProduto.Natureza.Text
Case Is = "A VISTA"
     LcNatureza = "VENDAS A VISTA"
     Exibe = False
     ExibeMonetario = True
   
Case Is = "A PRAZO"
     LcNatureza = "VENDAS A PRAZO"
     Exibe = True
     ExibeMonetario = False
   
Case Is = "SR - Simples Remessa"
     LcNatureza = "Simples Remessa"
     ExibeMonetario = False
     Exibe = False
   
Case Is = "ND - Nota Devolucao"
    LcNatureza = "Nota Devolução"
    ExibeMonetario = False
    Exibe = False
  
End Select
 CarregaTipoMonetario
'TipoMonetario.Visible = ExibeMonetario
'Label1(10).Visible = ExibeMonetario

End Function
Function CarregaTipoMonetario()
Dim RsMoney As Recordset
TipoMonetario.Clear
AbreBase
Set RsMoney = Dbbase.OpenRecordset("Select * from alid008 where VENDA='S' order by XTPMONET", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsMoney.EOF
   TipoMonetario.AddItem RsMoney("XTPMONET")
   RsMoney.MoveNext
Loop
RsMoney.Close
Dbbase.Close
Set RsMoney = Nothing
Set Dbbase = Nothing


End Function


Function GeraValor() As Currency
Dim LcValor As Currency
If Vencimento(11).Text = "  /  /  " Then
  If Vencimento(10).Text = "  /  /  " Then
     If Vencimento(9).Text = "  /  /  " Then
       If Vencimento(8).Text = "  /  /  " Then
            If Vencimento(7).Text = "  /  /  " Then
             If Vencimento(6).Text = "  /  /  " Then
                   If Vencimento(5).Text = "  /  /  " Then
                      If Vencimento(4).Text = "  /  /  " Then
                         If Vencimento(3).Text = "  /  /  " Then
                            If Vencimento(2).Text = "  /  /  " Then
                               If Vencimento(1).Text = "  /  /  " Then
                                  If Vencimento(0).Text = "  /  /  " Then
                                  Else
                                     ValorParcela.Text = CCur(Valor.Text)
                                     LcQuant = 1
                                  End If
                              Else
                                 ValorParcela.Text = CCur(Valor.Text) / 2
                                 LcQuant = 2
                              End If
                          Else
                             ValorParcela.Text = CCur(Valor.Text) / 3
                             LcQuant = 3
                         End If
                       Else
                          ValorParcela.Text = CCur(Valor.Text) / 4
                          LcQuant = 4
                      End If
                    Else
                       ValorParcela.Text = CCur(Valor.Text) / 5
                       LcQuant = 5
                    End If
                   Else
                     ValorParcela.Text = CCur(Valor.Text) / 6
                     LcQuant = 6
                   End If
              Else
               ValorParcela.Text = CCur(Valor.Text) / 7
               LcQuant = 7
              End If
            Else
               ValorParcela.Text = CCur(Valor.Text) / 8
               LcQuant = 8
            End If
         Else
            ValorParcela.Text = CCur(Valor.Text) / 9
            LcQuant = 9
         End If
       Else
            ValorParcela.Text = CCur(Valor.Text) / 10
            LcQuant = 10
         End If
        Else
            ValorParcela.Text = CCur(Valor.Text) / 11
            LcQuant = 11
        End If
      Else
         ValorParcela.Text = CCur(Valor.Text) / 12
         LcQuant = 12
     End If
quantidade.Text = LcQuant
ValorParcela.Text = FormatNumber(CCur(ValorParcela.Text), 2)
End Function

Private Sub Form_Load()
Dim LcVer, a As Integer
Valor.Text = 0 'FrmEntradaProduto.Txt(16).Text

ValorParcela.Text = FormatNumber(CCur(Valor.Text), 2)
Valor.Text = FormatNumber(CCur(Valor.Text), 2)
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
HabilitaPag
Select Case FrmEntradaProduto.Natureza.Text
    Case Is = "A VISTA"
         LcVer = False
    Case Is = "A PRAZO"
         LcVer = True
End Select
For a = 1 To 11
    Vencimento(a).Visible = LcVer
Next
Busca_Dados
quantidade.Visible = LcVer
'Label1(0).Visible = LcVer
End Sub

Private Sub Vencimento_LostFocus(Index As Integer)
If Vencimento(Index).Text = "  /  /  " Then Exit Sub
If Not IsDate(Vencimento(Index).Text) Then
   MsgBox "O Valor digitado deve Ser uma Data...", 64, "Aviso"
   Vencimento(Index).Text = "  /  /  "
   Vencimento(Index).SetFocus
Else
  GeraValor
End If
End Sub
