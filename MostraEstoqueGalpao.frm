VERSION 5.00
Begin VB.Form MostraEstoqueGalpao 
   BackColor       =   &H00CBB19C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque Nos Galpões"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Salvar F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox C 
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Uns2 
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
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox UnS1 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox California 
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Santa2 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox santa 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   0
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Aberto (Unidades)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   4680
      TabIndex        =   13
      Top             =   0
      Width           =   3300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque Fechado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "California"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Santa Maria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Galpao 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "MostraEstoqueGalpao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Saldo   As Double
Private Saldou  As Double
Private a As Integer
Private Estoque As ControleDb
Private LcSq As String
Private Sub Command1_Click()
On Error Resume Next
Unload Me

End Sub


Private Sub Command2_Click()
Dim LcSanta As Currency
Dim LcsantaU As Currency
Dim LcSanta2 As Currency
Dim Lcsanta2U As Currency
Dim LcCalifornia As Currency
Dim LcCaliforniaU As Currency
Dim LcTotal As Currency
Dim LcModifico As Boolean
Dim LcMsg As String
Dim LcCodigo As Long
Dim LcTipo As String
Dim LcCom As Integer
Dim StrSql As String
If Not IsNumeric(Santa2.Text) Then
   MsgBox "Informe um valor numérico para a Quantidade em santa Maria.", 64, "Aviso"
   Exit Sub
End If
If Not IsNumeric(Uns2.Text) Then
   MsgBox "Informe um valor numérico para a Quantidade em unidade em santa Maria.", 64, "Aviso"
   Exit Sub
End If

If Not IsNumeric(california.Text) Then
   MsgBox "Informe um valor numérico para a Quantidade em California.", 64, "Aviso"
   Exit Sub
End If
If Not IsNumeric(C.Text) Then
   MsgBox "Informe um valor numérico para a Quantidade em unidade em California.", 64, "Aviso"
   Exit Sub
End If

With FrmProduto
    LcSanta = 0 ' CDbl(santa.Text) - CDbl(.EstSanta)
    LcsantaU = 0 ' CDbl(UnS1.Text) - CDbl(.MinSanta.Text)
    LcSanta2 = CDbl(Santa2.Text) - CDbl(.EstSanta2.Text)
    Lcsanta2U = CDbl(Uns2.Text) - CDbl(.MinSanta2.Text)
    LcCalifornia = CDbl(california.Text) - CDbl(.EstCalifornia.Text)
    LcCaliforniaU = CDbl(C.Text) - CDbl(.MinCalifornia.Text)
    If IsNumeric(.Txt(16).Text) Then
       LcCom = .Txt(16).Text
    Else
       LcCom = 1
    End If
    LcCodigo = .Txt(0).Text
End With

'==> Verifica as modificações
'Estoque.ArmazenaEmGalpao = True
If LcSanta + LcsantaU + (LcSanta2 * LcCom) + Lcsanta2U + (LcCalifornia * LcCom) + LcCaliforniaU <> 0 Then
    LcMsg = "Confima a Atualização do Estoque?"
    LcResp = MsgBox(LcMsg, vbExclamation + vbYesNo, "Aviso")
    If lcrep = vbNo Then Exit Sub
    Dim ClEstoque As New ControleEstoque
    Dim LcNovoSaldoSanta As Currency
    Dim LcNovoSAldoCalifirnia As Currency
    Dim LcNovoSaldoGeral As Currency
    If LcSanta + LcsantaU + (LcSanta2 * LcCom) + Lcsanta2U + (LcCalifornia * LcCom) + LcCaliforniaU > 0 Then
       LcTipo = "E"
    Else
      LcTipo = "S"
      If LcSanta2 < 0 Then LcSanta2 = LcSanta2 * -1
      If Lcsanta2U < 0 Then Lcsanta2U = Lcsanta2U * -1
      If LcCalifornia < 0 Then LcCalifornia = LcCalifornia * -1
      If LcCaliforniaU < 0 Then LcCaliforniaU = LcCaliforniaU * -1
    End If
    ClEstoque.MovimentacaoManual LcCodigo, (LcSanta2 * LcCom) + Lcsanta2U, (LcCalifornia * LcCom) + LcCaliforniaU, LcTipo & "M Manual-" & GlUsuario, IIf(LcTipo = "E", "Entrada", "Saida") & " Manual", FrmProduto.Unidade, FrmProduto.Txt(13).Text, LcTipo
   ' Estoque.AtualizaEstoque , 0, (CDbl(Santa2.Text) * Estoque.QuantidadeDaUnidade) + CDbl(Uns2.Text), (CDbl(California.Text) * Estoque.QuantidadeDaUnidade) + CDbl(C.Text)
Else
  If LcSanta <> 0 Or LcsantaU <> 0 Or (LcSanta2 * LcCom) <> 0 Or Lcsanta2U <> 0 Or (LcCalifornia * LcCom) <> 0 Or LcCaliforniaU <> 0 Then
    
    Dim LcSaldo As Currency
    
    LcSanta = (CDbl(Santa2.Text) * LcCom) + CDbl(Uns2.Text)
    LcCalifornia = (CDbl(california.Text) * LcCom) + CDbl(C.Text)
    
    StrSql = "Update Produtos set Santa2=" & Replace(CStr(LcSanta), ",", ".") & ",California=" & Replace(CStr(LcCalifornia), ",", ".")
    StrSql = StrSql & " where codigo=" & LcCodigo
    ExecutaSql StrSql
  End If
End If
'==> Acerta o tipo

'LcSanta = (LcSanta * Estoque.QuantidadeDaUnidade) + LcsantaU
'LcSanta2 = (LcSanta2 * Estoque.QuantidadeDaUnidade) + Lcsanta2U
'LcCalifornia = (LcCalifornia * Estoque.QuantidadeDaUnidade) + LcCaliforniaU

'If LcSanta < 0 Then LcSanta = LcSanta * -1
'If LcSanta2 < 0 Then LcSanta2 = LcSanta2 * -1
'If LcCalifornia < 0 Then LcCalifornia = LcCalifornia * -1

'LcSq = "insert into HistoricoProduto (produto,descricao,santa,santa2,california,nf,data,tipo,unidade,codunid,ClienteForn) values ('"
'LcSq = LcSq & Estoque.CodProduto & "','" & Estoque.RetiraCaracter(Estoque.DescricaoProduto) & "'," & LcSanta & "," & LcSanta2 & "," & LcCalifornia
'LcSq = LcSq & ",'" & GlUsuario & "','" & Format(Date, "yyyy-mm-dd") & "','" & LcTipo & "','" & FrmProduto.Txt(13).Text & "','0','" & GlNomeMaquina & "')"
'ExecutaSql LcSq

Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then Command1_Click
If KeyCode = 113 Then Command2_Click
If KeyCode = 13 Then
   SendKeys "{tab}"
   SendKeys "{HOME}"
   SendKeys "+{END}"
End If
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
     Label1(0).Caption = Rs!Nome
  End If
  If a = 1 Then
     Label2.Caption = Rs!Nome
  End If
  If a = 2 Then
     Label1(1).Caption = Rs!Nome
  End If
  a = a + 1
  Rs.MoveNext
Loop
Set db = Nothing
Set Rs = Nothing

End Sub

Private Sub Form_Load()
Dim LcSql As String
carrega
'BuscaNomeGalpao
End Sub
Function carrega()
On Error Resume Next

Set Estoque = New ControleDb
Estoque.ArmazenaEmGalpao = True
Estoque.CodProduto = FrmProduto.Txt(0).Text
'santa.Text = Estoque.Santa1Fechado
'UnS1.Text = Estoque.Santa1Unitario
Santa2.Text = Estoque.Santa2Fechado
Uns2.Text = Estoque.Santa2Unitario
california.Text = Estoque.QuantidadeCaliforniaFechado
C.Text = Estoque.QuantidadeCaliforniaUnitario

End Function


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set Estoque = Nothing
FrmProduto.VinculaDados FrmProduto.Txt(0).Text
End Sub

