VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form detalhanota 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exibe os Itens da Nota Fiscal"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox MantemPreco 
      BackColor       =   &H00E6E4D2&
      Caption         =   "Mantem o Preço da Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8281
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12975841
      BackColorBkg    =   13296034
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Mostrar Todos Pedidos do Cliente F2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Default         =   -1  'True
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton CmdLanca 
      BackColor       =   &H00C0C000&
      Caption         =   "Lançar Produtos Selecionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
End
Attribute VB_Name = "detalhanota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LcDados
        Ordem As Long
        Codigo As Long
        Nome As String
        Unidade As String
        Com As Long
        Quantidade As Long
        Unitario As Currency
        total As Currency
End Type
Private MtOrdem() As LcDados

Private a As Integer
Private Ordem As Integer
Private Sub GeraGrid()
MSFlexGrid1.Cols = 8
MSFlexGrid1.ColWidth(0) = 800
MSFlexGrid1.ColWidth(1) = 800
MSFlexGrid1.ColWidth(2) = 800
MSFlexGrid1.ColWidth(3) = 5300
MSFlexGrid1.ColWidth(4) = 1100
MSFlexGrid1.ColWidth(5) = 1100
MSFlexGrid1.ColWidth(6) = 1100
MSFlexGrid1.ColWidth(7) = 1100


MSFlexGrid1.TextMatrix(0, 0) = "Selec."
MSFlexGrid1.TextMatrix(0, 1) = "Item"
MSFlexGrid1.TextMatrix(0, 2) = "Codigo"
MSFlexGrid1.TextMatrix(0, 3) = "Descrição do Produto"
MSFlexGrid1.TextMatrix(0, 4) = "Embalagem"
MSFlexGrid1.TextMatrix(0, 5) = "C/"
MSFlexGrid1.TextMatrix(0, 6) = "Quantidade"
MSFlexGrid1.TextMatrix(0, 7) = "V.Unitario"

MSFlexGrid1.ColAlignment(0) = 4
MSFlexGrid1.ColAlignment(1) = 4
MSFlexGrid1.ColAlignment(2) = 4
MSFlexGrid1.ColAlignment(4) = 4
MSFlexGrid1.ColAlignment(5) = 4
MSFlexGrid1.ColAlignment(6) = 4

ProcExit:
Exit Sub
ProcError:
 'CErr.Pop: End
Resume ProcExit
Resume Next
End Sub

Private Sub CmdLanca_Click()
Dim b As Long
For x = 0 To UBound(MtOrdem)
    If MtOrdem(x).Codigo <> 0 Then
       b = b + 1
       FrmProposta.MondaGridAutomatico MtOrdem(x).Codigo, MtOrdem(x).Quantidade, MtOrdem(x).Unitario, MtOrdem(x).Com, MtOrdem(x).Unidade, b
    End If
Next

'For x = 0 To MSFlexGrid1.Rows - 1
'    If MSFlexGrid1.TextMatrix(x, 0) = "S" Then
'        'Dim Dados As DadosEntrada
'        'Dados.CodPro = BuscaProduto(MSFlexGrid1.TextMatrix(x, 2))
'        'Dados.Qut = MSFlexGrid1.TextMatrix(x, 6)
'        Dim PrecoUnitario As Currency
'        Dim com As Long
'        Dim LcUnidade As String
'
'        If MantemPreco.Value Then
'           PrecoUnitario = CCur(MSFlexGrid1.TextMatrix(x, 7))
'           com = CLng(MSFlexGrid1.TextMatrix(x, 5))
'           LcUnidade = MSFlexGrid1.TextMatrix(x, 4)
'        Else
'            PrecoUnitario = 0
'        End If
'        FrmProposta.MondaGridAutomatico MSFlexGrid1.TextMatrix(x, 2), CLng(MSFlexGrid1.TextMatrix(x, 6)), PrecoUnitario, com, LcUnidade, x + 1
'    End If
'Next
Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function

Private Sub Command2_Click()
On Error Resume Next
Dim LcCl As String
Dim Rsa As ADODB.Recordset
If GlFormA.Name <> "Orcamento" Then
    LcSql = "SELECT alid050.CLiente, alid052.CodProd, alid052.qtde, alid052.Descricao, alid052.ValUnit "
    LcSql = LcSql & "FROM alid052 INNER JOIN alid050 ON alid052.numnf = alid050.numnf "
    LcSql = LcSql & "WHERE (((alid050.CLiente)='" & GlFormA.Txt(8).Text & "')) order by  alid052.numnf"
Else
   LcCl = orcamento.CodigoCliente.Text
    LcSql = "SELECT * "
    LcSql = LcSql & "FROM DadosOrcamento  INNER JOIN Orcamento ON DadosOrcamento.doc = Orcamento.doc "
    LcSql = LcSql & "WHERE Orcamento.CLiente='" & LcCl & "' order by dadosorcamento.doc asc"
    'MsgBox LcSql
End If

'"FROM Funcionários INNER JOIN(Pedidos " _
'        & "INNER JOIN [Detalhes do Pedido] " _
'        & "ON [Detalhes do Pedido].NúmeroDoPedido = " _
'        & "Pedidos.NúmeroDoPedido ) "
'MsgBox LcSql
Set Rsa = AbreRecordset(LcSql, True)
montagrid Rsa

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub



Private Sub Form_Load()
On Error Resume Next
Dim Rsa As ADODB.Recordset
Ordem = 0
ReDim MtOrdem(0)
'abreconexao
Dim LcSql As String
If GlFormA.Name <> "Orcamento" Then
   LcSql = "select * from alid052 where NUMNF='" & UltimasComprasCliente.Tag & "' order by cast(ITEM as decimal)"
Else
   LcSql = "select *,doc as NUMNF from DadosOrcamento where doc='" & UltimasComprasCliente.Tag & "' order by cast(ITEM as decimal)"
End If
'Debug.Print LcSql
Set Rsa = AbreRecordsetRel(LcSql, Rsa)
GeraGrid
montagrid Rsa


Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
Function montagrid(Rs As ADODB.Recordset)
Dim a As Long
MSFlexGrid1.Rows = 1

On Error GoTo errMonta
LcCap = Me.Caption
Me.Caption = "Filtrando registros. Aguarde..."
MSFlexGrid1.Rows = 1
If Rs.RecordCount + 1 > 14000 Then
    MSFlexGrid1.Rows = 14001
Else
    MSFlexGrid1.Rows = Rs.RecordCount + 1
End If
a = 1
With Rs
    Do Until .EOF
         DoEvents
         MSFlexGrid1.TextMatrix(a, 0) = "N"
         MSFlexGrid1.TextMatrix(a, 1) = Rs!Item & ""
         MSFlexGrid1.TextMatrix(a, 2) = Rs!codProd & ""
         MSFlexGrid1.TextMatrix(a, 3) = Rs!Descricao & ""
         MSFlexGrid1.TextMatrix(a, 4) = Rs!UNIMED & ""
         MSFlexGrid1.TextMatrix(a, 5) = Rs!QTDUM
         MSFlexGrid1.TextMatrix(a, 6) = Rs!Qtde
         MSFlexGrid1.TextMatrix(a, 7) = FormatNumber(Rs!VALUNIT, 2)
         
sai:
        .MoveNext
        a = a + 1
        If a > 14000 Then Exit Do
    Loop
End With
Me.Caption = LcCap

Exit Function

errMonta:
MsgBox err.Description & err.Number
cmdProcurar.Enabled = True
Resume Next

End Function

Private Sub MSFlexGrid1_Click()
Dim a As Long
a = MSFlexGrid1.Row
If MSFlexGrid1.TextMatrix(a, 0) <> "N" Then
    LimpaMatriz a
    MSFlexGrid1.TextMatrix(a, 0) = "N"
    For x = 0 To 7
        MSFlexGrid1.Col = x
        MSFlexGrid1.CellBackColor = &HC5FEE1
     Next
Else
    Ordem = Ordem + 1
    MSFlexGrid1.TextMatrix(a, 0) = Ordem
    LancaMatriz a, Ordem
    For x = 0 To 7
         MSFlexGrid1.Col = x
        MSFlexGrid1.CellBackColor = &HE6E4D2
     Next
End If
End Sub
Sub LimpaMatriz(linha As Long)
Dim MtTemp() As LcDados
Dim PosicaoTemp As Long
If MSFlexGrid1.TextMatrix(linha, 0) <> "N" Then
    '===> procura pelo lancamento
    Dim PosicaoExcluir As Long
    PosicaoExcluir = MSFlexGrid1.TextMatrix(linha, 0)
    For x = 0 To UBound(MtOrdem)
       If MtOrdem(x).Ordem <> PosicaoExcluir Then
         
         On Error Resume Next
         PosicaoTemp = UBound(MtTemp)
         If err.Number <> 0 Then
               PosicaoTemp = 0
                err.Number = 0
          Else
               PosicaoTemp = PosicaoTemp + 1
          End If
          ReDim Preserve MtTemp(PosicaoTemp)
          MtTemp(PosicaoTemp) = MtOrdem(x)
       End If
    Next
End If
MtOrdem = MtTemp
End Sub
Sub LancaMatriz(linha As Long, LcOrdem As Integer)
On Error Resume Next
Dim Posicao As Long
Posicao = UBound(MtOrdem)
If err.Number <> 0 Then
   Posicao = 0
    err.Number = 0
Else
   Posicao = Posicao + 1
End If
ReDim Preserve MtOrdem(Posicao)
MtOrdem(Posicao).Codigo = MSFlexGrid1.TextMatrix(linha, 2)
MtOrdem(Posicao).Nome = MSFlexGrid1.TextMatrix(linha, 3)
MtOrdem(Posicao).Com = MSFlexGrid1.TextMatrix(linha, 5)
MtOrdem(Posicao).Unidade = MSFlexGrid1.TextMatrix(linha, 4)
MtOrdem(Posicao).Quantidade = MSFlexGrid1.TextMatrix(linha, 6)
MtOrdem(Posicao).Unitario = MSFlexGrid1.TextMatrix(linha, 7)
MtOrdem(Posicao).total = MtOrdem(Posicao).Unitario * MtOrdem(Posicao).Quantidade
MtOrdem(Posicao).Ordem = LcOrdem
End Sub
