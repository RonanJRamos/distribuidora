VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDuplicaPedido 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicar Pedido"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFiltar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar"
      Height          =   615
      Left            =   2760
      TabIndex        =   7
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton CmdSelecionar 
      Caption         =   "Duplicar"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid Nota 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.TextBox NomeCliente 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.TextBox CodCliente 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmDuplicaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
Unload Me
End Sub
Sub GeraGrid()
On Error Resume Next
Nota.ColAlignment(0) = 3
Nota.ColAlignment(1) = 3
Nota.ColAlignment(2) = 3
Nota.ColAlignment(3) = 7

Nota.ColWidth(0) = 1800
Nota.ColWidth(1) = 1800
Nota.ColWidth(2) = 1800
Nota.ColWidth(3) = 1900

Nota.TextMatrix(0, 0) = "Selecionar"
Nota.TextMatrix(0, 1) = "Nº Pedido"
Nota.TextMatrix(0, 2) = "Emissão"
Nota.TextMatrix(0, 3) = "Valor"


Item.ColAlignment(0) = 3
Item.ColAlignment(1) = 1
Item.ColAlignment(2) = 3
Item.ColAlignment(3) = 3

Item.ColWidth(0) = 1000
Item.ColWidth(1) = 3000
Item.ColWidth(2) = 1200
Item.ColWidth(3) = 1200


Item.TextMatrix(0, 0) = "Cod. Prod"
Item.TextMatrix(0, 1) = "Nome"
Item.TextMatrix(0, 2) = "Quant"
Item.TextMatrix(0, 3) = "Unitario"
Item.TextMatrix(0, 4) = "Total"

LcTamanhoGrid = 1
End Sub
Function VerificaDisponivelGrid(LcCodProduto As String, LcQuantidade As Double, LcComG As Double) As Double
On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim LcCom As Long
Dim RsNota As ADODB.Recordset
LcSql = "Select * from Produtos where codigo=" & LcCodProduto
AbreBase
VerificaDisponivelGrid = 0
Set RsNota = AbreRecordset(LcSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not IsNull(RsNota("QuantEstoque")) Then LcQuantEstoque = RsNota("QuantEstoque") Else LcQuantEstoque = 0
If LcQuantEstoque < (CDbl(LcQuantidade) * CDbl(LcComG)) Then
   VerificaDisponivelGrid = 2
End If
End Function

Function ConferePrecoGrid(LcCodProduto As String, LcValor As Currency, LcComG As Currency) As Long
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim LcPRecoAntigo As Currency
Dim LcLcqM As Currency
Dim LcPreconovo As Currency

ConferePrecoGrid = 0
Set Rs = AbreRecordset("select * from Produtos where codigo=" & LcCodProduto, True)
If Not Rs.EOF Then
   If IsNull(Rs!QtdMedida) Then
      LcLcqM = 1
   Else
      If Rs!QtdMedida = 0 Then
          LcLcqM = 1
      Else
         LcLcqM = Rs!QtdMedida
      End If
   End If
   LcPRecoAntigo = CDec(Rs!MinimoVenda) / LcLcqM
Else
  LcPRecoAntigo = 0
End If
LcPreconovo = CDec(LcValor) / CDec(LcComG)
GlEscolha = True

If LcPreconovo < LcPRecoAntigo Then
 '   Comissao.Text = 1
    ConferePrecoGrid = 1
Else
  GlLibera = True
  'If CLng(Comissao.Text) <> 1 Then
'     Comissao.Text = 1.5
 ' End If
End If
End Function
Sub DesmarcaTodos()
For a = 0 To Nota.Rows - 1
  Nota.TextMatrix(a, 0) = "Não"
Next
End Sub
Private Sub CmdFiltar_Click()
Dim RsOrc As Recordset
Dim b As Integer
If Not IsNumeric(CodCliente.Text) Then
   MsgBox "Selecione o Clente Antes de filtrar os pedidos", 64, "Aviso"
   Exit Sub
End If
LcSql1 = "Select top 10 * from proposta where Cliente='" & Right("00000" & CodCliente.Text, 5) & "' and dtemis>= #07/04/2016# order by numNf desc"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
b = 1
Nota.Rows = 1

Do Until RsOrc.EOF
   Nota.Rows = b + 1
   Nota.TextMatrix(b, 0) = "Não"
   Nota.TextMatrix(b, 1) = RsOrc!NumNf & ""
   Nota.TextMatrix(b, 2) = Format(RsOrc!DTEMIS, "dd/mm/yy")
   Nota.TextMatrix(b, 3) = FormatNumber(RsOrc!ValorNota, 2)
   b = b + 1
   RsOrc.MoveNext
Loop

End Sub
Sub MostraItens(NF As String)
Dim RsOrc As Recordset
Dim b As Integer

LcSql1 = "Select * from subproposta where NUMNF='" & NF & "' order by codigo"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
b = 1
Item.Rows = 1

Do Until RsOrc.EOF
   Item.Rows = b + 1
   
   Item.TextMatrix(b, 0) = RsOrc!codProd & ""
   Item.TextMatrix(b, 1) = RsOrc!Descricao & ""
   Item.TextMatrix(b, 2) = FormatNumber(RsOrc!Qtde, 2)
   Item.TextMatrix(b, 3) = FormatNumber(RsOrc!VALUNIT, 2)
   Item.TextMatrix(b, 4) = FormatNumber(RsOrc!Qtde * RsOrc!VALUNIT, 2)
   b = b + 1
   RsOrc.MoveNext
Loop
End Sub

Private Sub CmdSelecionar_Click()
Dim LcNovoNf As String
Dim LcCap As String
LcCap = Me.Caption
Me.Caption = "aguarde, duplicando pedido..."
DoEvents
For a = 0 To Nota.Rows - 1
  If Nota.TextMatrix(a, 0) = "Sim" Then
     LcNovoNf = DuplicaCabecalho(Nota.TextMatrix(a, 1))
      Duplicaitens Nota.TextMatrix(a, 1), LcNovoNf
  End If
Next
Me.Caption = LcCap
MsgBox "Pedido gerado com o Nº:" & LcNovoNf & "!", 64, "Aviso"
FrmProposta.BuscaNota LcNovoNf
Unload Me
End Sub
Function CalculaNumeroNota() As String
On Error Resume Next
Dim LcSql As String, LcNumeroNota As String
Dim RsNota As Recordset
LcSql = "Select * from proposta order by NUMNF"
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
CalculaNumeroNota = LcNumeroNota

RsNota.Close
Dbbase.Close
Set RsNota = Nothing
Set Dbbase = Nothing

End Function
Function Duplicaitens(NF As String, NovoNf As String) As String
Dim RsOrc As Recordset
Dim RsOrcNovo As Recordset
Dim LcNovoCodigo As String
Dim BloqueioQuant As Boolean
Dim BloqueioValor As Boolean
Dim TemBloqueio As Boolean
TemBloqueio = False

LcSql1 = "Select * from subproposta where numNf='" & NF & "' order by Codigo"
LcSql2 = "Select * from subproposta"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOrcNovo = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)

Do Until RsOrc.EOF
   BloqueioQuant = VerificaDisponivelGrid(RsOrc!codProd, RsOrc!Qtde, RsOrc!QTDUM)
   BloqueioValor = ConferePrecoGrid(RsOrc!codProd, CDec(RsOrc!VALUNIT), CDec(RsOrc!QTDUM))
    RsOrcNovo.AddNew
    For C = 0 To RsOrc.Fields.Count - 1
        LcNome = RsOrc.Fields(C).Name
        If UCase(LcNome) = "NUMNF" Then
            RsOrcNovo(LcNome) = NovoNf
        ElseIf UCase(LcNome) = UCase("Bloqueado") Then
           If BloqueioQuant Or BloqueioValor Then
              RsOrcNovo(LcNome) = True
              TemBloqueio = True
           Else
              RsOrcNovo(LcNome) = False
           End If
        Else
            RsOrcNovo(LcNome) = RsOrc.Fields(C)
        End If
        DoEvents
    Next
    
    RsOrcNovo.Update
    RsOrc.MoveNext
    DoEvents
Loop
StrSql = "Update proposta set Bloqueado=" & TemBloqueio & " where numNf='" & NovoNf & "'"
Dbbase.Execute StrSql

End Function
Function DuplicaCabecalho(NF As String) As String
Dim RsOrc As Recordset
Dim RsOrcNovo As Recordset
Dim LcNovoCodigo As String
Dim StrSql As String
LcNovoCodigo = CalculaNumeroNota

LcSql1 = "Select  * from proposta where numNf='" & NF & "'"
LcSql2 = "Select  * from proposta order by numNf desc"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOrcNovo = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Len(GlNomeMaquina) = 0 Then
  NomeMaquina
End If
Nome_Maquina = GlNomeMaquina

Do Until RsOrc.EOF
    RsOrcNovo.AddNew
    For C = 0 To RsOrc.Fields.Count - 1
        LcNome = RsOrc.Fields(C).Name
        If UCase(LcNome) <> "CODIGO" Then
            If UCase(LcNome) = "NUMNF" Then
                RsOrcNovo(LcNome) = LcNovoCodigo
            Else
                RsOrcNovo(LcNome) = RsOrc.Fields(C)
            End If
            DoEvents
        End If
        
    Next
    RsOrcNovo("Maquina") = Nome_Maquina
    RsOrcNovo.Update
    RsOrc.MoveNext
    DoEvents
Loop
Dim RsProposta As DAO.Recordset
    Set RsProposta = Dbbase.OpenRecordset("Select top 1 codigo From proposta where Maquina='" & Nome_Maquina & "' order by codigo desc")
    If Not RsProposta.EOF Then
       Dim NumeroNFe As String
       LcNovoCodigo = RsProposta!Codigo
       'Txt(0).Text = NumeroNFe
       LcSq = "Update proposta set NUMNF='" & LcNovoCodigo & "' where codigo=" & RsProposta!Codigo
       Dbbase.Execute LcSq, Processados
    End If
StrSql = "Update proposta set Previsao=Null,Liberado=0,faturado=0,Validade='',OrdemCompra='',Bloqueado=0,dataliberacao=Null,JaEsteveBloqueado=0,MaquinaLiberacao='',HoraLiberacao='',pendente=0,Usuario='',Romaneio=0,DTEmis=#" & Format(Date, "mm/dd/yy") & "# where numNf='" & LcNovoCodigo & "'"
Dbbase.Execute StrSql
DuplicaCabecalho = LcNovoCodigo
End Function
Private Sub CodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    FrmPesquisaCliente.Txt.Text = NomeCliente.Text
    FrmPesquisaCliente.Show , Me
End If
End Sub

Private Sub CodCliente_LostFocus()
Dim RsOrc As Recordset
Dim b As Integer
If Not IsNumeric(CodCliente.Text) Then
   Exit Sub
End If
LcSql1 = "Select  * from ALID001 where Codigo='" & Right("00000" & CodCliente.Text, 5) & "'"
AbreBase
Set RsOrc = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not RsOrc.EOF Then
   CodCliente.Text = RsOrc!Codigo & ""
   NomeCliente.Text = RsOrc!RAZAOSOC & ""
Else
   CodCliente.Text = ""
   NomeCliente.Text = ""
End If

End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
End Sub

Private Sub Form_Load()
GeraGrid
End Sub

Private Sub NomeCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    GlCriterioSql = ""
    Load FrmPesquisaCliente
    FrmPesquisaCliente.Txt.Text = NomeCliente.Text
    FrmPesquisaCliente.ExibePesquisa
    FrmPesquisaCliente.Show , Me
End If
End Sub

Private Sub Nota_DblClick()
Dim LcRow As Integer
Dim lccol As Integer
On Error Resume Next
If Nota.Rows = 1 Then Exit Sub
LcColuna = Item.Col
linha = Nota.Row
DesmarcaTodos
If Nota.TextMatrix(linha, 0) = "Não" Then
   Nota.TextMatrix(linha, 0) = "Sim"
   MostraItens Nota.TextMatrix(linha, 1)
Else
   Nota.TextMatrix(linha, 0) = "Não"
   Item.Rows = 1
End If
 
End Sub
