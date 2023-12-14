VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControleDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarLimiteVenda                 As Double
Private mvarQuantidadeSanta1            As Double 'local copy
Private mvarQuantidadeSanta1Unitario    As Double 'local copy
Private mvarQuantidadeSanta2            As Double 'local copy
Private mvarQuantidadeSanta2Unitario    As Double 'local copy
Private mvarQuantidadeCalifornia        As Double 'local copy
Private mvarPrecoVenda                  As Double 'local copy
Private mvarPrecoMinimo                 As Double 'local copy
Private mvarMiminoEstoque               As Double 'local copy
Private mvarMaximoEstoque               As Double
Private mvarQuantidadeTotalUnitario     As Double 'local copy
Private mvarQuantidadeTotalEstoque      As Double 'local copy
Private mvarQuantidadeTotalEstoqueSeg   As Double
Private LcInicializado                  As Boolean
Private mvarArmazenaEmGalpao            As Boolean
Private mvarEstoqueGeralUnidade         As Boolean

Private mvarCodProduto                  As Integer 'local copy
Private mvarCodClien_forn               As String 'local copy

Private LcNomeFor_cli                   As String
Private mvarDescricaoProduto            As String 'local copy
Private LcNome_Unidade                  As String
Private mvarNF                          As String 'local copy
'local variable(s) to hold property value(s)
Private mvarQuantidadeDaUnidade As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarLançaFichaDeEstoque As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarPrecoDeCusto As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarCodigoDaUnidade As String 'local copy
'local variable(s) to hold property value(s)
Public Property Let ArmazenaEmGalpao(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ArmazenaEmGalpao = 5
    mvarArmazenaEmGalpao = vData
End Property


Public Property Get ArmazenaEmGalpao() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ArmazenaEmGalpao
    ArmazenaEmGalpao = mvarArmazenaEmGalpao
End Property



Public Property Let CodigoDaUnidade(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodigoDaUnidade = 5
    mvarCodigoDaUnidade = vData
End Property


Public Property Get CodigoDaUnidade() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodigoDaUnidade
    CodigoDaUnidade = mvarCodigoDaUnidade
End Property
Public Property Get maximoEstoque() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaximoEstoque
    maximoEstoque = mvarMaximoEstoque
End Property





Public Property Get PrecoDeCusto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrecoDeCusto
    PrecoDeCusto = mvarPrecoDeCusto
End Property



Public Sub AtualizaEstoque(Optional QuantidadeGeral As Double = 0, Optional TotalSanta As Double = 0, Optional TotalSanta1 As Double = 0, Optional TotalCalifornia As Double = 0)
Attribute AtualizaEstoque.VB_Description = "Informe a Quantidade total do Estoque de Cada Galpao em Unidades"
On Error GoTo errat
'abreconexao
conexaoAdo.BeginTrans
If Not mvarArmazenaEmGalpao Then
    LcQuantidade = QuantidadeGeral
Else
    LcQuantidade = TotalSanta + TotalSanta1 + TotalCalifornia
End If
LcSql = "UPDATE Produtos SET QuantEstoque=" & LcQuantidade & ",Santa1=" & TotalSanta
LcSql = LcSql & ",santa2=" & TotalSanta1 & ",california=" & TotalCalifornia & " where codigo=" & mvarCodProduto

ExecutaSql LcSql
conexaoAdo.CommitTrans

Exit Sub
errat:
Dim LcRespErr As Integer
LcRespErr = ProcessaErro(err.Number, err.Description)
If LcRespErr = 4 Then
   Resume 0
Else
  conexaoAdo.RollbackTrans
End If
End Sub

Public Property Let LançaFichaDeEstoque(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LançaFichaDeEstoque = 5
    mvarLançaFichaDeEstoque = vData
End Property


Public Property Get LançaFichaDeEstoque() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LançaFichaDeEstoque
    LançaFichaDeEstoque = mvarLançaFichaDeEstoque
End Property



Public Property Get QuantidadeDaUnidade() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QuantidadeDaUnidade
    QuantidadeDaUnidade = mvarQuantidadeDaUnidade
End Property

Private Sub BuscaCliente()
Dim RsClientes As ADODB.Recordset
On Error Resume Next
    'abreconexao
    Set RsClientes = AbreRecordset("Select * from clientes where codigo='" & mvarCodClien_forn & "'", True)
    If Not RsClientes.EOF Then
       LcNomeFor_cli = rsCliente!RAZAOSOC & ""
    Else
       LcNomeFor_cli = ""
    End If
     LcNomeFor_cli = RetiraCaracter(LcNomeFor_cli)
    Set rsCliente = Nothing

End Sub
Private Sub BuscaFornecedor()
Dim RsFornecedor As ADODB.Recordset
On Error Resume Next
    'abreconexao
    Set RsFornecedor = AbreRecordset("Select * from fornecedor where codigo='" & mvarCodClien_forn & "'", True)
    If Not RsFornecedor.EOF Then
       LcNomeFor_cli = RsFornecedor!RAZAOSOC & ""
    Else
       LcNomeFor_cli = ""
    End If
     LcNomeFor_cli = RetiraCaracter(LcNomeFor_cli)
    Set RsFornecedor = Nothing

End Sub
Public Function RetiraCaracter(LcStr As String) As String
On Error Resume Next
LcStr = Replace(LcStr, "'", "")
LcStr = Replace(LcStr, Chr(43), "")
RetiraCaracter = LcStr

End Function
Private Sub buscaNomeUnidade(Codigo As Integer)
Dim RsUnidades As ADODB.Recordset
On Error Resume Next
    'abreconexao
    Set RsUnidades = AbreRecordset("Select * from unidades where codigo=" & Codigo, True)
    If Not RsUnidades.EOF Then
       LcNome_Unidade = RsUnidades!RAZAOSOC & ""
    Else
       LcNome_Unidade = ""
    End If
    LcNome_Unidade = RetiraCaracter(LcNome_Unidade)
    Set RsUnidades = Nothing

End Sub

Public Property Let CodClien_forn(ByVal vData As String)
    
    mvarCodClien_forn = vData
    
End Property

Public Property Get CodClien_forn() As String
    CodClien_forn = mvarCodClien_forn
End Property


Public Property Let NF(ByVal vData As String)
    mvarNF = vData
End Property


Public Property Get NF() As String
    NF = mvarNF
End Property


Private Function Inicializa() As Boolean
'Ao inicializar a Clasee, deve-se saber se o sistema utilza armazemanento
'em galpao ou não, e preenche as variaveis de informação.
On Error GoTo errinIcializa
Dim RsProduto   As ADODB.Recordset
Dim RsConf      As ADODB.Recordset
If mvarCodProduto = 0 Then
    Inicializa = False
    Exit Function
End If
'abreconexao
'conexaoAdo.BeginTrans
LcSql = "Select * from produtos where codigo=" & mvarCodProduto
Set RsProduto = AbreRecordset(LcSql, True)
'Set RsConf = AbreRecordsetLeitura("select * from configuracoes")mvarArmazenaEmGalpao = GlArmazenaGalpao


'===> Inicializa as Variaveis
If Not RsProduto.EOF Then
'    DoEvents
   mvarDescricaoProduto = RetiraCaracter(RsProduto!Nome & "")
   If Not IsNull(RsProduto!Preco) Then mvarPrecoVenda = RsProduto!Preco Else mvarPrecoVenda = 0
   If Not IsNull(RsProduto!LimiteVenda) Then mvarLimiteVenda = RsProduto!LimiteVenda Else mvarLimiteVenda = 0
   If Not IsNull(RsProduto!MinimoVenda) Then mvarPrecoMinimo = RsProduto!MinimoVenda Else mvarPrecoMinimo = 0
   If Not IsNull(RsProduto!MinimoEst) Then mvarMiminoEstoque = RsProduto!MinimoEst Else mvarMiminoEstoque = 0
   If Not IsNull(RsProduto!maximoEstoque) Then mvarMaximoEstoque = RsProduto!maximoEstoque Else mvarMaximoEstoque = 0
   
   If Not IsNull(RsProduto!UnidMedida) Then mvarCodigoDaUnidade = RsProduto!UnidMedida Else mvarCodigoDaUnidade = 0
   If Not IsNull(RsProduto!QuantEstoque) Then mvarQuantidadeTotalEstoque = RsProduto!QuantEstoque Else mvarQuantidadeTotalEstoque = 0
   If Not IsNull(RsProduto!EstoqueSeguranca) Then mvarQuantidadeTotalEstoqueSeg = RsProduto!EstoqueSeguranca Else mvarQuantidadeTotalEstoqueSeg = 0

   If Not IsNull(RsProduto!QtdMedida) Then mvarQuantidadeDaUnidade = RsProduto!QtdMedida Else mvarQuantidadeDaUnidade = 0
   If Not IsNull(RsProduto!Custo) Then mvarPrecoDeCusto = RsProduto!Custo Else mvarPrecoDeCusto = 0
   '==> A Quantidade será usada para divisão se for zero passa para 1
   If mvarQuantidadeDaUnidade = 0 Then mvarQuantidadeDaUnidade = 1
   
  '==> Caso Armazene em galpoes seta as vairaveis
   If mvarArmazenaEmGalpao Then
       If Not IsNull(RsProduto!santa1) Then mvarQuantidadeSanta1 = RsProduto!santa1 Else mvarQuantidadeSanta1 = 0
       If Not IsNull(RsProduto!Santa2) Then mvarQuantidadeSanta2 = RsProduto!Santa2 Else mvarQuantidadeSanta2 = 0
       If Not IsNull(RsProduto!california) Then mvarQuantidadeCalifornia = RsProduto!california Else mvarQuantidadeCalifornia = 0
        
        
        'mvarQuantidadeSanta1Unitario
       ' mvarQuantidadeSanta2Unitario
       ' mvarQuantidadeSanta1Unitario
   End If
Else
   mvarDescricaoProduto = ""
   mvarPrecoVenda = 0
   mvarLimiteVenda = 0
   mvarPrecoMinimo = 0
   mvarMiminoEstoque = 0
   mvarMiminoEstoque = 0
   mvarQuantidadeTotalUnitario = 0
   mvarQuantidadeTotalEstoque = 0
   mvarQuantidadeSanta1 = 0
   mvarQuantidadeSanta1Unitario = 0
   mvarQuantidadeSanta2 = 0
   mvarQuantidadeSanta2Unitario = 0
   mvarQuantidadeCalifornia = 0
   mvarQuantidadeSanta1 = 0
   mvarQuantidadeSanta1Unitario = 0
End If
Inicializa = True
LcInicializado = True
'conexaoAdo.CommitTrans

Saida:
On Error Resume Next
Set RsConf = Nothing
Set RsProduto = Nothing
Exit Function

errinIcializa:
'conexaoAdo.RollbackTrans
If err.Number = 3265 Then Resume Next
'MsgBox err.Description & err.Number
Inicializa = False
LcInicializado = False
GoTo Saida


End Function
Private Function AtualizaFicha(LcValor As Double, EBaixa As Boolean, Quantidade As Double, CodUnidade As String, Optional Com As Double = 1) As Boolean
On Error GoTo Errfichaestoque
Dim RsFicha As ADODB.Recordset
Dim LcSql As String
Dim LcSaldoAnterior As Double
Dim LcQuantidade As Double
Dim LcTotal As Double
Dim LcSaldoemcaixa As Double
Dim LcSaldoUnitario As Double
Dim LcSaldo As Double
Dim LcTipo As String
Dim LcVaorTotalVenda As Double
Dim TotalM As String

AtualizaFicha = True
Exit Function

LcSql = "Select * from fichadeestoque where codProduto=" & mvarCodProduto & " Order by codigo desc"
'abreconexao
Set RsFicha = AbreRecordset(LcSql, True)

LcQuantidade = Quantidade * Com

LcTotal = Quantidade * LcValor

If Not RsFicha.EOF Then
   LcSaldoAnterior = RsFicha!SaldoGeralUnitario
Else
  LcSaldoAnterior = 0
End If

'===>Acerta o saldo do Estoque

If EBaixa Then
  LcSaldo = LcSaldoAnterior - LcQuantidade
  BuscaCliente
  LcTipo = "S"
Else
  LcSaldo = LcSaldoAnterior + LcQuantidade
  BuscaFornecedor
  LcTipo = "E"
End If

Set RsFicha = Nothing
'==> Monta a sql de atualizacao
LcSaldoemcaixa = TransformaCaixa(LcSaldo)
LcSaldoUnitario = TransformaUnidade(LcSaldo)
LcVaorTotalVenda = (LcValor / LcQuantidade) * LcSaldo

LcSql = "insert into fichadeestoque (nf,codproduto,data,quantidade,unitario,total,saldo,"
LcSql = LcSql & "tipo,clifor,unidade,com,quantunit,valorestoquevenda,saldogeralunitario) values "
LcSql = LcSql & "('" & mvarNF & "'," & mvarCodProduto & ",#" & Format(Date, "mm/dd/yy") & "#,"
LcSql = LcSql & Quantidade & "," & LcValor & "," & LcTotal & "," & LcSaldoemcaixa
LcSql = LcSql & ",'" & LcTipo & "','" & mvarCodClien_forn & "'," & LcVaorTotalVenda & "," & LcSaldo & ")"
'MsgBox LcSql

'LcSql = "UPDATE fichadeestoque SET nf='" & mvarNF & "', codProduto=" & mvarCodProduto
'LcSql = LcSql & ", data=#" & Format(Date, "mm/dd/yy") & "#, quantidade=" & Quantidade
'LcSql = LcSql & ", unitario=" & LcValor & ", total=" & LcTotal & ", saldo=" & LcSaldoemcaixa
'LcSql = LcSql & ", tipo='" & Lctipo & "', clifor='" & mvarCodClien_forn & "', unidade='" & LcNome_Unidade & "'"
'LcSql = LcSql & ",com='" & com & "',QuantUnit=" & LcSaldoUnitario & ", valorEstoqueVenda=" & LcVaorTotalVenda
'LcSql = LcSql & ", SaldoGeralUnitario=" & LcSaldo
TotalM = ExecutaSql(LcSql)

If TotalM > 0 Then
   AtualizaFicha = True
Else
   AtualizaFicha = False
End If
'Inicializa
Exit Function
Saida:

Exit Function

Errfichaestoque:
MsgBox err.Description
Resume 0
AtualizaFicha = False
GoTo Saida
End Function

Public Function BaixaEstoque(ByVal Quantidade As Double, LcValor As Double, DescUnidade As String, Optional Com As Double = 1, Optional Alternativo As Boolean = False) As Boolean
On Error GoTo ErrAtualizando
Dim LcSql As String
Dim LcQuatSanta1 As Double
Dim LcQuatSanta2 As Double
Dim LcQuantCalifornia As Double
Dim LcQuantidadeEstoque As Double
Dim LcquantBaixar As Double
Dim LcModificados As Integer
Dim Baixou As Boolean
'==> Transforma a Quantidade para saber a quantidade a baixar

LcquantBaixar = Quantidade * Com
If GlNaoVerificaEstoque Then
   Baixou = BaixaEstoqueNegativo(Quantidade, LcValor, DescUnidade, Com, Alternativo)
   BaixaEstoque = Baixou
   Exit Function
End If

'===> Verifica se está usando o galpão
If Not mvarArmazenaEmGalpao Then
   '==> Não esta usando o galpao, então vamos baixar direto
    LcQuantidadeEstoque = mvarQuantidadeTotalEstoque - LcquantBaixar
    LcSql = "UPDATE Produtos SET  QuantEstoque=" & LcQuantidadeEstoque
     If Alternativo Then
        LcSql = LcSql & ",EstoqueSeguranca=EstoqueSeguranca-" & LcquantBaixar
     End If
     
     LcSql = LcSql & " where codigo=" & mvarCodProduto
    'abreconexao

   LcModificados = ExecutaSql(LcSql)
    If LcModificados < 1 Then
        BaixaEstoque = False
    Else
        'AtualizaFicha LcValor, True, Quantidade, DescUnidade, com
        BaixaEstoque = True
    End If
    GoTo Saida
End If
'==> Efetua a baixa no santa maria
LcQuatSanta1 = mvarQuantidadeSanta1 - LcquantBaixar

'==> verifica se sobrou do santa maria
'==> Caso o valor de LcquatSanta1 for menor que zero, é que o santa não tinha
'==> Estoque o sufciente
If LcQuatSanta1 < 0 Then
   '==> zera o santa maria, e passa a quantidade a baixar para o proximo
   LcquantBaixar = LcQuatSanta1 * -1
   LcQuatSanta1 = 0
Else
   If LcQuatSanta1 >= 0 Then
       LcquantBaixar = 0
   End If

End If

'==> se a quantidade do santa for zero ou maior que zero, significa que a quantidade foi
'==> o suficiente, Então zeramos a quantidade a baixar

'==> Efetua a baixa no santa maria2
LcQuatSanta2 = mvarQuantidadeSanta2 - LcquantBaixar

'==> verifica se sobrou do santa maria2
'==> Caso o valor de LcquatSanta2 for menor que zero, é que o santa não tinha
'==> Estoque o sufciente
If LcQuatSanta2 < 0 Then
   '==> zera o santa maria2, e passa a quantidade a baixar para o proximo
   LcquantBaixar = LcQuatSanta2 * -1
   LcQuatSanta2 = 0
Else
    If LcQuatSanta2 >= 0 Then
       LcquantBaixar = 0
    End If

End If

'==> se a quantidade do santa2 for zero ou maior que zero, significa que a quantidade foi
'==> o suficiente, Então zeramos a quantidade a baixar

'==> Efetua a baixa no california.
LcQuantCalifornia = mvarQuantidadeCalifornia - LcquantBaixar

'==> verifica se sobrou do California
'==> Caso o valor de LcquatSanta2 for menor que zero, é que o santa não tinha
'==> Estoque o sufciente
If LcQuantCalifornia < 0 Then
   '==> zera o California, e passa a quantidade a baixar para o proximo
   LcquantBaixar = LcQuantCalifornia * -1
   LcQuantCalifornia = 0
Else
    If LcQuantCalifornia >= 0 Then
       LcquantBaixar = 0
    End If
End If

'==> se a quantidade do California for zero ou maior que zero, significa que a quantidade foi
'==> o suficiente, Então zeramos a quantidade a baixar


'==>Agora vamos atualizar a Quantidade Geral
LcQuantidadeEstoque = LcQuatSanta1 + LcQuatSanta2 + LcQuantCalifornia

'==> Cria a sql de Atualização
LcSql = "UPDATE Produtos SET santa1=" & Replace(CCur(LcQuatSanta1), ",", ".") & ", santa2=" & Replace(CCur(LcQuatSanta2), ",", ".") & ", california=" & Replace(CCur(LcQuantCalifornia), ",", ".")
LcSql = LcSql & ", QuantEstoque=" & Replace(CCur(LcQuantidadeEstoque), ",", ".")
If Alternativo Then
   LcSql = LcSql & ",EstoqueSeguranca=EstoqueSeguranca-" & LcquantBaixar
End If

LcSql = LcSql & " where codigo=" & mvarCodProduto

LcModificados = ExecutaSql(LcSql)
If LcModificados < 1 Then
    
        BaixaEstoque = False
   
Else
   If mvarLançaFichaDeEstoque Then AtualizaFicha LcValor, True, Quantidade, DescUnidade, Com
   BaixaEstoque = True
End If

'==> Reatualiza os dados
Inicializa
Saida:
Exit Function
ErrAtualizando:
BaixaEstoque = False
GoTo Saida

End Function
Public Function BaixaEstoqueNegativo(ByVal Quantidade As Double, LcValor As Double, DescUnidade As String, Optional Com As Double = 1, Optional Alternativo As Boolean = False) As Boolean
On Error GoTo ErrAtualizando
Dim LcSql As String
Dim LcQuatSanta1 As Double
Dim LcQuatSanta2 As Double
Dim LcQuantCalifornia As Double
Dim LcQuantidadeEstoque As Double
Dim LcquantBaixar As Double
Dim LcModificados As Integer
Dim LcBaixou As Boolean
'==> Transforma a Quantidade para saber a quantidade a baixar

LcquantBaixar = Quantidade * Com
mvarLançaFichaDeEstoque = True
'LcQuantidadeEstoque = LcQuatSanta1 + LcQuatSanta2 + LcQuantCalifornia
'==> Cria a sql de Atualização
LcQuantidadeEstoque = (mvarQuantidadeSanta1 + mvarQuantidadeSanta2 + mvarQuantidadeCalifornia) - LcquantBaixar

LcSql = "UPDATE Produtos SET santa1=" & Replace(CCur(LcQuantidadeEstoque), ",", ".") & ", santa2=0 , california=0" ' & Replace(CCur(LcQuantCalifornia), ",", ".")
If Alternativo Then
   LcSql = LcSql & ",EstoqueSeguranca=EstoqueSeguranca-" & Replace(CCur(LcquantBaixar), ",", ".")
End If
LcSql = LcSql & ", QuantEstoque=" & Replace(CCur(LcQuantidadeEstoque), ",", ".") & " where codigo=" & mvarCodProduto

LcModificados = ExecutaSql(LcSql)
If LcModificados < 1 Then
    
        BaixaEstoqueNegativo = False
   
Else
   If mvarLançaFichaDeEstoque Then AtualizaFicha LcValor, True, Quantidade, DescUnidade, Com
   BaixaEstoqueNegativo = True
End If
Exit Function


'===> Verifica se está usando o galpão
If Not mvarArmazenaEmGalpao Then
   '==> Não esta usando o galpao, então vamos baixar direto
    LcQuantidadeEstoque = mvarQuantidadeTotalEstoque - LcquantBaixar
    LcSql = "UPDATE Produtos SET  QuantEstoque=" & LcQuantidadeEstoque
    If Alternativo Then
       LcSql = LcSql & ",EstoqueSeguranca=EstoqueSeguranca-" & LcquantBaixar
    End If
    LcSql = LcSql & " where codigo=" & mvarCodProduto
    'abreconexao

   LcModificados = ExecutaSql(LcSql)
    If LcModificados < 1 Then
        BaixaEstoqueNegativo = False
    Else
        'AtualizaFicha LcValor, True, Quantidade, DescUnidade, com
        BaixaEstoqueNegativo = True
    End If
    GoTo Saida
End If

'==> Efetua a baixa no santa maria
If mvarQuantidadeSanta1 >= LcquantBaixar Then
   LcQuatSanta1 = mvarQuantidadeSanta1 - LcquantBaixar
   LcBaixou = True
End If
If mvarQuantidadeSanta2 >= LcquantBaixar And Not LcBaixou Then
   LcQuatSanta2 = mvarQuantidadeSanta2 - LcquantBaixar
   LcBaixou = True
End If
If mvarQuantidadeCalifornia >= LcquantBaixar And Not LcBaixou Then
   LcQuantCalifornia = mvarQuantidadeCalifornia - LcquantBaixar
   LcBaixou = True
End If
If Not LcBaixou Then
   LcQuatSanta1 = mvarQuantidadeSanta1 - LcquantBaixar
   LcBaixou = True
End If

LcQuantidadeEstoque = (mvarQuantidadeSanta1 + mvarQuantidadeSanta2 + mvarQuantidadeCalifornia) - LcquantBaixar
'LcQuantidadeEstoque = LcQuatSanta1 + LcQuatSanta2 + LcQuantCalifornia
'==> Cria a sql de Atualização
LcSql = "UPDATE Produtos SET santa1=" & Replace(CCur(LcQuatSanta1), ",", ".") & ", santa2=" & Replace(CCur(LcQuatSanta2), ",", ".") & ", california=" & Replace(CCur(LcQuantCalifornia), ",", ".")
If Alternativo Then
   LcSql = LcSql & ",EstoqueSeguranca=EstoqueSeguranca-" & Replace(CCur(LcquantBaixar), ",", ".")
End If
LcSql = LcSql & ", QuantEstoque=" & Replace(CCur(LcQuantidadeEstoque), ",", ".") & " where codigo=" & mvarCodProduto

LcModificados = ExecutaSql(LcSql)
If LcModificados < 1 Then
    
        BaixaEstoqueNegativo = False
   
Else
   If mvarLançaFichaDeEstoque Then AtualizaFicha LcValor, True, Quantidade, DescUnidade, Com
   BaixaEstoqueNegativo = True
End If

'==> Reatualiza os dados
Inicializa
Saida:
Exit Function
ErrAtualizando:
BaixaEstoqueNegativo = False
GoTo Saida

End Function
Public Function AcrescentaEstoque(Optional Quantidade As Double = 0, Optional santa1 As Double = 0, Optional Santa2 As Double = 0, Optional california As Double = 0, Optional LcValorUnitario As Double = 0, Optional CodUnidade As String = "0", Optional JaConvertidoNaUnidadeBasica As Boolean = False, Optional CompraSem As Boolean = False) As Boolean
On Error GoTo ErrAcrescimo
Dim LcSql As String
Dim LcQuantidadeGeral As Double
Dim SaldoSem As Double
Dim SaldoGeral As Double
Dim SaldoSanta1 As Double
Dim SaldoSanta2 As Double
Dim SaldoCalifornia As Double
Dim LcNovoSaldoSanta1 As Double
Dim LcNovoSaldoSanta2 As Double
Dim LcNovoSaldoCalifornia As Double
Dim LcNovoSaldoGeral As Double
If mvarCodProduto = 0 Then
    AcrescentaEstoque = False
    Exit Function
End If
Dim Rs As ADODB.Recordset
If Not JaConvertidoNaUnidadeBasica Then
    LcQuantidadeGeral = Quantidade + santa1 + Santa2 + california
    LcQuantidadeGeral = LcQuantidadeGeral * mvarQuantidadeDaUnidade
Else
    LcQuantidadeGeral = Quantidade + santa1 + Santa2 + california
End If

'abreconexao
Set Rs = AbreRecordset("Select * from produtos where codigo=" & mvarCodProduto, True)

If Not Rs.EOF Then
   If CompraSem Then
       If Not IsNull(Rs!EstoqueSeguranca) Then
          SaldoSem = LcQuantidadeGeral + Rs!EstoqueSeguranca
       Else
          SaldoSem = Rs!EstoqueSeguranca
       End If
    Else
      SaldoSem = Rs!EstoqueSeguranca
    End If
   If Not IsNull(Rs!QuantEstoque) Then LcSaldoGeral = Rs!QuantEstoque Else LcSaldoGeral = 0
   If mvarArmazenaEmGalpao Then
        If Not IsNull(Rs!santa1) Then SaldoSanta1 = Rs!santa1 Else SaldoSanta1 = 0
        If Not IsNull(Rs!Santa2) Then SaldoSanta2 = Rs!Santa2 Else SaldoSanta2 = 0
        If Not IsNull(Rs!california) Then SaldoCalifornia = Rs!california Else SaldoCalifornia = 0
   End If
Else
    LcSaldoGeral = 0
    SaldoSanta1 = 0
    SaldoSanta2 = 0
    SaldoCalifornia = 0
End If

'===> Verifica se é galpão
If mvarArmazenaEmGalpao Then
    If Not JaConvertidoNaUnidadeBasica Then
        LcNovoSaldoSanta1 = SaldoSanta1 + (santa1 * mvarQuantidadeDaUnidade)
        LcNovoSaldoSanta2 = SaldoSanta2 + (Santa2 * mvarQuantidadeDaUnidade)
        LcNovoSaldoCalifornia = SaldoCalifornia + (california * mvarQuantidadeDaUnidade)
    Else
        LcNovoSaldoSanta1 = SaldoSanta1 + santa1
        LcNovoSaldoSanta2 = SaldoSanta2 + Santa2
        LcNovoSaldoCalifornia = SaldoCalifornia + california

    End If
    LcNovoSaldoGeral = LcSaldoGeral + ((santa1 + Santa2 + california) * mvarQuantidadeDaUnidade)
Else
    LcNovoSaldoGeral = LcSaldoGeral + LcQuantidadeGeral
End If

'==> Cria a sql para a atualização
LcSql = "UPDATE Produtos SET "
If mvarArmazenaEmGalpao Then
      LcSql = LcSql & "santa1=" & Replace(CStr(LcNovoSaldoSanta1), ",", ".") & ", santa2=" & Replace(CStr(LcNovoSaldoSanta2), ",", ".") & ", california=" & Replace(CStr(LcNovoSaldoCalifornia), ",", ".") & ", "
End If
LcSql = LcSql & " EstoqueSeguranca=" & Replace(CStr(SaldoSem), ",", ".") & ", QuantEstoque=" & Replace(CStr(LcNovoSaldoGeral), ",", ".")


'==> Se informou o Valor , Atualiza o Custo

If LcValorUnitario > 0 Then
   LcV = Replace(LcValorUnitario, ",", ".")
   LcSql = LcSql & ",Custo=" & LcV
End If

LcSql = LcSql & " where codigo=" & mvarCodProduto

'abreconexao
'MsgBox LcSql
LcModificados = ExecutaSql(LcSql)
If LcModificados < 1 Then
   AcrescentaEstoque = False
Else
   If mvarLançaFichaDeEstoque Then AtualizaFicha LcValorUnitario, False, LcQuantidadeGeral, CodUnidade
   AcrescentaEstoque = True
End If

AcrescentaEstoque = True
'==> Reatualiza os dados
Inicializa
Saida:
Exit Function
ErrAcrescimo:
AcrescentaEstoque = False
GoTo Saida


End Function
Public Function EstornaEstoque(Optional Quantidade As Double = 0, Optional santa1 As Double = 0, Optional Santa2 As Double = 0, Optional california As Double = 0, Optional LcValor As Double = 0, Optional CodUnidade As String = "0") As Boolean
On Error GoTo ErrAcrescimo
Dim LcSql As String
Dim LcQuantidadeGeral As Double
Dim SaldoGeral As Double
Dim SaldoSanta1 As Double
Dim SaldoSanta2 As Double
Dim SaldoCalifornia As Double
Dim LcNovoSaldoSanta1 As Double
Dim LcNovoSaldoSanta2 As Double
Dim LcNovoSaldoCalifornia As Double
Dim LcNovoSaldoGeral As Double

Dim Rs As ADODB.Recordset
LcQuantidadeGeral = Quantidade + santa1 + Santa2 + california
'LcQuantidadeGeral = LcQuantidadeGeral * mvarQuantidadeDaUnidade
'abreconexao
Set Rs = AbreRecordset("Select * from produtos where codigo=" & mvarCodProduto, True)

If Not Rs.EOF Then
   If Not IsNull(Rs!QuantEstoque) Then LcSaldoGeral = Rs!QuantEstoque Else LcSaldoGeral = 0
   If Not IsNull(Rs!santa1) Then SaldoSanta1 = Rs!santa1 Else SaldoSanta1 = 0
   If Not IsNull(Rs!Santa2) Then SaldoSanta2 = Rs!Santa2 Else SaldoSanta2 = 0
   If Not IsNull(Rs!california) Then SaldoCalifornia = Rs!california Else SaldoCalifornia = 0
Else
    LcSaldoGeral = 0
    SaldoSanta1 = 0
    SaldoSanta2 = 0
    SaldoCalifornia = 0
End If

'===> Verifica se é galpão
If mvarArmazenaEmGalpao Then
    LcNovoSaldoSanta1 = SaldoSanta1 + (santa1)
    LcNovoSaldoSanta2 = SaldoSanta2 + (Santa2)
    LcNovoSaldoCalifornia = SaldoCalifornia + (california)
    LcNovoSaldoGeral = LcNovoSaldoSanta1 + LcNovoSaldoSanta2 + LcNovoSaldoCalifornia
Else
    LcNovoSaldoGeral = LcSaldoGeral + (LcQuantidadeGeral)
End If

'==> Cria a sql para a atualização
LcSql = "UPDATE Produtos SET santa1=" & LcNovoSaldoSanta1 & ", santa2=" & LcNovoSaldoSanta2 & ", california=" & LcNovoSaldoCalifornia
LcSql = LcSql & ", QuantEstoque=" & LcNovoSaldoGeral & " where codigo=" & mvarCodProduto

'abreconexao
'MsgBox LcSql
LcModificados = ExecutaSql(LcSql)
If LcModificados < 1 Then
   EstornaEstoque = False
Else
   If mvarLançaFichaDeEstoque Then AtualizaFicha LcValor, False, LcQuantidadeGeral, CodUnidade
   EstornaEstoque = True
End If

EstornaEstoque = True
'==> Reatualiza os dados
Inicializa
Saida:
Exit Function
ErrAcrescimo:
If err.Number = 3265 Then Resume Next
MsgBox err.Description & err.Number
EstornaEstoque = False
GoTo Saida


End Function

Private Function TransformaCaixa(ByVal Quantidade As Double) As Double
On Error GoTo errtrans
    Dim LcQuantidade As Double
    If Not LcInicializado Then Inicializa
    If mvarQuantidadeDaUnidade = 0 Then mvarQuantidadeDaUnidade = 1
    LcQuantidade = Fix(Quantidade / mvarQuantidadeDaUnidade)
    TransformaCaixa = LcQuantidade
    Exit Function
errtrans:
Dim LcRespErr As Integer
TransformaCaixa = 0
Exit Function


    
End Function
Private Function TransformaUnidade(Quantidade As Double) As Double
On Error GoTo errtr
    Dim LcQuantidade As Double
    Dim LcQuantUnitario As Double
    On Error Resume Next
    If Not LcInicializado Then Inicializa
    LcQuantidade = Quantidade / mvarQuantidadeDaUnidade
    LcQuantUnitario = LcQuantidade - Fix(LcQuantidade)
    LcQuantUnitario = LcQuantUnitario * mvarQuantidadeDaUnidade
    
    TransformaUnidade = Round(LcQuantUnitario, 0)
    Exit Function
    
    
   '  TransformaUnidade = Round((((Quantidade / mvarQuantidadeDaUnidade) - Fix(Quantidade / mvarQuantidadeDaUnidade)) * mvarQuantidadeDaUnidade), 0)
    
errtr:
TransformaUnidade = 0
Exit Function

End Function
Public Property Get EstoqueSegurancaTotalFechado() As Double
    If Not LcInicializado Then Inicializa
    EstoqueSegurancaTotalFechado = TransformaCaixa(mvarQuantidadeTotalEstoqueSeg)
    
End Property
Public Property Get EstoqueTotalFechado() As Double
    If Not LcInicializado Then Inicializa
    If Not mvarArmazenaEmGalpao Then
        EstoqueTotalFechado = TransformaCaixa(mvarQuantidadeTotalEstoque)
    Else
        EstoqueTotalFechado = Santa1Fechado + Santa2Fechado + QuantidadeCaliforniaFechado
    End If
End Property
Public Property Get EstoqueGeral() As Double
    If Not LcInicializado Then Inicializa
    EstoqueGeral = mvarQuantidadeTotalEstoque
End Property

Public Property Get EstoqueTotalUnitario() As Double
    If Not LcInicializado Then Inicializa
    If Not mvarArmazenaEmGalpao Then
        EstoqueTotalUnitario = TransformaUnidade(mvarQuantidadeTotalEstoque)
    Else
        EstoqueTotalUnitario = Santa1Unitario + Santa2Unitario + QuantidadeCaliforniaUnitario
    End If
End Property

Public Property Get MiminoEstoque() As Double
'==> Informa a quantidade Minima do Estoque deste produto
    If Not LcInicializado Then Inicializa
    MiminoEstoque = mvarMiminoEstoque
End Property

Public Property Get PrecoMinimo() As Double
'==> informa o preço minimo de venda para este produto
    If Not LcInicializado Then Inicializa
    PrecoMinimo = mvarPrecoMinimo
End Property

Public Property Get LimiteVenda() As Double
'==> Informa o preço de venda para este produto
    If Not LcInicializado Then Inicializa
    LimiteVenda = mvarLimiteVenda
End Property
Public Property Get PrecoVenda() As Double
'==> Informa o preço de venda para este produto
    If Not LcInicializado Then Inicializa
    PrecoVenda = mvarPrecoVenda
End Property
Public Property Get Santa1Fechado() As Double
'==>Informa a quantidade em caixas do Santa1
    If Not LcInicializado Then Inicializa
    Santa1Fechado = TransformaCaixa(mvarQuantidadeSanta1)
End Property
Public Property Get Santa2Fechado() As Double
'==>Informa a quantidade em caixas do Santa1
    If Not LcInicializado Then Inicializa
    Santa2Fechado = TransformaCaixa(mvarQuantidadeSanta2)
End Property
Public Property Get QuantidadeCaliforniaFechado() As Double
'==>Informa a quantidade em caixas do california
    If Not LcInicializado Then Inicializa
   QuantidadeCaliforniaFechado = TransformaCaixa(mvarQuantidadeCalifornia)
    
End Property
Public Property Get Santa1Unitario() As Double
    If Not LcInicializado Then Inicializa
    Santa1Unitario = TransformaUnidade(mvarQuantidadeSanta1)
End Property

Public Property Get Santa2Unitario() As Double
    If Not LcInicializado Then Inicializa
    Santa2Unitario = TransformaUnidade(mvarQuantidadeSanta2)
End Property
Public Property Get QuantidadeCaliforniaUnitario() As Double
'==>Informa a quantidade Unitario em santa Maria
    If Not LcInicializado Then Inicializa
QuantidadeCaliforniaUnitario = TransformaUnidade(mvarQuantidadeCalifornia)
End Property



Public Property Get DescricaoProduto() As String
'==> Informa a descrição do Produto
    If Not LcInicializado Then Inicializa
    DescricaoProduto = mvarDescricaoProduto
End Property


Public Property Let CodProduto(ByVal vData As Integer)
    mvarCodProduto = vData
    '==> Chama o Registro para a inicialização do sistema
    Inicializa
End Property


Public Property Get CodProduto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodProduto
    CodProduto = mvarCodProduto
End Property

Private Sub Class_Initialize()
mvarArmazenaEmGalpao = GlArmazenaGalpao
End Sub
