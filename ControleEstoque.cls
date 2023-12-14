VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControleEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private StrSql As String
Public Function MovimentacaoManual(Codigo As Long, santa As Currency, california As Currency, Doc As String, ClienteFornecedor As String, UNID As String, CodUnidade As String, Tipo As String) As String
Dim LcSaldoAnterior As Currency
Dim LcSaldo As Currency
Dim LcSaldoSanta As Currency
Dim LcSaldoCalifornia As Currency
Dim LcSaldoSantaAnterior As Currency
Dim LcSaldoCaliforniaAnterior As Currency
Dim LcSaldoBaixar As Currency
Dim afetados As Integer
Dim Erro As String
Dim LancamentoEstoqueOK As Boolean
Dim LancamentoFichaOK As Boolean
Dim QuantidadeVerificacao As Integer
Dim LCDescricao As String

Dim LcQuantBaixadoSanta As Currency
Dim LCQuantBaixadoCalifornia As Currency
Dim LcNovoSaldo As Currency

On Error GoTo ErroBaixa

'==> Inicia as Variaveis
afetados = 0
LcSaldoAnterior = 0
LcSaldo = 0
LcSaldoSanta = 0
LcSaldoCalifornia = 0
LcSaldoSantaAnterior = 0
LcSaldoCaliforniaAnterior = 0
LcSaldoBaixar = 0
LcQuantBaixadoSanta = 0
LCQuantBaixadoCalifornia = 0

'==> Recupera os saldos
RecuperaSaldoEstoque Codigo, LcSaldoSantaAnterior, LcSaldoCaliforniaAnterior, 0, 0, 0, 0, 0, 0, LcSaldoAnterior

If Tipo = "E" Then
    LcSaldoSanta = LcSaldoSantaAnterior + santa
    LcSaldoCalifornia = LcSaldoCaliforniaAnterior + california
    LcNovoSaldo = LcSaldoSanta + LcSaldoCalifornia
Else
    LcSaldoSanta = LcSaldoSantaAnterior - santa
    LcSaldoCalifornia = LcSaldoCaliforniaAnterior - california
    LcNovoSaldo = LcSaldoSanta + LcSaldoCalifornia
End If

'====> Atualiza o Estoque
StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
StrSql = StrSql & " where codigo=" & Codigo

afetados = ExecutaSql(StrSql)
If afetados = 0 Then
    '==> Dispara um erro
    err.Raise 1020, DEscricaoErro, DEscricaoErro
End If

'==> Passou entao vamos conferir se Esta ok.
LancamentoEstoqueOK = VerificaLancamentoEstoque(Codigo, LcNovoSaldo, LcSaldoSanta, LcSaldoCalifornia)
QuantidadeVerificacao = 0
Do While Not LancamentoEstoqueOK
   If QuantidadeVerificacao >= 10 Then
      err.Raise 1021, "Falha Lançando o Estoque do Produto:" & Codigo, "Falha Lançando o Estoque do Produto:" & Codigo
   End If
   StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
   StrSql = StrSql & " where codigo=" & Codigo
    
   afetados = ExecutaSql(StrSql)
   If afetados = 0 Then
        '==> Dispara um erro
        err.Raise 1020, DEscricaoErro, DEscricaoErro
   End If
   QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop
'===> O Estoque Esta ok Vamos Lancar a Ficha de Estoque
LancaFicha Codigo, LCDescricao, santa, california, Tipo, GlUsuario, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo
'===> Verifica se a ficha foi lancada.
LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, GlUsuario, Tipo)
QuantidadeVerificacao = 0
Do While Not LancamentoFichaOK
    If QuantidadeVerificacao >= 10 Then
          err.Raise 1021, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo
    End If
    LancaFicha Codigo, LCDescricao, santa, california, Tipo, Doc, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo

    LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, Doc, Tipo)
    QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop

Saida:
MovimentacaoManual = Erro
Exit Function

ErroBaixa:
Erro = err.Description


End Function
Public Function EntradaEstoque(Codigo As Long, santa As Currency, california As Currency, Doc As String, ClienteFornecedor As String, UNID As String, CodUnidade As String) As String
Dim LcSaldoAnterior As Currency
Dim LcSaldo As Currency
Dim LcSaldoSanta As Currency
Dim LcSaldoCalifornia As Currency
Dim LcSaldoSantaAnterior As Currency
Dim LcSaldoCaliforniaAnterior As Currency
Dim LcSaldoBaixar As Currency
Dim afetados As Integer
Dim Erro As String
Dim LancamentoEstoqueOK As Boolean
Dim LancamentoFichaOK As Boolean
Dim QuantidadeVerificacao As Integer
Dim LCDescricao As String

Dim LcQuantBaixadoSanta As Currency
Dim LCQuantBaixadoCalifornia As Currency
Dim LcNovoSaldo As Currency

On Error GoTo ErroBaixa

'==> Inicia as Variaveis
afetados = 0
LcSaldoAnterior = 0
LcSaldo = 0
LcSaldoSanta = 0
LcSaldoCalifornia = 0
LcSaldoSantaAnterior = 0
LcSaldoCaliforniaAnterior = 0
LcSaldoBaixar = 0
LcQuantBaixadoSanta = 0
LCQuantBaixadoCalifornia = 0

'==> Recupera os saldos
RecuperaSaldoEstoque Codigo, LcSaldoSantaAnterior, LcSaldoCaliforniaAnterior, 0, 0, 0, 0, 0, 0, LcSaldoAnterior
LcSaldoSanta = LcSaldoSantaAnterior + santa
LcSaldoCalifornia = LcSaldoCaliforniaAnterior + california

LcNovoSaldo = LcSaldoSanta + LcSaldoCalifornia

'====> Atualiza o Estoque
StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
StrSql = StrSql & " where codigo=" & Codigo

afetados = ExecutaSql(StrSql)
If afetados = 0 Then
    '==> Dispara um erro
    err.Raise 1020, DEscricaoErro, DEscricaoErro
End If

'==> Passou entao vamos conferir se Esta ok.
LancamentoEstoqueOK = VerificaLancamentoEstoque(Codigo, LcNovoSaldo, LcSaldoSanta, LcSaldoCalifornia)
QuantidadeVerificacao = 0
Do While Not LancamentoEstoqueOK
   If QuantidadeVerificacao >= 10 Then
      err.Raise 1021, "Falha Lançando o Estoque do Produto:" & Codigo, "Falha Lançando o Estoque do Produto:" & Codigo
   End If
   StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
   StrSql = StrSql & " where codigo=" & Codigo
    
   afetados = ExecutaSql(StrSql)
   If afetados = 0 Then
        '==> Dispara um erro
        err.Raise 1020, DEscricaoErro, DEscricaoErro
   End If
   QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop
'===> O Estoque Esta ok Vamos Lancar a Ficha de Estoque
LancaFicha Codigo, LCDescricao, santa, california, "E", Doc, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo
'===> Verifica se a ficha foi lancada.
LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, Doc, "E")
QuantidadeVerificacao = 0
Do While Not LancamentoEstoqueOK
    If QuantidadeVerificacao >= 10 Then
          err.Raise 1021, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo
    End If
    LancaFicha Codigo, LCDescricao, LcQuantBaixadoSanta, LCQuantBaixadoCalifornia, "S", Doc, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo
    LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, Doc, "E")
    QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop

Saida:
EntradaEstoque = Erro
Exit Function

ErroBaixa:
Erro = err.Description


End Function
Public Function BaixaEstoque(Codigo As Long, Quantidade As Currency, Doc As String, ClienteFornecedor As String, UNID As String, CodUnidade As String) As String
Dim LcSaldoAnterior As Currency
Dim LcSaldo As Currency
Dim LcSaldoSanta As Currency
Dim LcSaldoCalifornia As Currency
Dim LcSaldoSantaAnterior As Currency
Dim LcSaldoCaliforniaAnterior As Currency
Dim LcSaldoBaixar As Currency
Dim afetados As Integer
Dim Erro As String
Dim LancamentoEstoqueOK As Boolean
Dim LancamentoFichaOK As Boolean
Dim QuantidadeVerificacao As Integer
Dim LCDescricao As String

Dim LcQuantBaixadoSanta As Currency
Dim LCQuantBaixadoCalifornia As Currency
Dim LcNovoSaldo As Currency

On Error GoTo ErroBaixa

'==> Inicia as Variaveis
afetados = 0
LcSaldoAnterior = 0
LcSaldo = 0
LcSaldoSanta = 0
LcSaldoCalifornia = 0
LcSaldoSantaAnterior = 0
LcSaldoCaliforniaAnterior = 0
LcSaldoBaixar = 0
LcQuantBaixadoSanta = 0
LCQuantBaixadoCalifornia = 0

'==> Recupera os saldos
RecuperaSaldoEstoque Codigo, LcSaldoSantaAnterior, LcSaldoCaliforniaAnterior, 0, 0, 0, 0, 0, 0, LcSaldoAnterior
'==> Calcula o novo Saldo
LcSaldoBaixar = Quantidade
LcSaldo = LcSaldoAnterior - Quantidade

'==> Verifica se tem o saldo no Santa
If LcSaldoSantaAnterior >= Quantidade Then
   '==> No santa Maria tem a Quantidade para Baixar
   LcSaldoSanta = LcSaldoSantaAnterior - LcSaldoBaixar
   LcQuantBaixadoSanta = LcSaldoBaixar
   LcSaldoBaixar = 0
 ElseIf LcSaldoSantaAnterior > 0 Then
  '===> Tem Parte no santa
   LcSaldoSanta = 0
   LcSaldoBaixar = LcSaldoBaixar - LcSaldoSantaAnterior
   LcQuantBaixadoSanta = LcQuantBaixadoSanta + LcSaldoSantaAnterior
End If

If LcSaldoBaixar > 0 Then
   '==> Nao foi baixado tudo no santa, entao processa no california
    If LcSaldoCaliforniaAnterior >= LcSaldoBaixar Then
        '==> No California tem a Quantidade para Baixar
        LcSaldoCalifornia = LcSaldoCaliforniaAnterior - LcSaldoBaixar
        LCQuantBaixadoCalifornia = LcSaldoBaixar
        LcSaldoBaixar = 0
    ElseIf LcSaldoCaliforniaAnterior > 0 Then
        '===> Tem Parte no santa
        LcSaldoCalifornia = 0
        LcSaldoBaixar = LcSaldoBaixar - LcSaldoCalifornia
        LcSaldoCalifornia = LcSaldoCalifornia
    End If
Else
   LcSaldoCalifornia = LcSaldoCaliforniaAnterior
End If

If LcSaldoBaixar > 0 Then
    '===> Ainda tem Quantidade de Estoque, vai baixar no Santa mesmo deixando negativo.
    LcSaldoSanta = LcSaldoSantaAnterior - LcSaldoBaixar
    LcQuantBaixadoSanta = LcQuantBaixadoSanta + LcSaldoBaixar
    LcSaldoBaixar = 0
End If
LcNovoSaldo = LcSaldoSanta + LcSaldoCalifornia

'====> Atualiza o Estoque
StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
StrSql = StrSql & " where codigo=" & Codigo

afetados = ExecutaSql(StrSql)
If afetados = 0 Then
    '==> Dispara um erro
    err.Raise 1020, DEscricaoErro, DEscricaoErro
End If

'==> Passou entao vamos conferir se Esta ok.
LancamentoEstoqueOK = VerificaLancamentoEstoque(Codigo, LcNovoSaldo, LcSaldoSanta, LcSaldoCalifornia)
QuantidadeVerificacao = 0
Do While Not LancamentoEstoqueOK
   If QuantidadeVerificacao >= 10 Then
      err.Raise 1021, "Falha Lançando o Estoque do Produto:" & Codigo, "Falha Lançando o Estoque do Produto:" & Codigo
   End If
   StrSql = "Update Produtos set QuantEstoque=" & Replace(CStr(LcNovoSaldo), ",", ".") & ",Santa2=" & Replace(CStr(LcSaldoSanta), ",", ".") & ",California=" & Replace(CStr(LcSaldoCalifornia), ",", ".")
   StrSql = StrSql & " where codigo=" & Codigo
    
   afetados = ExecutaSql(StrSql)
   If afetados = 0 Then
        '==> Dispara um erro
        err.Raise 1020, DEscricaoErro, DEscricaoErro
   End If
   QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop
'===> O Estoque Esta ok Vamos Lancar a Ficha de Estoque
LancaFicha Codigo, LCDescricao, LcQuantBaixadoSanta, LCQuantBaixadoCalifornia, "S", Doc, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo
'===> Verifica se a ficha foi lancada.
LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, Doc, "S")
QuantidadeVerificacao = 0
Do While Not LancamentoEstoqueOK
    If QuantidadeVerificacao >= 10 Then
          err.Raise 1021, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo, "Falha Lançando a Ficha de Estoque do Produto:" & Codigo
    End If
    LancaFicha Codigo, LCDescricao, LcQuantBaixadoSanta, LCQuantBaixadoCalifornia, "S", Doc, ClienteFornecedor, UNID, CodUnidade, LcSaldoAnterior, LcNovoSaldo
    LancamentoFichaOK = VerificaLancamentoFicha(Codigo, LcNovoSaldo, LcSaldoAnterior, Doc, "S")
    QuantidadeVerificacao = QuantidadeVerificacao + 1
Loop

Saida:
BaixaEstoque = Erro
Exit Function

ErroBaixa:
Erro = err.Description


End Function

Private Function VerificaLancamentoEstoque(Codigo As Long, Saldo As Currency, SaldoSanta As Currency, SaldoCalifornia As Currency) As Boolean
Dim Resposta As Boolean
On Erro GoTo ErroVerifica
Resposta = False
Dim Rs As ADODB.Recordset
StrSql = "Select QuantEstoque,Santa2,California from Produtos where codigo=" & Codigo
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
    If CDec(Rs!QuantEstoque) = CDec(Saldo) And CDec(Rs!Santa2) = CDec(SaldoSanta) And CDec(Rs!california) = CDec(SaldoCalifornia) Then
       Resposta = True
    Else
        Resposta = False
    End If
Else
   Resposta = False
End If

Saida:
VerificaLancamentoEstoque = Resposta
Exit Function

ErroVerifica:
Resposta = False
GoTo Saida
End Function
Private Function VerificaLancamentoFicha(Codigo As Long, Saldo As Currency, Anterior As Currency, Nota As String, Tipo As String) As Boolean
Dim Resposta As Boolean
On Erro GoTo ErroVerifica
Resposta = False
Dim Rs As ADODB.Recordset

StrSql = "Select Produto,Anterior,Saldo,NF,Data,Tipo  from historicoproduto where Produto=" & Codigo & " and NF='" & Nota & "' and data='" & Format(Date, "yyyy-mm-dd") & "' and Tipo='" & Tipo & "' order by codigo desc LIMIT 1 "
'Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
   If Rs!Saldo = Saldo And Rs!Anterior = Anterior Then
       Resposta = True
   Else
       Resposta = False
   End If
Else
   Resposta = False
End If

VerificaLancamentoFicha = Resposta
Exit Function
ErroVerifica:
VerificaLancamentoFicha = Resposta

End Function
Public Function LancaFicha(Codigo As Long, Descricao As String, QuantSanta As Currency, QuantCalifornia As Currency, Tipo As String, Doc As String, Nome As String, Unidade As String, CodUnidade As String, Anterior As Currency, Saldo As Currency) As Boolean
Dim afetados As Integer

StrSql = "insert into historicoproduto (Produto,Descricao,Santa2,California,NF,Data,Tipo,"
StrSql = StrSql & "Unidade,CodUnid,ClienteForn,Anterior,Saldo) values "
StrSql = StrSql & "(" & Codigo & ",'" & Replace(Descricao, "'", "''") & "'," & Replace(QuantSanta, ",", ".") & "," & Replace(QuantCalifornia, ",", ".") & ","
StrSql = StrSql & "'" & Doc & "','" & Format(Date, "yyyy-mm-dd") & "',"
StrSql = StrSql & "'" & Tipo & "','" & Unidade & "'," & CodUnidade & ",'" & Replace(Nome, "'", "''") & "',"
StrSql = StrSql & Replace(Anterior, ",", ".") & "," & Replace(Saldo, ",", ".") & ")"
'Debug.Print StrSql
afetados = ExecutaSql(StrSql)
End Function

Public Function RecuperaSaldoEstoque(Codigo As Long, ByRef QtdSanta As Currency, ByRef QtdCalifornia As Currency, ByRef QtdSantaBase As Currency, ByRef QtdCaliforniaBase As Currency, ByRef QtdSAntaUnidadeRestante As Currency, ByRef QtdCaliforniaRestante As Currency, ByRef SaldoBase As Currency, ByRef SaldoBaseUnidadeRestante As Currency, ByRef SaldoGeral As Currency)
Dim Rs As ADODB.Recordset
Dim StrSql As String
StrSql = "Select QuantEstoque,Santa2,California,QtdMedida from Produtos where codigo=" & Codigo
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
    Dim LcQtmedida As Currency
    QtdSanta = Rs!Santa2
    QtdCalifornia = Rs!california
    LcQtmedida = Rs!QtdMedida
    
    QtdSantaBase = AcertaNumero(QtdSanta / LcQtmedida, 0)
    QtdCaliforniaBase = AcertaNumero(QtdCalifornia / LcQtmedida, 0)
    
    '===> Calcula a Quantidade Restante
    QtdSAntaUnidadeRestante = QtdSanta - (QtdSantaBase * LcQtmedida)
    QtdCaliforniaRestante = QtdCalifornia - (QtdCaliforniaBase * LcQtmedida)
    
    '===> Calcula o Saldo Total
    SaldoBase = QtdSantaBase + QtdCaliforniaBase
    SaldoBaseUnidadeRestante = QtdSAntaUnidadeRestante + QtdCaliforniaRestante
    SaldoGeral = QtdSanta + QtdCalifornia
Else
    QtdSanta = 0
    QtdCalifornia = 0
    LcQtmedida = R0
    
    QtdSantaBase = 0
    QtdCaliforniaBase = 0
    
    '===> Calcula a Quantidade Restante
    QtdSAntaUnidadeRestante = 0
    QtdCaliforniaRestante = 0
    
    '===> Calcula o Saldo Total
    SaldoBase = 0
    SaldoBaseUnidadeRestante = 0
    SaldoGeral = 0
End If

End Function
Sub TransfornaUnidade(Lc_Quantidade As Currency, Com As Integer, ByRef QuantCaixa As Currency, ByRef QuantUnidade As Currency)
Dim LcCaixa As Currency
Dim LcQuantUnidade As Currency
Dim LcTemp As Currency
LcTemp = Lc_Quantidade / Com
QuantCaixa = Fix(LcTemp)
LcQuantUnidade = LcTemp - QuantCaixa
QuantUnidade = LcQuantUnidade * Com

End Sub





