Attribute VB_Name = "MovimentoFiscal"
Function SaidaFiscal(codigoproduto As Long, Quantidade As Double, Data As Date, Nome As String, Optional CustoUnit As Double = 0, Optional CustoTotal As Double = 0)
'==> Esta função recebe os dados dos produtos ja com os valores referentes
'==> a unidade principal do produto, e salva na tabela de estoque fiscal.
Dim StrSql      As String
Dim RsEstoque   As ADODB.Recordset
Dim Saldo       As Double

StrSql = "Select * from estoquefiscal where codigoproduto=" & codigoproduto & " order by codigo desc"

Set RsEstoque = AbreRecordset(StrSql, True)

If Not RsEstoque.EOF Then
   Saldo = RsEstoque!Saldo
Else
   Saldo = 0
End If
Set RsEstoque = Nothing
Saldo = Saldo - Quantidade
StrSql = "insert into estoquefiscal (codigoproduto,nome,quantidadeSaida,valorcustomediounitario,vcustototal,saldo,data,quantidade) values (" & _
         codigoproduto & ",'" & _
         Nome & "'," & _
         Replace(Quantidade, ",", ".") & "," & _
         Replace(CustoUnit, ",", ".") & "," & _
         Replace(CustoTotal, ",", ".") & "," & _
         Replace(Saldo, ",", ".") & ",'" & _
         Format(Data, "yyyy-mm-dd") & "',0)"
         
afetados = ExecutaSql(StrSql)
End Function
Function EntradaFiscal(codigoproduto As Long, Quantidade As Double, Data As Date, Nome As String, Optional CustoUnit As Double = 0, Optional CustoTotal As Double = 0)
'==> Esta função recebe os dados dos produtos ja com os valores referentes
'==> a unidade principal do produto, e salva na tabela de estoque fiscal.
Dim StrSql      As String
Dim RsEstoque   As ADODB.Recordset
Dim Saldo       As Double

StrSql = "Select * from estoquefiscal where codigoproduto=" & codigoproduto & " order by codigo desc"

Set RsEstoque = AbreRecordset(StrSql, True)

If Not RsEstoque.EOF Then
   Saldo = IIf(Not IsNull(RsEstoque!Saldo), RsEstoque!Saldo, 0)
Else
   Saldo = 0
End If
Set RsEstoque = Nothing
Saldo = Saldo + Quantidade
StrSql = "insert into estoquefiscal (codigoproduto,nome,quantidade,valorcustomediounitario,vcustototal,saldo,data,quantidadeSaida) values (" & _
         codigoproduto & ",'" & _
         Nome & "'," & _
         Replace(Quantidade, ",", ".") & "," & _
         Replace(CustoUnit, ",", ".") & "," & _
         Replace(CustoTotal, ",", ".") & "," & _
         Replace(Saldo, ",", ".") & ",'" & _
         Format(Data, "yyyy-mm-dd") & "',0)"
         
afetados = ExecutaSql(StrSql)
End Function
Function BuscarSaldoAnterior(codigoproduto As Long) As String

Dim StrSql As String
Dim Rs As ADODB.Recordset
Dim Saldo As Double
Dim StrSaldo As String

StrSql = "Select * from estoquefiscal where CodigoProduto=" & codigoproduto & " and data<'" & Format(Inventario.Datai.Text, "yyyy-mm-dd") & "' order by codigo desc"

Set Rs = AbreRecordset(StrSql, True)

If Not Rs.EOF Then
   Saldo = IIf(Not IsNull(Rs!Saldo), Rs!Saldo, 0)
Else
   Saldo = 0
End If
Set Rs = Nothing
StrSaldo = AcertaNumero(CStr(Saldo), 2)
BuscarSaldoAnterior = StrSaldo


End Function

Function BuscarSaldoUltimo(codigoproduto As Long) As String

Dim StrSql As String
Dim Rs As ADODB.Recordset
Dim Saldo As Double
Dim StrSaldo As String

StrSql = "Select * from estoquefiscal where CodigoProduto=" & codigoproduto & " and data<='" & Format(Inventario.Dataf.Text, "yyyy-mm-dd") & "' order by codigo desc"

Set Rs = AbreRecordset(StrSql, True)

If Not Rs.EOF Then
   Saldo = IIf(Not IsNull(Rs!Saldo), Rs!Saldo, 0)
Else
   Saldo = 0
End If
Set Rs = Nothing
StrSaldo = AcertaNumero(CStr(Saldo), 2)
BuscarSaldoUltimo = StrSaldo


End Function
