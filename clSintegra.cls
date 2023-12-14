VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClSintegra"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private ErrosP As String
Private ClUtil As New Utilitario
Private mvarOrigem As String
Private Type DadosInventario
    Codigo As Long
    Quantidade As Double
    ValorUnitario As Double
End Type

Public Property Let Cliente_Origem(ByVal vData As String)
    mvarOrigem = vData
End Property


Public Property Get Origem() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Origem
   Origem = mvarOrigem
   
End Property

Public Property Get erroS() As Variant
    
        erroS = ErrosP
    
End Property

Private Sub Inicio()
'NomedoArquivo = "Sintegra"
'abreconexao

End Sub
Public Sub BuscaRegistro74Lidis(DataInicio As Date, DataFim As Date)
Dim RsEntrada As ADODB.Recordset
Dim RsSaida As ADODB.Recordset
Dim RsProd As ADODB.Recordset
Dim RsI     As ADODB.Recordset
Dim StrSql As String
Dim MtI() As DadosInventario
Dim a As Integer
Dim Quantidade As Double
Dim Achou As Boolean
Dim b As Integer

a = 0

'==> Abre a tb de entrada
StrSql = "SELECT entradanf.DATA, itensentradanf.ITEM, itensentradanf.QTDE, itensentradanf.VALUNIT, itensentradanf.descricao, itensentradanf.QTDUM " & _
         "FROM entradanf INNER JOIN itensentradanf ON entradanf.codigo = itensentradanf.CodigoNota " & _
         "WHERE (((entradanf.DATA) Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "#)) " & _
         "ORDER BY itensentradanf.ITEM;"

'Debug.Print StrSql

Set RsEntrada = AbreRecordset(StrSql, True)
'==> Abre a tb de Saida
StrSql = "SELECT alid050.DTEMIS, alid052.codProd, alid052.descricao, alid052.QTDE, alid052.QTDUM, alid052.VALUNIT " & _
         "FROM alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF " & _
         "WHERE (((alid050.DTEMIS) Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "#)) " & _
         "ORDER BY alid052.codProd;"
'Debug.Print StrSql
Set RsSaida = AbreRecordset(StrSql, True)

'==> Processa as entradas
Do Until RsEntrada.EOF
    Sintegra.Caption = "Recuperando Entradas do Produto:" & RsEntrada!Item & " " & RsEntrada!Descricao
    DoEvents

   Achou = False
   Quantidade = 0
   If a = 0 Then
      '==> é o primeiro registro
      ReDim Preserve MtI(a)
      Quantidade = RsEntrada!Qtde
      MtI(a).Codigo = RsEntrada!Item
      MtI(a).Quantidade = Quantidade
      MtI(a).ValorUnitario = RsEntrada!VALUNIT
      a = a + 1
   Else
      Quantidade = RsEntrada!Qtde
      '==> Ja existe, vamos ver se tem o codigo na matriz
      For b = 0 To UBound(MtI)
          If MtI(b).Codigo = RsEntrada!Item Then
             Achou = True
             Exit For
          End If
      Next
      
      If Achou Then
         MtI(b).Quantidade = MtI(b).Quantidade + Quantidade
         MtI(b).ValorUnitario = RsEntrada!VALUNIT
      Else
        ReDim Preserve MtI(a)
        Quantidade = RsEntrada!Qtde
        MtI(a).Codigo = RsEntrada!Item
        MtI(a).Quantidade = Quantidade
        MtI(a).ValorUnitario = RsEntrada!VALUNIT
        a = a + 1
      End If
   End If
   RsEntrada.MoveNext
Loop
'==> Processa as Saidas
Do Until RsSaida.EOF
    Sintegra.Caption = "Recuperando saidas do Produto:" & RsSaida!codProd & " " & RsSaida!Descricao
    DoEvents
 
    StrSql = "Select * from PRodutos where codigo=" & RsSaida!codProd
    
    Set RsProd = AbreRecordset(StrSql, True)
   ' MsgBox DEscricaoErro
    
    If Not RsProd.EOF Then
        QuantidadeUn = RsProd!QtdMedida
    Else
         QuantidadeUn = 1
    End If
     If IsNull(QuantidadeUn) Then QuantidadeUn = 1
    If QuantidadeUn = 0 Then QuantidadeUn = 1
   Achou = False
   Quantidade = 0
   If a = 0 Then
      '==> é o primeiro registro
      ReDim Preserve MtI(a)
      Quantidade = RsSaida!Qtde * IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)
      Quantidade = Quantidade / QuantidadeUn
      MtI(a).Codigo = RsSaida!codProd
      MtI(a).Quantidade = Quantidade * -1
      MtI(a).ValorUnitario = (RsSaida!VALUNIT / IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)) * QuantidadeUn
      a = a + 1
   Else
      Quantidade = RsSaida!Qtde * IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)
      Quantidade = Quantidade / QuantidadeUn
      '==> Ja existe, vamos ver se tem o codigo na matriz
      For b = 0 To UBound(MtI)
          If MtI(b).Codigo = RsSaida!codProd Then
             Achou = True
             Exit For
          End If
      Next
      
      If Achou Then
         MtI(b).Quantidade = MtI(b).Quantidade - Quantidade
         If MtI(b).ValorUnitario = 0 Then MtI(b).ValorUnitario = (RsSaida!VALUNIT / IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)) * QuantidadeUn
      Else
        ReDim Preserve MtI(a)
        Quantidade = RsSaida!Qtde * IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)
        Quantidade = Quantidade / QuantidadeUn
        MtI(a).Codigo = RsSaida!codProd
        MtI(a).Quantidade = Quantidade * -1
        MtI(a).ValorUnitario = (RsSaida!VALUNIT / IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)) * QuantidadeUn
        a = a + 1
      End If
   End If
   RsSaida.MoveNext
Loop

RsEntrada.Close
RsSaida.Close

'==> Acrescenta Dados inventario anterior

For a = 0 To UBound(MtI)
    Sintegra.Caption = "Verificando dados gravados do produto:" & MtI(a).Codigo
    DoEvents
    '==> Verifica se a quantidade e menor que um , se for zera
    If MtI(a).Quantidade < 1 Then MtI(a).Quantidade = 0
    
    '==> Verifica se tem dados no arquivo do sintegra
    If Sintegra.OptArquivo.Value = 1 Then
        StrSql = "Select * from BaseSintegra where CodigoProduto=" & MtI(a).Codigo
               
        Set RsI = AbreRecordset(StrSql, True)
        If Not RsI.EOF Then
           MtI(a).Quantidade = MtI(a).Quantidade + RsI!Quantidade
        End If
    End If
    If MtI(a).Quantidade < 1 Then MtI(a).Quantidade = 0
Next
'==> Gera arquivo Sintegra
StrSql = "Delete from inventariosintegra"
ExecutaSql StrSql
'Dim LcArq As Integer
'LcArq = FreeFile
'Open "C:\Inventario.txt" For Output As #LcArq
For a = 0 To UBound(MtI)
    Sintegra.Caption = "gerando Inventario Produto:" & MtI(a).Codigo
    DoEvents
    '==> Grava no arquivo do inventario do Sintegra
    StrSql = "Insert into inventariosintegra(data,CodigoProduto,Quantidade,ValorProduto)values("
    StrSql = StrSql & "#" & Format(DataFim, "mm/dd/yy") & "#,"
    StrSql = StrSql & MtI(a).Codigo & ","
    StrSql = StrSql & Replace(MtI(a).Quantidade, ",", ".") & ","
    StrSql = StrSql & Replace(MtI(a).ValorUnitario * MtI(a).Quantidade, ",", ".") & ")"
    ExecutaSql StrSql
    
    '==> Gera dados do Sintegra
    If Sintegra.IncluirZero.Value = 0 Then
       If MtI(a).Quantidade > 0 Then
           GeraRegistro74 CStr(DataFim), CStr(MtI(a).Codigo), MtI(a).Quantidade, MtI(a).ValorUnitario * MtI(a).Quantidade, "1"
       End If
    Else
           GeraRegistro74 CStr(DataFim), CStr(MtI(a).Codigo), MtI(a).Quantidade, MtI(a).ValorUnitario * MtI(a).Quantidade, "1"
    End If
Next
'Close #LcArq
End Sub
Public Sub BuscaRegistro53Lidis(DataInicio As Date, DataFinal As Date)
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim Cnpj As String
Dim Inscricao As String
Dim Estado As String
Dim DespesasAcessorias As Double

StrSql = "Select * from entradanf where DATA Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFinal, "mm/dd/yy") & "# and (BaseIcmsSubst>0 or IcmsSubst>0)"
Set Rs = AbreRecordset(StrSql, True)
'MsgBox DEscricaoErro

'==> Recupera os dados e manda para ser gravado na matriz
Do Until Rs.EOF
   Cnpj = AcertaCnpj(Rs!clicred, False)
   Inscricao = AcertaInscricao(Rs!clicred, False)
   Estado = BuscaEstado(Rs!clicred, False)
   DespesasAcessorias = Rs!Frete + Rs!Seguro + Rs!PIS_COFINS + Rs!NaoTributado + Rs!DespesasAcessorias
   GeraRegistro53 Cnpj, Inscricao, Format(Rs!Data, "dd/mm/yyyy"), Estado, "1", Rs!NF, Rs!CFOP, "T", Rs!BaseIcmsSubst, Rs!IcmsSubst, _
   DespesasAcessorias, "N", " "
  
   Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing

End Sub
Public Sub GeraRegistro74(Data As String, codigoproduto As String, Quantidade As Double _
                        , ValorProduto As Double, CodigoPosse As String)

Dim Registro74 As Dados74
Dim StrSql As String
Dim Unidade As String
Dim db As Database
Dim RsUnid As Recordset
Dim RsProd As ADODB.Recordset
Dim cst As String
Dim icms As String
Dim Registro75 As Dados75

StrSql = "Select * from PRodutos where codigo=" & codigoproduto

Set RsProd = AbreRecordset(StrSql, True)
If RsProd.EOF Then
  Set RsProd = Nothing
  Exit Sub
End If
Set db = OpenDatabase(GLBase)
 StrSql = "Select * from alid004 where cod='" & RsProd!UnidMedida & "'"
 Set RsUnid = db.OpenRecordset(StrSql)
 If Not RsUnid.EOF Then
    Unidade = RsUnid!Simbolo & ""
 Else
    Unidade = "UN"
 End If
 Cnpj = "00000000000000"
   '==> Inscricao
 Inscricao = String(14, " ")
 'If Quantidade = 0 Then Stop
 Registro74.Cnpj = Cnpj
 Registro74.Codigo_Posse = "1"
 Registro74.codigoproduto = Left(codigoproduto & "              ", 14)
 Registro74.Data = Format(Data, "yyyymmdd")
 Registro74.Quantidade = Quantidade
 Registro74.Estado = "MG"
 Registro74.Inscricao = Inscricao
 If CCur(ValorProduto) < 0.01 Then ValorProduto = 0.01
 Registro74.ValorProduto = CCur(ValorProduto)
' MsgBox CCur(ValorProduto)
 CriaRegistro74 Registro74
 If RsProd!cst = "060" Then
    icms = "0000"
 Else
    icms = "18"
 End If
 Registro75.Aliquota_Icms = icms
 Registro75.Aliquota_Ipi = 0
 Registro75.Base_Icms_subst = "000000000000"
 Registro75.Codigo = codigoproduto
 Registro75.Nome = RsProd!Nome
 Registro75.Reducao_Base = "00000"
 Registro75.Unidade = "UN"
        
 '==> Cria o Registro 75
 CriarDados75 Registro75
End Sub

Public Sub GeraRegistro53(Cnpj As String, Inscricao As String, Data As String, UF As String, Serie As String _
, NumeroNF As String, CFOP As String, Emitente As String, BaseCalculoIcms As Double, IcmsRetido As Double, DespesasAcess As Double _
, Situacao As String, CodigoAntecipacao As String)
Dim Registro53 As Dados53


 Cnpj = Replace(Cnpj, ".", "")
 Cnpj = Replace(Cnpj, "-", "")
 Cnpj = Replace(Cnpj, "/", "")
 Cnpj = Replace(Cnpj, "\", "")
 Cnpj = Replace(Cnpj, " ", "")
 Cnpj = Right("00000000000000" & Cnpj, 14)
   '==> Inscricao
 Inscricao = Replace(Inscricao, ".", "")
 Inscricao = Replace(Inscricao, "-", "")
 Inscricao = Replace(Inscricao, "/", "")
 Inscricao = Replace(Inscricao, "\", "")
 Inscricao = Replace(Inscricao, " ", "")
 Inscricao = Left(Inscricao & "              ", 14)
 
 Serie = Left("1" & "   ", 3)
 
'==> Grava os dados na Variavel
Registro53.Base_Cal_Subst = BaseCalculoIcms
Registro53.CFOP = CFOP
Registro53.Cnpj = Cnpj
Registro53.Codigo_Antecipacao = CodigoAntecipacao
Registro53.Data = Format(Data, "yyyymmdd")
Registro53.Despesas_Acessorias = DespesasAcess
Registro53.Emitente = Emitente
Registro53.Estado = UF
Registro53.Icms_Retido = IcmsRetido
Registro53.Inscricao = Inscricao
Registro53.Modelo = "01"
Registro53.Numero_Nf = Right("000000" & NumeroNF, 6)
Registro53.Serie = Serie
Registro53.Situacao = Situacao

CriarRegistro53 Registro53

End Sub

Public Sub AbreFormulario(Optional Esconde As Boolean = False)
Load Sintegra
If Esconde Then
    Sintegra.SSTab1.TabVisible(2) = False
    Sintegra.SSTab1.TabVisible(3) = False
End If
Sintegra.Show
End Sub
Public Sub BuscaSaidaCupomBhForte()
On Error GoTo ErrBuscaSaida
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim a As Integer
Dim Valor_Total As Double
Dim Valor_Icms  As Double
Dim strSeparaDevolucao As String
Dim LcTotal As Integer
Dim x As Integer

strSeparaDevolucao = SeparaDevolucao

StrSql = "SELECT ALID001.RAZAOSOC, prenota.* FROM prenota INNER JOIN ALID001 ON prenota.CLiente = ALID001.CODIGO where dtemis Between #" & Format(Sintegra.DataI.Text, "mm/dd/yy") & "# And #" & Format(Sintegra.DataF.Text, "mm/dd/yy") & "# and prenota.Sintegra=false and CondPag='Cupom Fiscal' "

If Len(strSeparaDevolucao) > 0 Then
    StrSql = StrSql & " and (" & SeparaDevolucao & ") "
End If
StrSql = StrSql & " order by dtemis"
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
   Rs.MoveLast
   LcTotal = Rs.RecordCount
   Rs.MoveFirst
End If
Sintegra.Cupom.Rows = 1
Do Until Rs.EOF
  DoEvents
  Sintegra.Caption = "Buscando a Cupom Fiscal  " & x & " de " & LcTotal & " Nº:" & Rs!NF & ""
  x = x + 1

  DoEvents
  a = Sintegra.Cupom.Rows
  Sintegra.Cupom.Rows = a + 1
  Sintegra.Cupom.TextMatrix(a, 0) = Rs!NF & ""
  Sintegra.Cupom.TextMatrix(a, 1) = Format(Rs!DtEmis, "dd/mm/yy") & ""
  Sintegra.Cupom.TextMatrix(a, 2) = Rs!razaosoc & ""
  Sintegra.Cupom.TextMatrix(a, 3) = Rs!TotalGeral & ""
  Sintegra.Cupom.TextMatrix(a, 4) = IIf(Not IsNull(Rs!icms), Rs!icms, 0) & ""
  Sintegra.Cupom.TextMatrix(a, 5) = ClUtil.AcertaNumero(CStr((IIf(Not IsNull(Rs!icms), Rs!icms, 0) / 100) * Rs!TotalGeral), 2) & ""
  Sintegra.Cupom.TextMatrix(a, 6) = Rs!Cliente & ""
  Sintegra.Cupom.TextMatrix(a, 7) = LcModelo & ""
  Sintegra.Cupom.TextMatrix(a, 8) = Rs!Serie & ""
  Sintegra.Cupom.TextMatrix(a, 9) = Rs!CFOP & "" '==> Codigo Fiscal
  Sintegra.Cupom.TextMatrix(a, 10) = IIf(Rs!Status = "Cancelado", "S", "N")
  Sintegra.Cupom.TextMatrix(a, 11) = Rs!CodNota & ""
  Valor_Total = Valor_Total + Rs!TotalGeral
  Valor_Icms = Valor_Icms + CDbl(Sintegra.Cupom.TextMatrix(a, 5))
  Rs.MoveNext
Loop
Set Rs = Nothing
Sintegra.TotalCupom.Text = ClUtil.AcertaNumero(CStr(Valor_Total), 2)
Sintegra.IcmsCupom.Text = ClUtil.AcertaNumero(CStr(Valor_Icms), 2)

Exit Sub
ErrBuscaSaida:
MsgBox err.Description & err.Number
'Resume 0
End Sub
Public Function ProcessaEntradaBhForte(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim RsDados         As ADODB.Recordset
Dim Registro50      As Dados50
Dim Registro54      As Dados54
Dim Registro75      As Dados75
Dim RegistroComp    As DadosComplementares
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim ValorDebitar    As Double
Dim ipi             As Double
Dim Frete           As Double
Dim Seguro          As Double
Dim Pis             As Double
Dim NaoTributado    As Double
Dim DespesasAcessorias  As Double
Dim Custo           As Double
Dim Serie           As String
Dim Situacao        As String
Dim Total54         As Integer
Dim a               As Integer
Dim Outras          As Double

StrSql = "SELECT EntradaNf.outras, EntradaNf.NUMLANCTO, EntradaNf.NF, EntradaNf.RECDESP, EntradaNf.CLICRED," & _
        "EntradaNf.VALOR, EntradaNf.DATA, EntradaNf.CONTR, EntradaNf.VP, EntradaNf.codigo," & _
        "EntradaNf.ValorProduto, EntradaNf.BaseIcms, EntradaNf.Icms, EntradaNf.Ipi," & _
        "EntradaNf.Geral, EntradaNf.Complementar, EntradaNf.cnpj, EntradaNf.codigofor," & _
        "EntradaNf.custos, EntradaNf.ContasLancadas, EntradaNf.parcelas, EntradaNf.primeiro," & _
        "EntradaNf.intermediario, EntradaNf.entrada, EntradaNf.Especie, EntradaNf.Serie," & _
        "EntradaNf.CodigoFiscal AS Cfop, EntradaNf.Tipo, EntradaNf.CodigoFormaDePagamento," & _
        "EntradaNf.QuantidadeDeItem, EntradaNf.vencimento0, EntradaNf.vencimento1," & _
        "EntradaNf.vencimento2, EntradaNf.vencimento3, EntradaNf.vencimento4," & _
        "EntradaNf.vencimento5, EntradaNf.vencimento6, EntradaNf.vencimento7, EntradaNf.vencimento8, EntradaNf.vencimento9, EntradaNf.valor1," & _
        "EntradaNf.valor2, EntradaNf.valor3, EntradaNf.valor4, EntradaNf.valor5, EntradaNf.valor6, EntradaNf.valor7," & _
        "EntradaNf.valor8 , EntradaNf.valor9, EntradaNf.valo0, EntradaNf.Usuario, EntradaNf.Modelo , EntradaNf.Sintegra," & _
        "EntradaNf.Status, EntradaNf.natureza, EntradaNf.Frete, EntradaNf.Seguro, EntradaNf.PIS_COFINS, EntradaNf.Complemento," & _
        "EntradaNf.NaoTributado, EntradaNf.DespesasAcessorias, EntradaNf.BaseCalculoIcmsSubstituicao, EntradaNf.ValorIcmsSubstituicao, EntradaNf.SubSerie, EntradaNf.TipoFrete, EntradaNf.TotalDesconto, EntradaNf.E_Cliente" & _
        " FROM EntradaNf" & _
        " where entrada Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "# and codigofiscal<>'1353' and codigofiscal<>'2353' order by entrada"

Set RsEntrada = AbreRecordset(StrSql, True)
Set RsDados = AbreRecordset("Select * from ItensEntradaNf order by CodigoNota,item")

Do Until RsEntrada.EOF
  ' If RsEntrada!nf = "28" Then Stop
   Sintegra.Caption = "processando Entrada Registro 50 NF:" & RsEntrada!NF
   DoEvents
   Frete = IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
    Seguro = IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
    Pis = IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
    Complemento = IIf(Not IsNull(RsEntrada!Complemento), RsEntrada!Complemento, 0)
    NaoTributado = IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
    DespesasAcessorias = IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
    Custo = IIf(Not IsNull(RsEntrada!Custos), RsEntrada!Custos, 0)
    ipi = IIf(Not IsNull(RsEntrada!ipi), RsEntrada!ipi, 0)
    Outras = IIf(Not IsNull(RsEntrada!Outras), RsEntrada!Outras, 0)
    ValorDebitar = ipi + Frete + Seguro + Pis + NaoTributado + DespesasAcessorias + Custo + Outras

   '==>Filtra os registros dos itens
   RsDados.Filter = "CodigoNota=" & RsEntrada!Codigo
   '==> Calcula o total de registros
   If Not RsDados.EOF Then
      RsDados.MoveLast
      Total54 = RsDados.RecordCount
      RsDados.MoveFirst
   Else
      Total54 = 0
   End If
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   Inscricao = AcertaInscricao(RsEntrada!codigofor, RsEntrada!E_Cliente)
   Estado = BuscaEstado(RsEntrada!codigofor, RsEntrada!E_Cliente)
   Cnpj = AcertaCnpj(RsEntrada!codigofor, RsEntrada!E_Cliente)
   
   '==> Acerta valores dos lancamentos
   Valor_Outras = IIf(Not IsNull(RsEntrada!ValorIcmsSubstituicao), RsEntrada!ValorIcmsSubstituicao, 0)
   'Valor_Outras = Valor_Outras + IIf(Not IsNull(RsEntrada!Outras), RsEntrada!Outras, 0)
   Serie = Left("1" & "   ", 3)
   Situacao = IIf(RsEntrada!Status = "Cancelado", "S", "N")
   Desconto = IIf(Not IsNull(RsEntrada!TotalDesconto), RsEntrada!TotalDesconto, 0)
   If Desconto > 0 Then Desconto = IIf(Total54 > 0, Desconto / Total54, Desconto)
   '****** Processa o registro 50
   ReDim MtIcms(0)
   If IcmsDiferente(RsDados) Then
      Valor_t = IIf(Total54 > 0, RsEntrada!valor / Total54, RsEntrada!valor)
      For a = 0 To UBound(MtIcms)
            Registro50.Aliquota = MtIcms(a).icms
            If RsDados!icms > 0 Then
              Registro50.Base_Calculo_Icms = MtIcms(0).valor
            Else
              Registro50.Base_Calculo_Icms = 0
            End If
            Registro50.CFOP = RsEntrada!CFOP
            Registro50.Cnpj = Cnpj
            Registro50.Data = Format(RsEntrada!Entrada, "yyyymmdd")
            Registro50.Emitente = "T"
            Registro50.Estado = Estado
            Registro50.Inscricao = Inscricao
            Registro50.Isenta_Nao_Tributada = "0000000000000"
            Registro50.Modelo = "01"
            Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
            Registro50.Outras = IIf(MtIcms(a).icms = 0, Valor_Outras, 0)
            Registro50.Serie = Serie
            Registro50.Situacao = Situacao
            Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
            Registro50.Valor_Total = Valor_t
            
            CriarDados50 Registro50
      Next
   Else
     Registro50.Aliquota = IIf(Not RsDados.EOF, RsDados!icms, 0)
     If Not RsDados.EOF Then
        If RsDados!icms > 0 Then
          Registro50.Base_Calculo_Icms = MtIcms(0).valor
        Else
          Registro50.Base_Calculo_Icms = 0
        End If
     Else
       Registro50.Base_Calculo_Icms = 0
     End If
     Registro50.CFOP = RsEntrada!CFOP
     Registro50.Cnpj = Cnpj
     Registro50.Data = Format(RsEntrada!Entrada, "yyyymmdd")
     Registro50.Emitente = "T"
     Registro50.Estado = Estado
     Registro50.Inscricao = Inscricao
     Registro50.Isenta_Nao_Tributada = "0000000000000"
     Registro50.Modelo = "01"
     Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
     Registro50.Outras = Valor_Outras
     Registro50.Serie = Serie
     Registro50.Situacao = Situacao
     Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
     Registro50.Valor_Total = RsEntrada!valor
     CriarDados50 Registro50
   End If
        
   '===> Processa os registros 54
   Do Until RsDados.EOF
        '==> Calcula valores referentes ao item
        ValorIpi = (IIf(IsNumeric(RsDados!ipi), RsDados!ipi, 0) / 100) * RsDados!Valor_Total_Compra

        Registro54.Aliquota_Icms = RsDados!icms
        If RsDados!icms > 0 Then
           Registro54.Base_Calculo_Icms = RsDados!Valor_Total_Compra
        Else
           Registro54.Base_Calculo_Icms = 0
        End If
        Registro54.Base_Calculo_subs_Trib = "000000000000"
        Registro54.CFOP = RsEntrada!CFOP
        Registro54.Cnpj = Cnpj
        Registro54.Codigo_Produto = RsDados!codigopesquisa
        Registro54.cst = RsDados!cst
        Registro54.Modelo = "01"
        Registro54.Numero_Item = RsDados!Item
        Registro54.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
        Registro54.Quantidade = RsDados!Qtde
        Registro54.Serie = Serie
        Registro54.Valor_Desconto = Desconto
        Registro54.Valor_Ipi = ValorIpi
        Registro54.Valor_Produto = RsDados!Valor_Total_Compra
        
        '==> cria o registro 54
        CriarDados54 Registro54
        
        '==>Valos criar o registro75
        
        Registro75.Aliquota_Icms = RsDados!icms
        Registro75.Aliquota_Ipi = RsDados!ipi
        Registro75.Base_Icms_subst = "000000000000"
        Registro75.Codigo = RsDados!codigopesquisa
        Registro75.Nome = RsDados!Descricao
        Registro75.Reducao_Base = "00000"
        Registro75.Unidade = RsDados!Unimed
        
        '==> Cria o Registro 75
        CriarDados75 Registro75
        
        RsDados.MoveNext
   Loop
   '==> Gera dados para registro complementares
   RegistroComp.CFOP = RsEntrada!CFOP
   RegistroComp.Cnpj = Cnpj
   RegistroComp.Modelo = "01"
   RegistroComp.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
   RegistroComp.Serie = Serie
   RegistroComp.Valor_Compl = IIf(Not IsNull(RsEntrada!Complemento), RsEntrada!Complemento, 0)
   RegistroComp.Valor_Despesas = IIf(Not IsNull(RsEntrada!Custos), RsEntrada!Custos, 0)
   RegistroComp.Valor_Frete = IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
   RegistroComp.Valor_pis = IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
   RegistroComp.Valor_Seguro = IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
   RegistroComp.Valor_Servicos = IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
   RegistroComp.Total_Itens = Total54
   Processar_Dados_complementares RegistroComp
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

End Function
Public Function ProcessaEntradaLidis(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim RsDados         As ADODB.Recordset
Dim Registro50      As Dados50
Dim Registro54      As Dados54
Dim Registro75      As Dados75
Dim RegistroComp    As DadosComplementares
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim ValorDebitar    As Double
Dim ipi             As Double
Dim Frete           As Double
Dim Seguro          As Double
Dim Pis             As Double
Dim NaoTributado    As Double
Dim DespesasAcessorias  As Double
Dim Custo           As Double
Dim Serie           As String
Dim Situacao        As String
Dim Total54         As Integer
Dim db              As Database
Dim RsUnid          As Recordset
Dim Unidade         As String
Dim CFOP            As String
Dim a               As Long
Dim NItem           As Integer

StrSql = "SELECT * FROM EntradaNf" & _
        " where data Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "# and cfop<>'1353' and cfop<>'2353' order by data"

Set RsEntrada = AbreRecordset(StrSql, True)
Set RsDados = AbreRecordset("Select  * from ItensEntradaNf order by CodigoNota,item")
'MsgBox RsDados("Valor_Total_Compra")
Set db = OpenDatabase(GLBase)

Do Until RsEntrada.EOF
   'If RsEntrada!Nf = "331756" Or RsEntrada!Nf = "331756" Then Stop
   
   Sintegra.Caption = "processando Entrada Registro 50 NF:" & RsEntrada!NF
   DoEvents
    Frete = IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
    Seguro = IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
    Pis = IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
    Complemento = IIf(Not IsNull(RsEntrada!Complementar), RsEntrada!Complementar, 0)
    NaoTributado = IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
    DespesasAcessorias = IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
    Custo = IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
    ipi = IIf(Not IsNull(RsEntrada!ipi), RsEntrada!ipi, 0)
    ValorDebitar = ipi + Frete + Seguro + Pis + NaoTributado + DespesasAcessorias + Custo

   '==>Filtra os registros dos itens
   RsDados.Filter = "CodigoNota=" & RsEntrada!Codigo
   '==> Calcula o total de registros
   If Not RsDados.EOF Then
      RsDados.MoveLast
      Total54 = RsDados.RecordCount
      RsDados.MoveFirst
   Else
      Total54 = 0
   End If
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   Inscricao = AcertaInscricao(RsEntrada!clicred, False)
   Estado = BuscaEstado(RsEntrada!clicred, False)
   Cnpj = AcertaCnpj(RsEntrada!clicred, False)
   
   '==> Acerta valores dos lancamentos
   Valor_Outras = IIf(Not IsNull(RsEntrada!IcmsSubst), RsEntrada!IcmsSubst, 0)
   Serie = Left("1" & "   ", 3)
   Situacao = "N"
   Desconto = IIf(Not IsNull(RsEntrada!Desconto), RsEntrada!Desconto, 0)
   If Desconto > 0 Then Desconto = IIf(Total54 > 0, Desconto / Total54, Desconto)
   '****** Processa o registro 50
   ReDim MtIcms(0)
'   If RsEntrada!NF = "000476" Then Stop
   If IcmsDiferente(RsDados) Then
      Valor_t = IIf(Total54 > 0, RsEntrada!valor / Total54, RsEntrada!valor)
      For a = 0 To UBound(MtIcms)
            Registro50.Aliquota = MtIcms(a).icms
            If RsDados!icms > 0 Then
              Registro50.Base_Calculo_Icms = MtIcms(0).valor
            Else
              Registro50.Base_Calculo_Icms = 0
            End If
            
            'CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro50.Aliquota)
            CFOP = RsEntrada!CFOP
            Registro50.CFOP = CFOP
            Registro50.Cnpj = Cnpj
            Registro50.Data = Format(RsEntrada!Data, "yyyymmdd")
            Registro50.Emitente = "T"
            Registro50.Estado = Estado
            Registro50.Inscricao = Inscricao
            Registro50.Isenta_Nao_Tributada = "0000000000000"
            Registro50.Modelo = "01"
            Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
            Registro50.Outras = IIf(MtIcms(a).icms = 0, Valor_Outras, 0)
            Registro50.Serie = Serie
            Registro50.Situacao = Situacao
            Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
            Registro50.Valor_Total = Valor_t
            
            CriarDados50 Registro50
      Next
   Else
     Registro50.Aliquota = IIf(Not RsDados.EOF, RsDados!icms, 0)
     If Not RsDados.EOF Then
        If RsDados!icms > 0 Then
          Registro50.Base_Calculo_Icms = MtIcms(0).valor
        Else
          Registro50.Base_Calculo_Icms = 0
        End If
     Else
       Registro50.Base_Calculo_Icms = 0
     End If
     'CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro50.Aliquota)
     CFOP = RsEntrada!CFOP
     Registro50.CFOP = CFOP
     Registro50.Cnpj = Cnpj
     Registro50.Data = Format(RsEntrada!Data, "yyyymmdd")
     Registro50.Emitente = "T"
     Registro50.Estado = Estado
     Registro50.Inscricao = Inscricao
     Registro50.Isenta_Nao_Tributada = "0000000000000"
     Registro50.Modelo = "01"
     Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
     Registro50.Outras = Valor_Outras
     Registro50.Serie = Serie
     Registro50.Situacao = Situacao
     Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
     Registro50.Valor_Total = RsEntrada!valor
     CriarDados50 Registro50
   End If
        
   '===> Processa os registros 54
   NItem = 0
   Do Until RsDados.EOF
        NItem = NItem + 1
        '==> Calcula valores referentes ao item
        ValorIpi = (IIf(IsNumeric(RsDados!ipi), RsDados!ipi, 0) / 100) * RsDados!ValorTotal

        Registro54.Aliquota_Icms = RsDados!icms
        If RsDados!icms > 0 Then
           Registro54.Base_Calculo_Icms = RsDados!ValorTotal
           cst = "000"
        Else
           Registro54.Base_Calculo_Icms = 0
           cst = "060"
        End If
        'CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro54.Aliquota_Icms)
        CFOP = RsEntrada!CFOP
        Registro54.Base_Calculo_subs_Trib = "000000000000"
        Registro54.CFOP = CFOP
        Registro54.Cnpj = Cnpj
        Registro54.Codigo_Produto = RsDados!Item
        Registro54.cst = cst
        Registro54.Modelo = "01"
        Registro54.Numero_Item = NItem
        Registro54.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
        Registro54.Quantidade = RsDados!Qtde
        Registro54.Serie = Serie
        Registro54.Valor_Desconto = Desconto
        Registro54.Valor_Ipi = ValorIpi
        Registro54.Valor_Produto = RsDados!ValorTotal
        
        '==> cria o registro 54
        CriarDados54 Registro54
        
        '==>Valos criar o registro75
        StrSql = "Select * from alid004 where cod='" & RsDados!Unimed & "'"
        Set RsUnid = db.OpenRecordset(StrSql)
        If Not RsUnid.EOF Then
           Unidade = RsUnid!Simbolo & ""
        Else
           Unidade = "UN"
        End If
        Registro75.Aliquota_Icms = RsDados!icms
        Registro75.Aliquota_Ipi = RsDados!ipi
        Registro75.Base_Icms_subst = "000000000000"
        Registro75.Codigo = RsDados!Item
        Registro75.Nome = RsDados!Descricao
        Registro75.Reducao_Base = "00000"
        Registro75.Unidade = Unidade
        
        '==> Cria o Registro 75
        CriarDados75 Registro75
        
        RsDados.MoveNext
   Loop
   '==> Gera dados para registro complementares
   'CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro54.Aliquota_Icms)
   CFOP = RsEntrada!CFOP
   RegistroComp.CFOP = CFOP
   RegistroComp.Cnpj = Cnpj
   RegistroComp.Modelo = "01"
   RegistroComp.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
   RegistroComp.Serie = Serie
   RegistroComp.Valor_Compl = IIf(Not IsNull(RsEntrada!Complementar), RsEntrada!Complementar, 0)
   RegistroComp.Valor_Despesas = IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
   RegistroComp.Valor_Frete = IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
   RegistroComp.Valor_pis = IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
   RegistroComp.Valor_Seguro = IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
   RegistroComp.Valor_Servicos = IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
   RegistroComp.Total_Itens = Total54
   Processar_Dados_complementares RegistroComp
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

End Function
Public Function ProcessaDevolucaoClienteBhForte(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim RsDados         As ADODB.Recordset
Dim Registro50      As Dados50
Dim Registro54      As Dados54
Dim Registro75      As Dados75
Dim RegistroComp    As DadosComplementares
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim ValorDebitar    As Double
Dim ipi             As Double
Dim Frete           As Double
Dim Seguro          As Double
Dim Pis             As Double
Dim NaoTributado    As Double
Dim DespesasAcessorias  As Double
Dim Custo           As Double
Dim Serie           As String
Dim Situacao        As String
Dim Total54         As Integer
Dim a               As Integer


StrSql = "SELECT * from devolucaocliente" & _
        " where dtemis Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "#  order by dtemis"

Set RsEntrada = AbreRecordset(StrSql, True)
Set RsDados = AbreRecordset("Select *, dadosDevolucao.Total AS Valor_Total_Compra from dadosdevolucao order by doc,item")

Do Until RsEntrada.EOF
   Sintegra.Caption = "processando Devolução Registro 50 NF:" & RsEntrada!NF
   DoEvents
   
   '==>Filtra os registros dos itens
   RsDados.Filter = "doc='" & RsEntrada!NF & "'"
   '==> Calcula o total de registros
   If Not RsDados.EOF Then
      RsDados.MoveLast
      Total54 = RsDados.RecordCount
      RsDados.MoveFirst
   Else
      Total54 = 0
   End If
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   Inscricao = AcertaInscricao(RsEntrada!Cliente, True)
   Estado = BuscaEstado(RsEntrada!Cliente, True)
   Cnpj = AcertaCnpj(RsEntrada!Cliente, True)
   
   '==> Acerta valores dos lancamentos
   Valor_Outras = 0 ' no caso da bh forte, usa-se zero
   Serie = Left("1" & "   ", 3)
   Situacao = IIf(RsEntrada!Status = "Cancelado", "S", "N")
   Desconto = 0
   If Desconto > 0 Then Desconto = IIf(Total54 > 0, Desconto / Total54, Desconto)
   '****** Processa o registro 50
   ReDim MtIcms(0)
   If IcmsDiferente(RsDados) Then
      Valor_t = IIf(Total54 > 0, RsEntrada!total / Total54, RsEntrada!total)
      For a = 0 To UBound(MtIcms)
            Registro50.Aliquota = MtIcms(a).icms
            If RsDados!icms > 0 Then
              Registro50.Base_Calculo_Icms = MtIcms(0).valor
            Else
              Registro50.Base_Calculo_Icms = 0
            End If
            Registro50.CFOP = RsEntrada!CFOP
            Registro50.Cnpj = Cnpj
            Registro50.Data = Format(RsEntrada!DtEmis, "yyyymmdd")
            Registro50.Emitente = "T"
            Registro50.Estado = Estado
            Registro50.Inscricao = Inscricao
            Registro50.Isenta_Nao_Tributada = "0000000000000"
            Registro50.Modelo = "01"
            Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
            Registro50.Outras = IIf(MtIcms(a).icms = 0, Valor_Outras, 0)
            Registro50.Serie = Serie
            Registro50.Situacao = Situacao
            Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
            Registro50.Valor_Total = Valor_t
            
            CriarDados50 Registro50
      Next
   Else
     Registro50.Aliquota = IIf(Not RsDados.EOF, RsDados!icms, 0)
     If Not RsDados.EOF Then
        If RsDados!icms > 0 Then
          Registro50.Base_Calculo_Icms = MtIcms(0).valor
        Else
          Registro50.Base_Calculo_Icms = 0
        End If
     Else
       Registro50.Base_Calculo_Icms = 0
     End If
     Registro50.CFOP = RsEntrada!CFOP
     Registro50.Cnpj = Cnpj
     Registro50.Data = Format(RsEntrada!DtEmis, "yyyymmdd")
     Registro50.Emitente = "T"
     Registro50.Estado = Estado
     Registro50.Inscricao = Inscricao
     Registro50.Isenta_Nao_Tributada = "0000000000000"
     Registro50.Modelo = "01"
     Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
     Registro50.Outras = Valor_Outras
     Registro50.Serie = Serie
     Registro50.Situacao = Situacao
     Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
     Registro50.Valor_Total = RsEntrada!total
     CriarDados50 Registro50
   End If
        
   '===> Processa os registros 54
   Do Until RsDados.EOF
        '==> Calcula valores referentes ao item
        ValorIpi = (IIf(IsNumeric(RsDados!ipi), RsDados!ipi, 0) / 100) * RsDados!Valor_Total_Compra

        Registro54.Aliquota_Icms = RsDados!icms
        If RsDados!icms > 0 Then
           Registro54.Base_Calculo_Icms = RsDados!Valor_Total_Compra
        Else
           Registro54.Base_Calculo_Icms = 0
        End If
        Registro54.Base_Calculo_subs_Trib = "000000000000"
        Registro54.CFOP = RsEntrada!CFOP
        Registro54.Cnpj = Cnpj
        Registro54.Codigo_Produto = RsDados!Codigo
        Registro54.cst = RsDados!cst
        Registro54.Modelo = "01"
        Registro54.Numero_Item = RsDados!Item
        Registro54.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
        Registro54.Quantidade = RsDados!quant
        Registro54.Serie = Serie
        Registro54.Valor_Desconto = Desconto
        Registro54.Valor_Ipi = ValorIpi
        Registro54.Valor_Produto = RsDados!Valor_Total_Compra
        
        '==> cria o registro 54
        CriarDados54 Registro54
        
        '==>Valos criar o registro75
        
        Registro75.Aliquota_Icms = RsDados!icms
        Registro75.Aliquota_Ipi = IIf(Not IsNull(RsDados!ipi), RsDados!ipi, 0)
        Registro75.Base_Icms_subst = "000000000000"
        Registro75.Codigo = RsDados!Codigo
        Registro75.Nome = RsDados!Descricao
        Registro75.Reducao_Base = "00000"
        Registro75.Unidade = RsDados!unid
        
        '==> Cria o Registro 75
        CriarDados75 Registro75
        
        RsDados.MoveNext
   Loop
   '==> Gera dados para registro complementares
   'RegistroComp.CFOP = RsEntrada!CFOP
   'RegistroComp.CNPJ = CNPJ
   'RegistroComp.Modelo = "01"
   'RegistroComp.Numero_Nf = Right("000000" & RsEntrada!nf, 6)
   'RegistroComp.Serie = Serie
   'RegistroComp.Valor_Compl = IIf(Not IsNull(RsEntrada!Complemento), RsEntrada!Complemento, 0)
   'RegistroComp.Valor_Despesas = IIf(Not IsNull(RsEntrada!Custos), RsEntrada!Custos, 0)
   'RegistroComp.Valor_Frete = IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
   'RegistroComp.Valor_pis = IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
   'RegistroComp.Valor_Seguro = IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
   'RegistroComp.Valor_Servicos = IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
   'RegistroComp.Total_Itens = Total54
   'Processar_Dados_complementares RegistroComp
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

End Function
Public Function ProcessaSaidaLidis(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim RsDados         As ADODB.Recordset
Dim Registro50      As Dados50
Dim Registro54      As Dados54
Dim Registro75      As Dados75
Dim RegistroComp    As DadosComplementares
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim ValorDebitar    As Double
Dim ipi             As Double
Dim Frete           As Double
Dim Seguro          As Double
Dim Pis             As Double
Dim NaoTributado    As Double
Dim ValorTotalItem  As Double
Dim DespesasAcessorias  As Double
Dim Custo           As Double
Dim Serie           As String
Dim Situacao        As String
Dim Total54         As Integer
Dim NItem           As Integer
Dim TotalPro        As Long
Dim TProce          As Long
Dim CFOP            As String
Dim a               As Long
StrSql = "SELECT * from alid050" & _
        " where dtemis Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "#  order by dtemis"

Debug.Print StrSql
Sintegra.Caption = "Abrindo tabela Notas de Saida..."
DoEvents
Set RsEntrada = AbreRecordset(StrSql, True)

StrSql = "SELECT * from alid052 order by numnf,item"

Set RsDados = AbreRecordset(StrSql, True)
If Not RsEntrada.EOF Then
   RsEntrada.MoveLast
   TotalPro = RsEntrada.RecordCount
   RsEntrada.MoveFirst
End If
Do Until RsEntrada.EOF
   TProce = TProce + 1
   Sintegra.Caption = "processando Saida Tipo 50 NF:" & RsEntrada!numnf & " - Registros " & TProce & " de " & TotalPro
   DoEvents
   Frete = 0
   Seguro = 0
   Pis = 0
   Complemento = 0
   NaoTributado = 0
   DespesasAcessorias = 0 'IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
   Custo = 0
   ipi = 0
   ValorDebitar = ipi + Frete + Seguro + Pis + NaoTributado + DespesasAcessorias + Custo

   '==>Filtra os registros dos itens
   RsDados.Filter = "numnf='" & RsEntrada!numnf & "'"
   '==> Calcula o total de registros
   If Not RsDados.EOF Then
      RsDados.MoveLast
      Total54 = RsDados.RecordCount
      RsDados.MoveFirst
   Else
      Total54 = 0
   End If
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   
   Inscricao = AcertaInscricao(RsEntrada!Cliente, True)
   Estado = BuscaEstado(RsEntrada!Cliente, True)
   Cnpj = AcertaCnpj(RsEntrada!Cliente, True)
   
   '==> Acerta valores dos lancamentos
   Valor_Outras = 0 'IIf(Not IsNull(RsEntrada!ValorIcmsSubstituicao), RsEntrada!ValorIcmsSubstituicao, 0)
   Serie = Left("1" & "   ", 3)
   Situacao = IIf(UCase(RsEntrada!Status) = UCase("Cancelada"), "S", "N")
   Desconto = IIf(Not IsNull(RsEntrada!Desconto), RsEntrada!Desconto, 0)
   If Desconto > 0 Then Desconto = IIf(Total54 > 0, Desconto / Total54, Desconto)
   '****** Processa o registro 50
   ReDim MtIcms(0)
   MtIcms(0).icms = 0
   MtIcms(0).valor = 0
   If IcmsDiferenteSaidaLidis(RsDados) Then
      Valor_t = IIf(Total54 > 0, RsEntrada!ValorNota / Total54, RsEntrada!ValorNota)
      For a = 0 To UBound(MtIcms)
            Registro50.Aliquota = MtIcms(a).icms
            If MtIcms(a).icms > 0 Then
              Registro50.Base_Calculo_Icms = MtIcms(a).valor
            Else
              Registro50.Base_Calculo_Icms = 0
            End If
            If Not IsNull(RsEntrada!CFOP) Then
               CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro50.Aliquota)
            Else
               CFOP = RetornaCfop("5405", Registro50.Aliquota)
            End If
            Registro50.CFOP = CFOP
            Registro50.Cnpj = Cnpj
            Registro50.Data = Format(RsEntrada!DtEmis, "yyyymmdd")
            Registro50.Emitente = "P"
            Registro50.Estado = Estado
            Registro50.Inscricao = Inscricao
            Registro50.Isenta_Nao_Tributada = "0000000000000"
            Registro50.Modelo = "01"
            Registro50.Numero_Nf = Right("000000" & RsEntrada!numnf, 6)
            Registro50.Outras = IIf(MtIcms(a).icms = 0, Valor_Outras, 0)
            Registro50.Serie = Serie
            Registro50.Situacao = Situacao
            Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
            Registro50.Valor_Total = MtIcms(a).valor
            
            CriarDados50 Registro50
      Next
   Else
     Registro50.Aliquota = IIf(Not RsDados.EOF, RsDados!icms, 0)
     If Not RsDados.EOF Then
        If RsDados!icms > 0 Then
          Registro50.Base_Calculo_Icms = MtIcms(0).valor
        Else
          Registro50.Base_Calculo_Icms = 0
        End If
     Else
       Registro50.Base_Calculo_Icms = 0
     End If
     If Not IsNull(RsEntrada!CFOP) Then
        CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro50.Aliquota)
     Else
       CFOP = RetornaCfop("5405", Registro50.Aliquota)
     End If
     Registro50.CFOP = CFOP
     Registro50.Cnpj = Cnpj
     Registro50.Data = Format(RsEntrada!DtEmis, "yyyymmdd")
     Registro50.Emitente = "P"
     Registro50.Estado = Estado
     Registro50.Inscricao = Inscricao
     Registro50.Isenta_Nao_Tributada = "0000000000000"
     Registro50.Modelo = "01"
     Registro50.Numero_Nf = Right("000000" & RsEntrada!numnf, 6)
     Registro50.Outras = Valor_Outras
     Registro50.Serie = Serie
     Registro50.Situacao = Situacao
     Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
     Registro50.Valor_Total = RsEntrada!ValorNota
     CriarDados50 Registro50
   End If
        
   '===> Processa os registros 54
   NItem = 0
   Do Until RsDados.EOF
        '==> Calcula valores referentes ao item
        NItem = NItem + 1
        ValorIpi = 0 '(IIf(IsNumeric(RsDados!ipi), RsDados!ipi, 0) / 100) * RsDados!Valor_Total_Compra
        ValorTotalItem = RsDados!Qtde * RsDados!VALUNIT
        Registro54.Aliquota_Icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
        If IIf(Not IsNull(RsDados!icms), RsDados!icms, 0) > 0 Then
           Registro54.Base_Calculo_Icms = ValorTotalItem
        Else
           Registro54.Base_Calculo_Icms = 0
        End If
        If Not IsNull(RsEntrada!CFOP) Then
           CFOP = RetornaCfop(Replace(RsEntrada!CFOP, ".", ""), Registro54.Aliquota_Icms)
        Else
           CFOP = RetornaCfop("5405", Registro54.Aliquota_Icms)
        End If
        Registro54.Base_Calculo_subs_Trib = "000000000000"
        Registro54.CFOP = CFOP
        Registro54.Cnpj = Cnpj
        Registro54.Codigo_Produto = RsDados!codProd
        Registro54.cst = IIf(Not IsNull(RsDados!cst), RsDados!cst, "000")
        Registro54.Modelo = "01"
        Registro54.Numero_Item = NItem
        Registro54.Numero_Nf = Right("000000" & RsEntrada!numnf, 6)
        Registro54.Quantidade = RsDados!Qtde
        Registro54.Serie = Serie
        Registro54.Valor_Desconto = Desconto
        Registro54.Valor_Ipi = ValorIpi
        Registro54.Valor_Produto = ValorTotalItem
        
        '==> cria o registro 54
        CriarDados54 Registro54
        
        '==>Valos criar o registro75
        
        Registro75.Aliquota_Icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
        Registro75.Aliquota_Ipi = 0
        Registro75.Base_Icms_subst = "000000000000"
        Registro75.Codigo = RsDados!codProd
        Registro75.Nome = RsDados!Descricao
        Registro75.Reducao_Base = "00000"
        Registro75.Unidade = RsDados!Unimed
        
        '==> Cria o Registro 75
        CriarDados75 Registro75
        
        RsDados.MoveNext
   Loop
   '==> Gera dados para registro complementares
   RegistroComp.CFOP = CFOP
   RegistroComp.Cnpj = Cnpj
   RegistroComp.Modelo = "01"
   RegistroComp.Numero_Nf = Right("000000" & RsEntrada!numnf, 6)
   RegistroComp.Serie = Serie
   RegistroComp.Valor_Compl = 0 'IIf(Not IsNull(RsEntrada!Complemento), RsEntrada!Complemento, 0)
   RegistroComp.Valor_Despesas = 0 'IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
   RegistroComp.Valor_Frete = 0 'IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
   RegistroComp.Valor_pis = 0 'IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
   RegistroComp.Valor_Seguro = 0 '= IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
   RegistroComp.Valor_Servicos = 0 '= IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
   RegistroComp.Total_Itens = Total54
   Processar_Dados_complementares RegistroComp
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

End Function
Function ProcessaSaidaBhForte(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim RsDados         As ADODB.Recordset
Dim Registro50      As Dados50
Dim Registro54      As Dados54
Dim Registro75      As Dados75
Dim RegistroComp    As DadosComplementares
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim ValorDebitar    As Double
Dim ipi             As Double
Dim Frete           As Double
Dim Seguro          As Double
Dim Pis             As Double
Dim NaoTributado    As Double
Dim DespesasAcessorias  As Double
Dim Custo           As Double
Dim Serie           As String
Dim Situacao        As String
Dim Total54         As Integer
Dim NItem           As Integer
Dim a               As Integer
Dim TotalPro        As Long
Dim TProce          As Long
Dim DescontoBrutoIcms As Double
Dim DescontoBruto   As Double
Dim AcrescimoBruto  As Double
Dim AcrescimoAc  As Double
StrSql = "SELECT prenota.doc, prenota.DTEMIS AS entrada, prenota.NATUREZA, prenota.TRANSP," & _
        "prenota.TIPOTRANS, prenota.acrecimo, prenota.PLACATRANS, prenota.UFTRANS, prenota.CGCCPFTRAN," & _
        "prenota.ENDTRANS, prenota.MUNICTRANS, prenota.UFMUNIC, prenota.INSCEST, prenota.OBS1," & _
        "prenota.OBS2, prenota.OBS3, prenota.OBS4, prenota.OBS5, prenota.OBS6, prenota.OBS7," & _
        "prenota.CONTR, prenota.codigo, prenota.Desconto, prenota.TotalDesconto, prenota.TotalGeral AS valor," & _
        "prenota.TotalProduto, prenota.Fabricante, prenota.CondPag, prenota.FoneTransp, prenota.CLiente," & _
        "prenota.OrcVenda, prenota.Vendedor, prenota.cidade, prenota.cep, prenota.formapag, prenota.dias," & _
        "prenota.vencimento1, prenota.codnota, prenota.vencimento2, prenota.vencimento3, prenota.vencimento4, prenota.vencimento5," & _
        "prenota.acrecimo AS DespesasAcessorias, prenota.DetalhaDesconto, prenota.DetahaAcrecimo, prenota.Status," & _
        "prenota.comissao, prenota.ipi, prenota.Industria, prenota.Representada, prenota.PrevComissao, prenota.prenota," & _
        "prenota.nf, prenota.naobaixaestoque, prenota.custo, prenota.lucro, prenota.Faturado, prenota.OBS, prenota.APrazo," & _
        "prenota.GeraFinanceiro, prenota.icms, prenota.informacoes, prenota.ValorIcms, prenota.QutVol, prenota.Especie, prenota.Marca," & _
        "prenota.Numero, prenota.PesoB, prenota.Pesol, prenota.Sintegra, prenota.CFOP, prenota.Serie, prenota.E_Cliente " & _
        "FROM prenota" & _
        " where dtemis Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "# and CondPag<>'Cupom Fiscal' order by dtemis"

Debug.Print StrSql
Sintegra.Caption = "Abrindo tabela Notas de Saida..."
DoEvents
Set RsEntrada = AbreRecordset(StrSql, True)

StrSql = "SELECT dadosPrenota.Codbase, dadosPrenota.CodNota, dadosPrenota.Quant AS QTDE, dadosPrenota.Descricao, dadosPrenota.Volt," & _
       "dadosPrenota.Unit, dadosPrenota.Ipi, dadosPrenota.Total AS Valor_Total_Compra, dadosPrenota.doc," & _
       "dadosPrenota.unid AS UNIMED, dadosPrenota.com, dadosPrenota.item, dadosPrenota.valorUnitarioReal," & _
       "dadosPrenota.controle, dadosPrenota.comissao, dadosPrenota.acomodacao, dadosPrenota.acrescimo," & _
       "dadosPrenota.codigounidade, dadosPrenota.desconto, dadosPrenota.comissaofabricante, dadosPrenota.DescricaoDesconto," & _
       "dadosPrenota.codsufixo, dadosPrenota.codprefixo, dadosPrenota.codigo AS codigopesquisa, dadosPrenota.custo, dadosPrenota.Lucro," & _
       "dadosPrenota.LucroPerc, dadosPrenota.prenota, dadosPrenota.nf, dadosPrenota.codpeca, dadosPrenota.Icms, dadosPrenota.CST, dadosPrenota.cancelado " & _
       "FROM dadosPrenota order by nf,item"

Set RsDados = AbreRecordset(StrSql, True)
If Not RsEntrada.EOF Then
   RsEntrada.MoveLast
   TotalPro = RsEntrada.RecordCount
   RsEntrada.MoveFirst
End If
Do Until RsEntrada.EOF
  ' If RsEntrada!nf = "005912" Then Stop

   TProce = TProce + 1
   Sintegra.Caption = "processando Saida Tipo 50 NF:" & RsEntrada!NF & " - Registros " & TProce & " de " & TotalPro
   DoEvents
    Frete = 0
    Seguro = 0
    Pis = 0
    Complemento = 0
    NaoTributado = 0
    DespesasAcessorias = 0 'IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
    Custo = 0
    ipi = 0
    ValorDebitar = ipi + Frete + Seguro + Pis + NaoTributado + DespesasAcessorias + Custo

   '==>Filtra os registros dos itens
   RsDados.Filter = "CodNota=" & RsEntrada!CodNota
   '==> Calcula o total de registros
   If Not RsDados.EOF Then
      RsDados.MoveLast
      Total54 = RsDados.RecordCount
      RsDados.MoveFirst
   Else
      Total54 = 0
   End If
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   Inscricao = AcertaInscricao(RsEntrada!Cliente, ECliente(RsEntrada!CFOP))
   Estado = BuscaEstado(RsEntrada!Cliente, ECliente(RsEntrada!CFOP))
   Cnpj = AcertaCnpj(RsEntrada!Cliente, ECliente(RsEntrada!CFOP))
   DescontoBrutoIcms = 0
   AcrescimoAc = 0
   AcrescimoBruto = 0
   '==> Acerta valores dos lancamentos
   Valor_Outras = 0 'IIf(Not IsNull(RsEntrada!ValorIcmsSubstituicao), RsEntrada!ValorIcmsSubstituicao, 0)
   Serie = IIf(CDbl(RsEntrada!NF) < 7592, Left("2" & "   ", 3), Left("1" & "   ", 3))
   Situacao = IIf(RsEntrada!Status = "Cancelado", "S", "N")
   Desconto = IIf(Not IsNull(RsEntrada!TotalDesconto), RsEntrada!TotalDesconto, 0)
   DescontoBruto = Desconto
   AcrescimoBruto = IIf(Not IsNull(RsEntrada!Acrecimo), RsEntrada!Acrecimo, 0)
'==> Rateia o desconto
   If Desconto > 0 Then Desconto = IIf(Total54 > 0, Desconto / Total54, Desconto)
   '****** Processa o registro 50
   ReDim MtIcms(0)
   MtIcms(0).icms = 0
   MtIcms(0).valor = 0
   
   If IcmsDiferente(RsDados, CDbl(RsEntrada!icms), False) Then
      Valor_t = IIf(Total54 > 0, RsEntrada!valor / Total54, RsEntrada!valor)
      TotalRegMatriz = UBound(MtIcms)
     If TotalRegMatriz = 0 Then TotalRegMatriz = 1
     DescontoBrutoIcms = DescontoBruto / TotalRegMatriz
     AcrescimoAc = AcrescimoBruto / TotalRegMatriz
      For a = 0 To UBound(MtIcms)
            Registro50.Aliquota = MtIcms(a).icms
            If RsDados!icms > 0 Then
              Registro50.Base_Calculo_Icms = MtIcms(0).valor - DescontoBrutoIcms + AcrescimoAc
            Else
              Registro50.Base_Calculo_Icms = 0
            End If
            Registro50.CFOP = RsEntrada!CFOP
            Registro50.Cnpj = Cnpj
            Registro50.Data = Format(RsEntrada!Entrada, "yyyymmdd")
            Registro50.Emitente = "P"
            Registro50.Estado = Estado
            Registro50.Inscricao = Inscricao
            Registro50.Isenta_Nao_Tributada = "0000000000000"
            Registro50.Modelo = "01"
            Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
            Registro50.Outras = IIf(MtIcms(a).icms = 0, Valor_Outras, 0)
            Registro50.Serie = Serie
            Registro50.Situacao = Situacao
            Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
            Registro50.Valor_Total = Valor_t
            
            CriarDados50 Registro50
      Next
   Else
     Registro50.Aliquota = RsEntrada!icms '************************IIf(Not RsDados.EOF, RsDados!Icms, 0)
     TotalRegMatriz = UBound(MtIcms)
     If TotalRegMatriz = 0 Then TotalRegMatriz = 1
     DescontoBrutoIcms = DescontoBruto / TotalRegMatriz
      AcrescimoAc = AcrescimoBruto / TotalRegMatriz
     If Not RsDados.EOF Then
        If RsEntrada!icms > 0 Then '*******************************RsDados!Icms > 0 Then
          Registro50.Base_Calculo_Icms = MtIcms(0).valor - DescontoBrutoIcms + AcrescimoAc
        Else
          Registro50.Base_Calculo_Icms = 0
        End If
     Else
       Registro50.Base_Calculo_Icms = 0
     End If
     Registro50.CFOP = RsEntrada!CFOP
     Registro50.Cnpj = Cnpj
     Registro50.Data = Format(RsEntrada!Entrada, "yyyymmdd")
     Registro50.Emitente = "P"
     Registro50.Estado = Estado
     Registro50.Inscricao = Inscricao
     Registro50.Isenta_Nao_Tributada = "0000000000000"
     Registro50.Modelo = "01"
     Registro50.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
     Registro50.Outras = Valor_Outras
     Registro50.Serie = Serie
     Registro50.Situacao = Situacao
     Registro50.Valor_Icms = Registro50.Base_Calculo_Icms * (Registro50.Aliquota / 100)
     Registro50.Valor_Total = RsEntrada!valor
     CriarDados50 Registro50
   End If
        
   '===> Processa os registros 54
   NItem = 0
   Do Until RsDados.EOF
        '==> Calcula valores referentes ao item
        NItem = NItem + 1
        ValorIpi = (IIf(IsNumeric(RsDados!ipi), RsDados!ipi, 0) / 100) * RsDados!Valor_Total_Compra

        Registro54.Aliquota_Icms = RsEntrada!icms '**************************IIf(Not IsNull(RsDados!Icms), RsDados!Icms, 0)
        If RsEntrada!icms > 0 Then '*******************************************IIf(Not IsNull(RsDados!Icms), RsDados!Icms, 0) > 0 Then
           Registro54.Base_Calculo_Icms = RsDados!Valor_Total_Compra - Desconto
        Else
           Registro54.Base_Calculo_Icms = 0
        End If
        Registro54.Base_Calculo_subs_Trib = "000000000000"
        Registro54.CFOP = RsEntrada!CFOP
        Registro54.Cnpj = Cnpj
        Registro54.Codigo_Produto = RsDados!codigopesquisa
        Registro54.cst = IIf(Not IsNull(RsDados!cst), RsDados!cst, "000")
        Registro54.Modelo = "01"
        Registro54.Numero_Item = NItem
        Registro54.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
        Registro54.Quantidade = RsDados!Qtde
        Registro54.Serie = Serie
        Registro54.Valor_Desconto = Desconto
        Registro54.Valor_Ipi = ValorIpi
        Registro54.Valor_Produto = RsDados!Valor_Total_Compra
        
        '==> cria o registro 54
        CriarDados54 Registro54
        
        '==>Valos criar o registro75
        
        Registro75.Aliquota_Icms = RsEntrada!icms '***************************** IIf(Not IsNull(RsDados!Icms), RsDados!Icms, 0)
        Registro75.Aliquota_Ipi = RsDados!ipi
        Registro75.Base_Icms_subst = "000000000000"
        Registro75.Codigo = RsDados!codigopesquisa
        Registro75.Nome = RsDados!Descricao
        Registro75.Reducao_Base = "00000"
        Registro75.Unidade = RsDados!Unimed
        
        '==> Cria o Registro 75
        CriarDados75 Registro75
        
        RsDados.MoveNext
   Loop
   '==> Gera dados para registro complementares
   RegistroComp.CFOP = RsEntrada!CFOP
   RegistroComp.Cnpj = Cnpj
   RegistroComp.Modelo = "01"
   RegistroComp.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
   RegistroComp.Serie = Serie
   RegistroComp.Valor_Compl = 0 'IIf(Not IsNull(RsEntrada!Complemento), RsEntrada!Complemento, 0)
   RegistroComp.Valor_Despesas = IIf(Not IsNull(RsEntrada!DespesasAcessorias), RsEntrada!DespesasAcessorias, 0)
   RegistroComp.Valor_Frete = 0 'IIf(Not IsNull(RsEntrada!Frete), RsEntrada!Frete, 0)
   RegistroComp.Valor_pis = 0 'IIf(Not IsNull(RsEntrada!PIS_COFINS), RsEntrada!PIS_COFINS, 0)
   RegistroComp.Valor_Seguro = 0 '= IIf(Not IsNull(RsEntrada!Seguro), RsEntrada!Seguro, 0)
   RegistroComp.Valor_Servicos = 0 '= IIf(Not IsNull(RsEntrada!NaoTributado), RsEntrada!NaoTributado, 0)
   RegistroComp.Total_Itens = Total54
   Processar_Dados_complementares RegistroComp
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

End Function
Public Function ECliente(CFOP As String) As Boolean

On Error Resume Next
Dim RsNatureza As ADODB.Recordset
Dim StrSql As String

StrSql = "Select * from naturezanotasaida where cfop='" & CFOP & "'"
Set RsNatureza = AbreRecordset(StrSql, True)

If Not RsNatureza.EOF Then
   If RsNatureza!ImprimeCodigoFornecedor Then
      ECliente = False
   Else
      ECliente = True
   End If
Else
  ECliente = True
End If
Set RsNatureza = Nothing
End Function

Public Function ProcessaConhecimentoFreteBhForte(DataInicio As Date, DataFim As Date)
Dim StrSql          As String
Dim RsEntrada       As ADODB.Recordset
Dim Registro70      As Dados70
Dim Cnpj            As String
Dim Inscricao       As String
Dim Estado          As String
Dim Valor_Outras    As Double
Dim Valor_t         As Double
Dim Desconto        As Double
Dim ValorIpi        As Double
Dim baseCalculo     As Double
Dim Serie           As String
Dim SubSerie        As String
Dim Situacao        As String
Dim Total54         As Integer

On Error GoTo err70

StrSql = "SELECT EntradaNf.NUMLANCTO, EntradaNf.NF, EntradaNf.RECDESP, EntradaNf.CLICRED," & _
        "EntradaNf.VALOR, EntradaNf.DATA, EntradaNf.CONTR, EntradaNf.VP, EntradaNf.codigo," & _
        "EntradaNf.ValorProduto, EntradaNf.BaseIcms, EntradaNf.Icms, EntradaNf.Ipi," & _
        "EntradaNf.Geral, EntradaNf.Complementar, EntradaNf.cnpj, EntradaNf.codigofor," & _
        "EntradaNf.custos, EntradaNf.ContasLancadas, EntradaNf.parcelas, EntradaNf.primeiro," & _
        "EntradaNf.intermediario, EntradaNf.entrada, EntradaNf.Especie, EntradaNf.Serie," & _
        "EntradaNf.CodigoFiscal AS Cfop, EntradaNf.Tipo, EntradaNf.CodigoFormaDePagamento," & _
        "EntradaNf.QuantidadeDeItem, EntradaNf.vencimento0, EntradaNf.vencimento1," & _
        "EntradaNf.vencimento2, EntradaNf.vencimento3, EntradaNf.vencimento4," & _
        "EntradaNf.vencimento5, EntradaNf.vencimento6, EntradaNf.vencimento7, EntradaNf.vencimento8, EntradaNf.vencimento9, EntradaNf.valor1," & _
        "EntradaNf.valor2, EntradaNf.valor3, EntradaNf.valor4, EntradaNf.valor5, EntradaNf.valor6, EntradaNf.valor7," & _
        "EntradaNf.valor8 , EntradaNf.valor9, EntradaNf.valo0, EntradaNf.Usuario, EntradaNf.Modelo , EntradaNf.Sintegra," & _
        "EntradaNf.Status, EntradaNf.natureza, EntradaNf.Frete, EntradaNf.Seguro, EntradaNf.PIS_COFINS, EntradaNf.Complemento," & _
        "EntradaNf.NaoTributado, EntradaNf.DespesasAcessorias, EntradaNf.BaseCalculoIcmsSubstituicao, EntradaNf.ValorIcmsSubstituicao, EntradaNf.SubSerie, EntradaNf.TipoFrete, EntradaNf.TotalDesconto, EntradaNf.E_Cliente" & _
        " FROM EntradaNf" & _
        " WHERE EntradaNf.entrada Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFim, "mm/dd/yy") & "#"

Debug.Print StrSql

Set RsEntrada = AbreRecordset(StrSql, True)
RsEntrada.Filter = "cfop='1353' OR cfop='2353'"

'RsEntrada.MoveLast
'MsgBox RsEntrada.RecordCount
'RsEntrada.MoveFirst
Do Until RsEntrada.EOF
   Sintegra.Caption = "processando Entrada Registro 70 NF:" & RsEntrada!NF
   DoEvents
   
   'GeraReg50 True, RsEntrada, RsDados
   
   '==> Buisca dados referente ao fornecedor
   Inscricao = AcertaInscricao(RsEntrada!codigofor, RsEntrada!E_Cliente)
   Estado = BuscaEstado(RsEntrada!codigofor, RsEntrada!E_Cliente)
   Cnpj = AcertaCnpj(RsEntrada!codigofor, RsEntrada!E_Cliente)
   
   '==> Acerta valores dos lancamentos
   Valor_Outras = IIf(Not IsNull(RsEntrada!ValorIcmsSubstituicao), RsEntrada!ValorIcmsSubstituicao, 0)
   Serie = Left("1" & "   ", 3)
   SubSerie = IIf(IsNull(RsEntrada!SubSerie), "", RsEntrada!SubSerie)
   SubSerie = Left(SubSerie & "  ", 2)
   
   If RsEntrada!icms = 0 Then
        baseCalculo = 0
   Else
        baseCalculo = RsEntrada!valor - RsEntrada!ipi
   End If
   Situacao = IIf(RsEntrada!Status = "Cancelado", "S", "N")
   
   Registro70.Cnpj = Cnpj
   Registro70.Inscricao = Inscricao
   Registro70.Data = Format(RsEntrada!Entrada, "yyyymmdd")
   Registro70.Estado = Estado
   Registro70.Modelo = BuscaModeloTransporte
   Registro70.Serie = IIf(RsEntrada!Serie = "1", "U", IIf(RsEntrada!Serie = "", "U", RsEntrada!Serie)) & ""
   Registro70.SubSerie = SubSerie
   Registro70.Numero_Nf = Right("000000" & RsEntrada!NF, 6)
   Registro70.CFOP = RsEntrada!CFOP
   Registro70.Valor_Total = RsEntrada!valor
   Registro70.Base_Calculo_Icms = baseCalculo
   Registro70.Valor_Icms = Registro70.Base_Calculo_Icms * (RsEntrada!icms / 100)
   Registro70.Isenta_Nao_Tributada = "00000000000000"
   Registro70.Outras = Valor_Outras
   Registro70.CifFob = IIf(IsNull(RsEntrada!TipoFrete), "0", RsEntrada!TipoFrete)
   Registro70.Situacao = Situacao
   
   
   CriarDados70 Registro70
   
        
   
   RsEntrada.MoveNext
Loop
Set RsEntrada = Nothing
Set RsDados = Nothing

Exit Function
err70:
MsgBox err.Description & err.Number
Resume 0
End Function
Function BuscaModeloTransporte() As String
On Error Resume Next
Dim NomeArq As String
Dim a As Integer
Dim Lca As Integer
Dim valor As String
GLBase = AcessoAdo.LocalBaseDados
For a = Len(GLBase) To 1 Step -1
   If Mid(GLBase, a, 1) = "\" Then Exit For
Next
NomeArq = Mid(GLBase, 1, a) & "modelotrans.txt"
Lca = FreeFile
Open NomeArq For Input As #Lca
Line Input #Lca, valor
 
Close #Lca
BuscaModeloTransporte = valor
End Function
Function IcmsDiferente(RsDados As ADODB.Recordset, Optional icms As Double = 0, Optional ProcessaIcms As Boolean = True) As Boolean

Dim a           As Integer
Dim b           As Integer
Dim primeiro    As Boolean
Dim Achou       As Boolean
primeiro = True
ReDim MtIcms(0)
If Not RsDados.EOF Or Not RsDados.BOF Then RsDados.MoveFirst
If ProcessaIcms Then
    Do Until RsDados.EOF
       Achou = False
       If primeiro Then
          a = 0
          MtIcms(a).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
          If mvarOrigem = "BHFORTE" Then
             MtIcms(a).valor = RsDados!Valor_Total_Compra
          Else
            MtIcms(a).valor = RsDados!ValorTotal
          End If
          primeiro = False
       Else
          a = UBound(MtIcms) + 1
          '==> Procura se ja teve outro
          For b = 0 To UBound(MtIcms)
             If MtIcms(b).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0) Then
                Achou = True
                Exit For
             End If
          Next
          If Achou Then
            If mvarOrigem = "BHFORTE" Then
               MtIcms(b).valor = MtIcms(b).valor + RsDados!Valor_Total_Compra
            Else
              MtIcms(b).valor = MtIcms(b).valor + RsDados!ValorTotal
            End If
             
          Else
            a = UBound(MtIcms) + 1
            ReDim Preserve MtIcms(a)
            MtIcms(a).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
            If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
            Else
              MtIcms(a).valor = RsDados!ValorTotal
            End If
          End If
          
       End If
       RsDados.MoveNext
    Loop
Else
  Do Until RsDados.EOF
       If primeiro Then
          a = 0
          MtIcms(a).icms = icms
          If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
          Else
              MtIcms(a).valor = RsDados!ValorTotal
          End If
          primeiro = False
       Else
          a = UBound(MtIcms) + 1
          '==> Procura se ja teve outro
          For b = 0 To UBound(MtIcms)
             If MtIcms(b).icms = icms Then
                Achou = True
                Exit For
             End If
          Next
          If Achou Then
            If mvarOrigem = "BHFORTE" Then
               MtIcms(b).valor = MtIcms(b).valor + RsDados!Valor_Total_Compra
            Else
                MtIcms(b).valor = MtIcms(b).valor + RsDados!ValorTotal
            End If
             
          Else
            a = UBound(MtIcms) + 1
            ReDim Preserve MtIcms(a)
            MtIcms(a).icms = icms
            If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
            Else
              MtIcms(a).valor = RsDados!ValorTotal
            End If
          End If
          
       End If
       RsDados.MoveNext
    Loop
End If
If UBound(MtIcms) > 0 Then
   IcmsDiferente = True
Else
   IcmsDiferente = False
End If
If Not RsDados.EOF Or Not RsDados.BOF Then RsDados.MoveFirst
End Function
Function IcmsDiferenteSaidaLidis(RsDados As ADODB.Recordset, Optional icms As Double = 0, Optional ProcessaIcms As Boolean = True) As Boolean

Dim a           As Integer
Dim b           As Integer
Dim primeiro    As Boolean
Dim Achou       As Boolean
primeiro = True
ReDim MtIcms(0)
If Not RsDados.EOF Or Not RsDados.BOF Then RsDados.MoveFirst
If ProcessaIcms Then
    Do Until RsDados.EOF
       Achou = False
       If primeiro Then
          a = 0
          MtIcms(a).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
          If mvarOrigem = "BHFORTE" Then
             MtIcms(a).valor = RsDados!Valor_Total_Compra
          Else
            MtIcms(a).valor = RsDados!Qtde * RsDados!VALUNIT
          End If
          primeiro = False
       Else
          a = UBound(MtIcms) + 1
          '==> Procura se ja teve outro
          For b = 0 To UBound(MtIcms)
             If MtIcms(b).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0) Then
                Achou = True
                Exit For
             End If
          Next
          If Achou Then
            If mvarOrigem = "BHFORTE" Then
               MtIcms(b).valor = MtIcms(b).valor + RsDados!Valor_Total_Compra
            Else
              MtIcms(b).valor = MtIcms(b).valor + (RsDados!Qtde * RsDados!VALUNIT)
            End If
             
          Else
            a = UBound(MtIcms) + 1
            ReDim Preserve MtIcms(a)
            MtIcms(a).icms = IIf(Not IsNull(RsDados!icms), RsDados!icms, 0)
            If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
            Else
              MtIcms(a).valor = RsDados!Qtde * RsDados!VALUNIT
            End If
          End If
          
       End If
       RsDados.MoveNext
    Loop
Else
  Do Until RsDados.EOF
       If primeiro Then
          a = 0
          MtIcms(a).icms = icms
          If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
          Else
              MtIcms(a).valor = RsDados!Qtde * RsDados!VALUNIT
          End If
          primeiro = False
       Else
          a = UBound(MtIcms) + 1
          '==> Procura se ja teve outro
          For b = 0 To UBound(MtIcms)
             If MtIcms(b).icms = icms Then
                Achou = True
                Exit For
             End If
          Next
          If Achou Then
            If mvarOrigem = "BHFORTE" Then
               MtIcms(b).valor = MtIcms(b).valor + RsDados!Valor_Total_Compra
            Else
                MtIcms(b).valor = MtIcms(b).valor + (RsDados!Qtde * RsDados!VALUNIT)
            End If
             
          Else
            a = UBound(MtIcms) + 1
            ReDim Preserve MtIcms(a)
            MtIcms(a).icms = icms
            If mvarOrigem = "BHFORTE" Then
               MtIcms(a).valor = RsDados!Valor_Total_Compra
            Else
              MtIcms(a).valor = RsDados!Qtde * RsDados!VALUNIT
            End If
          End If
          
       End If
       RsDados.MoveNext
    Loop
End If
If UBound(MtIcms) > 0 Then
   IcmsDiferenteSaidaLidis = True
Else
   IcmsDiferenteSaidaLidis = False
End If
If Not RsDados.EOF Or Not RsDados.BOF Then RsDados.MoveFirst
End Function

Sub VerificaProdPara75(CodigoPro As String, Descricao As String, AliquotaIcms As String, AliquotaIpi As String, BaseSubstituicao As String, Unidade As String)
'******** Excluir metodo
On Error Resume Next
Dim a As Integer
Dim Existe As Boolean
'If CodigoPro = "BD2T10850A" Then Stop
If Not TemRegistro75 Then
   TemRegistro75 = True
   a = 0
   Existe = False
Else
'If Len(CodigoPro) = 0 Then Stop
    a = UBound(Mt75)
    If err.Number = 0 Then
        For a = 0 To UBound(Mt75)
            If Trim(Mt75(a).Codigo) = Trim(CodigoPro) Then
               Existe = True
               Exit For
            End If
        Next
    Else
      Existe = False
    End If

End If

If Not Existe Then
    AliquotaIcms = IIf(IsNumeric(AliquotaIcms), AliquotaIcms, 0)
    AliquotaIcms = ClUtil.AcertaNumero(AliquotaIcms, 2)
    AliquotaIcms = Replace(AliquotaIcms, ".", "")
    AliquotaIcms = Replace(AliquotaIcms, ",", "")
    AliquotaIcms = Right("0000" & AliquotaIcms, 4)
    
    AliquotaIpi = IIf(IsNumeric(AliquotaIpi), AliquotaIpi, 0)
    AliquotaIpi = ClUtil.AcertaNumero(AliquotaIpi, 2)
    AliquotaIpi = Replace(AliquotaIpi, ".", "")
    AliquotaIpi = Replace(AliquotaIpi, ",", "")
    AliquotaIpi = Right("00000" & AliquotaIpi, 5)
    
    ReDim Preserve Mt75(a)
    Mt75(a).Codigo = CodigoPro
    Mt75(a).Nome = Descricao
    Mt75(a).Aliquota_Icms = AliquotaIcms
    Mt75(a).Aliquota_Ipi = AliquotaIpi
    Mt75(a).Base_Icms_subst = BaseSubstituicao
    Mt75(a).Reducao_Base = "00000"
    Mt75(a).Unidade = Unidade
End If
End Sub

Sub Criar_Dados_complementares(CodigoNota As String, Cnpj As String, Modelo As String, Serie As String, CFOP As String, valor As Double, CodigoComplemento As String)
'******** Excluir metodo
Dim a           As Long
Dim StrSql      As String
Dim BaseIcms    As String

On Error Resume Next
err.Number = 0
a = UBound(Mt54)
If err.Number <> 0 Then a = 0 Else a = a + 1
ReDim Preserve Mt54(a)
Mt54(a).Cnpj = Cnpj '2
Mt54(a).Modelo = Modelo '3
Mt54(a).Serie = Serie '4
Mt54(a).Numero_Nf = CodigoNota '5
Mt54(a).CFOP = CFOP '6
Mt54(a).cst = Right("   ", 3)  '7
Mt54(a).Numero_Item = CodigoComplemento   '8
Mt54(a).Codigo_Produto = Left("                ", 14)  '9
Mt54(a).Quantidade = "00000000000"  '10
Mt54(a).Valor_Produto = "000000000000"  '11
                            
BaseIcms = ClUtil.AcertaNumero(CStr(valor), 2)
BaseIcms = Replace(BaseIcms, ".", "")
BaseIcms = Replace(BaseIcms, ",", "")
              
Mt54(a).Valor_Desconto = Right("000000000000" & BaseIcms, 12) '12
Mt54(a).Base_Calculo_Icms = "000000000000" '13
Mt54(a).Base_Calculo_subs_Trib = "000000000000" '14
Mt54(a).Valor_Ipi = "000000000000"  '15
Mt54(a).Aliquota_Icms = "0000" '16

End Sub
Public Function RetornaCfop(CFOP As String, Aliquota As Double, Optional IgnoraDiferenca As Boolean = False) As String
Dim Resposta As String

If IgnoraDiferenca Then
   RetornaCfop = CFOP
   Exit Function
End If
If Aliquota > 0 Then
  If CFOP <> "1603" Then
    Select Case CFOP
       Case Is = "1403"
           Resposta = "1102"
       Case Is = "2403"
           Resposta = "2102"
       Case Is = "5403"
           Resposta = "5102"
       Case Is = "6403"
          Resposta = "6102"
       Case Else
         Resposta = CFOP
    End Select
  End If
Else
  If CFOP <> "1603" Then
    Select Case CFOP
       Case Is = "1102"
           Resposta = "1403"
       Case Is = "2102"
           Resposta = "2403"
       Case Is = "5102"
           Resposta = "5403"
       Case Is = "6102"
          Resposta = "6403"
       Case Else
         Resposta = CFOP
    End Select
  End If
End If
RetornaCfop = Resposta
End Function
Public Function AcertaInscricao(Codigo As Long, Cliente As Boolean) As String

Dim rsCliente   As ADODB.Recordset
Dim StrSql      As String
Dim Inscricao   As String
Dim db          As Database
Dim Rs          As Recordset


If Cliente Then
   If UCase(mvarOrigem) = "BHFORTE" Then
        StrSql = "Select * from alid001 where codigo=" & Codigo
        Set rsCliente = AbreRecordset(StrSql, True)
        If Not rsCliente.EOF Then
           If Not IsNull(rsCliente!INSCEST) Then
                Inscricao = Replace(rsCliente!INSCEST, "/", "")
                Inscricao = Replace(Inscricao, "-", "")
                Inscricao = Replace(Inscricao, ".", "")
                Inscricao = Replace(Inscricao, " ", "")
                If Len(Inscricao) > 0 Then
                   Inscricao = rsCliente!INSCEST & ""
                Else
                   Inscricao = "ISENTO"
                End If
           Else
                Inscricao = "ISENTO"
           End If
        End If
    End If
    If UCase(mvarOrigem) = "LIDIS" Then
        StrSql = "Select * from alid001 where codigo='" & Right("00000" & Codigo, 5) & "'"
        Set db = OpenDatabase(GLBase)
        Set Rs = db.OpenRecordset(StrSql)
        If Not Rs.EOF Then
           If Not IsNull(Rs!INSCEST) Then
                Inscricao = Replace(Rs!INSCEST, "/", "")
                Inscricao = Replace(Inscricao, "-", "")
                Inscricao = Replace(Inscricao, ".", "")
                Inscricao = Replace(Inscricao, " ", "")
                If Len(Inscricao) > 0 Then
                   Inscricao = Rs!INSCEST & ""
                Else
                   Inscricao = "ISENTO"
                End If
           Else
                Inscricao = "ISENTO"
           End If
        End If
    End If
Else
    If UCase(mvarOrigem) = "BHFORTE" Then
        StrSql = "Select * from fornecedor where codigo=" & Codigo
        Set rsCliente = AbreRecordset(StrSql, True)
        If Not rsCliente.EOF Then
           If Not IsNull(rsCliente!INSCEST) Then
                Inscricao = Replace(rsCliente!INSCEST, "/", "")
                Inscricao = Replace(Inscricao, "-", "")
                Inscricao = Replace(Inscricao, ".", "")
                Inscricao = Replace(Inscricao, " ", "")
                If Len(Inscricao) > 0 Then
                   Inscricao = rsCliente!INSCEST & ""
                Else
                   Inscricao = "ISENTO"
                End If
           Else
                Inscricao = "ISENTO"
           End If
        End If
    End If
    If UCase(mvarOrigem) = "LIDIS" Then
        StrSql = "Select * from ALID002 where codigo='" & Right("00000" & Codigo, 5) & "'"
        Set db = OpenDatabase(GLBase)
        Set Rs = db.OpenRecordset(StrSql)
        If Not Rs.EOF Then
           If Not IsNull(Rs!INSCEST) Then
                Inscricao = Replace(Rs!INSCEST, "/", "")
                Inscricao = Replace(Inscricao, "-", "")
                Inscricao = Replace(Inscricao, ".", "")
                Inscricao = Replace(Inscricao, " ", "")
                If Len(Inscricao) > 0 Then
                   Inscricao = Rs!INSCEST & ""
                Else
                   Inscricao = "ISENTO"
                End If
           Else
                Inscricao = "ISENTO"
           End If
        End If
    End If
   
  
End If

Inscricao = Replace(Inscricao, ",", "")
Inscricao = Replace(Inscricao, ".", "")
Inscricao = Replace(Inscricao, "-", "")
Inscricao = Replace(Inscricao, "/", "")
Inscricao = Replace(Inscricao, "\", "")
Inscricao = Replace(Inscricao, " ", "")
Inscricao = Left(Inscricao & "              ", 14)
AcertaInscricao = Inscricao
End Function

Public Function AcertaCnpj(Codigo As Long, Cliente As Boolean) As String
Dim rsCliente   As ADODB.Recordset
Dim StrSql      As String
Dim Cnpj        As String
Dim db          As Database
Dim Rs          As Recordset

If Cliente Then
  If UCase(mvarOrigem) = "BHFORTE" Then
        StrSql = "Select * from alid001 where codigo=" & Codigo
        Set rsCliente = AbreRecordset(StrSql, True)
        If Not rsCliente.EOF Then
           If Not IsNull(rsCliente!CGC) Then
                Cnpj = Replace(rsCliente!CGC, "/", "")
                Cnpj = Replace(Cnpj, "-", "")
                Cnpj = Replace(Cnpj, ".", "")
                Cnpj = Replace(Cnpj, " ", "")
                If Len(Cnpj) > 0 Then
                   Cnpj = rsCliente!CGC & ""
                Else
                   Cnpj = rsCliente!cpf & ""
                End If
           Else
                Cnpj = rsCliente!cpf & ""
           End If
        End If
  End If
  If UCase(mvarOrigem) = "LIDIS" Then
     StrSql = "Select * from alid001 where codigo='" & Right("00000" & Codigo, 5) & "'"
        Set db = OpenDatabase(GLBase)
        Set Rs = db.OpenRecordset(StrSql)
        If Not Rs.EOF Then
           If Not IsNull(Rs!CGC) Then
                Cnpj = Replace(Rs!CGC, "/", "")
                Cnpj = Replace(Cnpj, "-", "")
                Cnpj = Replace(Cnpj, ".", "")
                Cnpj = Replace(Cnpj, " ", "")
                If Len(Cnpj) > 0 Then
                   Cnpj = Rs!CGC & ""
                Else
                   Cnpj = Rs!cpf & ""
                End If
           Else
                Cnpj = Rs!cpf & ""
           End If
        End If
  End If

   
Else
  If UCase(mvarOrigem) = "BHFORTE" Then
     StrSql = "Select * from fornecedor where codigo=" & Codigo
     Set rsCliente = AbreRecordset(StrSql, True)
     If Not rsCliente.EOF Then
      If Not IsNull(rsCliente!CGC) Then
           Cnpj = Replace(rsCliente!CGC, "/", "")
           Cnpj = Replace(Cnpj, "-", "")
           Cnpj = Replace(Cnpj, ".", "")
           Cnpj = Replace(Cnpj, " ", "")
       End If
    End If
  End If
  If UCase(mvarOrigem) = "LIDIS" Then
     StrSql = "Select * from ALID002 where codigo='" & Right("00000" & Codigo, 5) & "'"
     Set db = OpenDatabase(GLBase)
     Set Rs = db.OpenRecordset(StrSql)
     If Not Rs.EOF Then
      If Not IsNull(Rs!CGC) Then
           Cnpj = Replace(Rs!CGC, "/", "")
           Cnpj = Replace(Cnpj, "-", "")
           Cnpj = Replace(Cnpj, ".", "")
           Cnpj = Replace(Cnpj, " ", "")
           If Len(Cnpj) = 0 Then
                Cnpj = Replace(Rs!cpf, "/", "")
                Cnpj = Replace(Cnpj, "-", "")
                Cnpj = Replace(Cnpj, ".", "")
                Cnpj = Replace(Cnpj, " ", "")
           End If
      Else
           Cnpj = Replace(Rs!cpf, "/", "")
           Cnpj = Replace(Cnpj, "-", "")
           Cnpj = Replace(Cnpj, ".", "")
           Cnpj = Replace(Cnpj, " ", "")
      
      End If
      
   End If
  End If
  
  
End If

Cnpj = Replace(Cnpj, ",", "")
Cnpj = Replace(Cnpj, ".", "")
Cnpj = Replace(Cnpj, "-", "")
Cnpj = Replace(Cnpj, "/", "")
Cnpj = Replace(Cnpj, "\", "")
Cnpj = Replace(Cnpj, " ", "")
Cnpj = Right("00000000000000" & Cnpj, 14)
AcertaCnpj = Cnpj
End Function
Public Function BuscaEstado(Codigo As Long, Cliente As Boolean) As String
Dim rsCliente   As ADODB.Recordset
Dim StrSql      As String
Dim Cnpj        As String
Dim db          As Database
Dim Rs          As Recordset

If Cliente Then
   If UCase(mvarOrigem) = "BHFORTE" Then
        StrSql = "Select * from alid001 where codigo=" & Codigo
        Set rsCliente = AbreRecordset(StrSql, True)
        If Not rsCliente.EOF Then
           BuscaEstado = rsCliente!Estado
        End If
   End If
   If UCase(mvarOrigem) = "LIDIS" Then
        StrSql = "Select * from alid001 where  codigo='" & Right("00000" & Codigo, 5) & "'"
        Set db = OpenDatabase(GLBase)
        Set Rs = db.OpenRecordset(StrSql)
        If Not Rs.EOF Then
           BuscaEstado = Rs!Estado
        End If
   End If
Else
  If UCase(mvarOrigem) = "BHFORTE" Then
    StrSql = "Select * from fornecedor where codigo=" & Codigo
    Set rsCliente = AbreRecordset(StrSql, True)
    If Not rsCliente.EOF Then
       BuscaEstado = rsCliente!Estado
    End If
  End If
  If UCase(mvarOrigem) = "LIDIS" Then
       StrSql = "Select * from ALID002 where codigo='" & Right("00000" & Codigo, 5) & "'"
        Set db = OpenDatabase(GLBase)
        Set Rs = db.OpenRecordset(StrSql)
        If Not Rs.EOF Then
           BuscaEstado = Rs!Estado
        End If
  End If
End If

End Function

Public Function BuscaDadosBhForte(DataInicio As Date, DataFinal As Date)
Dim StrSql As String
'==>processa a Entrada
StrSql = "Select * from entradanf where entrada Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFinal, "mm/dd/yy") & "# and Sintegra=False order by entrada"
BuscaEntrada StrSql
BuscaDevolucaoCliente
'==>Processa a saida

StrSql = "SELECT * FROM prenota where dtemis Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFinal, "mm/dd/yy") & "# and prenota.Sintegra=false and CondPag<>'Cupom Fiscal'"
StrDevolucao = SeparaDevolucao

If Len(StrDevolucao) > 0 Then
    StrSql = StrSql & " And (" & SeparaDevolucao & ")"
End If
StrSql = StrSql & " order by dtemis"

BuscaSaida StrSql

BuscaSaidaCupomBhForte

ErrosP = ClUtil.erroS


End Function



Private Sub Class_Initialize()
Inicio
End Sub
Function SeparaDevolucao() As String
On Error Resume Next

Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim Resposta As String

StrSql = "Select * from NaturezaNotaSaida where Devolucao=True"

Set Rs = AbreRecordset(StrSql, True)

Do Until Rs.EOF
    If Len(Resposta) > 0 Then Resposta = Resposta & " And "
    Resposta = Resposta & "orcvenda<>'" & Rs!Nome & "'"
    Rs.MoveNext
Loop

Set Rs = Nothing
SeparaDevolucao = Resposta

End Function
Public Sub BuscaEntrada(StrSql As String)
On Error GoTo ErrBuscaEntrada
Dim Rs As ADODB.Recordset
Dim a As Integer
Dim Valor_Total As Double
Dim Valor_Icms  As Double
Dim LcTotal As Integer
Dim Frete       As Double
Dim Seguro      As Double
Dim Pis         As Double
Dim Complemento As Double
Dim NaoTributado As Double
Dim DespesasAcessorias As Double
Dim Custo       As Double
Dim ipi         As Double
Dim ValorDebitar As Double

Dim x As Integer
 
Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
'Debug.Print DEscricaoErro
If Not Rs.EOF Then
    Rs.MoveLast
    LcTotal = Rs.RecordCount
    Rs.MoveFirst
End If
x = 1
With Sintegra
    
    .Entrada.Rows = 1
    Do Until Rs.EOF
      DoEvents
      Frete = IIf(Not IsNull(Rs!Frete), Rs!Frete, 0)
      Seguro = IIf(Not IsNull(Rs!Seguro), Rs!Seguro, 0)
      Pis = IIf(Not IsNull(Rs!PIS_COFINS), Rs!PIS_COFINS, 0)
      Complemento = IIf(Not IsNull(Rs!Complemento), Rs!Complemento, 0)
      NaoTributado = IIf(Not IsNull(Rs!NaoTributado), Rs!NaoTributado, 0)
      DespesasAcessorias = IIf(Not IsNull(Rs!DespesasAcessorias), Rs!DespesasAcessorias, 0)
      Custo = IIf(Not IsNull(Rs!Custos), Rs!Custos, 0)
      ipi = IIf(Not IsNull(Rs!ipi), Rs!ipi, 0)
      ValorDebitar = ipi + Frete + Seguro + Pis + NaoTributado + DespesasAcessorias + Custo
      
      .Caption = "Buscando a NF de Entrada  " & x & " de " & LcTotal & " Nº:" & Rs!NF & ""
      x = x + 1
      ClUtil.VerificaCnpj Rs!clicred, CLng(Rs!codigofor), False, mvarOrigem
      DoEvents
      a = .Entrada.Rows
      .Entrada.Rows = a + 1
      .Entrada.TextMatrix(a, 0) = Rs!NF & ""
      .Entrada.TextMatrix(a, 1) = Format(Rs!Entrada, "dd/mm/yy") & ""
      .Entrada.TextMatrix(a, 2) = Rs!clicred & ""
     ' Entrada.TextMatrix(a, 3) = IIf(Rs!ValorProduto > 0, Rs!ValorProduto, Rs!Valor) & ""
      .Entrada.TextMatrix(a, 3) = Rs!valor & ""
      .Entrada.TextMatrix(a, 4) = Rs!icms & ""
      'Entrada.TextMatrix(a, 5) = AcertaNumero(CStr((Rs!Icms / 100) * IIf(Rs!ValorProduto > 0, Rs!ValorProduto, Rs!Valor)), 2) & ""
      .Entrada.TextMatrix(a, 5) = ClUtil.AcertaNumero(CStr((Rs!icms / 100) * Rs!valor), 2) & ""
      .Entrada.TextMatrix(a, 6) = Rs!codigofor & ""
      .Entrada.TextMatrix(a, 7) = Rs!Modelo & ""
      .Entrada.TextMatrix(a, 8) = Rs!Serie & ""
      .Entrada.TextMatrix(a, 9) = Rs!codigofiscal & "" '==> Codigo Fiscal
      .Entrada.TextMatrix(a, 10) = IIf(Rs!Status = "Cancelado", "S", "N")
      .Entrada.TextMatrix(a, 11) = Rs!NF & ""
      .Entrada.TextMatrix(a, 12) = Rs!ValorIcmsSubstituicao & ""
      .Entrada.TextMatrix(a, 13) = Rs!Codigo & ""
      .Entrada.TextMatrix(a, 14) = ValorDebitar & ""
      .Entrada.TextMatrix(a, 15) = IIf(IsNull(Rs!SubSerie), "", Rs!SubSerie) & ""
      .Entrada.TextMatrix(a, 16) = IIf(IsNull(Rs!TipoFrete), "0", Rs!TipoFrete) & ""
      'Valor_Total = Valor_Total + IIf(Rs!ValorProduto > 0, Rs!ValorProduto, Rs!Valor) & ""
      Valor_Total = Valor_Total + Rs!valor '& ""
      Valor_Icms = Valor_Icms + CDbl(.Entrada.TextMatrix(a, 5))
      If UCase(mvarOrigem) = "BHFORTE" Then ClUtil.VerificaValorProduto CDbl(Rs!ValorProduto), CStr(Rs!NF), CLng(Rs!codigofor), Rs!Codigo
      
      Rs.MoveNext
    Loop
    Set Rs = Nothing
    .TotalEntrada.Text = ClUtil.AcertaNumero(CStr(Valor_Total), 2)
    .IcmsEntrada.Text = ClUtil.AcertaNumero(CStr(Valor_Icms), 2)
End With
ClUtil.VerificaDigitacaoSubstituicao CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
ErrosP = ClUtil.erroS
Exit Sub
ErrBuscaEntrada:
MsgBox err.Description & err.Number
'Resume 0
End Sub
Sub BuscaDevolucaoCliente()
On Error GoTo ErrBuscaSaida
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim a As Integer
Dim Valor_Total As Double
Dim Valor_Icms  As Double
Dim LcTotal As Integer
Dim x As Integer
StrSql = "SELECT ALID001.RAZAOSOC,DevolucaoCliente.Status,DevolucaoCliente.CFOP, DevolucaoCliente.Serie,DevolucaoCliente.doc,DevolucaoCliente.modelo, DevolucaoCliente.DTEMIS, DevolucaoCliente.OBS, DevolucaoCliente.CONTR, DevolucaoCliente.codigo, DevolucaoCliente.CLiente, DevolucaoCliente.OrcVenda, DevolucaoCliente.Vendedor, DevolucaoCliente.cidade, DevolucaoCliente.cep, DevolucaoCliente.nf, DevolucaoCliente.naobaixaestoque, DevolucaoCliente.ncotacao, DevolucaoCliente.maquina, DevolucaoCliente.Sintegra, DevolucaoCliente.Total, DevolucaoCliente.icms " & _
         "FROM DevolucaoCliente INNER JOIN ALID001 ON DevolucaoCliente.CLiente = ALID001.CODIGO where dtemis Between #" & Format(Sintegra.DataI.Text, "mm/dd/yy") & "# And #" & Format(Sintegra.DataF.Text, "mm/dd/yy") & "# order by dtemis"

Set Rs = AbreRecordset(StrSql, True)
Sintegra.DevCliente.Rows = 1
If Not Rs.EOF Then
   Rs.MoveLast
   LcTotal = Rs.RecordCount
   Rs.MoveFirst
End If
x = 1
Do Until Rs.EOF
  DoEvents
  Sintegra.Caption = "Buscando a NF de entrada " & x & " de " & LcTotal & " Nº:" & Rs!NF & ""
  x = x + 1
  a = Sintegra.DevCliente.Rows
  Sintegra.DevCliente.Rows = a + 1
  Sintegra.DevCliente.TextMatrix(a, 0) = Rs!NF & ""
  Sintegra.DevCliente.TextMatrix(a, 1) = Format(Rs!DtEmis, "dd/mm/yy") & ""
  Sintegra.DevCliente.TextMatrix(a, 2) = Rs!razaosoc & ""
  Sintegra.DevCliente.TextMatrix(a, 3) = Rs!total & ""
  Sintegra.DevCliente.TextMatrix(a, 4) = Rs!icms & ""
  Sintegra.DevCliente.TextMatrix(a, 5) = AcertaNumero(CStr((IIf(Not IsNull(Rs!icms), Rs!icms, 0) / 100) * IIf(Not IsNull(Rs!total), Rs!total, 0)), 2) & ""
   
  Sintegra.DevCliente.TextMatrix(a, 6) = Rs!Cliente & ""
  Sintegra.DevCliente.TextMatrix(a, 7) = Rs!Modelo & ""
  Sintegra.DevCliente.TextMatrix(a, 8) = Rs!Serie & ""
  Sintegra.DevCliente.TextMatrix(a, 9) = Rs!CFOP & "" '==> Codigo Fiscal
  Sintegra.DevCliente.TextMatrix(a, 10) = IIf(Rs!Status = "Cancelado", "S", "N")
  Sintegra.DevCliente.TextMatrix(a, 11) = Rs!doc & ""
  Valor_Total = Valor_Total + IIf(Not IsNull(Rs!total), Rs!total, 0)
  Valor_Icms = Valor_Icms + CDbl(Sintegra.DevCliente.TextMatrix(a, 5))
  Rs.MoveNext
Loop
 Set Rs = Nothing
Sintegra.IcmsCliente.Text = AcertaNumero(CStr(Valor_Icms), 2)
Sintegra.TotalCliente.Text = AcertaNumero(CStr(Valor_Total), 2)

Exit Sub
ErrBuscaSaida:
MsgBox err.Description & err.Number
'Resume 0
End Sub
Sub BuscaSaida(StrSql As String)
On Error GoTo ErrBuscaSaida
Dim Rs As ADODB.Recordset
Dim RsNatureza As ADODB.Recordset
Dim RsDados As ADODB.Recordset
Dim a As Integer
Dim Valor_Total As Double
Dim Valor_Icms  As Double
Dim StrDevolucao As String
Dim x As Integer
Dim LcTotal As Integer
Dim ECliente As Boolean

StrSql = "SELECT * FROM prenota where dtemis Between #" & Format(Sintegra.DataI.Text, "mm/dd/yy") & "# And #" & Format(Sintegra.DataF.Text, "mm/dd/yy") & "# and prenota.Sintegra=false and CondPag<>'Cupom Fiscal'"
StrDevolucao = SeparaDevolucao

If Len(StrDevolucao) > 0 Then
    StrSql = StrSql & " And (" & SeparaDevolucao & ")"
End If
StrSql = StrSql & " order by dtemis"
'Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
If Not Rs.EOF Then
   Rs.MoveLast
   LcTotal = Rs.RecordCount
   Rs.MoveFirst
End If
x = 1
Sintegra.Saida.Rows = 1
Do Until Rs.EOF
  DoEvents
  Sintegra.Caption = "Buscando a NF de Saida  " & x & " de " & LcTotal & " Nº:" & Rs!NF & ""
  x = x + 1
  DoEvents
  '==> verifica se é fornecedor ou cliente
 ' If Rs!nf = "001464" Then Stop
  StrSql = "Select * from NaturezaNotaSaida where cfop='" & Rs!CFOP & "'"
  Set RsNatureza = AbreRecordset(StrSql, True)
  If Not RsNatureza.EOF Then
     If RsNatureza!ImprimeCodigoFornecedor Then
        StrSql = "Select * from fornecedor where codigo=" & Rs!Cliente
        ECliente = False
     Else
        StrSql = "Select * from alid001 where codigo=" & Rs!Cliente
        ECliente = True
     End If
  Else
        StrSql = "Select * from alid001 where codigo=" & Rs!Cliente
        ECliente = True
  End If
  Set RsDados = AbreRecordset(StrSql, True)
  If RsDados.EOF Then
     Set RsDados = Nothing
     StrSql = "Select * from alid001 where codigo=" & Rs!Cliente
     Set RsDados = AbreRecordset(StrSql, True)
  End If
  ClUtil.VerificaCnpj RsDados!razaosoc, CLng(Rs!Cliente), ECliente
  
  Set RsNatureza = Nothing
  a = Sintegra.Saida.Rows
  Sintegra.Saida.Rows = a + 1
  Sintegra.Saida.TextMatrix(a, 0) = Rs!NF & ""
  Sintegra.Saida.TextMatrix(a, 1) = Format(Rs!DtEmis, "dd/mm/yy") & ""
  Sintegra.Saida.TextMatrix(a, 2) = RsDados!razaosoc & ""
  Sintegra.Saida.TextMatrix(a, 3) = Rs!TotalGeral & ""
  Sintegra.Saida.TextMatrix(a, 4) = Rs!icms & ""
  Sintegra.Saida.TextMatrix(a, 5) = ClUtil.AcertaNumero(CStr((Rs!icms / 100) * Rs!TotalGeral), 2) & ""
  
  Sintegra.Saida.TextMatrix(a, 6) = Rs!Cliente & ""
  Sintegra.Saida.TextMatrix(a, 7) = LcModelo & ""
  Sintegra.Saida.TextMatrix(a, 8) = Rs!Serie & ""
  Sintegra.Saida.TextMatrix(a, 9) = Rs!CFOP & "" '==> Codigo Fiscal
  Sintegra.Saida.TextMatrix(a, 10) = IIf(Rs!Status = "Cancelado", "S", "N")
  Sintegra.Saida.TextMatrix(a, 11) = Rs!CodNota
  Sintegra.Saida.TextMatrix(a, 12) = IIf(Not IsNull(Rs!TotalDesconto), Rs!TotalDesconto, 0)
  Sintegra.Saida.TextMatrix(a, 13) = Rs!Acrecimo
  Sintegra.Saida.TextMatrix(a, 14) = Rs!ipi
  Valor_Total = Valor_Total + Rs!TotalGeral
  Valor_Icms = Valor_Icms + CDbl(Sintegra.Saida.TextMatrix(a, 5))
  Rs.MoveNext
  Set RsDados = Nothing
Loop
Set Rs = Nothing
Sintegra.TotalSaida.Text = ClUtil.AcertaNumero(CStr(Valor_Total), 2)
Sintegra.IcmsSaida.Text = ClUtil.AcertaNumero(CStr(Valor_Icms), 2)
ErrosP = ClUtil.erroS

Exit Sub
ErrBuscaSaida:
MsgBox err.Description & err.Number
'Resume 0
End Sub

Public Function EscreverSintegraBhForte(NomeSintegra As String)
'******** Excluir metodo
Dim ClEscrever As New EscreverTexto
Dim LcArq As Integer

LcArq = FreeFile
If Dir(NomeSintegra, vbArchive) <> "" Then Kill NomeSintegra
Sintegra.Caption = "Gravando gerando dados do sintegra."

Open NomeSintegra For Output As #LcArq

ClEscrever.NumeroArquivo = LcArq

ClEscrever.EscreverRegistro10 CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
ClEscrever.EscreverRegistro11
ClEscrever.EscreverRegistro50
ClEscrever.EscreverRegistro54
ClEscrever.CriarRegistro60M
ClEscrever.EscreverRegistro70
ClEscrever.EscreverRegistro75 CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
ClEscrever.EscreverRegistro90

Close #LcArq
Set ClEscrever = Nothing

End Function

Public Sub GerarResumo()
On Error Resume Next
Dim a As Long
Dim b As Long
Dim C As Long
Dim d As Long
Dim Achou As Boolean
Dim MtResumo() As DadosResumo
Dim primeiro As Boolean
Dim ClUtil As New Utilitario
Dim Valor_Entrada As Double
Dim Valor_Saida As Double
Dim Base_entrada As Double
Dim Base_Saida As Double
Dim Icms_entrada As Double
Dim Icms_saida As Double
Dim Rs60A As ADODB.Recordset


'==> Resumo do 50
primeiro = True
For a = 0 To UBound(Mt50)
   If Mt50(a).Situacao = "N" Then
    If primeiro Then
           primeiro = False
           ReDim Preserve MtResumo(b)
           MtResumo(b).baseCalculo = Mt50(a).Base_Calculo_Icms
           MtResumo(b).CFOP = Mt50(a).CFOP
           MtResumo(b).valorIcms = Mt50(a).Valor_Icms
           MtResumo(b).ValorIsentas = Mt50(a).Isenta_Nao_Tributada
           MtResumo(b).ValorOutras = Mt50(a).Outras
           MtResumo(b).ValorTotal = Mt50(a).Valor_Total
           b = b + 1
        Else
           Achou = False
           For C = 0 To UBound(MtResumo)
               If MtResumo(C).CFOP = Mt50(a).CFOP Then
                  Achou = True
                  Exit For
               End If
           Next
           If Achou Then
             
             MtResumo(C).baseCalculo = MtResumo(C).baseCalculo + Mt50(a).Base_Calculo_Icms
             MtResumo(C).valorIcms = MtResumo(C).valorIcms + Mt50(a).Valor_Icms
             MtResumo(C).ValorIsentas = MtResumo(C).ValorIsentas + Mt50(a).Isenta_Nao_Tributada
             MtResumo(C).ValorOutras = MtResumo(C).ValorOutras + Mt50(a).Outras
             MtResumo(C).ValorTotal = MtResumo(C).ValorTotal + Mt50(a).Valor_Total
           Else
             ReDim Preserve MtResumo(b)
             MtResumo(b).baseCalculo = Mt50(a).Base_Calculo_Icms
             MtResumo(b).CFOP = Mt50(a).CFOP
             MtResumo(b).valorIcms = Mt50(a).Valor_Icms
             MtResumo(b).ValorIsentas = Mt50(a).Isenta_Nao_Tributada
             MtResumo(b).ValorOutras = Mt50(a).Outras
             MtResumo(b).ValorTotal = Mt50(a).Valor_Total
             b = b + 1
           End If
        End If
    End If
Next

'==> Escreve no grid resumocfop

With Sintegra.ResumoCfop
   .Rows = 1
   For a = 0 To UBound(MtResumo)
       d = .Rows
       .Rows = d + 1
       .TextMatrix(d, 0) = MtResumo(a).CFOP
       .TextMatrix(d, 1) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorTotal), 2)
       .TextMatrix(d, 2) = ClUtil.AcertaNumero(CStr(MtResumo(a).baseCalculo), 2)
       .TextMatrix(d, 3) = ClUtil.AcertaNumero(CStr(MtResumo(a).valorIcms), 2)
       .TextMatrix(d, 4) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorIsentas), 2)
       .TextMatrix(d, 5) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorOutras), 2)
       If Mid(MtResumo(a).CFOP, 1, 1) = "1" Or Mid(MtResumo(a).CFOP, 1, 1) = "2" Then
          Valor_Entrada = Valor_Entrada + MtResumo(a).ValorTotal
          Base_entrada = Base_entrada + MtResumo(a).baseCalculo
          Icms_entrada = Icms_entrada + MtResumo(a).valorIcms
       Else
          Valor_Saida = Valor_Saida + MtResumo(a).ValorTotal
          Base_Saida = Base_Saida + MtResumo(a).baseCalculo
          Icms_saida = Icms_saida + MtResumo(a).valorIcms
       End If
   Next
End With

'==> Escreve o resumo
Sintegra.Resumo.Rows = 5
Sintegra.Resumo.TextMatrix(1, 0) = "Registro 50 - Entrada"
Sintegra.Resumo.TextMatrix(1, 1) = ClUtil.AcertaNumero(CStr(Valor_Entrada), 2)
Sintegra.Resumo.TextMatrix(1, 2) = ClUtil.AcertaNumero(CStr(Base_entrada), 2)
Sintegra.Resumo.TextMatrix(1, 3) = ClUtil.AcertaNumero(CStr(Icms_entrada), 2)

'==> Saidas
Sintegra.Resumo.TextMatrix(2, 0) = "Registro 50 - Saida"
Sintegra.Resumo.TextMatrix(2, 1) = ClUtil.AcertaNumero(CStr(Valor_Saida), 2)
Sintegra.Resumo.TextMatrix(2, 2) = ClUtil.AcertaNumero(CStr(Base_Saida), 2)
Sintegra.Resumo.TextMatrix(2, 3) = ClUtil.AcertaNumero(CStr(Icms_saida), 2)


'===> resumo do 70
Valor_Entrada = 0
Base_entrada = 0
Icms_entrada = 0
Valor_Saida = 0
Base_Saida = 0
Icms_saida = 0

ReDim MtResumo(0)
primeiro = True
For a = 0 To UBound(Mt70)
   If primeiro Then
      primeiro = False
      ReDim Preserve MtResumo(b)
      MtResumo(b).baseCalculo = Mt70(a).Base_Calculo_Icms
      MtResumo(b).CFOP = Mt70(a).CFOP
      MtResumo(b).valorIcms = Mt70(a).Valor_Icms
      MtResumo(b).ValorIsentas = Mt70(a).Isenta_Nao_Tributada
      MtResumo(b).ValorOutras = Mt70(a).Outras
      MtResumo(b).ValorTotal = Mt70(a).Valor_Total
      b = b + 1
   Else
      Achou = False
      For C = 0 To UBound(MtResumo)
          If MtResumo(C).CFOP = Mt70(a).CFOP Then
             Achou = True
             Exit For
          End If
      Next
      If Achou Then
        
        MtResumo(C).baseCalculo = MtResumo(C).baseCalculo + Mt70(a).Base_Calculo_Icms
        MtResumo(C).valorIcms = MtResumo(C).valorIcms + Mt70(a).Valor_Icms
        MtResumo(C).ValorIsentas = MtResumo(C).ValorIsentas + Mt70(a).Isenta_Nao_Tributada
        MtResumo(C).ValorOutras = MtResumo(C).ValorOutras + Mt70(a).Outras
        MtResumo(C).ValorTotal = MtResumo(C).ValorTotal + Mt70(a).Valor_Total
      Else
        ReDim Preserve MtResumo(b)
        MtResumo(b).baseCalculo = Mt70(a).Base_Calculo_Icms
        MtResumo(b).CFOP = Mt70(a).CFOP
        MtResumo(b).valorIcms = Mt70(a).Valor_Icms
        MtResumo(b).ValorIsentas = Mt70(a).Isenta_Nao_Tributada
        MtResumo(b).ValorOutras = Mt70(a).Outras
        MtResumo(b).ValorTotal = Mt70(a).Valor_Total
        b = b + 1
      End If
   End If
Next

With Sintegra.ResumoCfop
    For a = 0 To UBound(MtResumo)
       d = .Rows
       If Len(MtResumo(a).CFOP) > 0 Then
       .Rows = d + 1
       .TextMatrix(d, 0) = MtResumo(a).CFOP
       .TextMatrix(d, 1) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorTotal), 2)
       .TextMatrix(d, 2) = ClUtil.AcertaNumero(CStr(MtResumo(a).baseCalculo), 2)
       .TextMatrix(d, 3) = ClUtil.AcertaNumero(CStr(MtResumo(a).valorIcms), 2)
       .TextMatrix(d, 4) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorIsentas), 2)
       .TextMatrix(d, 5) = ClUtil.AcertaNumero(CStr(MtResumo(a).ValorOutras), 2)
       If Mid(MtResumo(a).CFOP, 1, 1) = "1" Or Mid(MtResumo(a).CFOP, 1, 1) = "2" Then
          Valor_Entrada = Valor_Entrada + MtResumo(a).ValorTotal
          Base_entrada = Base_entrada + MtResumo(a).baseCalculo
          Icms_entrada = Icms_entrada + MtResumo(a).valorIcms
       Else
          Valor_Saida = Valor_Saida + MtResumo(a).ValorTotal
          Base_Saida = Base_Saida + MtResumo(a).baseCalculo
          Icms_saida = Icms_saida + MtResumo(a).valorIcms
       End If
    End If
   Next
End With


Sintegra.Resumo.TextMatrix(3, 0) = "Registro 70 - Entrada"
Sintegra.Resumo.TextMatrix(3, 1) = ClUtil.AcertaNumero(CStr(Valor_Entrada), 2)
Sintegra.Resumo.TextMatrix(3, 2) = ClUtil.AcertaNumero(CStr(Base_entrada), 2)
Sintegra.Resumo.TextMatrix(3, 3) = ClUtil.AcertaNumero(CStr(Icms_entrada), 2)

'==> Saidas
Sintegra.Resumo.TextMatrix(4, 0) = "Registro 70 - Saida"
Sintegra.Resumo.TextMatrix(4, 1) = ClUtil.AcertaNumero(CStr(Valor_Saida), 2)
Sintegra.Resumo.TextMatrix(4, 2) = ClUtil.AcertaNumero(CStr(Base_Saida), 2)
Sintegra.Resumo.TextMatrix(4, 3) = ClUtil.AcertaNumero(CStr(Icms_saida), 2)
End Sub


