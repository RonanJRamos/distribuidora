VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lidis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mvarLocalArmazenar As Variant 'local copy
Private ErrosP As String
Private NomeArquivo As String


Public Property Get erroS() As Variant
      erroS = ErrosP
    
End Property
Public Property Let LocalArmazenar(ByVal vData As Variant)
    mvarLocalArmazenar = vData
End Property


Public Property Set LocalArmazenar(ByVal vData As Variant)
    Set mvarLocalArmazenar = vData
End Property


Public Property Get LocalArmazenar() As Variant
    If IsObject(mvarLocalArmazenar) Then
        Set LocalArmazenar = mvarLocalArmazenar
    Else
        LocalArmazenar = mvarLocalArmazenar
    End If
End Property
Private Sub NomedoArquivo()
Dim Arquivo As String
Dim LocalPdv As String
Dim NomeSintegra As String

Dim a As Integer
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then
       Exit For
    End If
Next
Arquivo = Mid(GLBase, 1, a)
Arquivo = Arquivo & "Configuracaopdv.txt"
LcLocalSintegra = LeIni("Pdv", "Sintegra", Arquivo)
'MsgBox GLBase
If Right(LcLocalSintegra, 1) <> "\" Then LcLocalSintegra = LcLocalSintegra & "\"
NomeSintegra = LcLocalSintegra & "Sintegra" & Format(Sintegra.DataF.Text, "ddmmyyyy") & ".txt"

NomeArquivo = NomeSintegra
End Sub
Public Sub ProcessarSintegra()
    Dim ClGerarSintegra As New ClSintegra
    Dim Rs As ADODB.Recordset
    Dim StrSql As String
    '==> Limpa as matrizes
    ReDim Mt50(0)
    ReDim Mt54(0)
    ReDim Mt75(0)
    ReDim Mt70(0)
    ReDim Mt74(0)
    ReDim Mt53(0)
    TemRegistro50 = False
    TemRegistro54 = False
    TemRegistro75 = False
    TemRegistro70 = False
    TotalReg75 = 0
    TotalReg50 = 0
    TotalReg54 = 0
    TotalReg60 = 0
    TotalReg70 = 0
    Dim DataInicio As Date
    '==> Busca a entrada
    ClGerarSintegra.Cliente_Origem = "Lidis"
    ClGerarSintegra.ProcessaEntradaLidis CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
    ClGerarSintegra.BuscaRegistro53Lidis CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
    '==> Busca Saida
    ClGerarSintegra.ProcessaSaidaLidis CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
    If Sintegra.Inventario.Value = 1 Then
        
       If Sintegra.OptApartir.Value Then DataInicio = CDate(Sintegra.DataInventario.Text)
       If Sintegra.OPtSintegraAnterior.Value Then
          StrSql = "Select * from inventariosintegra LIMIT 1"
          Set Rs = AbreRecordset(StrSql, True)
          If Not Rs.EOF Then
             DataInicio = CDate(Rs!Data + 1)
          Else
             DataInicio = CDate(Sintegra.DataInventario.Text)
          End If
       End If
       ClGerarSintegra.BuscaRegistro74Lidis DataInicio, CDate(Sintegra.DataF.Text)
    End If
       '==> Busca Frete
    'ClGerarSintegra.ProcessaConhecimentoFreteBhForte CDate(Sintegra.Datai.Text), CDate(Sintegra.Dataf.Text)
    '==> Escreve sintegra
   ' ClGerarSintegra.EscreverSintegra LocalArmazenar
     
     
End Sub
Public Sub BuscaDados()
ErrosP = ""
   BuscaEntrada CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
   BuscaSaida
  ' ErrosP = ClGerarSintegra.Erros
    
End Sub
Public Sub EscreverSintegra()
Dim ClEscrever As New EscreverTexto
Dim LcArq As Integer
NomedoArquivo
LcArq = FreeFile
If Dir(NomeArquivo, vbArchive) <> "" Then Kill NomeArquivo
Sintegra.Caption = "Gravando gerando dados do sintegra."

Open NomeArquivo For Output As #LcArq

ClEscrever.NumeroArquivo = LcArq

EscreverRegistro10 CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text), LcArq
EscreverRegistro11 LcArq
ClEscrever.EscreverRegistro50
ClEscrever.EscreverRegistro53
ClEscrever.EscreverRegistro54
If Sintegra.Inventario.Value = 1 Then ClEscrever.EscreveRegistro74
ClEscrever.EscreverRegistro75 CDate(Sintegra.DataI.Text), CDate(Sintegra.DataF.Text)
ClEscrever.EscreverRegistro90

Close #LcArq
Set ClEscrever = Nothing
MsgBox NomeArquivo
End Sub
Public Sub GerarResumo()
  Dim ClGerarSintegra As New ClSintegra
  
  ClGerarSintegra.GerarResumo
End Sub

Sub BuscaSaida()
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
Dim db As Database
Dim RsDado As Recordset
Dim StrSql As String
Dim ClUtil As New Utilitario

StrSql = "SELECT * FROM alid050 where dtemis Between #" & Format(Sintegra.DataI.Text, "mm/dd/yy") & "# And #" & Format(Sintegra.DataF.Text, "mm/dd/yy") & "#"
Set db = OpenDatabase(GLBase)
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
  Sintegra.Caption = "Buscando a NF de Saida  " & x & " de " & LcTotal & " Nº:" & Rs!numnf & ""
  x = x + 1
  DoEvents
  '==> verifica se é fornecedor ou cliente
 ' If Rs!nf = "001464" Then Stop
  StrSql = "Select * from alid001 where codigo='" & Rs!Cliente & "'"
  ECliente = True
  
  Set RsDado = db.OpenRecordset(StrSql)
  If RsDado.EOF Then
     Set RsDado = Nothing
     StrSql = "Select * from alid001 where codigo='" & Right("00000" & Rs!Cliente, 5) & "'"
     Set RsDado = db.OpenRecordset(StrSql)
  End If
  If Not RsDado.EOF Then
     ClUtil.VerificaCnpj RsDado!razaosoc, CLng(Rs!Cliente), ECliente, "LIDIS"
  Else
     ErrosP = ErrosP & Chr(13) & "Cliente da nota saida nº:" & Rs!numnf & " não foi encontrado."
  End If
  Set RsNatureza = Nothing
  a = Sintegra.Saida.Rows
  Sintegra.Saida.Rows = a + 1
  Sintegra.Saida.TextMatrix(a, 0) = Rs!numnf & ""
  Sintegra.Saida.TextMatrix(a, 1) = Format(Rs!DtEmis, "dd/mm/yy") & ""
  Sintegra.Saida.TextMatrix(a, 2) = IIf(Not RsDado.EOF, RsDado!razaosoc, "") & ""
  Sintegra.Saida.TextMatrix(a, 3) = Rs!ValorNota & ""
  Sintegra.Saida.TextMatrix(a, 4) = Rs!valorIcms & ""
  Sintegra.Saida.TextMatrix(a, 5) = ClUtil.AcertaNumero(CStr(Rs!valorIcms), 2) & ""
  
  Sintegra.Saida.TextMatrix(a, 6) = Rs!Cliente & ""
  Sintegra.Saida.TextMatrix(a, 7) = LcModelo & ""
  Sintegra.Saida.TextMatrix(a, 8) = "1"
  If Not IsNull(Rs!CFOP) Then
     Sintegra.Saida.TextMatrix(a, 9) = Replace(Rs!CFOP, ".", "") & ""  '==> Codigo Fiscal
  Else
     Sintegra.Saida.TextMatrix(a, 9) = "5405"
  End If

  Sintegra.Saida.TextMatrix(a, 10) = IIf(UCase(Rs!Status) = UCase("Cancelada"), "S", "N")
  Sintegra.Saida.TextMatrix(a, 11) = Rs!numnf
  Sintegra.Saida.TextMatrix(a, 12) = IIf(Not IsNull(Rs!Desconto), Rs!Desconto, 0)
  Sintegra.Saida.TextMatrix(a, 13) = 0
  Sintegra.Saida.TextMatrix(a, 14) = 0
  Valor_Total = Valor_Total + Rs!ValorNota
  Valor_Icms = Valor_Icms + CDbl(Sintegra.Saida.TextMatrix(a, 5))
  Rs.MoveNext
  Set RsDados = Nothing
Loop
Set Rs = Nothing
Sintegra.TotalSaida.Text = ClUtil.AcertaNumero(CStr(Valor_Total), 2)
Sintegra.IcmsSaida.Text = ClUtil.AcertaNumero(CStr(Valor_Icms), 2)
ErrosP = ErrosP & ClUtil.erroS

Exit Sub
ErrBuscaSaida:
MsgBox err.Description & err.Number
Resume 0
End Sub

Public Function BuscaEntrada(DataInicio As Date, DataFinal As Date)
Dim StrSql As String
Dim CSint As New ClSintegra
'==>processa a Entrada
'CSint.erroS = ""
CSint.Cliente_Origem = "LIDIS"
StrSql = "SELECT EntradaNf.ValorProduto,EntradaNf.NF,EntradaNf.CLICRED, EntradaNf.CLICRED as codigofor, EntradaNf.VALOR, EntradaNf.DATA AS entrada, EntradaNf.codigo, EntradaNf.Icms, EntradaNf.Ipi, EntradaNf.Complementar as Complemento,EntradaNf.Complementar as custos, EntradaNf.Serie, EntradaNf.cfop, EntradaNf.cfop as CodigoFiscal, EntradaNf.Frete, EntradaNf.Seguro, EntradaNf.PIS_COFINS, " & _
         "EntradaNf.NaoTributado, EntradaNf.DespesasAcessorias, EntradaNf.IcmsSubst as ValorIcmsSubstituicao, EntradaNf.SubSerie, EntradaNf.TipoFrete,'01' AS modelo,'Confirmado' as Status " & _
         "FROM EntradaNf " & _
         "where data Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFinal, "mm/dd/yy") & "# order by data"

CSint.BuscaEntrada StrSql
'==>Processa a saida

ErrosP = ErrosP & CSint.erroS

End Function
Public Sub EscreverRegistro10(DataI As Date, DataF As Date, LcArq As Integer)
Dim Rs As Recordset
Dim StrSql      As String
Dim StrResposta As String
Dim Cnpj        As String
Dim Inscricao   As String
Dim Fax         As String
Dim Fone        As String
Dim db          As Database
Set db = OpenDatabase(GLBase)

StrSql = "Select * from empresa"
Set Rs = db.OpenRecordset(StrSql)
Sintegra.Caption = "Criando o Registro 10 - Dados da empresa."
DoEvents
If Not Rs.EOF Then
   StrResposta = "10"
   '==> Monta o cgc
   Cnpj = Rs!CGC
   Cnpj = Replace(Cnpj, ".", "")
   Cnpj = Replace(Cnpj, "-", "")
   Cnpj = Replace(Cnpj, "/", "")
   Cnpj = Replace(Cnpj, "\", "")
   Cnpj = Replace(Cnpj, " ", "")
   Cnpj = Right("00000000000000" & Cnpj, 14)
   StrResposta = StrResposta & Cnpj
   
   '==> Inscricao
   Inscricao = Rs!inscricaoestadual & ""
   Inscricao = Replace(Inscricao, ".", "")
   Inscricao = Replace(Inscricao, "-", "")
   Inscricao = Replace(Inscricao, "/", "")
   Inscricao = Replace(Inscricao, "\", "")
   Inscricao = Replace(Inscricao, " ", "")
   Inscricao = Left(Inscricao & "              ", 14)
   StrResposta = StrResposta & Inscricao
  '==> Razao Social
   StrResposta = StrResposta & Left(Rs!Razao & String(35, " "), 35)
  '==> Municipio
   StrResposta = StrResposta & Left(Rs!Cidade & String(30, " "), 30)
  '==> Uf
   StrResposta = StrResposta & Left(Rs!Estado & String(2, " "), 2)
  '==> fax
   Fax = Rs!Fax & ""
   Fax = Replace(Fax, ".", "")
   Fax = Replace(Fax, "-", "")
   Fax = Replace(Fax, "/", "")
   Fax = Replace(Fax, "\", "")
   Fax = Replace(Fax, ")", "")
   Fax = Replace(Fax, "(", "")
   Fax = Replace(Fax, " ", "")
   Fax = Right("00000000000000" & Fax, 10)
   StrResposta = StrResposta & Fax
  '==>Data inicial
   StrResposta = StrResposta & Format(DataI, "yyyymmdd")
  '==>Data Final
   StrResposta = StrResposta & Format(DataF, "yyyymmdd")
   '==> Codigo do Convenio
   StrResposta = StrResposta & Rs!CodigoConvenio & ""
   '==> Codigo da identificacao da natureza infomada
   StrResposta = StrResposta & Rs!CodigoNaturezaInformacao & ""
    '==> Codigo da FINALIDADES DA APRESENTAÇÃO DO ARQUIVO MAGNÉTICO
   StrResposta = StrResposta & Left(Sintegra.Finalidade.Text, 1) & ""
   
End If
Set Rs = Nothing
Print #LcArq, StrResposta


End Sub

Public Sub EscreverRegistro11(LcArq As Integer)
Dim Rs          As Recordset
Dim StrSql      As String
Dim StrResposta As String
Dim Cnpj        As String
Dim Inscricao   As String
Dim Fax         As String
Dim Fone        As String
Dim Cep         As String

Dim db          As Database
Set db = OpenDatabase(GLBase)

StrSql = "Select * from empresa"
Set Rs = db.OpenRecordset("Select * from empresa")
Sintegra.Caption = "Criando o Registro 11 - complemento dos dados da empresa."
DoEvents
If Not Rs.EOF Then
   StrResposta = "11"
   '==> Logradouro
   StrResposta = StrResposta & Left(Rs!Endereco & String(34, " "), 34)
   '==> Numero
   StrResposta = StrResposta & Right("00000" & Rs!Numero, 5)
   '==> Complemento
   StrResposta = StrResposta & Left(Rs!Complemento & String(22, " "), 22)
   '==> Bairro
   StrResposta = StrResposta & Left(Rs!Bairro & String(15, " "), 15)
   '==> Cep
   Cep = Rs!Cep & ""
   Cep = Replace(Cep, ".", "")
   Cep = Replace(Cep, "-", "")
   Cep = Replace(Cep, "/", "")
   Cep = Replace(Cep, "\", "")
   Cep = Replace(Cep, " ", "")
   Cep = Right("00000000" & Cep, 8)
   StrResposta = StrResposta & Cep
  '==> Responsavel
   StrResposta = StrResposta & Left(Rs!Responsavel & String(28, " "), 28)
  '==> Fone
   Fone = Rs!Fone & ""
   Fone = Replace(Fone, ".", "")
   Fone = Replace(Fone, "-", "")
   Fone = Replace(Fone, "/", "")
   Fone = Replace(Fone, "\", "")
   Fone = Replace(Fone, ")", "")
   Fone = Replace(Fone, "(", "")
   Fone = Replace(Fone, " ", "")
   Fone = Right("0000000000000000" & Fone, 12)
   StrResposta = StrResposta & Fone
End If
Print #LcArq, StrResposta
Set Rs = Nothing
End Sub
