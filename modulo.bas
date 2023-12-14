Attribute VB_Name = "Module1"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileKeys% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal Section$, ByVal Zero&, ByVal Default$, ByVal ReturnBuffer$, ByVal LenReturnBuffer%, ByVal FileName$)
Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileDelKey% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section$, ByVal Entry As Any, ByVal Zero&, ByVal FileName$)
Declare Function WritePrivateProfileDelSect% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section$, ByVal Zero&, ByVal EmptyStr$, ByVal FileName$)
Declare Function WritePrivateProfileStringA% Lib "kernel32" (ByVal Section$, ByVal Entry As Any, ByVal CharStr As Any, ByVal FileName$)


Public Declare Function Extenso Lib "Extens32.dll" Alias "extenso" (ByVal Valor As String, ByVal Retorno As String) As Integer
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function ConsisteInscricaoEstadual Lib "DllInscE32.dll" (ByVal Insc As String, ByVal UF As String) As Integer

'Option Explicit
Type pesquisa
     campo As String
     Indice As String
     Tipo As String
End Type
Type Especies
     codigo As Long
     Descricao As String
      End Type
Type Dados
     Item As String
     codigo As String
     produto As String
     Valor As Double
     Desconto As Currency
     Devolucao As Date
     Brinde As Integer
     Lancamento As Integer
End Type
Public Type DadosEntrada
     Item As String
     CodPro As String
     produto As String
     Und As String
     Qut As Double
     VUnit As Currency
     Vtotal As Currency
     Venda1 As Currency
     Venda2 As Currency
     Venda3 As Currency
     ipi As Double
     cst As String
     precomim As Currency
     icms As Double
     Com As Long
     almox As String
     santamaria As Single
     santamaria1 As Single
     california As Single
     Bloqueado As Boolean
     tipoliberacao As Integer
     jaEsteveBloqueado As Boolean
     MaquinaLiberacao As String
     DataLiberacao As Date
     HoraLiberacao As String
     usuario    As String
     QuanTidadeBaixa As Double
     NumeroVale  As String
     ItemBaixo As Boolean
     CFOP As String
     NCM As String
End Type
Public Type DadosEtiqueta
       codigo As String
       Nome As String
       Endereco As String
       Bairro As String
       Cidade As String
       Cep As String
       UF As String
       Imprime As Integer
 End Type
 Public LcSq As String
Public GlVerificarIcmsDiferenciado As Boolean
Public GlImprimeValorCFOPNota As Boolean
Public GlNaoVerificaEstoque As Boolean
Public GlExibirLucratividade As Boolean
Public GLInformacaoNF     As String
Public GLAproveitamentoICMS As Boolean
Public GlBoletoCEF As Boolean
Public GlIntrucao1, GlIntrucao2, GlIntrucao3 As String
Public GlLinhasSaltarInicio, GlLinhasSaltarBFim As Double
Public GLNImprimeBaseC    As Boolean
Public GlContrato         As Boolean
Public GlGeraPorItem As Boolean
Public GlboletoA4 As Integer
Public GlLiberaIcms As Boolean
Public GlSaidaIcms As Boolean
Public LcSql        As String
Public GlLiberaSenhaAlteraPr As Boolean
Public FrmExibeSenha As Boolean
Public Glmargemnota As Integer
Public GlSenhaLiberacao, GlRec As String
Public LcPerguntaVendedor, LcPerguntaCliente, GlBuscaProdutoAgora As Integer
Public GlComissaoVelha As Currency
Public MtEtiqueta() As DadosEtiqueta
Public MtProduto() As Dados
Public acesso As New Collection
Public PrecoVendaNormal, PrecoVendaAlterado, PrecoMinimodeVenda, PrecoMimimodeVendaAlterado As Currency
Public GlCodigoProtecao As Long
Public Msg As String
Public MtPesquisa(31) As pesquisa
Public MtEspecie() As Especies
Public MtCliente() As Especies
Public GlAcrescimo, GlDescontos, GlaxaReb As Currency
Public GlBoleto, GlNota As String
Public GlUsuario As String
Public GlGrupo, GlFuncao, GlStringBase, GlordemAnterior As String
Public SacolasDevolvidas, GlMargem, LcTamanhoEtiqueta As Long
Public GlUtilizado, GlCredito As Currency
Public GlDataAtual, GlDataSistema As Date
Public GlFormAtual As Tabela
Public GlImprimeSemLinha As Integer
Public GLPesquisa, GlNreg, GlPergunta, GlEscolhe, GlArmazenaGalpao As Integer
Public GlConfirmaExclusao, GlConfirmaAlteracao, GLConfirmaNovo As Integer
Public GlLucroCad, GlLucroAlteracao, GlMinimoAlteracao, GlImprimeSpool, GlServidorImpressora As Integer
Public GlRepresentante, GlComercio, GlDadosTransportadora, GlEsclheVendedor, GlEscolheCliente As Integer
Public GLTamanhoEspecie, GlTamanhoCliente, GlLibera, GLCalculacodigoProduto As Integer
Public GLCalculacodigoCliente, GLCalculacodigoFornecedor As Integer
Public GlFormA As Form
Public Area As Workspace
Public Dbbase As Database, RsLogDisponivel As Recordset
Public RsAtualP As ADODB.Recordset
Public RsAtual As Recordset
Public RsFuncao As Recordset
Public GlPortaNota, GlPortaBoleto, GlPortaOrcamento As String
Public Gl40colunas, GlSaltoLinhasOrcamento As Integer
Public GlImpressao, GLSaltoLinhaNota, LcTamanho, GlLocado, GlDiaReserva, GlDiasDevolucao, GlDiasDevolucaoReserva As Integer
Public GlCalculaProduto, LcCodigoConvenio, GlVariasComissao As Integer
Public GLtipoDevolucao, GlEscolha, GlIniceAtual As Integer
Public LcResposta, LcTipoDados, LcRegAtual, LcTamaanho As Integer
Public GlFaturaSaida, GlVistaSaida, GlCaixaSaida, GlFaturaEntrada, GlVistaEntrada, GlCaixaEntrada As Integer
Public GlIpi, GlRateiaAcrecimo, GlDetalhaDesconto, GlImprimeDetalhaDesconto As Integer
Public GlMsg1 As String
'Public GlMsg2 As String
Public GlDescUnit, GLPadraoWindows, GlServidorImpressoraOrc, GlImprimeSpoolOrc As Integer
Public GlDecimais As Long
Public GlFuncCodigo, GlFuncNome, GlFuncEmpresa As Integer
Public GlCap, GlOpcaoEmpresa As String
Public GlBook As String, GlCodigoAnterior As String
Public GLBase As String
Public GlForm As String
Public GlMov, GlAlteraCodigo As Integer
Public GlCampo0 As String
Public GlCampo1 As String
Public GlCampo2 As String
Public GlCampo3 As String
Public GlCampo4 As String
Public GlCampo5 As String
Public GlCampo6 As String
Public GlCampo7 As String
Public GlCampo8 As String
Public GlCampo9 As String
Public GlCampo10 As String
Public GlCampo11 As String
Public GlCampo12 As String
Public GlCampo13 As String
Public GlCampo14 As String
Public GlCampo15 As String
Public GlCampo16 As String
Public GlCampo17 As String
Public GlCampo18 As String
Public GlCampo19 As String
Public GlCampo20 As String
Public GlCampo21 As String
Public GlCampo22 As String
Public GlCampo23 As String
Public GlCampo24 As String
Public GlCampo25 As String
Public GlCampo26 As String
Public GlCampo27 As String
Public GlCampo28 As String
Public GlCampo29 As String
Public GlCampo30 As String

Public GlCampo31 As String
Public GlCampo32 As String
Public GlCampo33 As String
Public GlCampo34 As String
Public GlCampo35 As String
Public GlCampo36 As String
Public GlCampo37 As String
Public GlCampo38 As String
Public GlCampo39 As String
Public GlCampo40 As String
Public GlCampo41 As String
Public GlCampo42 As String
Public GlCampo43 As String
Public GlCampo44 As String
Public GlCampo45 As String
Public GlCampo46 As String
Public GlCampo47 As String
Public GlCampo48 As String
Public GlCampo49 As String
Public GlCampo50 As String
Public GlCampo51 As String
Public GlCampo52 As String
Public GlCampo53 As String
Public GlCampo54 As String
Public GlCampo55 As String
Public GlCampo56 As String
Public GlCampo57 As String
Public GlCampo58 As String
Public GlCampo59 As String
Public GlCampo60 As String

Public GlCampo61 As String
Public GlCampo62 As String
Public GlCampo63 As String
Public GlCampo64 As String
Public GlCampo65 As String
Public GlCampo66 As String
Public GlCampo67 As String
Public GlCampo68 As String
Public GlCampo69 As String
Public GlCampo70 As String
Public GlCampo71 As String
Public GlCampo72 As String
Public GlCampo73 As String
Public GlCampo74 As String
Public GlCampo75 As String
Public GlCampo76 As String
Public GlCampo77 As String
Public GlCampo78 As String
Public GlCampo79 As String
Public GlCampo80 As String
Public GlCampo81 As String
Public GlCampo82                As String
Public GlCampo83                As String
Public GlCampo84                As String
Public GlCampo85                As String
Public GlCampo86                As String
Public GlCampo87                As String
Public GlCampo88                As String
Public GlCampo89                As String
Public GlSenhaCredito           As String
Public GlSenhaDebito            As String
Public GlLiberaPedidoVendas     As String
Public GlCodigoProduto          As String
Public GlRestricao              As String
Public GlNomeMaquina            As String
Public GlCliDeb                 As String

Public GlChave                  As Variant
Public GlEstoqueDisponivel      As String
Public GlNumeroLocacao          As Long
Public GlFormInicial            As String
Public NomeBanco                As String
Public LcCriterio               As String
Public LcIndice                 As String
Public GlCriterioSql            As String
Public GlMsg, GlMsg2, GlMsg3    As String
Public GlTab                    As String
Public GlSq                     As String
Public LcAlterado               As Integer
Public GlErroProt               As Integer
Public LcNovo                   As Integer
Public GlCarregado              As Integer
Public GlLiberaVenda            As Integer
Public GlInclusaoReceita        As Boolean
Public GlBaixaReceita           As Boolean
Public GlVendaVista             As Boolean
Public GlVendaPrazo             As Boolean
Public GlInclusaoDespesa        As Boolean
Public GlBaixaDespesa           As Boolean
Public GlEntradaVista           As Boolean
Public GlEntradaPrazo           As Boolean
Public GlInclusaoCheque         As Boolean
Public GlBaixaCheque            As Boolean
Public GlRecuperou              As Boolean
Public GlLiberaAtraso           As Boolean
Public GlSaiuAtraso             As Boolean
Public GlAtualizaPreco          As Boolean
Public GlImplanta               As Boolean
Public GlLimpaTelaOrc           As Boolean
Public GlSisCarregado           As Boolean
Public GlMeiaFolha              As Boolean

Public GlItensMeiaFolha         As Integer
Public GlMargemMeiaFolha        As Integer
Public GlSaltoFinalMeiaFolha    As Integer
Public GlImpressaoMeiaFolha     As Integer
Public GlCabecalhoMeiaFolha     As Boolean
Public GlUsaEstoqueSeguranca    As Boolean
Public GlBaixarEstoquenoPedido  As Boolean
Public GlMostraMsgClientePedido As Boolean

Public Const FundoAlterado = &HFFFFC0
Public Const FundoMornal = &HFFFFFF
Public GlPortaMala As String
Public GlPuloFim, GlExibeComissao, GlAchaBase As Integer
Public GlSenha, GlSoOrcamento, GlNaoBloqueia As Integer
Public LigaTitulo, DesligaTitulo, LigaNegrito, DesligaNegrito, LigaDraft As String
Public LcQSanta As Double
Public LcQUnSanta As Double
Public LcQUnSanta1 As Double
Public LcQUnCalifornia As Double
Public LcQUnSantas As Double
Public PesquisandoNota As Boolean
Public GlPermitirVendaEstoqueNegativo As Boolean
'Public a As Integer

Public Enum Movimentos
  enPrimeiro = 1
  enAnterior = 2
  
  enSeguinte = 3
  enultimo = 4
End Enum

Public Enum Tabela
        Cliente
        fornecedor
        produto
        Especie
        produtora
        Convenio
        Funcionario
        Receber
        pagar
        opcao
        usuario
        GrpUsuario
        Transportadora
        EntradaProduto
        Galpao
        Cidade
        monetario
        tiporec
        Unidade
        pedido
        Caixa
        Cheques
        PropostaCliente
        Custo
End Enum
Public Csintegra As New Lidis
Public GlNomeProjeto As String
Public Const GlSistemaImplementado As String = "Lidis"
Public g_InteropToolbox As InteropToolbox
Public GlPodeAbrirOBS As Boolean

Public Sub Transforna_Unidade(Lc_Quantidade As Currency, Com As Integer, ByRef QuantCaixa As Currency, ByRef QuantUnidade As Currency)
Dim LcCaixa As Currency
Dim LcQuantUnidade As Currency
Dim LcTemp As Currency
LcTemp = Lc_Quantidade / Com
QuantCaixa = Fix(LcTemp)
LcQuantUnidade = LcTemp - QuantCaixa
QuantUnidade = LcQuantUnidade * Com

End Sub

Function Consiste(Insc As String, UF As String) As Integer
  Consiste = ConsisteInscricaoEstadual(Insc, UF)
End Function

Function acertaestoque(LcProduto As String)
Dim Rsp As Recordset
Dim RsG As Recordset
Dim RsU As Recordset
Dim db  As Database
Dim LcSql As String
Dim LcSql1 As String
Dim LcTotalG    As Double
Dim LcTotalU    As Double
Dim LcUn        As String
Dim LcQunUn     As String
Dim LcCust      As Double

LcSql = "Select * from alid009 where cod like '" & LcProduto & "*'"

Set db = OpenDatabase(GLBase)
Set Rsp = db.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsU = db.OpenRecordset("Select * from alid004", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until Rsp.EOF
    LcSql1 = "Select * From alid013 where ITEM='" & Rsp!cod & "'"
    Set RsG = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
    LcTotalG = 0
    LcTotalU = 0
    Do Until RsG.EOF
       LcTotalG = LcTotalG + RsG!Estoque
       LcTotalU = LcTotalU + RsG!QuantUnidade
       RsG.MoveNext
    Loop
    Rsp.Edit
    Rsp("QuantEstoque") = LcTotalG
    Rsp("quantUnidade") = LcTotalU
    Rsp.Update
    LcPes = "codigo=" & Rsp!UNIMED & "'"
    RsU.FindFirst LcPes
    If Not RsU.NoMatch Then
       LcUn = RsU!Simbolo
    Else
       LcUn = ""
    End If
    If Not IsNull(Rsp!QTDUNIMED) Then
       LcQunUn = Rsp!QTDUNIMED
    Else
       LcQunUn = 0
    End If
    If Not IsNull(Rsp!Custo) Then
       LcCust = Rsp!Custo
    Else
       LcCust = 0
    End If
    If Not IsNull(Rsp!cod) Then
       If Not IsNull(Rsp!Nome) Then
          If GlImplanta Then
            
             LcQUnSanta = LcTotalU
             LcQSanta = LcTotalG
             Call Ficha("IMPLAT", Rsp!cod, Rsp!Nome, LcTotalG, LcCust, LcCust * LcTotalG, "E", "IMPLANTACAO", LcUn, LcQunUn)
          End If
       End If
    End If
    Rsp.MoveNext
    RsG.Close
    Set RsG = Nothing
Loop
Rsp.Close
RsU.Close
db.Close
Set Rsp = Nothing
Set RsU = Nothing
Set db = Nothing

End Function

Function ExibeCpf(LcCpf1 As String) As String
On Error Resume Next
Dim a As Integer
Dim LcCpf As String
LcCpf = ""
For a = 1 To Len(LcCpf1)
    If IsNumeric(Mid(LcCpf1, a, 1)) Then
       LcCpf = LcCpf & Mid(LcCpf1, a, 1)
    End If
Next
If Len(LcCpf) > 0 Then
   LcCpf = Mid(LcCpf, 1, 3) & "." & Mid(LcCpf, 4, 3) & "." & Mid(LcCpf, 7, 3) & "-" & Mid(LcCpf, 10)
Else
  LcCpf = "   .   .   -  "
End If
ExibeCpf = LcCpf
End Function

Function VerificaAtraso(LcCli As String) As Boolean
Dim Lca         As String
Dim bb          As Database
Dim RsConta     As ADODB.Recordset
Dim rsCliente   As Recordset
'abreconexao

Lca = "select * from alid015 where cliente='" & LcCli & "' and dtvenc <'" & Format(Date, "yyyy-mm-dd") & "' and valpago=0"
Set bb = OpenDatabase(GLBase, False, False)
Set RsConta = AbreRecordset(Lca)
Set rsCliente = bb.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsConta.EOF Then
   VerificaAtraso = False
Else
   LcPes = "codigo='" & LcCli & "'"
   rsCliente.FindFirst LcPes
   If Not rsCliente.NoMatch Then
      GlCliDeb = rsCliente!razaosoc
   Else
     GlCliDeb = ""
   End If
   GlLiberaAtraso = False
   GlSaiuAtraso = False
   AvisoDebito.Show , FrmPrincipal
   Do While Not GlSaiuAtraso
      DoEvents
   Loop
   VerificaAtraso = Not GlLiberaAtraso
End If

RsConta.Close
rsCliente.Close
bb.Close

Set RsConta = Nothing
Set rsCliente = Nothing
Set bb = Nothing
End Function

Function Zeracontas()
On Error Resume Next
Dim Lca     As String
Dim RsReceber   As ADODB.Recordset
Dim RsPagar     As Recordset

Set bb = OpenDatabase(GLBase, False, False)
Set RsPagar = bb.OpenRecordset("alid014", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsReceber = AbreRecordset("select * from alid015", RsReceber)
'abreconexao
LcSql = "Update alid015 SET VALPAGO = 0 where isnull(VALPAGO)"
LcRegistrosAfetados = ExecutaSql(LcSql)


err.Number = 0
'Do Until RsReceber.EOF
'   If err.Number > 0 Then Exit Do
'   If IsNull(RsReceber!VALPAGO) Then
'      RsReceber.Edit
'      RsReceber!VALPAGO = 0
'      RsReceber.Update
'   End If
'   RsReceber.MoveNext
'Loop
err.Number = 0
Do Until RsPagar.EOF
   If err.Number > 0 Then Exit Do
   If IsNull(RsPagar!VALPAGO) Then
      RsPagar.Edit
      RsPagar!VALPAGO = 0
      RsPagar.Update
   End If
   RsPagar.MoveNext
Loop
RsPagar.Close
'RsReceber.Close

Set RsPagar = Nothing
'Set RsReceber = Nothing
'FechaConexao
   

End Function
Function BuscaDirWin() As String
Dim LcDirWindows    As String
Dim LcCaracter      As String
Dim GlDirWinSystem  As String
Dim retValue        As Long
Dim i               As Integer
Dim GlBuffer        As Integer
Dim GlDevApi        As Integer
GlBuffer = 255
LcDirWindows = String(255, " ")

GlDevApi = GetWindowsDirectory(LcDirWindows, GlBuffer)

'=== Esta Sequencia Tem por Finalidade, Separar somente o
'=== Nome do Diretorio Windows, Separando os Caracters Indesejados
'=== Os Caracteres ascII de 47 a 126 são as Letras Validas

For i = 1 To 255
       LcCaracter = Mid(LcDirWindows, i, 1)
       If Asc(LcCaracter) >= 47 And Asc(LcCaracter) <= 126 Then
          GlDirWinSystem = GlDirWinSystem & LcCaracter
       Else
          Exit For
       End If
Next i
BuscaDirWin = GlDirWinSystem
End Function


Function logOrcamento()
On Error GoTo ErrLog
Dim LcNota As String, Nfi As String
Dim DbbaseOrc As Database
Dim RsImpressora As Recordset, RsLogNota As Recordset
Dim FnunNota, LcAbriu As Integer

If Not GlServidorImpressoraOrc Then Exit Function
GlImprimeSpoolOrc = True
Set DbbaseOrc = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsImpressora = DbbaseOrc.OpenRecordset("select * from impressoras where Impressora='" & GlPortaOrcamento & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If UCase(RsImpressora!Maquina) = "LOCAL" Then
  Set RsLogNota = DbbaseOrc.OpenRecordset("select * from LogImpressaoOrcamento where Maquina='" & "LOCAL" & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Else
  Set RsLogNota = DbbaseOrc.OpenRecordset("select * from LogImpressaoOrcamento where Maquina='" & GlNomeMaquina & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
End If
If Not RsImpressora.EOF Then
   LcNota = RsImpressora!EnderecoLocal
Else
   LcNota = "LPT1"
End If
'End If
While GlImprimeSpoolOrc
   While Not RsLogNota.EOF
     Nfi = RsLogNota!NF
     'logBoleta (Nfi)
     
     If Not LcAbriu Then
        FnunNota = FreeFile
        Open LcNota For Output Access Write As #FnunNota  'Abre Porta Nf
     End If
     Print #FnunNota, RsLogNota!Dados
     RsLogNota.Delete
     RsLogNota.MoveNext
     DoEvents
     LcAbriu = True
   Wend
   If LcAbriu Then
      Close #FnunNota
      LcAbriu = False
   End If
   If UCase(RsImpressora!Maquina) = "LOCAL" Then
      Set RsLogNota = DbbaseOrc.OpenRecordset("select * from LogImpressaoOrcamento where Maquina='" & "LOCAL" & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Else
      Set RsLogNota = DbbaseOrc.OpenRecordset("select * from LogImpressaoOrcamento where Maquina='" & GlNomeMaquina & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   End If
   DoEvents
Wend
SaidaLog:
If LcAbriu Then Close #FnunNota
'RsLogNota.Close
Exit Function

ErrLog:
If err.Number = 3021 Then
   GoTo SaidaLog
End If
If err = 3260 Then Resume 0
If err = 75 Then Resume Next
If err = 440 Then GoTo SaidaLog
If err = 3167 Then
   Set RsLogNota = DbbaseOrc.OpenRecordset("select * from LogImpressaoOrcamento where Maquina='" & GlNomeMaquina & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   RsLogNota.MoveLast
   RsLogNota.MoveFirst
   Resume 0
End If
MsgBox err.Description & err.Number
'Stop
'Resume 0

'Resume 0
GoTo SaidaLog
End Function

Function NomeMaquina()
' Read the computer's name
Dim compname As String * 255, cname As String
Dim x As Variant, LCLEtra As String
Dim a As Long

x = GetComputerName(compname, 255)
' Trim blank spaces and ending vbNullChar
cname = RTrim(LTrim(compname))
cname = Left(cname, Len(cname) - 1)
For a = 1 To 255
    LCLEtra = Mid$(cname, a, 1)
   ' MsgBox Asc(LcLetra)
    If Asc(LCLEtra) = 0 Then
       Exit For
    Else
      GlNomeMaquina = GlNomeMaquina & LCLEtra
    End If
 Next
'Text1.Text = cname
End Function
Function LogNotaFiscal()
Inicio:
On Error GoTo ErrLog
Dim LcNota As String, Nfi As String
Dim DbbaseLog As Database
Dim RsImpressora As Recordset, RsLogNota As Recordset
Dim FnunNota, LcAbriu As Integer

FnunNota = FreeFile
If Not GlServidorImpressora Then Exit Function
GlImprimeSpool = True
Set DbbaseLog = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsLogNota = DbbaseLog.OpenRecordset("select * from LogImpressaoNota where Maquina='" & GlNomeMaquina & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsImpressora = DbbaseLog.OpenRecordset("select * from impressoras where Impressora='" & GlPortaNota & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Not RsImpressora.EOF Then
   LcNota = RsImpressora!EnderecoLocal
Else
   LcNota = "LPT1"
End If
   While Not RsLogNota.EOF
     Nfi = RsLogNota!NF
     If GlboletoA4 = 0 Then logBoleta (Nfi)
     If Not LcAbriu Then Open LcNota For Output Access Write As #FnunNota  'Abre Porta Nf
    ' Print #FnunNota, Chr(27) + Chr(48) & RsLogNota!dados
    Print #FnunNota, RsLogNota!Dados
     RsLogNota.Delete
     RsLogNota.MoveNext
     DoEvents
     LcAbriu = True
   Wend
   If LcAbriu Then
      Close #FnunNota
      LcAbriu = False
   End If
  ' Set RsLogNota = DbbaseLog.OpenRecordset("select * from LogImpressaoNota where Maquina='" & GlNomeMaquina & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
  ' DoEvents
'Wend
SaidaLog:
If LcAbriu Then Close #FnunNota
RsLogNota.Close
Exit Function

ErrLog:
If err = 3021 Then GoTo SaidaLog
If err = 3260 Then Resume 0
If err = 75 Then Resume Next
If err = 440 Then GoTo SaidaLog
If err = 3167 Then
   Set RsLogNota = Dbbase.OpenRecordset("select * from LogImpressaoNota where Impressora='" & GlPortaNota & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   RsLogNota.MoveLast
   RsLogNota.MoveFirst
   Resume 0
End If
'MsgBox err.Description & err.Number
'Stop
'Resume 0
On Error Resume Next
Close #FnunNota
RsLogNota.Close
GoTo Inicio
End Function
Function ExibeCgc(LcCgc1 As String) As String
On Error Resume Next
Dim LcCgc As String
Dim a As Integer
LcCgc = ""
For a = 1 To Len(LcCgc1)
    If IsNumeric(Mid(LcCgc1, a, 1)) Then
        LcCgc = LcCgc & Mid(LcCgc1, a, 1)
    End If
Next
If Len(LcCgc) > 0 Then
   LcCgc = Mid(LcCgc, 1, 2) & "." & Mid(LcCgc, 3, 3) & "." & Mid(LcCgc, 6, 3) & "/" & Mid(LcCgc, 9, 4) & "-" & Mid(LcCgc, 13)
Else
   LcCgc = "  .   .   /    -  "
End If
ExibeCgc = LcCgc
End Function
Function logBoleta(Nfi)
On Error GoTo ErrLog
Dim LcBoleto As String
Dim DbbaseLogb As Database
Dim RsImpressora As Recordset, RsLogBoleto As Recordset
Dim FnunBoleto As Integer
FnunBoleto = FreeFile + 2
Set DbbaseLogb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsLogBoleto = DbbaseLogb.OpenRecordset("select * from LogBoleto where maquina='" & GlNomeMaquina & "' and nf='" & Nfi & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsImpressora = DbbaseLogb.OpenRecordset("select * from impressoras where Impressora='" & GlPortaBoleto & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)

   '=== Vai Imprimir
If Not RsImpressora.EOF Then
  LcBoleto = RsImpressora!EnderecoLocal
Else
  LcBoleto = "Lpt2"
End If
If RsLogBoleto.EOF Then GoTo exitsu
Open LcBoleto For Output Access Write As #FnunBoleto 'Abre Porta Nf
While Not RsLogBoleto.EOF
   Print #FnunBoleto, RsLogBoleto!Dados
   RsLogBoleto.Delete
   RsLogBoleto.MoveNext
   DoEvents
Wend


'RsLogBoleto.Close
RsImpressora.Close
'Set RsLogBoleto = Nothing
Set RsImpressora = Nothing

SaidaLog:
Close #FnunBoleto
exitsu:
Exit Function

ErrLog:
If err = 3021 Then GoTo SaidaLog
If err = 3260 Then Resume 0
If err = 75 Then Resume 0
If err = 3167 Then
   Set RsLogBoleto = Dbbase.OpenRecordset("select * from LogBoleto where Impressora='" & GlPortaBoleto & "' order by nf,Sequencia", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   RsLogBoleto.MoveLast
   RsLogBoleto.MoveFirst
   Resume 0
End If
If err = 440 Then GoTo SaidaLog
'MsgBox Err.Description & Err.Number

'Stop
'Resume 0
GoTo SaidaLog
End Function
Function GeraHistorico(produto, Descricao, NF, Tipo As String, Data As Date, santa, Santa2, califonia As Double, santau As Double, santa1u As Double, californiau As Double)
Dim db As Database
Dim RsHistorico As Recordset
Set db = OpenDatabase(GLBase)

Set RsHistorico = db.OpenRecordset("HistoricoProduto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
 
 RsHistorico.AddNew
 RsHistorico("Produto") = produto
 RsHistorico("descricao") = Descricao
 RsHistorico("unisanta") = santau
 RsHistorico("unsanta1") = santa1u
 RsHistorico("Uncalifornia") = californiau
 RsHistorico("santa") = santa
 RsHistorico("santa2") = Santa2
 RsHistorico("california") = califonia
 RsHistorico("nf") = NF
 RsHistorico("data") = Data
 RsHistorico("tipo") = Tipo
 RsHistorico.Update
 RsHistorico.Close
 db.Close
 Set RsHistorico = Nothing
 Set db = Nothing
End Function
Function GeraExtenso(LcExtenso As Currency) As String
Dim Retorno$, x%
Dim LcTamanho, LcGeraMaiusculo, a As Integer
Dim LCLEtra, LcCorreto, passaextenso As String
On Error GoTo Passa_Err
  Retorno$ = Space$(512)
  x% = Extenso(LcExtenso, Retorno$)
  passaextenso = Trim$(Retorno$)
LcGeraMaiusculo = True
LcTamanho = Len(passaextenso)
For a = 1 To LcTamanho
    LCLEtra = Mid(passaextenso, a, 1)
    If LcGeraMaiusculo Then
       LCLEtra = UCase(LCLEtra)
       LcGeraMaiusculo = False
    End If
    If Len(Trim(LCLEtra)) = 0 Then
       LcGeraMaiusculo = True
    End If
    LcCorreto = LcCorreto + LCLEtra
Next
GeraExtenso = LcCorreto
Passa_Fim:
  Exit Function
Passa_Err:
  MsgBox Error$(err)
  Resume Passa_Fim
End Function
Function VerificaReservaVencida()
Dim LcData As Date
Dim RsReserva As Recordset, RsProduto As Recordset
Dim LcCriterio, lcchave As String

LcData = Format(GlDataSistema, "dd/mm/yyyy")
LcCriterio = "select * From Reserva where DataMaxima < #" & GlDataSistema & "#"

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsReserva = Dbbase.OpenRecordset(LcCriterio, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsProduto = Dbbase.OpenRecordset("Produto", dbOpenTable, dbSeeChanges, dbOptimistic)
RsProduto.Index = "Codigo"

Do Until RsReserva.EOF
   lcchave = RsReserva!codigoproduto
   RsProduto.Seek "=", lcchave
   If Not RsProduto.NoMatch Then
      RsProduto.Edit
      RsProduto!Reservado = 0
      RsProduto.Update
   End If
   RsReserva.Delete
   RsReserva.MoveNext
Loop
RsReserva.Close
RsProduto.Close
Dbbase.Close

End Function
    
Public Function lancacaixa(LcTipo, LcDoc As String, LcTipoM As String, LcValor As Double)
On Error GoTo ErrLanca
Dim LcData          As Date
Dim RsCaixa         As Recordset
Dim RSLanca         As Recordset
Dim RsMov           As Recordset
Dim RsTipo          As Recordset
Dim LCLanca         As String
Dim LcCri           As String
Dim Lccr            As String
Dim LcCricaixa      As String
Dim LcPesCricaixa   As String
Dim LcPesquisa      As String
Dim LcRecPesp       As String
Dim LcSaldo         As Double
Dim LcEntradas      As Double
Dim LcSaida         As Double
Dim LcSaldoAnterior As Double
Dim LcSaldoDeletado As Double
Dim LcDataDia       As Date
Dim LcEditar        As Boolean
LcRecPesp = LcTipo
LcSaldoDeletado = 0
LcData = Format(GlDataSistema, "dd/mm/yyyy")
LcEditar = False
If LcRecPesp = "Receita" Then
    LcCri = " Select * from Alid015 where nf='" & LcDoc & "'"
    LCLanca = "R"
Else
    LcCri = " Select * from Alid014 where nf='" & LcDoc & "'"
    LCLanca = "D"
End If
Lccr = "Select * From MovimentacaoCaixa order by contador"
LcCricaixa = "Select * from Caixa"
LcTipo = "Select * from alid008 where TPMONET='" & LcTipoM & "'"
LcDataDia = Date
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsCaixa = Dbbase.OpenRecordset(LcCricaixa, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenTable, dbSeeChanges, dbOptimistic)
Set RSLanca = Dbbase.OpenRecordset(LcCri)
Set RsMov = Dbbase.OpenRecordset(Lccr, dbOpenDynaset, dbSeeChanges, dbOptimistic) ' ", dbOpenTable, dbSeeChanges, dbOptimistic)
Set RsTipo = Dbbase.OpenRecordset(LcTipo, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'====Busca Valor Doc
'If Not RsLanca.EOF Then
'   LcValor = RsLanca!valor
'Else
'   LcValor = 0
'End If
'===== Verifica se o tipo Monetario Pode Movimentar caixa

'=== Procura se Existe o Doc E Apaga

If RsTipo.EOF Then
   MsgBox "Tipo Monetario deste Documento Não Foi Encontrado." & Chr(13) & "O Caixa Não Será Atualizado.", 64, "Aviso"
   Exit Function
End If
If RsTipo!MOVCAIXA <> "S" Then
   Exit Function
End If
'==== Busca o Saldo Anterior
LcPesquisa = "Nf='" & LcDoc & "' and Rec_Pag='" & LCLanca & "'"
RsMov.FindFirst LcPesquisa
If Not RsMov.NoMatch Then
   LcEditar = True
   RsMov.Edit
   LcDataDia = RsMov!DataLancamento
   LcSaldoDeletado = RsMov!Valor
   If LCLanca = "R" Then
      LcSaldoAnterior = RsMov!Saldo + (LcValor - RsMov!Valor)
   Else
      LcSaldoAnterior = RsMov!Saldo - (LcValor + RsMov!Valor)
   End If
Else
   
   If Not RsMov.EOF Then
      RsMov.MoveLast
      LcSaldoAnterior = RsMov!Saldo
   Else
      LcSaldoAnterior = 0
   End If
   RsMov.AddNew
End If
'==== Gera o Movimento do Dia.
If LCLanca = "R" Then
   RsMov!Saldo = LcSaldoAnterior + LcValor
Else
   RsMov!Saldo = LcSaldoAnterior - LcValor
End If
RsMov!TipoMonetario = LcTipoM
RsMov!Rec_Pag = LCLanca
RsMov!Valor = LcValor
RsMov!DataLancamento = LcDataDia
RsMov!fechado = False
RsMov!NF = LcDoc

RsMov.Update
LcSaldoAnterior = RsMov!Saldo
RsMov.MoveNext
'=== Acerta Todo o Movimento
Do Until RsMov.EOF
    RsMov.Edit
    If RsMov!Rec_Pag = "R" Then
       RsMov("Saldo") = LcSaldoAnterior + RsMov!Valor
    Else
       RsMov("Saldo") = LcSaldoAnterior - RsMov!Valor
    End If
    RsMov.Update
    LcSaldoAnterior = RsMov!Saldo
    RsMov.MoveNext
Loop

'===Acerta o Caixa
  '=Busca o Caixa do Dia
If LcEditar Then
   LcPesCricaixa = "Data=#" & Format(LcDataDia, "mm/dd/yy") & "#"
Else
   LcPesCricaixa = "Data=#" & Format(Date, "mm/dd/yy") & "#"
End If
RsCaixa.FindFirst LcPesCricaixa
If Not RsCaixa.NoMatch Then
   RsCaixa.Edit
   If LCLanca = "R" Then
      RsCaixa!Recebimentos = RsCaixa!Recebimentos + LcValor - LcSaldoDeletado
      RsCaixa!SaldoAtual = RsCaixa!SaldoAtual + LcValor - LcSaldoDeletado
   Else
      RsCaixa!Pagamentos = RsCaixa!Pagamentos + LcValor + LcSaldoDeletado
      RsCaixa!SaldoAtual = RsCaixa!SaldoAtual - LcValor + LcSaldoDeletado
   End If
Else
   RsCaixa.AddNew
   If LCLanca = "R" Then
      RsCaixa!Recebimentos = LcValor
      RsCaixa!Data = Date
      RsCaixa!Pagamentos = 0
      RsCaixa!SaldoAnterior = 0
      RsCaixa!SaldoAtual = LcValor
      RsCaixa!fechado = False
   Else
      RsCaixa!Recebimentos = 0
      RsCaixa!Data = Date
      RsCaixa!Pagamentos = LcValor
      RsCaixa!SaldoAnterior = 0
      RsCaixa!SaldoAtual = 0 - LcValor
      RsCaixa!fechado = False
   End If
End If
RsCaixa.Update
LcSaldoAnterior = RsCaixa!SaldoAtual
RsCaixa.MoveNext
'==== Atualiza caixas posteriores
Do Until RsCaixa.EOF
   RsCaixa.Edit
   RsCaixa!SaldoAnterior = LcSaldoAnterior
   RsCaixa!SaldoAtual = LcSaldoAnterior + RsCaixa!Recebimentos - RsCaixa!Pagamentos
   RsCaixa.Update
   RsCaixa.MoveNext
Loop
RsCaixa.Close
RSLanca.Close
RsMov.Close
RsTipo.Close
Dbbase.Close
Set RsCaixa = Nothing
Set RSLanca = Nothing
Set Dbbase = Nothing
Exit Function
ErrLanca:
MsgBox err.Description & err.Number
Resume Next

End Function
Public Function VerificaDuplicado(Lcinde As Integer) As Integer
Dim LcIndiceAtual, LcIndex, LcVeriDuplic As String
'Exit Function
LcIndex = Mid(GlFormA.txt(Lcinde).Tag, 7, 2)
LcVeriDuplic = Mid$(GlFormA.txt(Lcinde).Tag, 10, 1)

If LcTipoDados <> 1 Then Exit Function
If LcVeriDuplic = "S" Then
 
   LcIndiceAtual = LcIndice
   LcIndice = MtPesquisa(Val(LcIndex)).Indice
   GlChave = GlFormA.txt(Lcinde).Text
   Call AbreBanco(GlFormAtual)
   If AchaReg(1) Then
      VerificaDuplicado = True
      MsgBox "O " & MtPesquisa(Val(LcIndex)).campo & " " & GlFormA.txt(Lcinde).Text & " Já foi Cadastrado...", 64, "Dados Duplicados."
   Else
      VerificaDuplicado = False
   End If
   LcIndice = LcIndiceAtual

End If


End Function
Function MontaMatrizCliente()
On Error GoTo Erromonta
Dim a As Long
LcIndice = "nome"
Call AbreBanco(Cliente)
a = 0

ReDim MtCliente(a)
Do Until RsAtual.EOF
     ReDim Preserve MtCliente(a)
     MtCliente(a).codigo = RsAtual!codigo
     MtCliente(a).Descricao = RsAtual!Nome
     GlFormA.cbo(0).AddItem RsAtual!Nome
     RsAtual.MoveNext
     a = a + 1
Loop

GlTamanhoCliente = a
FechaBanco
Exit Function
Erromonta:
'Stop
'MsgBox Err.Description & Err
Resume Next
End Function

Sub IniciaNFE()
    
    ' Instantiate the Toolbox
    Set g_InteropToolbox = New InteropToolbox
    g_InteropToolbox.Initialize
    
    ' Call Initialize method only when first creating the toolbox
    ' This aids in the debugging experience
    g_InteropToolbox.Initialize

    ' Signal Application Startup
    g_InteropToolbox.EventMessenger.RaiseApplicationStartedupEvent
    
    
    ' Do application logic
          
        
    ' Signal Application Shutdown
    g_InteropToolbox.EventMessenger.RaiseApplicationShutdownEvent
    
End Sub

Public Sub Main()
On Error Resume Next
Dim Arqini As String
IniciaNFE
Arqini = BuscaDirWin
Arqini = App.EXEName & ".ini"
GlNomeProjeto = App.EXEName & "Nfe"
If Dir(Arqini, vbArchive) = "" Then
Debug.Print Arqini
NomedoArquivo = Arqini
abreconexao
ListaIgnorar = "Nome"
ListaIgnorar = "nomereceita"
'Shell "net time \\servidor decisao /set /yes", vbHide
GlImplanta = False
ObtemPath
'VerificaVersao
ProcessaDDl
ProcessaDDlMySql
Verificatb
'Zeracontas


End If
frmSplash.Show
End Sub
Public Function sFormataCaminho(ByVal sCaminho As String) As String

    'Verifica se existe "\" no caminho do arquivo
    If Not Right(sCaminho, 1) = Chr(92) Then
        sCaminho = sCaminho & Chr(92)
    End If
    sFormataCaminho = sCaminho
    
End Function

Public Function Etch(fname As Form, a$, x, Y, LcCor)
    fname.CurrentX = x
    fname.CurrentY = Y
    fname.ForeColor = QBColor(Int(LcCor / 2))
    fname.Print a$
    fname.CurrentX = x - 38
    fname.CurrentY = Y - 30
    fname.ForeColor = QBColor(LcCor)
    fname.Print a$
End Function
Function montamatrizfornecedor()
On Error GoTo Erromonta
Dim a As Long
LcIndice = "nome"
Call AbreBanco(fornecedor)
a = 0

ReDim MtCliente(a)
Do Until RsAtual.EOF
     ReDim Preserve MtCliente(a)
     MtCliente(a).codigo = RsAtual!codigo
     MtCliente(a).Descricao = RsAtual!Nome
     GlFormA.cbo(0).AddItem RsAtual!Nome
     RsAtual.MoveNext
     a = a + 1
Loop

GlTamanhoCliente = a
FechaBanco
Exit Function
Erromonta:
'Stop
'MsgBox Err.Description & Err
Resume Next
End Function
Public Function RetornaNome(LcCodigo As String) As String
On Error Resume Next
Dim a As Long
 For a = 0 To GlTamanhoCliente
     If GlCampo3 = MtCliente(a).codigo Then
        
        RetornaNome = MtCliente(a).Descricao
        Exit For
     End If
 Next
End Function
Public Function Sincroniza()
On Error Resume Next
Exit Function
Dim a As Integer, LcCodigo, LcIndiceCampo As String
LcCodigo = GlCampo0
For a = 0 To 30
    If MtPesquisa(a).Indice = LcIndice Then
       LcIndiceCampo = Mid$(GlFormA.txt(a).Tag, 7, 2)
       GlChave = GlFormA.txt(LcIndiceCampo).Text
       Exit For
    End If
Next

Call AbreBanco(GlFormAtual)
Call AchaReg(1)
Do Until RsAtual!codigo = CLng(LcCodigo)
    If err <> 0 Then Exit Do
    RsAtual.MoveNext
Loop

End Function


Public Function ObtemPath()
On Error Resume Next
Dim NumeroDoArquivo As Integer, Achou As String
Dim LcObj As CommonDialog

    Achou = Dir(App.Path & "\BaseDados.txt")
    NumeroDoArquivo = FreeFile
    If Achou = "" Then
        'Arquivo não existe
       
           With FrmPrincipal.Abrird
            .InitDir = App.Path
            .FileName = "*.mdb"
            .ShowOpen
           End With
           GLBase = FrmPrincipal.Abrird.FileName
           If GLBase = "" Then End
        Open App.Path & "\BaseDados.txt" For Output As #NumeroDoArquivo
        Print #NumeroDoArquivo, GLBase
    Else
        'Arquivo já existe
        Open App.Path & "\BaseDados.txt" For Input As #NumeroDoArquivo
        Line Input #NumeroDoArquivo, GLBase
        If Dir(GLBase) = "" Then
            GlAchaBase = False
            ErroBase.Show
            While Not GlAchaBase
                  DoEvents
            Wend
           If Not GlRecuperou Then
              With FrmPrincipal.Abrird
                  .InitDir = App.Path
                  .FileName = "*.mdb"
                  .ShowOpen
              End With
              GLBase = FrmPrincipal.Abrird.FileName
        
              Close #NumeroDoArquivo
             GLBase = FrmPrincipal.Abrird.FileName
             Open App.Path & "\BaseDados.txt" For Output As #NumeroDoArquivo
             Print #NumeroDoArquivo, GLBase
          End If
      End If
    End If
    If GLBase = "*.mdb" Then
     'Arquivo não existe
      With FrmPrincipal.Abrird
         .InitDir = App.Path
         .FileName = "*.mdb"
         .ShowOpen
      End With
      GLBase = FrmPrincipal.Abrird.FileName
      If GLBase = "*.mdb" Then End
         Close #NumeroDoArquivo
         Open App.Path & "\BaseDados.txt" For Output As #NumeroDoArquivo
         Print #NumeroDoArquivo, GLBase
   End If
   Close #NumeroDoArquivo
End Function


Public Function FadeForm(frm As Form, pRed As Integer, pGreen As Integer, pBlue As Integer)
    Dim SaveScale As Integer, SaveStyle As Integer, SaveDraw As Integer
    Dim Y As Long, x As Long, i As Long, j As Long, pixels As Long
    'salvar as configurações atuais do form
    SaveScale = frm.ScaleMode
    SaveStyle = frm.DrawStyle
    SaveDraw = frm.AutoRedraw
    'pintar a tela
    frm.ScaleMode = 3
    pixels = Screen.Height / Screen.TwipsPerPixelY
    x = pixels / 64 + 0.5
    frm.DrawStyle = 5
    frm.AutoRedraw = True
    For j = 0 To pixels Step x
        Y = 240 - 245 * j / pixels
        If Y < 0 Then Y = 0
        frm.Line (-2, j - 2)-(Screen.Width + 2, j + x + 3), RGB(-pRed * Y, -pGreen * Y, -pBlue * Y + 1), BF
    Next j
    'restaura configurações do form
    frm.ScaleMode = SaveScale
    frm.DrawStyle = SaveStyle
    frm.AutoRedraw = SaveDraw
End Function
Function Teclas(tecla As Integer)
Dim LcCap As String
On Error GoTo ErTecla
LcCap = GlFormA.Caption
If GlFormA.Name <> "ContratoFornecimento" And GlFormA.Name <> "FrmOpcoes" And GlFormA.Name <> "FrmReajustaPreco" Then
   GlIniceAtual = Screen.ActiveControl.Index
End If
Select Case tecla
    Case 112
    Case 113
     If GlFormA.Name = "FrmBaixaReceita" Or GlFormA.Name = "FrmBaixaDespesas" Or GlFormA.Name = "FrmReajustaPreco" Then
        SendKeys "%+{O}"
     Else
        SendKeys "%+{S}"
     End If
     
    Case 114
      Select Case GlFormA.Name
       Case Is = "FrmSaidaProdutoAlternativo"
            SendKeys "%+{F}"
         Case Is = "FrmSaidaProduto"
            SendKeys "%+{F}"
         Case Is = "FrmEntradaProduto"
            SendKeys "%+{F}"
         Case Is = "FrmBaixaReceita"
            SendKeys "%+{C}"
         Case Is = "FrmBaixaDespesas"
            SendKeys "%+{C}"
         Case Is = "FrmVendaOrcam"
            SendKeys "%+{P}"
         Case Is = "FrmPedido"
            SendKeys "%+{P}"
         Case Else
            SendKeys "%+{E}"
      End Select
    Case 115
      Select Case GlFormA.Name
       Case Is = "FrmSaidaProdutoAlternativo"
            SendKeys "%+{E}"
         Case Is = "FrmSaidaProduto"
            SendKeys "%+{E}"
         Case Is = "FrmEntradaProduto"
            SendKeys "%+{E}"
         Case Is = "FrmPedido"
            SendKeys "%+{E}"
         Case Is = "FrmVendaOrcam"
            SendKeys "%+{E}"
         Case Is = "FrmFuncionario"
            SendKeys "%{m}"
         Case Else
            SendKeys "%+{T}"
            
      End Select
    Case 116
       GlFormA.Caption = "Aguarde, Executando Filtro..."
       Select Case GlFormA.Name
        Case Is = "FrmSaidaProdutoAlternativo"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmPesquisaCliente.Show , GlFormA
           Else '=== é Produto
              FrmPesquisaProdutos.Show , GlFormA
           End If
       Case Is = "ContratoFornecimento"
            FrmPesquisaProdutos.Show , GlFormA
       Case Is = "FrmTransportadora"
           ExibeCidade.Show , GlFormA
        Case Is = "FrmFornecedor"
           ExibeCidade.Show , GlFormA
        Case Is = "FrmGalpao"
           ExibeCidade.Show , GlFormA
        Case Is = "FrmCliente"
           ExibeCidade.Show , GlFormA
        Case Is = "FrmProduto"
           ExibeUnidade.Show , GlFormA
        Case Is = "FrmSaidaProduto"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmPesquisaCliente.Show , GlFormA
           Else '=== é Produto
              FrmPesquisaProdutos.Show , GlFormA
           End If
         Case Is = "FrmProposta"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmBuscaCliente.Show , GlFormA
           Else '=== é Produto
              FrmBuscaProduto.Show , GlFormA
           End If
        Case Is = "FrmEntradaProduto"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmPesquisaFornecedores.Show , GlFormA
           Else '=== é Produto
              FrmBuscaProduto.Tag = FrmEntradaProduto.txt(1).Text
              FrmBuscaProduto.Show , GlFormA
           End If
      
        Case Is = "FrmPedido"
           If GlEscolhe = 1 Then '=== é Fornecedor
              FrmPesquisaFornecedores.Show , GlFormA
           Else '=== é Produto
              If GlEscolhe = 2 Then '==== é Produto
                 FrmBuscaProduto.Show , GlFormA
              Else
                 FrmBuscaCliente.Show , GlFormA
              End If
           End If
        Case Is = "Receitas"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmBuscaCliente.Show , GlFormA
           Else '=== é Produto
              ExibeMonetario.Show , GlFormA
           End If
        Case Is = "Despesas"
           If GlEscolhe = 1 Then '=== é Cliente
              FrmPesquisaFornecedores.Show , GlFormA
           Else '=== é Produto
              ExibeMonetario.Show , GlFormA
           End If
       End Select
       GlFormA.Caption = LcCap
    Case 117
        SendKeys "%+{P}"
    Case 118
         If GlFormA.Name = "FrmVendaOrcam" Then
            SendKeys "%+{Q}"
         Else
            SendKeys "%+{A}"
         End If
    Case 119
        SendKeys "%+{G}"
    Case 120
        SendKeys "%+{U}"
    Case 121
        If GlFormA.Name = "FrmEntradaProduto" Or GlFormA.Name = "FrmSaidaProduto" Or GlFormA.Name = "FrmSaidaProdutoAlternativo" Then
           SendKeys "%+{C}"
        Else
           SendKeys "%+{F}"
        End If
    Case 122
        SendKeys "%+{Q}"
    Case Is = 123
        SendKeys "%+{O}"
End Select
Exit Function
ErTecla:
If err <> 343 Then
    MsgBox err.Description & " Nº: " & err
Else
    Resume Next
End If
    
'MsgBox Err.Description & Err

Exit Function
'Resume 0
End Function

Function AbreBase()

Set Area = DBEngine.Workspaces(0)
Set Dbbase = OpenDatabase(GLBase, False, False) ' "dBASE III;")
End Function
Function AcertaDecimal(LcNumero As String) As String
Dim LcPrefixo, LcSufixo, LCLEtra As String
Dim LcSaida, a As Integer
LcSaida = False
For a = Len(LcNumero) To 1 Step -1
    LCLEtra = Mid(LcNumero, a, 1)
    If LCLEtra = "." Then
       LcSaida = True
       Exit For
    End If
Next
If LcSaida Then
   LcPrefixo = Mid(LcNumero, 1, a - 1)
   LcSufixo = Left(Mid(LcNumero, a + 1) & "00", 2)
   AcertaDecimal = LcPrefixo & "," & LcSufixo
Else
   AcertaDecimal = LcNumero
End If

End Function

Public Function AbreBanco(LcTbl As Tabela)
 On Error GoTo ErroAbertura

 AbreBase

 Select Case LcTbl
        Case Is = Custo
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("DecricaoCusto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("DecricaoCusto", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
            
        Case Is = Cliente
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid001", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
            
        Case Is = PropostaCliente
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("PropostaCliente", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("PropostaCliente", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
         Case Is = Cheques
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("Cheques", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("Cheques", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
          Case Is = Transportadora
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("Transportadora", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("Transportadora", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
            
         Case Is = fornecedor
           If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid002", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
           
        Case Is = Galpao
           If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid012", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid012", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
        Case Is = Cidade
           If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
                Set RsAtual = Dbbase.OpenRecordset("alid005", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
        Case Is = produto
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid009", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid009", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
           
        Case Is = Unidade
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("alid004", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid004", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
        Case Is = pedido
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("Pedido", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("Pedido", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
        Case Is = produtora
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("Produtora", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("Produtora", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
       Case Is = Convenio
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("Convenio", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("Convenio", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
       Case Is = Funcionario
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid200", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
       Case Is = monetario
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid008", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
       Case Is = tiporec
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid007", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid007", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
      Case Is = Receber
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("alid015", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid015", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
      Case Is = pagar
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("alid014", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid014", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
      
      Case Is = EntradaProduto
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("EntradaProduto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("EntradaProduto", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
      Case Is = Caixa
            If GLPesquisa Then
               Set RsAtual = Dbbase.OpenRecordset("ALID016", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("ALID016", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
      Case Is = opcao
            If GLPesquisa Then
                Set RsAtual = Dbbase.OpenRecordset("alid202", dbOpenDynaset, dbSeeChanges, dbOptimistic)
            Else
               Set RsAtual = Dbbase.OpenRecordset("alid202", dbOpenTable, dbSeeChanges, dbOptimistic)
            End If
           
      Case Is = usuario
            Set RsAtual = Dbbase.OpenRecordset("Usuario", dbOpenDynaset, dbSeeChanges, dbOptimistic)
      Case Is = GrpUsuario
            Set RsAtual = Dbbase.OpenRecordset("GrpUsuario", dbOpenDynaset, dbSeeChanges, dbOptimistic)
 End Select
If GLPesquisa Then Exit Function

 If Len(Trim(LcIndice)) <> 0 Then
    RsAtual.Index = LcIndice
 Else
    RsAtual.Index = "Codigo"
 End If
    
 Exit Function
ErroAbertura:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
        MsgBox err.Description & err
End Select
End Function
Public Function MoveTecla(LcIndex As Integer, LcTecla As Integer)
On Error Resume Next

Dim LcStop, a As Integer, LcCriterio As String
LcStop = Screen.ActiveForm.txt(LcIndex).TabIndex + 1


If LcTecla = 13 Then 'Pressionou enter
   For a = 0 To 35
      
       If Screen.ActiveForm.txt(a).TabIndex = LcStop Then 'Achou
          Screen.ActiveForm.txt(a).SetFocus
          
          Exit Function
       End If
   Next
   'Stop
   Screen.ActiveForm.txt(1).SetFocus
End If
AbreBase
If LcTecla = 112 Then 'Chamou a ajuda
  ' Ajuda.Show , Screen.ActiveForm
   RsFuncao.Close
   Exit Function
End If
If LcTecla >= 113 And LcTecla <= 123 Then 'Pressionou Uma Tecla de Função
   Set RsFuncao = Dbbase.OpenRecordset("TeclasFuncao", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcCriterio = "CodigoTecla=" & LcTecla
   RsFuncao.FindFirst LcCriterio
   If Not RsFuncao.NoMatch Then
      Select Case RsFuncao!FuncãodaTecla
             Case Is = "Pesquisa"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  frmPesquisa.Show , Screen.ActiveForm
             Case Is = "Ordena"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  FrmOrdena.Show , Screen.ActiveForm
             Case Is = "Exclui"
                  If LcTipoDados = 1 Or LcTipoDados = 3 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  If Exclui(GlFormAtual) = 1 Then
                     GlFormA.VinculaDados
                  End If
             Case Is = "Primeiro"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  GlMov = True
                  If MovImentacao(enPrimeiro, GlFormAtual) Then GlFormA.VinculaDados
                  GlMov = False
             Case Is = "Anterior"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  GlMov = True
                  If MovImentacao(enAnterior, GlFormAtual) Then GlFormA.VinculaDados
                  GlMov = False
             Case Is = "Próximo"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  GlMov = True
                  If MovImentacao(enSeguinte, GlFormAtual) Then GlFormA.VinculaDados
                  GlMov = False
             Case Is = "Último"
                  If LcTipoDados = 1 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  GlMov = True
                  If MovImentacao(enSeguinte, GlFormAtual) Then GlFormA.VinculaDados
                  GlMov = False
             Case Is = "Salva"
                  If LcTipoDados = 3 Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  If GlFormA.CmdSalvar.Enabled = False Then MsgBox "Operação Não Disponivel.", 64, "Aviso": Exit Function
                  Call SalvaRegistro(GlFormAtual)
                  LcRegAtual = True
                  GlFormA.VinculaDados
                  LcRegAtual = False
                  'NovoReg
             Case Else
                  MsgBox "Esta Tecla Não Foi Programada...", 64, "Aviso"
      End Select
   End If
   RsFuncao.Close
End If

Exit Function
ErroMove:
'Stop
Resume Next
End Function
Public Function LeConfiguracaoMeiaFolha()
On Error Resume Next
Dim LcCamilho As String
Dim LcArq As Integer
Dim a As Integer

For a = Len(GLBase) To 1 Step -1
   If Mid(GLBase, a, 1) = "\" Then Exit For
Next
LcCaminho = Mid(GLBase, 1, a)
LcCaminho = LcCaminho & "meiaFolha.ini"
LcArq = FreeFile
Open LcCaminho For Input As #LcArq
a = 1
err.Number = 0
Do Until EOF(LcArq)
    If err.Number > 0 Then Exit Do
    Input #LcArq, integridade
    Select Case a
        Case Is = 1
            GlItensMeiaFolha = CInt(integridade)
        Case Is = 2
            GlMargemMeiaFolha = CInt(integridade)
        Case Is = 3
            GlSaltoFinalMeiaFolha = CInt(integridade)
        Case Is = 4
            GlImpressaoMeiaFolha = CInt(integridade)
        Case Is = 5
           If integridade = 1 Then GlCabecalhoMeiaFolha = True Else GlCabecalhoMeiaFolha = False
    End Select
    a = a + 1
Loop
Close #LcArq

End Function
Public Function VerificaOpcoes()
On Error Resume Next
Dim LcArq, NumeroMsg, a As Integer
Dim integridade, LcArqMsg, letra, glmsga, LcAspas As String
Dim RsAtual As Recordset, RsOp As Recordset

LcAspas = Chr(34)
'==Seta Vairaveis Para Uso do Sistema

LigaTitulo = Chr(27) & Chr(87) & Chr(1) & Chr(27) & Chr(71) & Chr(27) & Chr(69)
DesligaTitulo = Chr(27) & Chr(87) & Chr(0) & Chr(27) & Chr(72) & Chr(27) & Chr(70)
LigaNegrito = Chr(27) & Chr(71)
DesligaNegrito = Chr(27) & Chr(72)
LigaDraft = Chr(27) & "x" & Chr(0)

AbreBase
Set RsAtual = Dbbase.OpenRecordset("empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsOp = Dbbase.OpenRecordset("alid901", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcArq = FreeFile

Open App.Path & "\opcao.txt" For Input As #LcArq
a = 1
err.Number = 0
Do Until EOF(LcArq)
    If err.Number > 0 Then Exit Do
    Input #LcArq, integridade
    Select Case a
        Case Is = 1
            GLSaltoLinhaNota = CInt(integridade)
        Case Is = 2
            GlPortaNota = integridade
        Case Is = 3
            GlPortaBoleto = integridade
        Case Is = 4
            If integridade = "1" Then GLConfirmaNovo = True Else GLConfirmaNovo = False
        Case Is = 5
           If integridade = "1" Then GlConfirmaAlteracao = True Else GlConfirmaAlteracao = False
        Case Is = 6
            If integridade = "1" Then GlConfirmaExclusao = True Else GlConfirmaExclusao = False
        Case Is = 7
            If integridade = "1" Then GlFaturaSaida = True Else GlFaturaSaida = False
        Case Is = 8
            If integridade = "1" Then GlVistaSaida = True Else GlVistaSaida = False
        Case Is = 9
            If integridade = "1" Then GlCaixaSaida = True Else GlCaixaSaida = False
       Case Is = 10
            If integridade = "1" Then GlFaturaEntrada = True Else GlFaturaEntrada = False
       Case Is = 11
           If integridade = "1" Then GlVistaEntrada = True Else GlVistaEntrada = False
       Case Is = 12
           If integridade = "1" Then GlCaixaEntrada = True Else GlCaixaEntrada = False
      Case Is = 13
           If integridade = "1" Then GlLucroCad = True Else GlLucroCad = False
      Case Is = 14
           If integridade = "1" Then GlLucroAlteracao = True Else GlLucroAlteracao = False
      Case Is = 15
           If integridade = "1" Then GlMinimoAlteracao = True Else GlMinimoAlteracao = False
      Case Is = 16
           If integridade = "1" Then GlComercio = True Else GlComercio = False
      Case Is = 17
           If integridade = "1" Then GlRepresentante = True Else GlRepresentante = False
      Case Is = 18
          If integridade = "1" Then GLCalculacodigoProduto = True Else GLCalculacodigoProduto = False
      Case Is = 19
            GlPortaOrcamento = integridade
      Case Is = 20
            If integridade = "1" Then Gl40colunas = True Else Gl40colunas = False
      Case Is = 21
            GlSaltoLinhasOrcamento = CInt(integridade)
      Case Is = 22
            If integridade = "1" Then GLCalculacodigoCliente = True Else GLCalculacodigoCliente = False
      Case Is = 23
            If integridade = "1" Then GLCalculacodigoFornecedor = True Else GLCalculacodigoFornecedor = False
      Case Is = 24
            If integridade = "1" Then GlVariasComissao = True Else GlVariasComissao = False
      Case Is = 25
           GlMargem = CLng(integridade)
      Case Is = 26
          If integridade = "1" Then GlDadosTransportadora = True Else GlDadosTransportadora = False
      Case Is = 27
          If integridade = "1" Then GlEsclheVendedor = True Else GlEsclheVendedor = False
      Case Is = 28
          If integridade = "1" Then GlEscolheCliente = True Else GlEscolheCliente = False
      Case Is = 29
           If integridade = "1" Then GlIpi = True Else GlIpi = False
      Case Is = 30
            If integridade = "1" Then GlDetalhaDesconto = True Else GlDetalhaDesconto = False
      Case Is = 31
           If integridade = "1" Then GlImprimeDetalhaDesconto = True Else GlImprimeDetalhaDesconto = False
      Case Is = 32
            If integridade = "1" Then GlRateiaAcrecimo = True Else GlRateiaAcrecimo = False
      Case Is = 33
          GlEstoqueDisponivel = integridade
      Case Is = 34
         GlPuloFim = CInt(integridade)
      Case Is = 35
         If integridade = "1" Then GLPadraoWindows = True Else GLPadraoWindows = False
      Case Is = 36
         GlDecimais = CLng(integridade)
      Case Is = 37
         If integridade = "1" Then GlArmazenaGalpao = True Else GlArmazenaGalpao = False
      Case Is = 38
         If integridade = "1" Then GlServidorImpressora = True Else GlServidorImpressora = False
       Case Is = 39
         If integridade = "1" Then GlDescUnit = True Else GlDescUnit = False
       Case Is = 40
           If integridade = "1" Then GlSenha = True Else GlSenha = False
       Case Is = 41
           If integridade = "1" Then GlSoOrcamento = True Else GlSoOrcamento = False
       Case Is = 42
           If integridade = "1" Then GlNaoBloqueia = True Else GlNaoBloqueia = False
       Case Is = 43
           GlSenhaLiberacao = integridade
       Case Is = 44
           If integridade = 1 Then GlServidorImpressoraOrc = True Else GlServidorImpressoraOrc = 0
       ' Case Is = 45
      '     If integridade = 1 Then GlDescUnit = True Else GlDescUnit = False
       Case Is = 45
           If integridade = 1 Then GlFuncCodigo = True Else GlFuncCodigo = False
       Case Is = 46
           If integridade = 1 Then GlFuncNome = True Else GlFuncNome = False
       Case Is = 47
           If integridade = 1 Then GlFuncEmpresa = True Else GlFuncEmpresa = False
       Case Is = 48
           If integridade = 1 Then GlImprimeSemLinha = True Else GlImprimeSemLinha = False
       Case Is = 49
           If integridade = 1 Then GlInclusaoReceita = True Else GlInclusaoReceita = False
       Case Is = 50
           If integridade = 1 Then GlBaixaReceita = True Else GlBaixaReceita = False
       Case Is = 51
           If integridade = 1 Then GlVendaVista = True Else GlVendaVista = False
       Case Is = 52
           If integridade = 1 Then GlVendaPrazo = True Else GlVendaPrazo = False
       Case Is = 53
           If integridade = 1 Then GlInclusaoDespesa = True Else GlInclusaoDespesa = False
       Case Is = 54
           If integridade = 1 Then GlBaixaDespesa = True Else GlBaixaDespesa = False
       Case Is = 55
           If integridade = 1 Then GlEntradaVista = True Else GlEntradaVista = False
       Case Is = 56
           If integridade = 1 Then GlEntradaPrazo = True Else GlEntradaPrazo = False
       Case Is = 57
           If integridade = 1 Then GlInclusaoCheque = True Else GlInclusaoCheque = False
       Case Is = 58
           If integridade = 1 Then GlBaixaCheque = True Else GlBaixaCheque = False
        Case Is = 59
            GlSenhaDebito = integridade
        Case Is = 60
            GlSenhaCredito = integridade '60
        Case Is = 61
            If integridade = 1 Then GlAtualizaPreco = True Else GlAtualizaPreco = 0
        Case Is = 62
           If integridade = 1 Then GlLimpaTelaOrc = True Else GlLimpaTelaOrc = False
        Case Is = 63
           Glmargemnota = CInt(integridade)
        Case Is = 64
           If integridade = 1 Then GlMeiaFolha = True Else GlMeiaFolha = False
        Case Is = 65
           If integridade = 1 Then GlboletoA4 = True Else GlboletoA4 = False
        Case Is = 66
           If integridade = 1 Then GlGeraPorItem = True Else GlGeraPorItem = False
        Case Is = 67
           If integridade = 1 Then GlImprimeValorCFOPNota = True Else GlImprimeValorCFOPNota = False
        Case Is = 68
            If integridade = 1 Then GlContrato = True Else GlContrato = False
        Case Is = 69
            If integridade = 1 Then GlNaoVerificaEstoque = True Else GlNaoVerificaEstoque = False
        Case Is = 70
            If integridade = 1 Then GlExibirLucratividade = True Else GlExibirLucratividade = False
        Case Is = 71
            If integridade = 1 Then GlVerificarIcmsDiferenciado = True Else GlVerificarIcmsDiferenciado = False
        Case Is = 72
            If integridade = 1 Then GlComissaoBelclean = True Else GlComissaoBelclean = False
        Case Is = 73
            GLInformacaoNF = integridade
        Case Is = 74
            If integridade = 1 Then GLNImprimeBaseC = True Else GLNImprimeBaseC = False
        Case 75
            If integridade = 1 Then GLAproveitamentoICMS = True Else GLAproveitamentoICMS = False
        Case 76
            If integridade = 1 Then GlBoletoCEF = True Else GlBoletoCEF = False
        Case 77
            GlIntrucao1 = integridade
        Case 78
            GlIntrucao2 = integridade
        Case 79
            GlIntrucao3 = integridade
        Case 80
            GlLinhasSaltarInicio = CInt(integridade)
        Case 81
            GlLinhasSaltarBFim = CInt(integridade)
        Case 82
             If integridade = 1 Then GlUsaEstoqueSeguranca = True Else GlUsaEstoqueSeguranca = False
         Case Is = 83
            If integridade = 1 Then GlPermitirVendaEstoqueNegativo = True Else GlPermitirVendaEstoqueNegativo = False
         Case Is = 84
            If integridade = 1 Then GlBaixarEstoquenoPedido = True Else GlBaixarEstoquenoPedido = False
         Case Is = 85
            If integridade = 1 Then GlMostraMsgClientePedido = True Else GlMostraMsgClientePedido = False
      End Select
   a = a + 1
Loop
If Not RsOp.EOF Then
   GlSenhaLiberacao = RsOp!Fantasia
   GlSenhaDebito = RsOp("nome")
   GlSenhaCredito = RsOp("END")
   GlLiberaPedidoVendas = RsOp("cidade")
Else
   GlSenhaLiberacao = ""
End If

Close #LcArq

NumeroMsg = FreeFile
For a = Len(GLBase) To 1 Step -1
    letra = Mid(GLBase, a, 1)
    If letra = "\" Then
       LcArqMsg = Mid(GLBase, 1, a) & "msg.txt"
       Exit For
    End If
Next
'=== Busca a Menssagem de Pedido
Open LcArqMsg For Input As #NumeroMsg
GlMsg = ""
a = 1
err.Number = 0
Do Until EOF(NumeroMsg)
    If err.Number > 0 Then Exit Do
    Line Input #NumeroMsg, glmsga
    Select Case a
      Case Is = 1
       GlMsg = glmsga
      Case Is = 2
       'GlMsg1 = GlMsg1 & " " & glmsga
       GlMsg1 = glmsga
       Case Is = 3
       GlMsg2 = glmsga
    End Select
    a = a + 1
Loop
If Not RsAtual.EOF Then
   GlOpcaoEmpresa = RsAtual!opcao & ""
End If
Close #NumeroMsg

Exit Function
errVerif:
Close #LcArq
'MsgBox Err.Description & Err
'Resume 0
End Function
Public Function NovoReg() As Integer
LcNovo = Not LcNovo

GlCampo0 = ""
GlCampo1 = ""
GlCampo2 = ""
GlCampo3 = ""
GlCampo4 = ""
GlCampo5 = ""
GlCampo6 = ""
GlCampo7 = ""
GlCampo8 = ""
GlCampo9 = ""
GlCampo10 = ""
GlCampo11 = ""
GlCampo12 = ""
GlCampo13 = ""
GlCampo14 = ""
GlCampo15 = ""
GlCampo16 = ""
GlCampo17 = ""
GlCampo18 = ""
GlCampo19 = ""
GlCampo20 = ""
GlCampo21 = ""
GlCampo22 = ""
GlCampo23 = ""
GlCampo24 = ""
GlCampo25 = ""
GlCampo26 = ""
GlCampo27 = ""
GlCampo28 = ""
GlCampo29 = ""
GlCampo30 = ""
End Function
Function AcertaNumero(LcNumero As String, LcTamaDecimal As Long) As String
Dim LcTa, a As Long
Dim LcDecimal, LcInteiro, LCLEtra, LcZeros As String
Dim LcAchou As Integer
LcNumero = CDbl(LcNumero)
LcNumero = CStr(LcNumero)

If LcTamaDecimal = 0 Then LcTamaDecimal = 2
LcTa = Len(LcNumero)
For a = 1 To LcTa
    LCLEtra = Mid$(LcNumero, a, 1)
    If LCLEtra = "," Or LCLEtra = "." Then
       LcAchou = True
    End If
    If LcAchou Then
       If LCLEtra <> "," Or LCLEtra = "." Then
          LcDecimal = LcDecimal & LCLEtra
       End If
    Else
       LcInteiro = LcInteiro & LCLEtra
    End If
Next
If Len(LcInteiro) = 0 Then LcInteiro = "0"
For a = 1 To LcTamaDecimal
    LcZeros = LcZeros & "0"
Next
LcDecimal = Left(LcDecimal & LcZeros, LcTamaDecimal)
AcertaNumero = LcInteiro & "," & LcDecimal


End Function
Public Function RegistroAtual(LcTabl As Tabela)
On Error GoTo ErroAtual

LcRegAtual = True

If Not GlMov Then
 If RsAtual.BOF Then
   LcResposta = MsgBox("Não Existe Registros Cadastrados, deseja incluir um Novo ?" _
   , vbExclamation + vbYesNo, "Novo Registro")
   If LcResposta = 6 Then
      LcTipoDados = 1
      Call NovoReg
      GoTo SaiRegAtual
   Else
     NovoReg
    '  Unload GlFormA
    '  FrmPrincipal.Show
      GoTo SaiRegAtual
   End If
  End If
 End If
Select Case LcTabl
        Case Is = Cliente
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo1 = RsAtual!razaosoc & ""
            GlCampo2 = RsAtual!Fantasia & ""
            GlCampo3 = RsAtual!End & ""
            'GlCampo4 = RsAtual!numero & ""
            GlCampo15 = RsAtual!Contato & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone1 & ""
            GlCampo11 = RsAtual!Fone2 & ""
            GlCampo30 = RsAtual!CGC & ""
            GlCampo12 = RsAtual!INSCEST & ""
            GlCampo13 = RsAtual!Fax & ""
            GlCampo17 = RsAtual!LimiteCredito & ""
            GlCampo23 = RsAtual!CondicaoEspecial & ""
            GlCampo20 = RsAtual!CreditoUtilizado & ""
            GlCampo22 = RsAtual!TelemarketingAtende & ""
            GlCampo31 = RsAtual!Aniversario & ""
            GlCampo32 = RsAtual!Email & ""
            GlCampo26 = RsAtual!cpf & ""
            GlCampo27 = RsAtual!rg & ""
            GlCampo35 = RsAtual!dadosnota & ""
            'GlCampo19 = RsAtual!DataUltimaVisita & ""
            'GlCampo25 = RsAtual!Observacao & ""
            'GlCampo30 = RsAtual!cpf & ""
            
        Case Is = PropostaCliente
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo27 = RsAtual!nproposta & ""
            GlCampo4 = RsAtual!loja & ""
            GlCampo1 = RsAtual!razaosoc & ""
            GlCampo3 = RsAtual!End & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone & ""
            GlCampo21 = RsAtual!filial & ""
            GlCampo22 = RsAtual!nomeloja & ""
            GlCampo28 = RsAtual!foneloja & ""
            GlCampo30 = RsAtual!cpf & ""
            GlCampo12 = RsAtual!rg & ""
            GlCampo29 = RsAtual!dataemissao & ""
            GlCampo26 = RsAtual!vendedor & ""
            GlCampo24 = RsAtual!produto & ""
            GlCampo31 = RsAtual!orgao & ""
            GlCampo32 = RsAtual!ufrg & ""
            GlCampo33 = RsAtual!datanascimento & ""
            GlCampo15 = RsAtual!pai & ""
            GlCampo34 = RsAtual!mae & ""
            GlCampo35 = RsAtual!nacionalidade & ""
            GlCampo86 = RsAtual!fonedevsolid & ""
            GlCampo36 = RsAtual!cartaocredito & ""
            GlCampo37 = RsAtual!numerocartao & ""
            GlCampo38 = RsAtual!qtdeveiculo & ""
            GlCampo38 = RsAtual!outraspropried & ""
            GlCampo40 = RsAtual!endnumero & ""
            GlCampo41 = RsAtual!Complemento & ""
            GlCampo14 = RsAtual!temporesid & ""
            GlCampo11 = RsAtual!empresatrabalha & ""
            GlCampo13 = RsAtual!salarioliq & ""
            GlCampo42 = RsAtual!temposerv & ""
            GlCampo2 = RsAtual!cargo & ""
            GlCampo17 = RsAtual!cnpjemppropria & ""
            GlCampo49 = RsAtual!endempresa & ""
            GlCampo43 = RsAtual!numeroempresa & ""
            GlCampo20 = RsAtual!complempresa & ""
            GlCampo46 = RsAtual!ufempresa & ""
            GlCampo47 = RsAtual!cidadeempresa & ""
            GlCampo48 = RsAtual!bairroempresa & ""
            GlCampo45 = RsAtual!cepempresa & ""
            GlCampo44 = RsAtual!dddtelramalemp & ""
            GlCampo50 = RsAtual!nomeref1 & ""
            GlCampo51 = RsAtual!foneref1 & ""
            GlCampo52 = RsAtual!nomeref2 & ""
            GlCampo53 = RsAtual!foneref2 & ""
            GlCampo54 = RsAtual!tarifa & ""
            GlCampo55 = RsAtual!Tabela & ""
            GlCampo56 = RsAtual!nparcelas & ""
            GlCampo57 = RsAtual!datacontrato & ""
            GlCampo58 = RsAtual!carencia & ""
            GlCampo59 = RsAtual!vrcompra & ""
            GlCampo60 = RsAtual!tar & ""
            GlCampo61 = RsAtual!valorentrada & ""
            GlCampo62 = RsAtual!vrtarentrada & ""
            GlCampo63 = RsAtual!valorprestacao & ""
            GlCampo64 = RsAtual!valortotalprazo & ""
            GlCampo5 = RsAtual!primvencimento & ""
            GlCampo16 = RsAtual!ultimovencimento & ""
            GlCampo18 = RsAtual!taxaam & ""
            GlCampo25 = RsAtual!taxaaa & ""
            GlCampo19 = RsAtual!banco & ""
            GlCampo23 = RsAtual!Agencia & ""
            GlCampo66 = RsAtual!contacorrente & ""
            GlCampo65 = RsAtual!desde & ""
            GlCampo68 = RsAtual!primcheque & ""
            GlCampo67 = RsAtual!ultcheque & ""
            GlCampo69 = RsAtual!descrbem & ""
            GlCampo70 = RsAtual!nomeconjuge & ""
            GlCampo72 = RsAtual!naturalidadeconjuge & ""
            GlCampo73 = RsAtual!nacionalidadeconjuge & ""
            GlCampo74 = RsAtual!nascimconjuge & ""
            GlCampo75 = RsAtual!rgconjuge & ""
            GlCampo76 = RsAtual!orgaoconjuge & ""
            GlCampo77 = RsAtual!emissaoconjuge & ""
            GlCampo78 = RsAtual!empresaconjuge & ""
            GlCampo79 = RsAtual!telempconj & ""
            GlCampo80 = RsAtual!rendabrconj & ""
            GlCampo81 = RsAtual!nomedevsolid & ""
            GlCampo82 = RsAtual!cpfdevsolid & ""
            GlCampo83 = RsAtual!rgdevsolid & ""
            GlCampo84 = RsAtual!nascdevsolid & ""
            GlCampo85 = RsAtual!enddevsolid & ""
            
        Case Is = Galpao
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo1 = RsAtual!Nome & ""
            GlCampo2 = RsAtual!Contato & ""
            GlCampo3 = RsAtual!End & ""
            'GlCampo4 = RsAtual!numero & ""
            'GlCampo5 = RsAtual!complemento & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone1 & ""
            GlCampo11 = RsAtual!Fone2 & ""
            GlCampo12 = RsAtual!CGC & ""
            GlCampo13 = RsAtual!Fax & ""
            GlCampo30 = RsAtual!INSCEST & ""
            'GlCampo14 = RsAtual!Celular & ""
            'GlCampo15 = RsAtual!Email & ""
            'GlCampo16 = RsAtual!DataUltimaCompra & ""
            'GlCampo18 = RsAtual!ValorUltCompra & ""
            'GlCampo19 = RsAtual!DataUltimaVisita & ""
            'GlCampo25 = RsAtual!Observacao & ""
            'GlCampo30 = RsAtual!cpf & ""
    
        Case Is = fornecedor
         If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo1 = RsAtual!razaosoc & ""
            GlCampo2 = RsAtual!Fantasia & ""
            GlCampo3 = RsAtual!End & ""
            GlCampo4 = RsAtual!COMISSAOREPRESENTANTE & ""
            GlCampo5 = RsAtual!Complemento & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone1 & ""
            GlCampo11 = RsAtual!Fone2 & ""
            GlCampo30 = RsAtual!CGC & ""
            GlCampo13 = RsAtual!Fax & ""
            GlCampo12 = RsAtual!INSCEST & ""
            GlCampo31 = RsAtual!Numero & ""
            GlCampo32 = RsAtual!Email & ""
            'GlCampo16 = RsAtual!DataUltimaCompra & ""
            'GlCampo18 = RsAtual!ValorUltCompra & ""
            'GlCampo19 = RsAtual!DataUltimaVisita & ""
            'GlCampo25 = RsAtual!Observacao & ""
            'GlCampo30 = RsAtual!cpf & ""
             
          
        Case Is = produto
            If Len(Trim(RsAtual!cod)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!cod & ""
            End If
           GlCampo1 = RsAtual!Nome & ""
          ' GlCampo2 = RsAtual!Descricao & ""
           GlCampo3 = RsAtual!cst & ""
           GlCampo21 = RsAtual!percentualcusto & ""
           'GlCampo5 = RsAtual!Estoque & ""
           'GlCampo6 = RsAtual!Reservado & ""
           GlCampo8 = RsAtual!MPVENDA & ""
           GlCampo7 = RsAtual!ipi & ""
           GlCampo22 = RsAtual!minimo & ""
           GlCampo10 = RsAtual!maximo & ""
           GlCampo9 = RsAtual!QuantUnidade & ""
           'GlCampo15 = RsAtual!Observacoes & ""
           GlCampo18 = RsAtual!ComissaoFornecedor & ""
           'GlCampo19 = RsAtual!Fornecedor & ""
           GlCampo13 = RsAtual!UNIMED & ""
           GlCampo12 = RsAtual!PMU & ""
           GlCampo14 = RsAtual!Ptab & ""
           GlCampo16 = RsAtual!QTDUNIMED & ""
           GlCampo17 = RsAtual!Lucro & ""
           GlCampo19 = RsAtual!fornecedor & ""
           GlCampo20 = RsAtual!QuantEstoque & ""
           GlCampo23 = RsAtual!Custo & ""
           
       Case Is = Cidade
            If Len(Trim(RsAtual!cod)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!cod & ""
            End If
            GlCampo1 = RsAtual!Nome & ""
            GlCampo8 = RsAtual!Estado & ""
            FrmCidade.CodigoIbgeCidade.Text = RsAtual!CodigoIBGEMunicipio & ""
            FrmCidade.CodigoIBGEEstado.Text = RsAtual!CodigoIBGEEstado & ""
       Case Is = monetario
            If Len(Trim(RsAtual!TPMONET)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!TPMONET & ""
            End If
            GlCampo1 = RsAtual!XTPMONET & ""
            GlCampo8 = RsAtual!venda & ""
            GlCampo2 = RsAtual!COMPRA & ""
            GlCampo19 = RsAtual!vp & ""
            GlCampo30 = RsAtual!MOVCAIXA & ""
            FrmTipoMonetario.DescricaoNFE.Text = RsAtual!DescricaoNFE & ""
             
       Case Is = tiporec
       
            If Len(Trim(RsAtual!cod)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!cod & ""
            End If
           GlCampo1 = RsAtual!Nome & ""
           GlCampo8 = RsAtual!rd & ""
       Case Is = Custo
       
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
           GlCampo1 = RsAtual!Descricao & ""
           
      Case Is = Unidade
       
            If Len(Trim(RsAtual!cod)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!cod & ""
            End If
           GlCampo1 = RsAtual!Nome & ""
           GlCampo8 = RsAtual!Simbolo & ""
           GlCampo4 = RsAtual!quantidade & ""
     Case Is = Funcionario
     If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo1 = RsAtual!Nome & ""
            GlCampo2 = RsAtual!Comissao & ""
            GlCampo3 = RsAtual!End & ""
            GlCampo4 = RsAtual!datanascimento & ""
           ' GlCampo5 = RsAtual!complemento & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone & ""
           ' GlCampo11 = RsAtual!cpf & ""
           ' GlCampo12 = RsAtual!rg & ""
            GlCampo13 = RsAtual!CarteiraTrabalho & ""
           ' GlCampo14 = RsAtual!pai & ""
           ' GlCampo15 = RsAtual!mae & ""
           ' GlCampo16 = RsAtual!Funcao & ""
           ' GlCampo17 = RsAtual!Horario & ""
           ' GlCampo18 = RsAtual!Salario & ""
           ' GlCampo19 = RsAtual!Comissao & ""
            GlCampo20 = RsAtual!dataadmissao & ""
            GlCampo21 = RsAtual!datademissao & ""
           ' GlCampo22 = RsAtual!Observacao & ""
    Case Is = Receber

            
            If Len(Trim(RsAtual!NF)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!NF & ""
            End If
           GlCampo20 = RsAtual!Valor & ""
           GlCampo2 = RsAtual!Cliente & ""
           'GlCampo3 = RsAtual!CodigoCliente & ""
           GlCampo4 = RsAtual!DTVENC & ""
           GlCampo5 = RsAtual!TPMONET & ""
           'GlCampo6 = RsAtual!DataRec & ""
           GlCampo7 = RsAtual!DTPAGTO & ""
           GlCampo8 = RsAtual!VALPAGO & ""
           GlCampo9 = RsAtual!Data & ""
           GlCampo1 = RsAtual!Emitente & ""
           GlCampo10 = RsAtual!Obs & ""
           GlCampo11 = RsAtual!codDesp & ""
           GlCampo12 = RsAtual!NomeDesp & ""
          
    Case Is = Cheques

            
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            GlCampo5 = RsAtual!cheque & ""
           GlCampo2 = RsAtual!pedido & ""
           GlCampo1 = RsAtual!Cliente & ""
           GlCampo3 = RsAtual!Agencia & ""
           GlCampo6 = RsAtual!banco & ""
           GlCampo23 = RsAtual!emissao & ""
           GlCampo10 = RsAtual!Emitente & ""
           GlCampo8 = RsAtual!passadopara & ""
           'GlCampo7 = RsAtual!DTPAGTO & ""
           GlCampo22 = RsAtual!dataentrada & ""
           GlCampo9 = RsAtual!Valor & ""
           GlCampo24 = RsAtual!datadeposito & ""
           If RsAtual!Compensado Then
              GlCampo25 = "1"
           Else
              GlCampo25 = "0"
           End If
    Case Is = pagar
      
           If Len(Trim(RsAtual!NF)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!NF & ""
            End If
           GlCampo1 = RsAtual!Valor & ""
           GlCampo2 = RsAtual!credor & ""
           GlCampo4 = RsAtual!DTVENC & ""
           GlCampo5 = RsAtual!TPMONET & ""
           GlCampo7 = RsAtual!DTPAGTO & ""
           If GlCampo7 = "01/01/11" Then GlCampo7 = ""
           GlCampo8 = RsAtual!VALPAGO & ""
           GlCampo9 = RsAtual!Data & ""
           GlCampo10 = RsAtual!Obs & ""
           GlCampo11 = RsAtual!codDesp & ""
           GlCampo12 = RsAtual!NomeDesp & ""
     Case Is = Transportadora
            If Len(Trim(RsAtual!codigo)) = 0 Then
               GlCampo0 = 0
            Else
               GlCampo0 = RsAtual!codigo & ""
            End If
            
            GlCampo1 = RsAtual!razaosoc & ""
            GlCampo2 = RsAtual!Fantasia & ""
            GlCampo3 = RsAtual!End & ""
            GlCampo6 = RsAtual!Bairro & ""
            GlCampo7 = RsAtual!Cidade & ""
            GlCampo8 = RsAtual!Estado & ""
            GlCampo9 = RsAtual!Cep & ""
            GlCampo10 = RsAtual!Fone1 & ""
            GlCampo11 = RsAtual!Fone2 & ""
            GlCampo30 = RsAtual!CGC & ""
            GlCampo13 = RsAtual!Fax & ""
            GlCampo12 = RsAtual!INSCEST & ""
    
End Select

SaiRegAtual:
'FechaBanco
RetornaCorFundo
Exit Function
ErroAtual:
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      
      Resume Next
   Else
      'MsgBox Err.Description & " " & Err
   End If
   'MsgBox Err.Description & Err.Number
  ' Resume 0
   Resume Next
End If


End Function
Public Function SalvaRegistro(LcTabl As Tabela)
 On Error Resume Next
 Dim GlPes As Integer
 Dim lcpe As String
  Dim Resposta As Integer
    GlPes = GLPesquisa
    GlIniceAtual = Screen.ActiveControl.Index
    If LcAlterado Then
        
        
        If GlConfirmaAlteracao Then
           Resposta = MsgBox("Confirma a Alteração deste registro?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Aviso")
        Else
           Resposta = vbYes
        End If
        If Resposta = vbYes Then
            If LcNovo Then
                'Adiciona registro
                Call IncluiNovoRegistro(LcTabl, False)
            Else
                Call IncluiNovoRegistro(LcTabl, False)
            End If
        ElseIf Resposta = vbCancel Then
            LcAlterado = False
            Exit Function
        Else
            DesfazMudanca (LcTabl)
        End If
        
    Else
        If GlPergunta Then
           Resposta = MsgBox("Salva o Novo registro?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Aviso")
           If Resposta = vbYes Then
              Call IncluiNovoRegistro(LcTabl, True)
           End If
        Else
           Call IncluiNovoRegistro(LcTabl, True)
        End If
    End If
    LcAlterado = False
    'LcRegAtual = False
    Screen.ActiveForm.txt(1).SetFocus
    Screen.ActiveForm.CmdSalvar.Enabled = False
    GLPesquisa = GlPes
    
    If GlPes Then
       AbreBanco (LcTabl)
       Select Case LcTabl
              Case Is = produto
                  lcpe = "cod='" & Screen.ActiveForm.txt(0) & "'"
              Case Is = Unidade
                  lcpe = "cod='" & Screen.ActiveForm.txt(0) & "'"
              Case Is = Cidade
                  lcpe = "cod='" & Screen.ActiveForm.txt(0) & "'"
              Case Else
                  lcpe = "codigo='" & Screen.ActiveForm.txt(0) & "'"
       End Select
       RsAtual.Find lcpe
    End If
End Function
Public Function CalculaCodigo()
On Error Resume Next

   Select Case GlFormA.Name
   
   Case Is = "FrmCusto"
      If Not GLCalculacodigoProduto Then Exit Function
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
   Case Is = "FrmProduto"
      If Not GLCalculacodigoProduto Then Exit Function
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!cod) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   Case Is = "FrmVales"
      If Not GLCalculacodigoProduto Then Exit Function
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!cod) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   
   Case Is = "FrmCliente"
      If Not GLCalculacodigoCliente Then Exit Function
      If Not RsAtual.EOF Then
       RsAtual.MoveLast
       GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   Case Is = "FrmPropostaCompra"
      If Not GLCalculacodigoCliente Then Exit Function
      If Not RsAtual.EOF Then
       RsAtual.MoveLast
       GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   Case Is = "FrmFornecedor"
      If Not GLCalculacodigoFornecedor Then Exit Function
      If Not RsAtual.EOF Then
       RsAtual.MoveLast
       GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   Case Is = "FrmFuncionario"
      If Not RsAtual.EOF Then
       RsAtual.MoveLast
       GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
   Case Is = "FrmCidade"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!cod) + 1
      Else
         GlCampo0 = "0001"
      End If
      GlCampo0 = Right("0000" & GlCampo0, 4)
      
   Case Is = "FrmTipoMonetario"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!TPMONET) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
  Case Is = "FrmTiporeceita"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!cod) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
  Case Is = "FrmUnidade"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!cod) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
  Case Is = "FrmGalpao"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
 Case Is = "FrmUnidade"
      If Not RsAtual.EOF Then
         RsAtual.MoveLast
         GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "01"
      End If
      GlCampo0 = Right("00" & GlCampo0, 2)
      
 Case Is = "FrmTransportadora"
     
      If Not RsAtual.EOF Then
       RsAtual.MoveLast
       GlCampo0 = Val(RsAtual!codigo) + 1
      Else
         GlCampo0 = "00001"
      End If
      GlCampo0 = Right("00000" & GlCampo0, 5)
 End Select


Screen.ActiveForm.txt(0).Text = GlCampo0
End Function

Public Function IncluiNovoRegistro(LcTbl As Tabela, Adiciona As Integer)
On Error GoTo ErroIncl
Dim LcCodigo As Long, LcIndiceAnterior, LcNomeArquivo As String
Dim LcData As Date
Dim RsCaixa As Recordset
Dim LcNumero As Integer
 LcIndiceAnterior = LcIndice
 
 If GlFormA.Name <> "Receitas" And GlFormA.Name <> "Despesas" Then
    LcIndice = "CODIGO"
 End If
  
 If LcTipoDados = 1 Then
    AbreBanco (LcTbl)
    CalculaCodigo
    RsAtual.AddNew
 Else
    If GlAlteraCodigo Then
       GlChave = GlCodigoAnterior
    Else
       GlChave = GlCampo0
    End If
    GlAlteraCodigo = False
    If GlFormA.Name = "Frmcheques" Then
       LcIndice = "Cheque"
    End If
    'MsgBox GLPesquisa
    
    
    
    
    GLPesquisa = False
    
    AbreBanco (LcTbl)
   ' RsAtual.Index = LcIndice
    If Not AchaReg(1) Then Exit Function
    RsAtual.Edit
 End If
 If Not VerificaRequerido Then FechaBanco: Exit Function
 
  'A função VerificaTipo Converte o Formato dos Dados Antes de Gravar.
 'Deve-se observar Que o Valor inteiro corresponde ao numero do final do campo
  Maiuscula
  If LcTbl = 200 Then
     If LcTipoDados = 1 Then
        LcData = Format(GlDataSistema, "dd/mm/yyyy")
        Set RsCaixa = Dbbase.OpenRecordset("Caixa", dbOpenTable, dbSeeChanges, dbOptimistic)
        RsCaixa.Index = "DataDoLancamento"
        RsCaixa.Seek "=", LcData
        If Not RsCaixa.NoMatch Then
           RsCaixa.Edit
           If IsEmpty(RsCaixa!CLientesNovos) Then
              RsCaixa!CLientesNovos = 1
           Else
              RsCaixa!CLientesNovos = RsCaixa!CLientesNovos + 1
           End If
           RsCaixa.Update
        Else
           RsCaixa.AddNew
           RsCaixa!CLientesNovos = 1
           RsCaixa!DataMovimento = LcData
           RsCaixa!fechado = False
           RsCaixa.Update
        End If
        RsCaixa.Close
        
     End If
     
  End If
  
      
  Select Case LcTbl
        Case Is = Cliente
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!razaosoc = VerificaTipo(1, GlCampo1)
            RsAtual!Fantasia = VerificaTipo(2, GlCampo2)
            RsAtual!End = VerificaTipo(3, GlCampo3)
            RsAtual!Bairro = VerificaTipo(6, GlCampo6)
            RsAtual!Cidade = GlCampo7
            RsAtual!Estado = VerificaTipo(8, GlCampo8)
            RsAtual!Cep = GlCampo9
            RsAtual!Fone1 = VerificaTipo(10, GlCampo10)
            RsAtual!Fone2 = VerificaTipo(11, GlCampo11)
            RsAtual!CGC = GlCampo30
            RsAtual!INSCEST = GlCampo12
            RsAtual!CondicaoEspecial = VerificaTipo(23, GlCampo23)
            If Len(GlCampo17) > 0 Then If IsNumeric(GlCampo17) Then RsAtual!LimiteCredito = VerificaTipo(17, GlCampo17)
            RsAtual!CreditoUtilizado = VerificaTipo(20, GlCampo20)
            RsAtual!Fax = VerificaTipo(13, GlCampo13)
            RsAtual!TelemarketingAtende = GlCampo22
            RsAtual!Contato = VerificaTipo(15, GlCampo15)
            RsAtual!Aniversario = GlCampo31
            RsAtual!Email = GlCampo32
            RsAtual!cpf = GlCampo26
            RsAtual!rg = GlCampo27
            RsAtual!Numero = FrmCliente.Numero.Text
            If FrmCliente.Comodato.Value = 1 Then
               RsAtual!Comodato = True
            Else
               RsAtual!Comodato = False
            End If
            RsAtual!dadosnota = GlCampo35
            RsAtual!TipoContribuinte = FrmCliente.TipoContr.Text & ""
            RsAtual!InscricaoSuframa = FrmCliente.Suframa.Text & ""
            RsAtual!InscricaoMunicipal = FrmCliente.InscMunic.Text & ""
            RsAtual!EmailFinanceiro = FrmCliente.EmailFinanceiro.Text & ""
            RsAtual!GrupoEconomicoID = FrmCliente.Get_ID_Gr_Economico(FrmCliente.GrupoEconomico.Text)
            RsAtual!GrupoEconomicoNome = FrmCliente.GrupoEconomico.Text
            If IsDate(FrmCliente.DataCadastro.Text) Then
                RsAtual!DataCadastro = FrmCliente.DataCadastro.Text
            End If
            If FrmCliente.Bloqueado.Value = 1 Then
               RsAtual!Bloqueado = True
            Else
               RsAtual!Bloqueado = False
            End If
        Case Is = Galpao
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!Nome = VerificaTipo(1, GlCampo1)
            RsAtual!Contato = VerificaTipo(2, GlCampo2)
            RsAtual!End = VerificaTipo(3, GlCampo3)
            RsAtual!Bairro = VerificaTipo(6, GlCampo6)
            RsAtual!Cidade = GlCampo7
            RsAtual!Estado = VerificaTipo(8, GlCampo8)
            RsAtual!Cep = VerificaTipo(9, GlCampo9)
            RsAtual!Fone1 = VerificaTipo(10, GlCampo10)
            RsAtual!Fone2 = VerificaTipo(11, GlCampo11)
            RsAtual!CGC = VerificaTipo(12, GlCampo12)
            RsAtual!Fax = VerificaTipo(13, GlCampo13)
            RsAtual!INSCEST = VerificaTipo(30, GlCampo30)
        Case Is = fornecedor
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!razaosoc = VerificaTipo(1, GlCampo1)
            RsAtual!Fantasia = VerificaTipo(2, GlCampo2)
            RsAtual!End = VerificaTipo(3, GlCampo3)
            If Len(GlCampo4) Then If IsNumeric(GlCampo4) Then RsAtual!COMISSAOREPRESENTANTE = VerificaTipo(4, GlCampo4)
            RsAtual!Bairro = VerificaTipo(6, GlCampo6)
            RsAtual!Cidade = GlCampo7
            RsAtual!Estado = VerificaTipo(8, GlCampo8)
            RsAtual!Cep = VerificaTipo(9, GlCampo9)
            RsAtual!Fone1 = VerificaTipo(10, GlCampo10)
            RsAtual!Fone2 = VerificaTipo(11, GlCampo11)
            
            CNPJ = FrmFornecedor.txt(30).Text & ""
            CNPJ = Replace(CNPJ, ",", "")
            CNPJ = Replace(CNPJ, ".", "")
            CNPJ = Replace(CNPJ, "-", "")
            CNPJ = Replace(CNPJ, "/", "")
            CNPJ = Replace(CNPJ, "\", "")
            CNPJ = Replace(CNPJ, " ", "")
            
            cpf = FrmFornecedor.txt(17).Text & ""
            cpf = Replace(cpf, ",", "")
            cpf = Replace(cpf, ".", "")
            cpf = Replace(cpf, "-", "")
            cpf = Replace(cpf, "/", "")
            cpf = Replace(cpf, "\", "")
            cpf = Replace(cpf, " ", "")
            
            RsAtual!CGC = CNPJ
            RsAtual!INSCEST = FrmFornecedor.txt(12).Text & ""
            RsAtual!Fax = VerificaTipo(13, GlCampo13)
            RsAtual!cpf = cpf ' FrmFornecedor.Txt(17).Text
            RsAtual!Contato = FrmFornecedor.Contato.Text & ""
            RsAtual!Email = FrmFornecedor.Email.Text & ""
            RsAtual!Numero = FrmFornecedor.Numero.Text & ""
            RsAtual!Complemento = FrmFornecedor.txt(5).Text & ""
        Case Is = produto
         'With FrmProduto
          '  RsAtual!cod = .Txt(0).Text & ""
          '  RsAtual!Nome = .Txt(1).Text & ""
          '  RsAtual!cst = .Txt(3).Text & ""
          '  If Len(.Nome.Text) = 0 Then .Nome.Text = "."
          '  RsAtual!fornecedor = .Nome.Text & ""  'GlCampo19
          '  If Len(Trim(.Txt(20).Text)) = 0 Then .Txt(20).Text = 0
          '  RsAtual!QuantEstoque = CDbl(.Txt(20).Text)
          '  If Len(Trim(.Txt(7).Text)) = 0 Then .Txt(7).Text = 0
          '  RsAtual!ipi = CDbl(.Txt(7).Text)
          '  RsAtual!Percentualcusto = .valor(5).Text & ""
          '  If Len(Trim(.valor(2).Text)) = 0 Then .valor(2).Text = 0
          '  RsAtual!MPVENDA = CDbl(.valor(2).Text)
          '  If Len(Trim(.valor(4).Text)) = 0 Then .valor(4).Text = 0
          '  RsAtual!minimo = CDbl(.valor(4).Text)
          '  If Len(Trim(.Txt(9).Text)) = 0 Then .Txt(9).Text = 0
          '  RsAtual!quantUnidade = CDbl(.Txt(9).Text)
          '  If Len(Trim(.Txt(10).Text)) = 0 Then .Txt(10).Text = 0
          '  RsAtual!Maximo = CDbl(.Txt(10).Text)
          '  'If Len(Trim(GlCampo18)) = 0 Then GlCampo18 = 0
          '  RsAtual!ComissaoFornecedor = .Txt(18).Text & ""
          '  If Len(Trim(.valor(3).Text)) = 0 Then .valor(3).Text = 0
          '  RsAtual!PMU = CDbl(.valor(3).Text)
          '  RsAtual!Unimed = .Txt(13).Text & ""
          '  If Len(Trim(.valor(1).Text)) = 0 Then .valor(1).Text = 0
          '  RsAtual!Ptab = CDbl(.valor(1).Text)
          '  If Len(Trim(.valor(0).Text)) = 0 Then .valor(0).Text = 0
          '  RsAtual!Lucro = CDbl(.valor(0).Text)
          '  If Len(Trim(.Txt(16).Text)) = 0 Then .Txt(16).Text = 0
          '  RsAtual!QTDUNIMED = .Txt(16).Text & ""
          '  RsAtual!Percentualcusto = .valor(5).Text & ""
          '  If Len(Trim(.valor(6).Text)) = 0 Then .valor(6).Text = 0
          '  RsAtual!Custo = CDbl(.valor(6).Text)
          ' End With
            'RsAtual!cod = VerificaTipo(0, GlCampo0)
            'RsAtual!nome = VerificaTipo(1, GlCampo1)
            'RsAtual!cst = VerificaTipo(3, GlCampo3)
            'RsAtual!Fornecedor = GlCampo19
            'If Len(Trim(GlCampo20)) = 0 Then GlCampo20 = 0
            'RsAtual!QuantEstoque = GlCampo20
            'If Len(Trim(GlCampo7)) = 0 Then GlCampo7 = 0
            'RsAtual!ipi = GlCampo7
            'RsAtual!Percentualcusto = GlCampo21
            'If Len(Trim(FrmProduto.valor(2).Text)) = 0 Then GlCampo8 = 0
            'RsAtual!MPVENDA = GlCampo8
            'If Len(Trim(GlCampo22)) = 0 Then GlCampo22 = 0
            'RsAtual!minimo = GlCampo22
            'If Len(Trim(GlCampo9)) = 0 Then GlCampo9 = 0
            'RsAtual!quantUnidade = GlCampo9
            'If Len(Trim(GlCampo10)) = 0 Then GlCampo10 = 0
            'RsAtual!Maximo = CDbl(GlCampo10)
            'If Len(Trim(GlCampo18)) = 0 Then GlCampo18 = 0
            
            'RsAtual!ComissaoFornecedor = GlCampo18
            'If Len(Trim(GlCampo12)) = 0 Then GlCampo12 = 0
            'RsAtual!PMU = CDbl(GlCampo12)
            'RsAtual!Unimed = VerificaTipo(13, GlCampo13)
            'If Len(Trim(GlCampo14)) = 0 Then GlCampo14 = 0
            'RsAtual!Ptab = GlCampo14
            'If Len(Trim(GlCampo17)) = 0 Then GlCampo17 = 0
            'RsAtual!Lucro = GlCampo17
            'If Len(Trim(GlCampo16)) = 0 Then GlCampo16 = 0
            'RsAtual!QTDUNIMED = GlCampo16
            'If Len(Trim(GlCampo23)) = 0 Then GlCampo23 = 0
            'RsAtual!Custo = GlCampo23
            
        Case Is = Cidade
             If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            RsAtual!cod = GlCampo0
            RsAtual!CodigoIBGEMunicipio = FrmCidade.CodigoIbgeCidade.Text
            RsAtual!CodigoIBGEEstado = FrmCidade.CodigoIBGEEstado.Text
            RsAtual!Nome = GlCampo1
            RsAtual!Estado = GlCampo8
            
        Case Is = Custo
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            RsAtual!codigo = GlCampo0
            RsAtual!Descricao = GlCampo1
            
        Case Is = monetario
            RsAtual!TPMONET = VerificaTipo(0, GlCampo0)
            RsAtual!XTPMONET = VerificaTipo(1, GlCampo1)
            RsAtual!venda = VerificaTipo(8, GlCampo8)
            RsAtual!COMPRA = VerificaTipo(2, GlCampo2)
            RsAtual!vp = VerificaTipo(19, GlCampo19)
            RsAtual!MOVCAIXA = VerificaTipo(30, GlCampo30)
            RsAtual!DescricaoNFE = FrmTipoMonetario.DescricaoNFE.Text
        Case Is = Unidade
            RsAtual!cod = VerificaTipo(0, GlCampo0)
            RsAtual!Nome = VerificaTipo(1, GlCampo1)
            RsAtual!Simbolo = VerificaTipo(1, GlCampo8)
            If Len(GlCampo4) = 0 Then GlCampo4 = 0
            RsAtual!quantidade = GlCampo4
            
        Case Is = tiporec
            RsAtual!cod = VerificaTipo(0, GlCampo0)
            RsAtual!Nome = VerificaTipo(1, GlCampo1)
            RsAtual!rd = VerificaTipo(1, GlCampo8)
        Case Is = Convenio
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!Nome = VerificaTipo(1, GlCampo1)
            RsAtual!Contato = VerificaTipo(2, GlCampo2)
            RsAtual!Fone = VerificaTipo(3, GlCampo3)
            RsAtual!FoneOPC = VerificaTipo(4, GlCampo4)
            RsAtual!Fax = VerificaTipo(5, GlCampo5)
            RsAtual!RUA = VerificaTipo(6, GlCampo6)
            RsAtual!Numero = VerificaTipo(7, GlCampo7)
            RsAtual!Bairro = VerificaTipo(8, GlCampo8)
            RsAtual!Cidade = VerificaTipo(9, GlCampo9)
            RsAtual!Estado = VerificaTipo(10, GlCampo10)
            RsAtual!Cep = VerificaTipo(11, GlCampo11)
            If Len(GlCampo12) > 0 Then If IsDate(GlCampo21) Then RsAtual!DataInicio = VerificaTipo(12, GlCampo12)
            If Len(GlCampo13) > 0 Then If IsDate(GlCampo13) Then RsAtual!DiaVencimento = VerificaTipo(13, GlCampo13)
            If Len(GlCampo14) > 0 Then If IsNumeric(GlCampo14) Then RsAtual!Desconto = VerificaTipo(14, GlCampo14)
            RsAtual!Obs = VerificaTipo(15, GlCampo15)
      Case Is = Funcionario
            
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!Nome = VerificaTipo(1, GlCampo1)
            RsAtual!Comissao = VerificaTipo(2, GlCampo2)
            RsAtual!End = VerificaTipo(3, GlCampo3)
            If Len(GlCampo4) > 0 Then If IsDate(GlCampo4) Then RsAtual!datanascimento = VerificaTipo(4, GlCampo4)
            RsAtual!Bairro = VerificaTipo(6, GlCampo6)
            RsAtual!Cidade = GlCampo7
            RsAtual!Estado = VerificaTipo(8, GlCampo8)
            RsAtual!Cep = VerificaTipo(9, GlCampo9)
            RsAtual!Fone = VerificaTipo(10, GlCampo10)
            RsAtual!CarteiraTrabalho = VerificaTipo(13, GlCampo13)
            If Len(GlCampo20) > 0 Then If IsDate(GlCampo20) Then RsAtual!dataadmissao = VerificaTipo(20, GlCampo20)
            If Len(GlCampo21) > 0 Then If IsDate(GlCampo21) Then RsAtual!datademissao = VerificaTipo(21, GlCampo21)
      
      Case Is = Receber
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            RsAtual!NF = VerificaTipo(0, GlCampo0)
            
            If Len(GlCampo20) = 0 Then
               RsAtual!Valor = 0
            Else
               If IsNumeric(GlCampo20) Then RsAtual!Valor = GlCampo20
            End If
            RsAtual!Emitente = GlCampo1
            RsAtual!Cliente = VerificaTipo(2, GlCampo2)
            If Not IsNull(GlCampo4) Then If IsDate(GlCampo4) Then RsAtual!DTVENC = GlCampo4
            RsAtual!TPMONET = GlCampo5
            If Len(GlCampo7) > 0 Then If IsDate(GlCampo7) Then RsAtual!DTPAGTO = GlCampo7
            If Len(GlCampo8) > 0 Then
               If IsNumeric(GlCampo8) Then RsAtual!VALPAGO = GlCampo8
            Else
               RsAtual!VALPAGO = 0
            End If
            If Len(GlCampo9) > 0 Then If IsDate(GlCampo9) Then RsAtual!Data = GlCampo9
            RsAtual!Obs = GlCampo10
            RsAtual!codDesp = GlCampo11
            RsAtual!NomeDesp = GlCampo12
      Case Is = Cheques
            RsAtual!cheque = VerificaTipo(5, GlCampo5)
            
            RsAtual!pedido = VerificaTipo(2, GlCampo2)
            RsAtual!Cliente = GlCampo1
            RsAtual!Agencia = VerificaTipo(3, GlCampo3)
            RsAtual!banco = VerificaTipo(6, GlCampo6)
            If GlCampo23 <> "  /  /  " Then If IsDate(GlCampo23) Then RsAtual!emissao = CDate(GlCampo23)
                      
            RsAtual!Emitente = VerificaTipo(10, GlCampo10)
            If GlCampo22 <> " /  /  " Then If IsDate(GlCampo22) Then RsAtual!dataentrada = CDate(GlCampo22)
                                  
            RsAtual!passadopara = VerificaTipo(8, GlCampo8)
            If Len(GlCampo9) > 0 Then If IsNumeric(GlCampo9) Then RsAtual!Valor = GlCampo9
            If GlCampo24 <> "  /  /  " Then If IsDate(GlCampo24) Then RsAtual!datadeposito = CDate(GlCampo24)
            If Len(GlCampo25) > 0 Then RsAtual!Compensado = GlCampo25
      Case Is = pagar
            
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            RsAtual!NF = VerificaTipo(0, GlCampo0)
            If Len(GlCampo1) = 0 Then
               RsAtual!Valor = 0
            Else
               If IsNumeric(GlCampo1) Then RsAtual!Valor = GlCampo1
            End If
            RsAtual!credor = VerificaTipo(2, GlCampo2)
            If Len(GlCampo4) > 0 Then If IsDate(GlCampo4) Then RsAtual!DTVENC = GlCampo4
            RsAtual!TPMONET = GlCampo5
            If Len(GlCampo7) > 0 Then If IsDate(GlCampo7) Then RsAtual!DTPAGTO = VerificaTipo(7, GlCampo7)
            If Len(GlCampo8) > 0 Then
               If IsNumeric(GlCampo8) Then RsAtual!VALPAGO = GlCampo8 Else RsAtual!VALPAGO = 0
            Else
               RsAtual!VALPAGO = 0
            End If
            If Len(GlCampo9) > 0 Then If IsDate(GlCampo9) Then RsAtual!Data = GlCampo9
            RsAtual!Obs = GlCampo10
            RsAtual!codDesp = GlCampo11
            RsAtual!NomeDesp = GlCampo12
         Case Is = Transportadora
            If Len(Trim(GlCampo0)) = 0 Then GlCampo0 = 0
            
            RsAtual!codigo = VerificaTipo(0, GlCampo0)
            RsAtual!razaosoc = VerificaTipo(1, GlCampo1)
            RsAtual!End = VerificaTipo(3, GlCampo3)
            RsAtual!Bairro = VerificaTipo(6, GlCampo6)
            RsAtual!Cidade = GlCampo7
            RsAtual!Estado = VerificaTipo(8, GlCampo8)
            RsAtual!Cep = VerificaTipo(9, GlCampo9)
            RsAtual!Fone1 = VerificaTipo(10, GlCampo10)
            RsAtual!Fone2 = VerificaTipo(11, GlCampo11)
            RsAtual!CGC = VerificaTipo(30, GlCampo30)
            RsAtual!INSCEST = VerificaTipo(12, GlCampo12)
            RsAtual!Fax = VerificaTipo(13, GlCampo13)
     
  End Select

  RsAtual.Update
  'If LcTipoDados = 1 Then RsAtual.Seek "=", GlCampo0
  LcIndice = LcIndiceAnterior
Exit Function
ErroIncl:
  '  Stop
'MsgBox err.Description & err
If err.Number = 3021 Then Resume Next
MsgBox err.Description & err
Resume Next
LcNumero = FreeFile
LcNomeArquivo = "c:\ErrosVirtual.txt"

Open LcNomeArquivo For Append As #LcNumero      ' Open file for output.
 Write #LcNumero, "Função IncluiRegistro, Erro Nº " & err.Number
 Write #LcNumero, "Descrição:" & err.Description
 Write #LcNumero, "Data: " & Date & "Hora: " & Time()
 
Close #LcNumero

Resume Next
End Function
Function reparacnpj()
Dim a As Integer
On Error GoTo errcgc
Dim Lc1, Lc2, Lc3, Lc4, lc5, LcCep, LcCgc, LcInsc As String
AbreBase
AbreBanco (Cliente)
'LcCap = Me.Caption
'Me.Caption = "Processando Dados dos clientes, aguarde..."
Do Until RsAtual.EOF
   If Len(RsAtual!Cep) > 0 Then
      For a = 1 To Len(RsAtual!Cep)
          If IsNumeric(Mid(RsAtual!Cep, a, 1)) Then
             LcCep = LcCep & Mid(RsAtual!Cep, a, 1)
          End If
     Next
   End If
   If Len(RsAtual!CGC) > 0 Then
     For a = 1 To Len(RsAtual!CGC)
          If IsNumeric(Mid(RsAtual!CGC, a, 1)) Then
             LcCgc = LcCgc & Mid(RsAtual!CGC, a, 1)
          End If
     Next
   End If
   If Len(RsAtual!INSCEST) > 0 Then
     For a = 1 To Len(RsAtual!INSCEST)
          If IsNumeric(Mid(RsAtual!INSCEST, a, 1)) Then
             LcInsc = LcInsc & Mid(RsAtual!INSCEST, a, 1)
          End If
     Next
   End If
   RsAtual.Edit
   RsAtual!Cep = LcCep
   RsAtual!CGC = LcCgc
   RsAtual!INSCEST = LcInsc
   RsAtual.Update
   LcCep = ""
   LcCgc = ""
   LcInsc = ""
   RsAtual.MoveNext
Loop
RsAtual.MoveFirst
Do Until RsAtual.EOF
   If Len(RsAtual!Cep) > 0 Then
      LcCep = Mid(RsAtual!Cep, 1, 2) & "." & Mid(RsAtual!Cep, 3, 3) & "-" & Mid(RsAtual!Cep, 6)
   End If
   If Len(RsAtual!CGC) > 0 Then
      LcCgc = Mid(RsAtual!CGC, 1, 2) & "." & Mid(RsAtual!CGC, 3, 3) & "." & Mid(RsAtual!CGC, 6, 3) & "/" & Mid(RsAtual!CGC, 9, 4) & "-" & Mid(RsAtual!CGC, 13)
   End If
   If Len(RsAtual!INSCEST) > 0 Then
      LcInsc = Mid(RsAtual!INSCEST, 1, 3) & "." & Mid(RsAtual!INSCEST, 4, 3) & "." & Mid(RsAtual!INSCEST, 7, 3) & "." & Mid(RsAtual!INSCEST, 10)
   End If
   RsAtual.Edit
   RsAtual!Cep = Left(LcCep, 10)
   RsAtual!CGC = Left(LcCgc, 18)
   RsAtual!INSCEST = Left(LcInsc, 16)
   RsAtual.Update
   LcCep = ""
   LcCgc = ""
   LcInsc = ""
   RsAtual.MoveNext
Loop
RsAtual.Close
'Me.Caption = LcCap
Exit Function
errcgc:
Exit Function
End Function

Public Function FechaBanco()
On Error Resume Next
RsAtual.Close
Dbbase.Close
Set RsAtual = Nothing
Set Dbbase = Nothing
End Function
Public Function AchaReg(TipoPes As Integer) As Integer
On Error GoTo ErroAcha
If TipoPes = 1 Then
   RsAtual.Seek "=", GlChave
Else
   RsAtual.Seek ">=", GlChave
End If
If RsAtual.NoMatch Then
   AchaReg = False
Else
   AchaReg = True
End If
Exit Function
ErroAcha:

'MsgBox Err.Description

'MsgBox Err.Description
Resume Next
End Function
Public Function Alterado()
On Error Resume Next
Dim lcform As Form

Set lcform = Screen.ActiveForm

If LcRegAtual Then
  
   Exit Function
End If
If LcTipoDados = 2 Then
  Screen.ActiveForm.ActiveControl.BackColor = FundoAlterado
  LcAlterado = True
End If

GlCampo0 = LTrim(RTrim(lcform.txt(0).Text))
GlCampo1 = LTrim(RTrim(lcform.txt(1).Text))
GlCampo2 = LTrim(RTrim(lcform.txt(2).Text))
GlCampo3 = LTrim(RTrim(lcform.txt(3).Text))
GlCampo4 = LTrim(RTrim(lcform.txt(4).Text))
GlCampo5 = LTrim(RTrim(lcform.txt(5).Text))
GlCampo6 = LTrim(RTrim(lcform.txt(6).Text))
GlCampo7 = LTrim(RTrim(lcform.txt(7).Text))
GlCampo8 = LTrim(RTrim(lcform.txt(8).Text))
GlCampo9 = LTrim(RTrim(lcform.txt(9).Text))
GlCampo10 = LTrim(RTrim(lcform.txt(10).Text))
GlCampo11 = LTrim(RTrim(lcform.txt(11).Text))
GlCampo12 = LTrim(RTrim(lcform.Mask(2).Text))
GlCampo13 = LTrim(RTrim(lcform.txt(13).Text))
GlCampo14 = LTrim(RTrim(lcform.txt(14).Text))
GlCampo15 = LTrim(RTrim(lcform.txt(15).Text))
GlCampo16 = LTrim(RTrim(lcform.txt(16).Text))
GlCampo17 = LTrim(RTrim(lcform.txt(17).Text))
GlCampo18 = LTrim(RTrim(lcform.txt(18).Text))
GlCampo19 = LTrim(RTrim(lcform.txt(19).Text))
GlCampo20 = LTrim(RTrim(lcform.txt(20).Text))
GlCampo21 = LTrim(RTrim(lcform.txt(21).Text))
GlCampo22 = LTrim(RTrim(lcform.txt(22).Text))
GlCampo23 = LTrim(RTrim(lcform.txt(23).Text))
GlCampo24 = LTrim(RTrim(lcform.txt(24).Text))
GlCampo25 = LTrim(RTrim(lcform.txt(25).Text))
GlCampo26 = LTrim(RTrim(lcform.txt(26).Text))
GlCampo27 = LTrim(RTrim(lcform.txt(27).Text))
GlCampo28 = LTrim(RTrim(lcform.txt(28).Text))
GlCampo29 = LTrim(RTrim(lcform.txt(29).Text))
GlCampo30 = LTrim(RTrim(lcform.txt(30).Text))
GlCampo31 = LTrim(RTrim(lcform.txt(31).Text))
GlCampo32 = LTrim(RTrim(lcform.txt(32).Text))
GlCampo35 = LTrim(RTrim(lcform.txt(12).Text))
If LcTipoDados <> 3 Then
   Screen.ActiveForm.CmdSalvar.Enabled = True
   Screen.ActiveForm.MnSalvar.Enabled = True
End If
End Function
Public Function DesfazMudanca(LcTbl As Tabela)

 RegistroAtual (LcTbl)
 RetornaCorFundo
End Function
Public Function NovaPosicao(LcMovimento As Movimentos, LcTbl As Tabela)
'If Not RsAtual.NoMatch Then
 Msg = ""
 Select Case LcMovimento
    Case Is = enPrimeiro
        If RsAtual.BOF Then
           Msg = "Este é o Primeiro Registro"
        Else
           RsAtual.MoveFirst
           If RsAtual.BOF Then Msg = "Este é o Primeiro Registro"
        End If
    Case Is = enAnterior
        If RsAtual.BOF Then
           Msg = "Este é o Primeiro Registro"
        Else
            RsAtual.MovePrevious
            If RsAtual.BOF Then
               RsAtual.MoveNext
               Msg = "Este é o Primeiro Registro"
            End If
        End If
       
    Case Is = enSeguinte
         
         If RsAtual.EOF Then
           Msg = "Este é o Último Registro"
         Else
           
           
            RsAtual.MoveNext
            If RsAtual.EOF Then
               RsAtual.MovePrevious
               Msg = "Este é o Último Registro"
            End If
            
        End If
        
    Case Is = enultimo
         If RsAtual.EOF Then
           Msg = "Este é o Último Registro"
         Else
            RsAtual.MoveLast
            If RsAtual.EOF Then Msg = "Este é o Último Registro"
        End If
        
 End Select
End Function

Public Function Ficha(NF As String, codigo As String, produto As String, quantidade As Double, Unitario As Double, total As Double, Tipo As String, clifor As String, Unidade As String, Com As String)
Dim db          As Database
Dim RsFicha     As Recordset
Dim Rsp         As Recordset
Dim Rsun        As Recordset
Dim LcSaldo     As Double
Dim LcSaldoAnt  As Double
Dim LcSalUnAnt  As Double
Dim LcSalUn     As Double
Dim LcValEst    As Double
Dim LcValEstC   As Double
Dim LcComp      As Double
Dim LcCaixa     As Double
Dim LcQUnid     As Double
Dim LcVenda     As Double
Dim LcSql       As String
Dim LcSql1      As String
Dim LcSql2      As String
Dim LcPes       As String
Dim LcCodun     As String
Dim LcCodUnP    As String
Dim LcAchou     As Boolean
Dim Lca         As Long
Dim LcBUnidd    As Boolean
Dim LcAbre      As Boolean

Lca = NovoCodigoFicha()
DoEvents
LcSql = "Select * from fichadeestoque where codigo='" & codigo & "' order by auto"
LcSql1 = "Select * from alid009 where cod='" & codigo & "'"
LcSql2 = "Select * from alid004 where SIMBOLO='" & Unidade & "'"

Set db = OpenDatabase(GLBase)
Set RsFicha = db.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Rsp = db.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Rsun = db.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'===> Busca o Codigo da Unidade
lccaixa1 = 0
LcResto1 = 0
If Not Rsun.EOF Then
   LcCodun = Rsun!cod
Else
   LcCodun = ""
End If
If Not Rsp.EOF Then
   If IsNull(Rsp!UNIMED) Then LcCodUnP = "" Else LcCodUnP = Rsp!UNIMED
   If IsNull(Rsp!QTDUNIMED) Then LcComp = 0 Else LcComp = Rsp!QTDUNIMED
   If IsNull(Rsp!Ptab) Then LcVenda = 0 Else LcVenda = Rsp!Ptab
Else
   LcCodUnP = 0
   LcComp = 0
End If
'===> Verifica se a Unidade é a Mesma da Principal

If (LcCodun = LcCodUnP) And (LcComp = CDbl(Com)) Then
    '===> A Unidade é a mesma
    LcCaixa = quantidade
    LcQUnid = 0
Else
    LcBUnidd = True
    If (quantidade * CDbl(Com)) > LcComp Then
       '==> A quantidade Principal é Maior que a Quantidade Unitaria
       LcCaixa = Int(((quantidade * CDbl(Com)) / LcComp))
       LcQUnid = (quantidade * CDbl(Com)) - (LcCaixa * LcComp)
    Else
       '==> A Quantidade Unitaria é Menor
       
       LcCaixa = 0
       LcQUnid = quantidade * CDbl(Com)
    End If
    
End If
'==> Verifica se a Quantidade Disponivel Unitaria é Maior que a Baixada
If Not RsFicha.EOF Then
   RsFicha.MoveLast
   If RsFicha!QuantUnit < (LcQUnid) And (Tipo = "S") Then
        '===> Abre a caixa
        LcAbre = True
   End If

End If
If (LcCodun = LcCodUnP) And (LcComp = CDbl(Com)) And Tipo = "CS" Then
Else
   '==> Verifica se a Unidade e maioir
   'RsFicha.MoveLast
  ' If (LcQUnid + RsFicha!QuantUnit) >= LcComp Then
       lccaixa1 = 0
       LcResto1 = 0
  ' End If
   
End If

'==> Verifica o Saldo
If Not RsFicha.EOF Then
   RsFicha.MoveLast
   If LcAbre Then
        LcSaldoAnt = RsFicha!Saldo - 1
        LcSalUnAnt = RsFicha!QuantUnit + LcComp
   Else
        LcSaldoAnt = RsFicha!Saldo
        LcSalUnAnt = RsFicha!QuantUnit
        LcValEst = RsFicha!valorEstoqueVenda
  End If
Else
   If NF = "IMPLAT" Then
      LcSaldoAnt = LcQSanta
      LcSalUnAnt = LcQUnSanta
   Else
      LcSaldoAnt = 0
      LcSalUnAnt = 0
      LcValEst = 0
   End If
End If
'==> Calcula o Novo Saldo
'If codigo = "01719" Then Stop
Select Case Tipo
    Case Is = "E"
       If NF = "IMPLAT" Then
          LcSaldo = LcSaldoAnt
          LcSalUnAnt = LcSalUnAnt
          LcValEst = (quantidade * Unitario)
       Else
        LcSaldo = LcSaldoAnt + LcCaixa
        LcSalUnAnt = LcSalUnAnt + LcQUnid
        If LcValEst = 0 Then
           LcValEst = (quantidade * Unitario)
        End If
       End If
    Case Is = "S"
        LcSaldo = LcSaldoAnt - LcCaixa
        LcSalUnAnt = LcSalUnAnt - LcQUnid
    Case Is = "CS"
        LcSaldo = LcSaldoAnt + LcCaixa + lccaixa1
        If (LcResto1 > 0) Or (lccaixa1 > 0) Then
           LcSalUnAnt = LcResto1
        Else
           LcSalUnAnt = LcSalUnAnt + LcQUnid
        End If
           
End Select

RsFicha.AddNew
RsFicha!NF = NF
RsFicha!auto = Lca
RsFicha!codigo = codigo
RsFicha!Descricao = produto
RsFicha!Data = Date
RsFicha!quantidade = quantidade
RsFicha!Unitario = Unitario
RsFicha!total = total
RsFicha!Saldo = LcSaldo
RsFicha!Tipo = Tipo
RsFicha!clifor = clifor
RsFicha!Unidade = Unidade
RsFicha!Com = Com
If LcSalUnAnt >= LcComp Then
   RsFicha!QuantUnit = LcSalUnAnt - LcComp
   RsFicha!Saldo = RsFicha!Saldo + 1
Else
   RsFicha!QuantUnit = LcSalUnAnt
End If
If LcComp > 0 Then
 RsFicha!valorEstoqueVenda = (RsFicha!Saldo * LcVenda) + (RsFicha!QuantUnit * (LcVenda / LcComp))
Else
 RsFicha!valorEstoqueVenda = 0
End If
RsFicha.Update

RsFicha.Close
db.Close
Set RsFicha = Nothing
Set db = Nothing




End Function
Function NovoCodigoFicha() As Long
On Error Resume Next
DoEvents
Dim db2 As Database
Dim Rsa As Recordset
Set db2 = OpenDatabase(GLBase)
Set Rsa = db2.OpenRecordset("select * from fichadeestoque order by auto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Rsa.EOF Then
   NovoCodigoFicha = 1
Else
   Rsa.MoveLast
   NovoCodigoFicha = Rsa!auto + 1
End If
Rsa.Close
db2.Close
Set Rsa = Nothing
Set db2 = Nothing

End Function
Public Function MovImentacao(LcMovimento As Movimentos, LcTbl As Tabela) As Integer
Dim a, LcIndiceA As Integer
Dim LcCodigoAtual As Variant
Dim LcIndiceAnterior As String

On Error Resume Next
GlIniceAtual = Screen.ActiveControl.Index
If LcAlterado Then
   SalvaRegistro (LcTbl)
End If
LcCodigoAtual = Screen.ActiveForm.txt(0).Text
'LcCodigoAtual = GlCampo0

Call NovaPosicao(LcMovimento, LcTbl)

If Len(Msg) <> 0 Then
   MsgBox Msg, 64, "Impossível Movimentar."
   Msg = ""
   MovImentacao = False
Else

   RegistroAtual (LcTbl)
   MovImentacao = True
   'GlBook = RsAtual.Bookmark
End If

Screen.ActiveForm.CmdSalvar.Enabled = False
Screen.ActiveForm.MnSalvar.Enabled = False
Exit Function
ErroMovimento:
Select Case ErrosSistema
       Case Is = 6
          Resume 0
       Case Is = 7
          End
End Select
End Function
Public Function RetornaCorFundo()
On Error Resume Next
Dim lcform As Form, a As Integer
Set lcform = Screen.ActiveForm
For a = 0 To 32
    If a = 31 Then
       lcform.Check.BackColor = FundoMornal
    Else
       lcform.txt(a).BackColor = FundoMornal
    End If
Next
For a = 0 To 5
   lcform.Data(a).BackColor = FundoMornal
Next
For a = 0 To 8
   lcform.Valor(a).BackColor = FundoMornal
Next
lcform.deposito(1).BackColor = FundoMornal
lcform.dataentrada.BackColor = FundoMornal
lcform.emissao(0).BackColor = FundoMornal
lcform.Valor.BackColor = FundoMornal
lcform.fornecedor.BackColor = FundoMornal
'LcForm.Compensado.BackColor = FundoMornal
On Error GoTo 0
End Function
Public Function Exclui(LcTbl As Tabela) As Integer
On Error GoTo errEsclui
Dim LcInAnt As String
Dim LcPes As Integer
If GlConfirmaExclusao Then
   LcResposta = MsgBox("Confirma a Exclusão deste Registro ?", 36, "Confirmação")
Else
   LcResposta = 6 ' Força o Sim
End If
If LcResposta = 7 Then
   MsgBox "Operação Cancelada Pelo Usuário", 32, "Aviso"
   Exclui = 0
   GoTo SaiExclusao
End If
LcInAnt = LcIndice
LcIndice = "CODIGO"
GlChave = GlFormA.txt(0).Text
Call controleexclusao(GlTab, GlSq)
AbrePesquisa:
   
   LcPes = GLPesquisa
   GLPesquisa = False
   LcIndice = "codigo"
   Select Case GlFormA.Name
      Case Is = "FrmProduto"
           AbreBanco (produto)
      Case Is = "FrmCliente"
           AbreBanco (Cliente)
      Case Is = "FrmFornecedor"
           AbreBanco (fornecedor)
      Case Is = "FrmFuncionario"
           AbreBanco (Funcionario)
      Case Is = "FrmTransportadora"
           AbreBanco (Transportadora)
      Case Is = "FrmCidade"
           AbreBanco (Cidade)
      Case Is = "FrmTiporeceita"
           AbreBanco (tiporec)
      Case Is = "FrmUnidade"
           AbreBanco (Unidade)
      Case Is = "FrmGalpao"
           AbreBanco (Galpao)
      Case Is = "FrmTipoMonetario"
           AbreBanco (monetario)
           GlChave = GlFormA.txt(0).Text
           LcIndice = "TPMONET"
      Case Is = "Receitas"
           AbreBanco (Receber)
           GlChave = GlFormA.txt(0).Text
      Case Is = "Despesas"
           AbreBanco (pagar)
           GlChave = GlFormA.txt(0).Text
           LcIndice = "NF"
       Case Is = "Frmcheques"
           AbreBanco (Cheques)
           GlChave = GlFormA.txt(0).Text
           
    End Select
   
   'RsAtual.Index = LcIndice
   
'If Not AchaReg(1)
If AchaReg(1) Then
   
   RsAtual.Delete
   RsAtual.MoveNext
   If Not RsAtual.EOF Then
      RegistroAtual (LcTbl)
      Exclui = 1
   Else
     RsAtual.MovePrevious
     If Not RsAtual.BOF Then
       RegistroAtual (LcTbl)
       Exclui = 1
     Else
       NovoReg
       Exclui = 1
     End If
   End If
End If
LcIndice = LcInAnt
LcRegAtual = False
SaiExclusao:
'FechaBanco
Exit Function
errEsclui:

If err.Number = 3251 Then
   GLPesquisa = True
   GoTo AbrePesquisa
End If
'MsgBox Err.Description & Err
Resume Next
End Function

Public Function ErrosSistema() As Integer
Dim LcRepete, LcIcone As Integer, Msg, lctitulo, LcNomeArquivo As String
Dim LcExibemsg, LcNumero As Integer
LcIcone = 64
LcNumero = FreeFile
LcNomeArquivo = "c:\ErrosVirtual.txt"

Open LcNomeArquivo For Append As #LcNumero      ' Open file for output.
 Write #LcNumero, "Erro Geral, Erro Nº " & err.Number
 Write #LcNumero, "Descrição:" & err.Description
 Write #LcNumero, "Data: " & Date & "Hora: " & Time()
 
Close #LcNumero

Select Case err
    Case Is = 3045
         Msg = "O Banco de Dados Está Aberto em Modo Exclusivo Por outro Usuário..."
         Msg = Msg & Chr(13)
         Msg = Msg & "Deseja Tentar Abri-lo Novamente ? "
         LcRepete = True
         lctitulo = "Erro Abertura Banco"
         LcIcone = 4116
         LcExibemsg = True
   Case Is = 3261
         Msg = "O Banco de Dados Está Aberto em Modo Exclusivo Por outro Usuário..."
         Msg = Msg & Chr(13)
         Msg = Msg & "Deseja Tentar Abri-lo Novamente ? "
         LcRepete = True
         lctitulo = "Erro Abertura Banco"
         LcIcone = 4116
         LcExibemsg = True
    Case Is = 3021
         LcRepete = True
         LcExibemsg = False
    Case Else
         Msg = err.Description
         lctitulo = "Erro Nº " & err
 End Select

If LcRepete Then
   If LcExibemsg Then ErrosSistema = MsgBox(Msg, LcIcone, lctitulo) Else ErrosSistema = 0
   
Else
   'MsgBox msg, LcIcone, lctitulo
   ErrosSistema = 0
  
End If
End Function

Function CriaMascara()
Dim a As Integer, LcMascara As String
For a = 0 To 30
   LcMascara = Mid$(Screen.ActiveForm.txt(a).Tag, 1, 1)
   
   Select Case LcMascara
          Case Is = "N"
               Screen.ActiveForm.txt(a).Mask = "99999999"
          Case Is = "M"
               Screen.ActiveForm.txt(a).Mask = "999.999,99"
          Case Is = "C"
               Screen.ActiveForm.txt(a).Mask = "##.###.####-##"
          Case Is = "G"
               Screen.ActiveForm.txt(a).Mask = "##.###.####-##"
          Case Is = "F"
               Screen.ActiveForm.txt(a).Mask = "(##)###-####"
          Case Is = "D"
               Screen.ActiveForm.txt(a).Mask = "##/##/####"
          Case Is = "P"
               Screen.ActiveForm.txt(a).Mask = "##.###-###"
          Case Is = "S"
               Screen.ActiveForm.txt(a).Mask = ""
  End Select
Next

End Function

Function VerificaTipo(LcIndice As Integer, LcCampo As String) As Variant
On Error Resume Next
Dim LcNumero As Long, LcTipo, LcIndiceCampo As String, LcMoeda As Currency
Dim LcData As Date, a As Integer

For a = 0 To 30
    LcIndiceCampo = Mid$(Screen.ActiveForm.txt(a).Tag, 7, 2)
   
    LcTipo = Mid$(Screen.ActiveForm.txt(a).Tag, 3, 1)
    If Val(LcIndiceCampo) = LcIndice Then
         
         Select Case LcTipo
           Case Is = "N"
         
             If Len(Trim(LcCampo)) = 0 Then
                VerificaTipo = 0
             Else
               VerificaTipo = CLng(AcertaDecimal(LcCampo))
             End If
             
          Case Is = "D"
             
             If Len(LcCampo) = 0 Then
                VerificaTipo = Null
             Else
               If LcCampo = "" Then
                  VerificaTipo = Null
               Else
                  VerificaTipo = LcCampo
               End If
             End If
          Case Is = "T"
             VerificaTipo = LcCampo
          Case Is = "M"
           
             If LcCampo = "" Then
                VerificaTipo = 0
             Else
               
                VerificaTipo = CCur(AcertaDecimal(LcCampo))
             End If
             
         End Select
         Exit For
    End If
Next

End Function
Function VerificaRequerido() As Integer
On Error Resume Next
Dim a As Integer, LcRequerido, LcNome  As String
Dim LcIndex As Integer

VerificaRequerido = True
For a = 0 To 30
   LcIndex = Mid(Screen.ActiveForm.txt(a).Tag, 7, 2)
   LcRequerido = Mid$(Screen.ActiveForm.txt(a).Tag, 5, 1)
   LcNome = Mid$(Screen.ActiveForm.txt(a).Tag, 12)
   If LcIndex < a Then Exit For
   If LcRequerido = "S" Then
       If LcNome = "CODIGO" And LcTipoDados = 1 Then
         Else
         If Len(Screen.ActiveForm.txt(a).Text) = 0 Then
            MsgBox "O Campo " & LcNome & " é Necessário...", 64, "Aviso"
            Screen.ActiveForm.txt(a).SetFocus
            VerificaRequerido = False
            Exit Function
         End If
       End If
   End If
Next

End Function
Function lancacaixaDebito(LcValor As Currency)
On Error Resume Next
Dim LcData As Date
Dim RsCaixa As Recordset
Dim LcSaldo, LcEntradas, LcSaida As Currency
LcData = Format(GlDataSistema, "dd/mm/yyyy")

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsCaixa = Dbbase.OpenRecordset("Caixa", dbOpenTable, dbSeeChanges, dbOptimistic)
RsCaixa.Index = "DataDoLancamento"
RsCaixa.Seek "=", LcData
If Not RsCaixa.NoMatch Then
   If IsNull(RsCaixa!Saldo) Then LcSaldo = 0 - LcValor Else LcSaldo = RsCaixa!Saldo - LcValor
   If IsNull(RsCaixa!Saida) Then LcSaida = LcValor Else LcSaida = RsCaixa!Saida - LcValor
   LcEntradas = 0
   RsCaixa.Edit
Else
   LcSaldo = 0 - LcValor
   LcEntradas = 0
   LcSaida = LcValor
   RsCaixa.AddNew
End If
RsCaixa!Saldo = LcSaldo
RsCaixa!Entrada = LcEntradas
RsCaixa!DataMovimento = LcData
RsCaixa!Saida = LcSaida
RsCaixa!fechado = False
RsCaixa.Update
RsCaixa.Close
End Function
Function AtualizaCaixaNovo()
Dim LcData As Date
Dim RsCaixa As Recordset, RsHistorico As Recordset, RsHistoricoAnt As Recordset
Dim LcValor As Currency
Dim LcFitas, LcSacolas, a, LcPendentes As Long
Dim LcCriterio, LcCriterio1 As String
Dim LcAchou As Integer
LcCriterio = "select * From HistoricoLocGeral where DataDevolucao=#" & GlDataSistema & "#"
LcCriterio1 = "select * From HistoricoLocGeral where DataDevolucao<#" & GlDataSistema & "#"
LcValor = 0
LcFitas = 0
LcSacolas = 0
LcPendentes = 0
LcAchou = False

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsCaixa = Dbbase.OpenRecordset("Caixa", dbOpenTable, dbSeeChanges, dbOptimistic)
Set RsHistorico = Dbbase.OpenRecordset(LcCriterio, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsHistoricoAnt = Dbbase.OpenRecordset(LcCriterio1, dbOpenDynaset, dbSeeChanges, dbOptimistic)

RsCaixa.Index = "DataDoLancamento"
RsCaixa.Seek "=", GlDataSistema
'== Verifica se Já Foi Criado o caixa do dia
If Not RsCaixa.NoMatch Then
   RsCaixa.Close
   RsHistorico.Close
   Dbbase.Close
   Exit Function
End If
'==Busca as Fitas a Devolver Hoje
Do Until RsHistoricoAnt.EOF
   If RsHistoricoAnt!FitasDevolvidas < RsHistoricoAnt!Fitas Then
      LcPendentes = LcPendentes + (RsHistoricoAnt!Fitas - RsHistoricoAnt!FitasDevolvidas)
   End If
   RsHistoricoAnt.MoveNext
Loop
'==Busca as Fitas a Pendentes
Do Until RsHistorico.EOF
   
   LcAchou = True
   LcValor = RsHistorico!ValorPagar + LcValor
   LcFitas = RsHistorico!Fitas + LcFitas
   LcSacolas = RsHistorico!Sacolas + LcSacolas
   RsHistorico.MoveNext
   
Loop
If LcAchou Then
  RsCaixa.AddNew
  RsCaixa!DataMovimento = GlDataSistema
  RsCaixa!FitasaDev = LcFitas
  RsCaixa!SacolasMov = LcSacolas
  RsCaixa!ValoresReceber = LcValor
  RsCaixa!DevolucoesPendentes = LcPendentes
  RsCaixa.Update
End If
RsCaixa.Close
RsHistorico.Close
RsHistoricoAnt.Close
End Function


Function Maiuscula()
On Error Resume Next
GlCampo0 = UCase(GlCampo0)
GlCampo1 = UCase(GlCampo1)
GlCampo2 = UCase(GlCampo2)
GlCampo3 = UCase(GlCampo3)
GlCampo4 = UCase(GlCampo4)
GlCampo5 = UCase(GlCampo5)
GlCampo6 = UCase(GlCampo6)
GlCampo7 = UCase(GlCampo7)
GlCampo8 = UCase(GlCampo8)
GlCampo9 = UCase(GlCampo9)
GlCampo10 = UCase(GlCampo10)
GlCampo11 = UCase(GlCampo11)
GlCampo12 = UCase(GlCampo12)
GlCampo13 = UCase(GlCampo13)
GlCampo14 = UCase(GlCampo14)
GlCampo15 = UCase(GlCampo15)
GlCampo16 = UCase(GlCampo16)
GlCampo17 = UCase(GlCampo17)
GlCampo18 = UCase(GlCampo18)
GlCampo19 = UCase(GlCampo19)
GlCampo20 = UCase(GlCampo20)
GlCampo21 = UCase(GlCampo21)
GlCampo22 = UCase(GlCampo22)
GlCampo23 = UCase(GlCampo23)
GlCampo24 = UCase(GlCampo24)
GlCampo25 = UCase(GlCampo25)
GlCampo26 = UCase(GlCampo26)
GlCampo27 = UCase(GlCampo27)
GlCampo28 = UCase(GlCampo28)
GlCampo29 = UCase(GlCampo29)
GlCampo30 = UCase(GlCampo30)
GlCampo35 = UCase(GlCampo35)
End Function
Function controleexclusao(Tabela As String, criterio As String)
On Error GoTo errexcl
Dim Rs  As Recordset
Dim RsE As Recordset
Dim LcCampo As String
Dim campo   As Field
Dim LcHora As Date
AbreBase

Set Rs = Dbbase.OpenRecordset(criterio, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsE = Dbbase.OpenRecordset("Select * from deletada", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcHora = Time
Do Until Rs.EOF
   If err.Number > 0 Then Exit Do
   RsE.AddNew
   For Each campo In Rs.Fields
       LcCampo = campo.Name
       'RsE.Fields(LcCampo) = Rs.Fields(LcCampo)
   Next
   RsE.Fields("maquinaExclusao") = GlNomeMaquina
   RsE.Fields("usuarioExclusao") = GlUsuario
   RsE.Fields("dataexclusao") = Date
   RsE.Fields("horaexclusao") = LcHora
   RsE.Fields("tabelaexcluida") = Tabela
   RsE.Update
   Rs.MoveNext
Loop
Rs.Close
RsE.Close
Dbbase.Close
Set Rs = Nothing
Set RsE = Nothing
Set Dbbase = Nothing
Exit Function
errexcl:
MsgBox err.Description & err.Number
'Resume 0
Exit Function
End Function
Function LeIni(Seção$, Chave$, Arqini$) As String
Dim i&, Ret$, max_size As Byte
  max_size = 255
  Ret = Space(max_size)
  i = GetPrivateProfileStringA(Seção$, Chave$, "", Ret, max_size, Arqini$)
  LeIni = Left(Ret, InStr(Ret, Chr(0)) - 1)
End Function
'Write in Ini File =============================================
Function GravaIni(Seção$, Chave$, strValue$, Arqini$) As Boolean
On Error GoTo errGrvini
Dim i&
  i = WritePrivateProfileStringA(Seção$, Chave$, strValue$, Arqini$)
  GravaIni = Len(LeIni(Seção, Chave$, Arqini$)) > 0
Exit Function
errGrvini:
GravaIni = False
End Function
'Delete Keys in Ini File ==========================
Function DelKey(Section$, Key$, Arq$) As Boolean
On Error GoTo ErrDelKey
  WritePrivateProfileDelKey Section, Key, 0&, Arq
  DelKey = True
Exit Function
ErrDelKey:
  DelKey = False
End Function
'Delete Sections in Ini File =================
Function DelSection(Section$, Arq$) As Boolean
On Error GoTo ErrDelKey
  WritePrivateProfileDelSect Section, 0&, "", Arq
  DelSection = True
Exit Function
ErrDelKey:
  DelSection = False
End Function
'Lê os nomes de todas as chaves de uma dada seção========
Function KeysInSection(Section$, Arq$) As String
  Dim Buff As String * 1024, Result%
  Result = GetPrivateProfileKeys(Section, 0&, "", Buff, Len(Buff), Arq)
  KeysInSection = Left$(Buff, Result)
End Function
'Lê os nomes de todas as seções de um dado arquivo========
Function SectionsInFile(Arq$) As String
  Dim Rtn$, Result$, Pos&
  Result = Chr(255)
  Rtn = Space(1024)
  success = GetPrivateProfileStringSections(0, 0, "", Rtn, 1024, Arq$)
  Pos = InStr(1, Rtn, "  ")
  SectionsInFile = Left$(Rtn, (Pos - 2)) 'Result)
End Function


Public Function Calc_CNPJ1(Valor As String) As Boolean
' Inicializa variaveis
Dim Mult1 As String
Dim Mult2 As String
Dim dig1 As Integer
Dim dig2 As Integer
Dim x As Integer
Mult1 = "543298765432"
Mult2 = "6543298765432"
For a = 1 To Len(Valor)
    If IsNumeric(Mid(Valor, a, 1)) Then
       LcCNPJ = LcCNPJ & Mid(Valor, a, 1)
    End If
Next
Valor = LcCNPJ
For x = 1 To 12
dig1 = dig1 + (Val(Mid$(Valor, x, 1)) * Val(Mid$(Mult1, x, 1)))
Next

For x = 1 To 13
dig2 = dig2 + (Val(Mid$(Valor, x, 1)) * Val(Mid$(Mult2, x, 1)))
Next
dig1 = (dig1 * 10) Mod 11
dig2 = (dig2 * 10) Mod 11

If dig1 = 10 Then dig1 = 0
If dig2 = 10 Then dig2 = 0

Calc_CNPJ1 = False
If dig1 = Val(Mid$(Valor, 13, 1)) And dig2 = Val(Mid$(Valor, 14, 1)) Then Calc_CNPJ1 = True
'If dig2 = Val(Mid$(valor, 14, 1)) Then Calc_CNPJ = True

End Function
Public Function Calc_CPF(Valor As String) As Boolean
' Inicializa variaveis
Dim dig1 As Integer
Dim dig2 As Integer
Dim Mult1 As Integer
Dim Mult2 As Integer
Dim x As Integer
Dim LcCpf As String
Mult1 = 10
Mult2 = 11
For a = 1 To Len(Valor)
    If IsNumeric(Mid(Valor, a, 1)) Then
       LcCpf = LcCpf & Mid(Valor, a, 1)
    End If
Next
Valor = LcCpf
For x = 1 To 9
dig1 = dig1 + (Val(Mid$(Valor, x, 1)) * Mult1)
Mult1 = Mult1 - 1
Next

For x = 1 To 10
dig2 = dig2 + (Val(Mid$(Valor, x, 1)) * Mult2)
Mult2 = Mult2 - 1
Next

dig1 = (dig1 * 10) Mod 11
dig2 = (dig2 * 10) Mod 11
If dig1 = 10 Then dig1 = 0
If dig2 = 10 Then dig2 = 0

Calc_CPF = False

If Val(Mid$(Valor, 10, 1)) = dig1 And Val(Mid$(Valor, 11, 1)) = dig2 Then Calc_CPF = True
'If Val(Mid$(VALOR, 11, 1)) <> dig2 Then Calc_CPF = True
End Function
Public Function Calc_CNPJ(CGC As String) As Boolean
  Dim Retorno, a, j, i, d1, d2
  'TiraMascara CGC
  
  CGC = Replace(CGC, ".", "")
  CGC = Replace(CGC, ",", "")
  CGC = Replace(CGC, "-", "")
  CGC = Replace(CGC, "/", "")
  CGC = Replace(CGC, "\", "")
  CGC = Replace(CGC, " ", "")
  'Debug.Print CGC
  If Len(CGC) = 8 And Val(CGC) > 0 Then
     a = 0
     j = 0
     d1 = 0
     For i = 1 To 7
         a = Val(Mid(CGC, i, 1))
         If (i Mod 2) <> 0 Then
            a = a * 2
         End If
         If a > 9 Then
            j = j + Int(a / 10) + (a Mod 10)
         Else
            j = j + a
         End If
     Next i
     d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
     If d1 = Val(Mid(CGC, 8, 1)) Then
        ValidaCGC = True
     Else
        ValidaCGC = False
        'MsgBox "CNPJ inválido!,Verifique", vbCritical, "Valida CGC"
     End If
  Else
     If Len(CGC) = 14 And Val(CGC) > 0 Then
        a = 0
        i = 0
        d1 = 0
        d2 = 0
        j = 5
        For i = 1 To 12 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        d1 = IIf(a > 1, 11 - a, 0)
        a = 0
        i = 0
        j = 6
        For i = 1 To 13 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        d2 = IIf(a > 1, 11 - a, 0)
        If (d1 = Val(Mid(CGC, 13, 1)) And d2 = Val(Mid(CGC, 14, 1))) Then
           ValidaCGC = True
        Else
           ValidaCGC = False
          ' MsgBox "CNPJ inválido!, Verifique", vbCritical, "Aviso do Sistema"
        End If
     Else
        ValidaCGC = False
        'MsgBox "CNPJ inválido!,Verifique", vbCritical, "Aviso do Sistema"
     End If
  End If
  Calc_CNPJ = ValidaCGC
End Function
