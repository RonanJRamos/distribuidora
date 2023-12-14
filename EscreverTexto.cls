VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EscreverTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Type DadosReg60D
   Codigo       As String
   Descricao    As String
   icms         As String
   ipi          As String
   Unidade      As String
   Quantidade   As Double
   data         As Date
   valor        As Double
   Situacao     As String
End Type
Private LcArq As Integer
Public Property Let NumeroArquivo(ByVal vData As Integer)
    LcArq = vData
End Property
Public Sub EscreverRegistro54()
Dim StrSintegea As String
Dim a As Long
Dim total As String
Dim Desconto As String
Dim BaseIcms   As String
Dim BaseIcmsSubs As String
Dim ValordoIPI As String
Dim AliquotaIcms As String
Dim ValorQuant  As String
Dim ClUtil As New Utilitario
TotalReg54 = 0

For a = 0 To UBound(Mt54)
   'If Mt54(a).Numero_Nf = "294240" Then Stop
    ValorQuant = ClUtil.AcertaNumero(CStr(Mt54(a).Quantidade), 3)
    ValorQuant = Replace(ValorQuant, ".", "")
    ValorQuant = Replace(ValorQuant, ",", "")
    ValorQuant = Right("00000000000" & ValorQuant, 11)
    
    total = ClUtil.AcertaNumero(CStr(Mt54(a).Valor_Produto), 2)
    total = Replace(total, ".", "")
    total = Replace(total, ",", "")
    total = Right("000000000000" & total, 12)
    
    Desconto = ClUtil.AcertaNumero(CStr(Mt54(a).Valor_Desconto), 2)
    Desconto = Replace(Desconto, ".", "")
    Desconto = Replace(Desconto, ",", "")
    Desconto = Right("000000000000" & Desconto, 12)
    
    BaseIcms = ClUtil.AcertaNumero(CStr(Mt54(a).Base_Calculo_Icms), 2)
    BaseIcms = Replace(BaseIcms, ".", "")
    BaseIcms = Replace(BaseIcms, ",", "")
    BaseIcms = Right("000000000000" & BaseIcms, 12)
    
    BaseIcmsSubs = ClUtil.AcertaNumero(CStr(Mt54(a).Base_Calculo_subs_Trib), 2)
    BaseIcmsSubs = Replace(BaseIcmsSubs, ".", "")
    BaseIcmsSubs = Replace(BaseIcmsSubs, ",", "")
    BaseIcmsSubs = Right("000000000000" & BaseIcmsSubs, 12)
    
    ValordoIPI = ClUtil.AcertaNumero(CStr(Mt54(a).Valor_Ipi), 2)
    ValordoIPI = Replace(ValordoIPI, ".", "")
    ValordoIPI = Replace(ValordoIPI, ",", "")
    ValordoIPI = Right("000000000000" & ValordoIPI, 12)
    
    AliquotaIcms = ClUtil.AcertaNumero(CStr(Mt54(a).Aliquota_Icms), 2)
    AliquotaIcms = Replace(AliquotaIcms, ".", "")
    AliquotaIcms = Replace(AliquotaIcms, ",", "")
    AliquotaIcms = Right("0000" & AliquotaIcms, 4)
    'If Mt54(a).Numero_Nf = "027575" Then Stop
    StrSintegra = "54"
    StrSintegra = StrSintegra & Mt54(a).Cnpj
    StrSintegra = StrSintegra & Mt54(a).Modelo
    StrSintegra = StrSintegra & Mt54(a).Serie
    StrSintegra = StrSintegra & Mt54(a).Numero_Nf
    StrSintegra = StrSintegra & Mt54(a).CFOP
    StrSintegra = StrSintegra & Right("000" & Mt54(a).cst, 3)
    StrSintegra = StrSintegra & Right("000" & Mt54(a).Numero_Item, 3)
    StrSintegra = StrSintegra & Left(Mt54(a).Codigo_Produto & "              ", 14)
    StrSintegra = StrSintegra & ValorQuant
    StrSintegra = StrSintegra & total
    StrSintegra = StrSintegra & Desconto
    StrSintegra = StrSintegra & BaseIcms
    StrSintegra = StrSintegra & BaseIcmsSubs
    StrSintegra = StrSintegra & ValordoIPI
    StrSintegra = StrSintegra & AliquotaIcms
    TotalReg54 = TotalReg54 + 1
    Print #LcArq, StrSintegra
Next
End Sub
Public Sub EscreveRegistro74()
Dim StrSintegea As String
Dim a As Long
Dim ValorPrduto As String
Dim Desconto As String
Dim BaseIcms   As String
Dim BaseIcmsSubs As String
Dim ValordoIPI As String
Dim AliquotaIcms As String
Dim ValorQuant  As String
Dim ClUtil As New Utilitario
TotalReg74 = 0

For a = 0 To UBound(Mt74)
   'If Mt54(a).Numero_Nf = "294240" Then Stop
  If Sintegra.IncluirZero.Value = 1 Then
         ValorQuant = ClUtil.AcertaNumero(CStr(Mt74(a).Quantidade), 3)
         ValorQuant = Replace(ValorQuant, ".", "")
         ValorQuant = Replace(ValorQuant, ",", "")
         ValorQuant = Right("0000000000000" & ValorQuant, 13)
         
         ValorPrduto = ClUtil.AcertaNumero(CStr(Mt74(a).ValorProduto), 2)
         ValorPrduto = Replace(ValorPrduto, ".", "")
         ValorPrduto = Replace(ValorPrduto, ",", "")
         ValorPrduto = Right("0000000000000" & ValorPrduto, 13)
        
         StrSintegra = "74"
         StrSintegra = StrSintegra & Mt74(a).data
         StrSintegra = StrSintegra & Mt74(a).codigoproduto
         StrSintegra = StrSintegra & ValorQuant
         StrSintegra = StrSintegra & ValorPrduto
         StrSintegra = StrSintegra & Mt74(a).Codigo_Posse
         StrSintegra = StrSintegra & Mt74(a).Cnpj
         StrSintegra = StrSintegra & Mt74(a).Inscricao
         StrSintegra = StrSintegra & Mt74(a).estado
         StrSintegra = StrSintegra & String(45, " ")
         TotalReg74 = TotalReg74 + 1
         Print #LcArq, StrSintegra
    Else
       If Mt74(a).Quantidade > 0 Then
         ValorQuant = ClUtil.AcertaNumero(CStr(Mt74(a).Quantidade), 3)
         ValorQuant = Replace(ValorQuant, ".", "")
         ValorQuant = Replace(ValorQuant, ",", "")
         ValorQuant = Right("0000000000000" & ValorQuant, 13)
         
         ValorPrduto = ClUtil.AcertaNumero(CStr(Mt74(a).ValorProduto), 2)
         ValorPrduto = Replace(ValorPrduto, ".", "")
         ValorPrduto = Replace(ValorPrduto, ",", "")
         ValorPrduto = Right("0000000000000" & ValorPrduto, 13)
        
         StrSintegra = "74"
         StrSintegra = StrSintegra & Mt74(a).data
         StrSintegra = StrSintegra & Mt74(a).codigoproduto
         StrSintegra = StrSintegra & ValorQuant
         StrSintegra = StrSintegra & ValorPrduto
         StrSintegra = StrSintegra & Mt74(a).Codigo_Posse
         StrSintegra = StrSintegra & Mt74(a).Cnpj
         StrSintegra = StrSintegra & Mt74(a).Inscricao
         StrSintegra = StrSintegra & Mt74(a).estado
         StrSintegra = StrSintegra & String(45, " ")
         TotalReg74 = TotalReg74 + 1
         Print #LcArq, StrSintegra
       End If
    End If
Next

End Sub
Public Sub EscreverRegistro10(Datai As Date, Dataf As Date)
Dim Rs As ADODB.Recordset
Dim StrSql      As String
Dim StrResposta As String
Dim Cnpj        As String
Dim Inscricao   As String
Dim Fax         As String
Dim Fone        As String

StrSql = "Select * from empresa"
Set Rs = AbreRecordset(StrSql, True)
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
   StrResposta = StrResposta & Left(Rs!estado & String(2, " "), 2)
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
   StrResposta = StrResposta & Format(Datai, "yyyymmdd")
  '==>Data Final
   StrResposta = StrResposta & Format(Dataf, "yyyymmdd")
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

Public Sub EscreverRegistro11()
Dim Rs As ADODB.Recordset
Dim StrSql      As String
Dim StrResposta As String
Dim Cnpj        As String
Dim Inscricao   As String
Dim Fax         As String
Dim Fone        As String
Dim Cep         As String

StrSql = "Select * from empresa"
Set Rs = AbreRecordset(StrSql, True)
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
Public Sub EscreverRegistro50()
Dim StrSintegra     As String
Dim ClUtil          As New Utilitario
Dim a               As Long
Dim total           As String
Dim TotalBase       As String
Dim TotalIcms       As String
Dim TotalIsenta     As String
Dim TotalOutras     As String
Dim AliquotaIcms    As String
TotalReg50 = 0
For a = 0 To UBound(Mt50)
    
    total = ClUtil.AcertaNumero(CStr(Mt50(a).Valor_Total), 2)
    total = Replace(total, ".", "")
    total = Replace(total, ",", "")
    total = Right("0000000000000" & total, 13)
    
    TotalBase = ClUtil.AcertaNumero(CStr(Mt50(a).Base_Calculo_Icms), 2)
    TotalBase = Replace(TotalBase, ".", "")
    TotalBase = Replace(TotalBase, ",", "")
    TotalBase = Right("0000000000000" & TotalBase, 13)
    
    TotalIcms = ClUtil.AcertaNumero(CStr(Mt50(a).Valor_Icms), 2)
    TotalIcms = Replace(TotalIcms, ".", "")
    TotalIcms = Replace(TotalIcms, ",", "")
    TotalIcms = Right("0000000000000" & TotalIcms, 13)
    
    TotalIsenta = ClUtil.AcertaNumero(CStr(Mt50(a).Isenta_Nao_Tributada), 2)
    TotalIsenta = Replace(TotalIsenta, ".", "")
    TotalIsenta = Replace(TotalIsenta, ",", "")
    TotalIsenta = Right("0000000000000" & TotalIsenta, 13)
    
    TotalOutras = ClUtil.AcertaNumero(CStr(Mt50(a).Outras), 2)
    TotalOutras = Replace(TotalOutras, ".", "")
    TotalOutras = Replace(TotalOutras, ",", "")
    TotalOutras = Right("0000000000000" & TotalOutras, 13)
    
    AliquotaIcms = ClUtil.AcertaNumero(CStr(Mt50(a).Aliquota), 2)
    AliquotaIcms = Replace(AliquotaIcms, ".", "")
    AliquotaIcms = Replace(AliquotaIcms, ",", "")
    AliquotaIcms = Right("0000" & AliquotaIcms, 4)

    If CLng(Mt50(a).Numero_Nf) = 0 Then Stop
    StrSintegra = "50"
    StrSintegra = StrSintegra & Mt50(a).Cnpj
    StrSintegra = StrSintegra & Mt50(a).Inscricao
    StrSintegra = StrSintegra & Mt50(a).data
    StrSintegra = StrSintegra & Mt50(a).estado
    StrSintegra = StrSintegra & Mt50(a).Modelo
    StrSintegra = StrSintegra & Mt50(a).Serie
    StrSintegra = StrSintegra & Mt50(a).Numero_Nf
    StrSintegra = StrSintegra & Mt50(a).CFOP
    StrSintegra = StrSintegra & Mt50(a).Emitente
    StrSintegra = StrSintegra & total
    StrSintegra = StrSintegra & TotalBase
    StrSintegra = StrSintegra & TotalIcms
    StrSintegra = StrSintegra & TotalIsenta
    StrSintegra = StrSintegra & TotalOutras
    StrSintegra = StrSintegra & AliquotaIcms
    StrSintegra = StrSintegra & Mt50(a).Situacao
    'MsgBox Len(StrSintegra)
    
    TotalReg50 = TotalReg50 + 1
    Print #LcArq, StrSintegra
Next
End Sub


Public Sub EscreverRegistro70()
Dim StrSintegra As String
Dim ClUtil As New Utilitario
Dim a As Long
Dim b As Long
Dim valor As String
Dim x As Integer
Dim AliquotaIcms As String

TotalReg70 = 0
On Error Resume Next
err.Number = 0
a = UBound(Mt70)
If err.Number <> 0 Then Exit Sub

For a = 0 To UBound(Mt70)
  If Mt70(a).Valor_Total > 0 Then
    total = ClUtil.AcertaNumero(CStr(Mt70(a).Valor_Total), 2)
    total = Replace(total, ".", "")
    total = Replace(total, ",", "")
    total = Right("0000000000000" & total, 13)
    
    TotalBase = ClUtil.AcertaNumero(CStr(Mt70(a).Base_Calculo_Icms), 2)
    TotalBase = Replace(TotalBase, ".", "")
    TotalBase = Replace(TotalBase, ",", "")
    TotalBase = Right("00000000000000" & TotalBase, 14)
    
    TotalIcms = ClUtil.AcertaNumero(CStr(Mt70(a).Valor_Icms), 2)
    TotalIcms = Replace(TotalIcms, ".", "")
    TotalIcms = Replace(TotalIcms, ",", "")
    TotalIcms = Right("00000000000000" & TotalIcms, 14)
    
    TotalIsenta = ClUtil.AcertaNumero(CStr(Mt70(a).Isenta_Nao_Tributada), 2)
    TotalIsenta = Replace(TotalIsenta, ".", "")
    TotalIsenta = Replace(TotalIsenta, ",", "")
    TotalIsenta = Right("00000000000000" & TotalIsenta, 14)
    
    TotalOutras = ClUtil.AcertaNumero(CStr(Mt70(a).Outras), 2)
    TotalOutras = Replace(TotalOutras, ".", "")
    TotalOutras = Replace(TotalOutras, ",", "")
    TotalOutras = Right("00000000000000" & TotalOutras, 14)
    
    
    StrSintegra = "70"
    StrSintegra = StrSintegra & Mt70(a).Cnpj
    StrSintegra = StrSintegra & Mt70(a).Inscricao
    StrSintegra = StrSintegra & Mt70(a).data
    StrSintegra = StrSintegra & Mt70(a).estado
    StrSintegra = StrSintegra & Mt70(a).Modelo
    StrSintegra = StrSintegra & Mt70(a).Serie
    StrSintegra = StrSintegra & Mt70(a).SubSerie
    StrSintegra = StrSintegra & Mt70(a).Numero_Nf
    StrSintegra = StrSintegra & Mt70(a).CFOP
    StrSintegra = StrSintegra & total
    StrSintegra = StrSintegra & TotalBase
    StrSintegra = StrSintegra & TotalIcms
    StrSintegra = StrSintegra & TotalIsenta
    StrSintegra = StrSintegra & TotalOutras
    StrSintegra = StrSintegra & Mt70(a).CifFob
    StrSintegra = StrSintegra & Mt70(a).Situacao
    TotalReg70 = TotalReg70 + 1
    Print #LcArq, StrSintegra
  End If
Next

End Sub
Public Sub EscreverRegistro75(Datai As Date, Dataf As Date)

'On Error Resume Next
Dim StrSintegra As String
Dim ClUtil As New Utilitario
Dim a As Long
Dim b As Long
Dim valor As String
Dim x As Integer
Dim AliquotaIcms As String
Dim AliquotaIpi As String

x = 0

For a = 0 To UBound(Mt75)
    
   Sintegra.Caption = "Criando o Registro 75 - " & a & " de " & UBound(Mt75) & " Produto - " & Mt75(a).Nome
   DoEvents
   
   valor = ClUtil.AcertaNumero(IIf(Len(Mt75(a).Base_Icms_subst) > 0, Mt75(a).Base_Icms_subst, 0), 2)
   valor = Replace(valor, ",", "")
   valor = Replace(valor, ".", "")
   valor = Right(String(13, "0") & valor, 13)
   
   AliquotaIpi = ClUtil.AcertaNumero(CStr(Mt75(a).Aliquota_Ipi), 2)
   AliquotaIpi = Replace(AliquotaIpi, ".", "")
   AliquotaIpi = Replace(AliquotaIpi, ",", "")
   AliquotaIpi = Right("0000" & AliquotaIpi, 5)
             
   AliquotaIcms = ClUtil.AcertaNumero(CStr(Mt75(a).Aliquota_Icms), 2)
   AliquotaIcms = Replace(AliquotaIcms, ".", "")
   AliquotaIcms = Replace(AliquotaIcms, ",", "")
   AliquotaIcms = Right("0000" & AliquotaIcms, 4)
          
   StrSintegra = "75"
   StrSintegra = StrSintegra & Format(Datai, "yyyymmdd") '==> Data inicial
   StrSintegra = StrSintegra & Format(Dataf, "yyyymmdd") '==> Data Final
   StrSintegra = StrSintegra & Left(Mt75(a).Codigo & String(14, " "), 14) '==> Codigo do Produto
   StrSintegra = StrSintegra & String(8, " ") '==> Codigo NCM
   StrSintegra = StrSintegra & Left(Trim(Mt75(a).Nome) & String(53, " "), 53) '==> descricao do Produto
   StrSintegra = StrSintegra & Left(IIf(Len(Trim(Mt75(a).Unidade)) > 0, Mt75(a).Unidade, "PC") & String(6, " "), 6) '==> Unidade
   StrSintegra = StrSintegra & AliquotaIpi '==> Aliquota IPI
   StrSintegra = StrSintegra & AliquotaIcms '==> Aliquota ICMS
   StrSintegra = StrSintegra & "00000" '==> Reducao base de Calculo
             
   StrSintegra = StrSintegra & valor '===> Base de calculo da substituicao Tributaria
   TotalReg75 = TotalReg75 + 1
   Print #LcArq, StrSintegra
Next

End Sub

Public Sub EscreverRegistro53()
Dim StrSintegea As String
Dim a As Long
Dim total As String
Dim ValorIcmsRetido As String
Dim BaseIcmsSubs As String
Dim AliquotaIcms As String
Dim ValorQuant  As String
Dim ClUtil As New Utilitario
Dim DespAcess As String
TotalReg53 = 0

For a = 0 To UBound(Mt53)
  If Mt53(a).Icms_Retido > 0 Then
    ValorIcmsRetido = ClUtil.AcertaNumero(CStr(Mt53(a).Icms_Retido), 2)
    ValorIcmsRetido = Replace(ValorIcmsRetido, ".", "")
    ValorIcmsRetido = Replace(ValorIcmsRetido, ",", "")
    ValorIcmsRetido = Right("0000000000000" & ValorIcmsRetido, 13)
    
    BaseIcmsSubs = ClUtil.AcertaNumero(CStr(Mt53(a).Base_Cal_Subst), 2)
    BaseIcmsSubs = Replace(BaseIcmsSubs, ".", "")
    BaseIcmsSubs = Replace(BaseIcmsSubs, ",", "")
    BaseIcmsSubs = Right("0000000000000" & BaseIcmsSubs, 13)
    
    DespAcess = ClUtil.AcertaNumero(CStr(Mt53(a).Despesas_Acessorias), 2)
    DespAcess = Replace(DespAcess, ".", "")
    DespAcess = Replace(DespAcess, ",", "")
    DespAcess = Right("0000000000000" & DespAcess, 13)
    
    StrSintegra = "53"
    StrSintegra = StrSintegra & Mt53(a).Cnpj
    StrSintegra = StrSintegra & Mt53(a).Inscricao
    StrSintegra = StrSintegra & Mt53(a).data
    StrSintegra = StrSintegra & Mt53(a).estado
    StrSintegra = StrSintegra & Mt53(a).Modelo
    StrSintegra = StrSintegra & Mt53(a).Serie
    StrSintegra = StrSintegra & Mt53(a).Numero_Nf
    StrSintegra = StrSintegra & Mt53(a).CFOP
    StrSintegra = StrSintegra & Mt53(a).Emitente
    StrSintegra = StrSintegra & BaseIcmsSubs
    StrSintegra = StrSintegra & ValorIcmsRetido
    StrSintegra = StrSintegra & DespAcess
    StrSintegra = StrSintegra & Mt53(a).Situacao
    StrSintegra = StrSintegra & Mt53(a).Codigo_Antecipacao
    StrSintegra = StrSintegra & String(29, " ")
    TotalReg53 = TotalReg53 + 1
  '  MsgBox StrSintegra
    Print #LcArq, StrSintegra
  End If
Next
End Sub
Public Sub EscreverRegistro90()
On Error Resume Next
Dim StrSintegra As String
Dim db As Database
Dim Rs As Recordset
Dim StrSql As String
Dim Cnpj As String
Dim Inscricao As String
StrSql = "Select * from Empresa"
Set db = OpenDatabase(GLBase)

Set Rs = db.OpenRecordset(StrSql)

Sintegra.Caption = "Criando o Registro 90 - Finalização"
DoEvents

If Not Rs.EOF Then
   Cnpj = Rs!CGC
   Cnpj = Replace(Cnpj, ".", "")
   Cnpj = Replace(Cnpj, "-", "")
   Cnpj = Replace(Cnpj, "/", "")
   Cnpj = Replace(Cnpj, "\", "")
   Cnpj = Replace(Cnpj, " ", "")
   Cnpj = Right("00000000000000" & Cnpj, 14)
   '==> Inscricao
   Inscricao = Rs!inscricaoestadual & ""
   Inscricao = Replace(Inscricao, ".", "")
   Inscricao = Replace(Inscricao, "-", "")
   Inscricao = Replace(Inscricao, "/", "")
   Inscricao = Replace(Inscricao, "\", "")
   Inscricao = Replace(Inscricao, " ", "")
   Inscricao = Left(Inscricao & "              ", 14)
End If

StrSintegra = "90"
StrSintegra = StrSintegra & Cnpj
StrSintegra = StrSintegra & Inscricao

If TotalReg50 > 0 Then
   StrSintegra = StrSintegra & "50"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg50, 8)
End If
If TotalReg53 > 0 Then
   StrSintegra = StrSintegra & "53"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg53, 8)
End If

If TotalReg54 > 0 Then
   StrSintegra = StrSintegra & "54"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg54, 8)
End If
If TotalReg60 > 0 Then
   StrSintegra = StrSintegra & "60"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg60, 8)
End If
If TotalReg70 > 0 Then
   StrSintegra = StrSintegra & "70"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg70, 8)
End If
If TotalReg74 > 0 Then
   StrSintegra = StrSintegra & "74"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg74, 8)
End If
If TotalReg75 > 0 Then
   StrSintegra = StrSintegra & "75"
   StrSintegra = StrSintegra & Right("00000000" & TotalReg75, 8)
End If
StrSintegra = StrSintegra & "99"
'TotalRegs = TotalReg75 + TotalReg50 + TotalReg53 + TotalReg54 + TotalReg70 + TotalReg74 + TotalReg60 + 3
'MsgBox TotalRegs
StrSintegra = StrSintegra & Right("00000000" & TotalReg75 + TotalReg50 + TotalReg53 + TotalReg54 + TotalReg70 + TotalReg74 + TotalReg60 + 3, 8)
StrSintegra = Left(StrSintegra & String(125, " "), 125) & "1"


Print #LcArq, StrSintegra

End Sub

Sub CriarReg60D(Serie_Impressora As String, Modelo_Impressora As String)
Dim Mt()        As DadosReg60D
Dim StrSintegra As String
Dim Rs          As ADODB.Recordset
Dim StrSql      As String
Dim a           As Integer
Dim b           As Integer
Dim C           As Integer
Dim Achou       As Boolean
Dim data        As Date
Dim Quantidade  As String
Dim valor       As String
Dim Cutil       As New Utilitario
Dim CSint       As New ClSintegra

a = 0
b = 0
C = 0
For a = 1 To Sintegra.Cupom.Rows - 1
    Sintegra.Caption = "Criando o Registro 60D - " & a & " de " & Sintegra.Cupom.Rows - 1
    DoEvents

    StrSql = "Select * from dadosprenota where doc='" & Sintegra.Cupom.TextMatrix(a, 11) & "'"
    Set Rs = AbreRecordset(StrSql, True)
    Do Until Rs.EOF
       '==> é o primeiro registro
       If b = 0 Then
            ReDim Preserve Mt(b)
            Mt(b).Codigo = Rs!Codigo
            Mt(b).Descricao = Rs!Descricao
            Mt(b).icms = IIf(Not IsNull(Rs!icms), Rs!icms, 0)
            Mt(b).ipi = IIf(Not IsNull(Rs!ipi), Rs!ipi, 0)
            Mt(b).Unidade = Rs!unid
          
            Mt(b).data = CDate(Sintegra.Cupom.TextMatrix(a, 1))
            Mt(b).Quantidade = Rs!quant
            Mt(b).valor = Rs!total
            Mt(b).Situacao = Sintegra.Cupom.TextMatrix(a, 10)
            b = b + 1
       Else
         '==> ja existe regisro, vamos ver se existe o produto.
         For C = 0 To b - 1
           Achou = False
           If Mt(C).Codigo = Rs!Codigo And Mt(C).data = Sintegra.Cupom.TextMatrix(a, 1) And Mt(C).Situacao = Sintegra.Cupom.TextMatrix(a, 10) Then
              Achou = True
              Exit For
           End If
         Next
         If Achou Then
            Mt(C).Quantidade = Mt(C).Quantidade + Rs!quant
            Mt(C).valor = Mt(C).valor + Rs!total
         Else
            ReDim Preserve Mt(b)
            Mt(b).Codigo = Rs!Codigo
            Mt(b).Descricao = Rs!Descricao
            Mt(b).icms = IIf(Not IsNull(Rs!icms), Rs!icms, 0)
            Mt(b).ipi = IIf(Not IsNull(Rs!ipi), Rs!ipi, 0)
            Mt(b).Unidade = Rs!unid
            Mt(b).data = CDate(Sintegra.Cupom.TextMatrix(a, 1))
            Mt(b).Quantidade = Rs!quant
            Mt(b).valor = Rs!total
            Mt(b).Situacao = Sintegra.Cupom.TextMatrix(a, 10)
            b = b + 1
         End If
       End If
       Rs.MoveNext
    Loop
  
Next

'==> Vamos gerar o arquivo do sintegra
On Error Resume Next
err.Number = 0
a = UBound(Mt)
If err.Number <> 0 Then Exit Sub
For a = 0 To UBound(Mt)
    Sintegra.Caption = "Gravando o Registro 60D - do dia " & Mt(a).data & " Reg " & a + 1 & " de " & UBound(Mt) + 1
    DoEvents

   StrSintegra = "60"
   StrSintegra = StrSintegra & "D"
   StrSintegra = StrSintegra & Format(Mt(a).data, "yyyymmdd")
   StrSintegra = StrSintegra & Left(Serie_Impressora & "                    ", 20)
   StrSintegra = StrSintegra & Left(Mt(a).Codigo & String(14, " "), 14)
   Quantidade = Mt(a).Quantidade
   Quantidade = Cutil.AcertaNumero(Quantidade, 3)
   Quantidade = Replace(Quantidade, ",", "")
   Quantidade = Replace(Quantidade, ".", "")
   
   Quantidade = Right(String(13, "0") & Quantidade, 13)
   
   valor = Mt(a).valor
   valor = Cutil.AcertaNumero(valor, 2)
   valor = Replace(valor, ",", "")
   valor = Replace(valor, ".", "")
   valor = Right(String(16, "0") & valor, 16)
   
   StrSintegra = StrSintegra & Quantidade
   StrSintegra = StrSintegra & valor
   StrSintegra = StrSintegra & String(16, "0")
   StrSintegra = StrSintegra & IIf(Mt(a).Situacao = "S", "CANC", "F   ")
   StrSintegra = StrSintegra & String(13, "0")
   StrSintegra = StrSintegra & String(19, " ")
   Print #LcArq, StrSintegra
   TotalReg60 = TotalReg60 + 1
   CSint.VerificaProdPara75 Mt(a).Codigo, Mt(a).Descricao, Mt(a).icms, Mt(a).ipi, "0", Mt(a).Unidade

  ' VerificaProdPara75 Mt(a).Codigo
Next
End Sub

Sub CriarRegistro60I(Serie_Impressora As String, Modelo_Impressora As String)
On Error GoTo ErroCria60I
Dim Rs60I As ADODB.Recordset
Dim StrSql As String
Dim a As Integer
Dim StrSintegra As String
Dim Quantidade As String
Dim Valor_Liquido As String
Dim ClUtil As New Utilitario
Dim ClSinte As New ClSintegra

For a = 1 To Sintegra.Cupom.Rows - 1
    Sintegra.Caption = "Criando o Registro 60I - " & a & " de " & Sintegra.Cupom.Rows - 1
    DoEvents

   StrSql = "Select * from DadosPreNota where Doc='" & Sintegra.Cupom.TextMatrix(a, 11) & "'"
   Set Rs60I = AbreRecordset(StrSql, True)
   '==> Gera um para cada produto
   Do Until Rs60I.EOF
        Quantidade = Rs60I!quant
        Quantidade = ClUtil.AcertaNumero(Quantidade, 3)
        Quantidade = Replace(Quantidade, ".", "")
        Quantidade = Replace(Quantidade, ",", "")
        ValorDesconto = IIf(Not IsNull(Rs60I!Desconto), Rs60I!Desconto, 0)
        Valor_Liquido = Rs60I!Unit - (ValorDesconto / Rs60I!quant)
        Valor_Liquido = ClUtil.AcertaNumero(Valor_Liquido, 2)
        Valor_Liquido = Replace(Valor_Liquido, ".", "")
        Valor_Liquido = Replace(Valor_Liquido, ",", "")
        
        StrSintegra = "60"   '==>1
        StrSintegra = StrSintegra & "I" '==>2
        StrSintegra = StrSintegra & Format(Sintegra.Cupom.TextMatrix(a, 1), "yyyymmdd") '==>3
        StrSintegra = StrSintegra & Left(Serie_Impressora & "                    ", 20) '==>4
        StrSintegra = StrSintegra & Modelo_Impressora '==>5
        StrSintegra = StrSintegra & Right("000000" & Rs60I!NF, 6) '==>6
        StrSintegra = StrSintegra & Right("000" & Rs60I!Item, 3) '==>7
        StrSintegra = StrSintegra & Left(Rs60I!Codigo & String(14, " "), 14) '==>8
        StrSintegra = StrSintegra & Right("0000000000000" & Quantidade, 13) '==>9
        StrSintegra = StrSintegra & Right("0000000000000" & Valor_Liquido, 13) '==>10
        StrSintegra = StrSintegra & "000000000000" '==>11
        '==> Verifica o Status para marcar a situacao trib do cupom
        If Sintegra.Cupom.TextMatrix(a, 10) = "S" Then
           '==> Cancelado
           StrSintegra = StrSintegra & "CANC"
        Else
          '==> Pode ser que o Item esteje cancelado
          If Rs60I!Cancelado Then
             '==> Cancelado
            StrSintegra = StrSintegra & "CANC"
          Else
            '==> Normal
            StrSintegra = StrSintegra & "F   "
          End If
          
        End If
        StrSintegra = StrSintegra & "000000000000"
        StrSintegra = StrSintegra & String(16, " ")
        
        Print #LcArq, StrSintegra
        ClSinte.VerificaProdPara75 Rs60I!Codigo, Rs60I!Descricao, IIf(Not IsNull(Rs60I!icms), Rs60I!icms, 0), IIf(Not IsNull(Rs60I!ipi), Rs60I!ipi, 0), "0", Rs60I!unid
        
        'VerificaProdPara75 Rs60I!Codigo
        TotalReg60 = TotalReg60 + 1
        Rs60I.MoveNext
   Loop
   Set Rs60I = Nothing
Next

Exit Sub
ErroCria60I:
MsgBox err.Description & err.Number
Resume 0
End Sub


Sub CriarRegistro60M()
On Error GoTo erroReg60M
Dim a                   As Integer
Dim Rs                  As ADODB.Recordset
Dim RsEmpresa           As ADODB.Recordset
Dim StrSql              As String
Dim StrSintegra         As String
Dim Numero_serie        As String
Dim Modelo_Equipamento  As String
Dim Totalizador_geral   As String
Dim Venda_Bruta         As String
Dim ClUtil              As New Utilitario

StrSql = "Select * from Sintegra_Cupom where Data Between #" & Format(Sintegra.Datai.Text, "mm/dd/yy") & _
         "# And #" & Format(Sintegra.Dataf.Text, "mm/dd/yy") & "# order by Data"

Set Rs = AbreRecordset(StrSql, True)
'==> Abre a tb empresa
StrSql = "Select * from empresa"
Set RsEmpresa = AbreRecordset(StrSql, True)
If Not RsEmpresa.EOF Then
   Numero_serie = RsEmpresa!SerieImpressora & ""
   Modelo_Equipamento = Mid(RsEmpresa!Modelo_Equipamento, 1, 2) & ""
End If
Set RsEmpresa = Nothing
If Not Rs.EOF Then
   Rs.MoveLast
   LcTotal = Rs.RecordCount
   Rs.MoveFirst
End If
Dim x As Integer
TotalReg60 = 0
Do Until Rs.EOF
    x = x + 1
    Sintegra.Caption = "Criando o Registro 60M - " & x & " de " & LcTotal
    DoEvents

    Totalizador_geral = ClUtil.AcertaNumero(CStr(Rs!Totalizador_geral), 2)
    Totalizador_geral = Replace(Totalizador_geral, ",", "")
    Totalizador_geral = Replace(Totalizador_geral, ".", "")
    
    Venda_Bruta = ClUtil.AcertaNumero(CStr(Rs!Valor_Venda), 2)
    Venda_Bruta = Replace(Venda_Bruta, ",", "")
    Venda_Bruta = Replace(Venda_Bruta, ".", "")

    StrSintegra = "60"
    StrSintegra = StrSintegra & "M"
    StrSintegra = StrSintegra & Format(Rs!data, "yyyymmdd")
    StrSintegra = StrSintegra & Left(Numero_serie & "                    ", 20)
    StrSintegra = StrSintegra & "001"
    StrSintegra = StrSintegra & Modelo_Equipamento
    StrSintegra = StrSintegra & Right("000000" & Rs!Contador_inicio_dia, 6)
    StrSintegra = StrSintegra & Right("000000" & Rs!Contador_fim_Dia, 6)
    StrSintegra = StrSintegra & Right("000000" & Rs!Contador_Reducao_Z, 6)
    StrSintegra = StrSintegra & Right("000" & Rs!Contador_Reinicio_Operacao, 3)
    StrSintegra = StrSintegra & Right("0000000000000000" & Venda_Bruta, 16)
    StrSintegra = StrSintegra & Right("0000000000000000" & Totalizador_geral, 16)
    StrSintegra = StrSintegra & String(37, " ")
    Print #LcArq, StrSintegra
    GeraRegistro60A Rs, Numero_serie
    TotalReg60 = TotalReg60 + 1
    Rs.MoveNext
Loop
Set Rs = Nothing
CriarRegistro60I Numero_serie, Modelo_Equipamento
CriarReg60D Numero_serie, Modelo_Equipamento
Exit Sub
erroReg60M:
MsgBox err.Description & err.Number
'Resume 0
'Stop
End Sub
Sub GeraRegistro60A(RsReg60 As ADODB.Recordset, Serie_Impressora As String)
On Error GoTo errgerareg60A
Dim StrSintegra As String
Dim Rs60A As ADODB.Recordset
Dim Valor_Venda_Bruta As String
Dim StrSql As String
Dim Id_Situacao() As String
Dim Situacao As String
Dim x As Integer
Dim LcTotal As Integer
Dim ClUtil  As New Utilitario

StrSql = "Select * from Sintegra_Situacao_tributaria where Data=#" & Format(RsReg60!data, "mm/dd/yy") & "#"

Set Rs60A = AbreRecordset(StrSql, True)
If Not Rs60A.EOF Then
   Rs60A.MoveLast
   LcTotal = Rs60A.RecordCount
   Rs60A.MoveFirst
   
End If
Do Until Rs60A.EOF
    x = x + 1
    Sintegra.Caption = "Criando o Registro 60A - " & a & " de " & x & " Produto - " & LcTotal
    DoEvents

    Valor_Venda_Bruta = CStr(Rs60A!Valor_Acumulador_Parcial)
    Valor_Venda_Bruta = ClUtil.AcertaNumero(Valor_Venda_Bruta, 2)
    Valor_Venda_Bruta = Replace(Valor_Venda_Bruta, ",", "")
    Valor_Venda_Bruta = Replace(Valor_Venda_Bruta, ".", "")
    
    StrSintegra = "60"
    StrSintegra = StrSintegra & "A"
    StrSintegra = StrSintegra & Format(Rs60A!data, "yyyymmdd")
    StrSintegra = StrSintegra & Left(Serie_Impressora & "                    ", 20)
    Id_Situacao = Split(Rs60A!Indentificacao_Situacao, "-")
    Situacao = Id_Situacao(0)
    Situacao = Trim(Situacao)
    StrSintegra = StrSintegra & Left(Situacao & "    ", 4)
    StrSintegra = StrSintegra & Right("000000000000" & Valor_Venda_Bruta, 12)
    StrSintegra = StrSintegra & String(79, " ")
    Print #LcArq, StrSintegra
    TotalReg60 = TotalReg60 + 1
    Rs60A.MoveNext
Loop
Set Rs60A = Nothing



Exit Sub
errgerareg60A:
MsgBox err.Description & err.Number
End Sub

