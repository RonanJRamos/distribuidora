Attribute VB_Name = "Gerador"
Public Function CriarDados70(Registro As Dados70) As Boolean
Dim b As Long

If Not TemRegistro70 Then
   b = 0
   TemRegistro70 = True
Else
    b = UBound(Mt70) + 1
    ReDim Preserve Mt70(b)
End If
ReDim Preserve Mt70(b)
Mt70(b).Cnpj = Registro.Cnpj
Mt70(b).Inscricao = Registro.Inscricao
Mt70(b).Data = Registro.Data
Mt70(b).Estado = Registro.Estado
Mt70(b).Modelo = Registro.Modelo
Mt70(b).Serie = Registro.Serie 'Left("1" & "   ", 3)
Mt70(b).SubSerie = Registro.SubSerie
Mt70(b).Numero_Nf = Registro.Numero_Nf
Mt70(b).CFOP = Registro.CFOP
Mt70(b).Valor_Total = Registro.Valor_Total
Mt70(b).Base_Calculo_Icms = Registro.Base_Calculo_Icms
Mt70(b).Valor_Icms = Registro.Valor_Icms
Mt70(b).Isenta_Nao_Tributada = Registro.Isenta_Nao_Tributada
Mt70(b).Outras = Registro.Outras
Mt70(b).CifFob = Registro.CifFob
Mt70(b).Situacao = Registro.Situacao

End Function
Public Function CriarRegistro53(Registro As Dados53) As Boolean
Dim a As Long
On Error GoTo erroCriaDados53

If TemRegistro53 Then
   a = UBound(Mt53)
Else
   a = -1
   TemRegistro53 = True
End If
a = a + 1

ReDim Preserve Mt53(a)
Mt53(a).Base_Cal_Subst = Registro.Base_Cal_Subst
Mt53(a).CFOP = Registro.CFOP
Mt53(a).Cnpj = Registro.Cnpj
Mt53(a).Codigo_Antecipacao = Registro.Codigo_Antecipacao
Mt53(a).Data = Registro.Data
Mt53(a).Despesas_Acessorias = Registro.Despesas_Acessorias
Mt53(a).Emitente = Registro.Emitente
Mt53(a).Estado = Registro.Estado
Mt53(a).Icms_Retido = Registro.Icms_Retido
Mt53(a).Inscricao = Registro.Inscricao
Mt53(a).Modelo = Registro.Modelo
Mt53(a).Numero_Nf = Registro.Numero_Nf
Mt53(a).Serie = Registro.Serie
Mt53(a).Situacao = Registro.Situacao


CriarDados53 = True

Exit Function

erroCriaDados53:
CriarDados53 = False
End Function
Public Function CriarDados50(Registro As Dados50) As Boolean
Dim Achou       As Boolean
Dim Cnpj        As String
Dim ValorTotal  As String
Dim ValorBase   As String
Dim a           As Long
Dim b           As Long
Dim C           As Long

On Error GoTo errocria50

If TemRegistro50 Then
 '==> busca para saber se tem este registro 50
    Achou = False
     For C = 0 To UBound(Mt50)
         If Mt50(C).Aliquota = Registro.Aliquota And Mt50(C).CFOP = Registro.CFOP _
            And Mt50(C).Cnpj = Registro.Cnpj And Mt50(C).Emitente = Registro.Emitente _
            And Mt50(C).Numero_Nf = Registro.Numero_Nf Then
                Achou = True
                Exit For
         End If
     Next
     
      
     If Achou Then
        Mt50(C).Valor_Total = Mt50(C).Valor_Total + Registro.Valor_Total
        Mt50(C).Base_Calculo_Icms = Mt50(C).Base_Calculo_Icms + Registro.Base_Calculo_Icms
        Mt50(C).Valor_Icms = Mt50(C).Valor_Icms + (Registro.Base_Calculo_Icms * (Registro.Aliquota / 100))
        Mt50(C).Outras = Mt50(C).Outras + Registro.Outras
     Else
        b = UBound(Mt50) + 1
        ReDim Preserve Mt50(b)
        Mt50(b).Cnpj = Registro.Cnpj
        Mt50(b).Inscricao = Registro.Inscricao
        Mt50(b).Data = Registro.Data
        Mt50(b).Estado = Registro.Estado
        Mt50(b).Modelo = Registro.Modelo
        Mt50(b).Serie = Registro.Serie 'Left("1" & "   ", 3)
        Mt50(b).Numero_Nf = Registro.Numero_Nf
        Mt50(b).CFOP = Registro.CFOP
        Mt50(b).Emitente = Registro.Emitente
        Mt50(b).Valor_Total = Registro.Valor_Total
        Mt50(b).Base_Calculo_Icms = Registro.Base_Calculo_Icms
        Mt50(b).Valor_Icms = Registro.Base_Calculo_Icms * (Registro.Aliquota / 100)
        Mt50(b).Isenta_Nao_Tributada = Registro.Isenta_Nao_Tributada
        Mt50(b).Outras = Registro.Outras
        Mt50(b).Aliquota = Registro.Aliquota
        Mt50(b).Situacao = Registro.Situacao
   End If
  Else
     '==> Cria o registro 50
     TemRegistro50 = True
     b = 0
     ReDim Preserve Mt50(b)
     Mt50(b).Cnpj = Registro.Cnpj
     Mt50(b).Inscricao = Registro.Inscricao
     Mt50(b).Data = Registro.Data
     Mt50(b).Estado = Registro.Estado
     Mt50(b).Modelo = Registro.Modelo
     Mt50(b).Serie = Registro.Serie 'Left("1" & "   ", 3)
     Mt50(b).Numero_Nf = Registro.Numero_Nf
     Mt50(b).CFOP = Registro.CFOP
     Mt50(b).Emitente = Registro.Emitente
     Mt50(b).Valor_Total = Registro.Valor_Total
     Mt50(b).Base_Calculo_Icms = Registro.Base_Calculo_Icms
     Mt50(b).Valor_Icms = Registro.Base_Calculo_Icms * (Registro.Aliquota / 100)
     Mt50(b).Isenta_Nao_Tributada = Registro.Isenta_Nao_Tributada
     Mt50(b).Outras = Registro.Outras
     Mt50(b).Aliquota = Registro.Aliquota
     Mt50(b).Situacao = Registro.Situacao
End If
CriarDados50 = True
Exit Function

errocria50:
MsgBox err.Description & err.Number
Resume 0
CriarDados50 = False
End Function

Public Function CriarDados54(Registro As Dados54) As Boolean
Dim a As Long
On Error GoTo erroCriaDados54

If TemRegistro54 Then
   a = UBound(Mt54)
Else
   a = -1
   TemRegistro54 = True
End If
a = a + 1

ReDim Preserve Mt54(a)
Mt54(a).Modelo = Registro.Modelo
Mt54(a).Cnpj = Registro.Cnpj
Mt54(a).Serie = Registro.Serie ' Left("1" & "   ", 3)
Mt54(a).Numero_Nf = Registro.Numero_Nf
Mt54(a).CFOP = Registro.CFOP
Mt54(a).cst = Registro.cst ' Right("000" & RsDados!Cst, 3)
Mt54(a).Numero_Item = Registro.Numero_Item ' Right("000" & RsDados!Item, 3)
Mt54(a).Codigo_Produto = Registro.Codigo_Produto 'Left(RsDados!codigopesquisa & "              ", 14)
Mt54(a).Quantidade = Registro.Quantidade
Mt54(a).Valor_Produto = Registro.Valor_Produto
Mt54(a).Valor_Desconto = Registro.Valor_Desconto
Mt54(a).Base_Calculo_Icms = Registro.Base_Calculo_Icms 'BaseIcms
Mt54(a).Base_Calculo_subs_Trib = Registro.Base_Calculo_subs_Trib '"000000000000"
Mt54(a).Valor_Ipi = Registro.Valor_Ipi
Mt54(a).Aliquota_Icms = Registro.Aliquota_Icms
CriarDados54 = True

Exit Function

erroCriaDados54:
CriarDados54 = False
End Function
Public Function CriaRegistro74(Registro As Dados74) As Boolean
Dim a As Long
On Error GoTo erroCriaDados74

If TemRegistro74 Then
   a = UBound(Mt74)
Else
   a = -1
   TemRegistro74 = True
End If
a = a + 1

ReDim Preserve Mt74(a)
Mt74(a).Cnpj = Registro.Cnpj
Mt74(a).Codigo_Posse = Registro.Codigo_Posse
Mt74(a).CodigoProduto = Registro.CodigoProduto
Mt74(a).Data = Registro.Data
Mt74(a).Estado = Registro.Estado
Mt74(a).Inscricao = Registro.Inscricao
Mt74(a).Quantidade = Registro.Quantidade
Mt74(a).ValorProduto = Registro.ValorProduto

CriaRegistro74 = True

Exit Function
erroCriaDados74:
CriaRegistro74 = False
End Function
Public Function CriarDados75(Registro As Dados75) As Boolean
On Error GoTo ErroCria75

Dim a As Integer
Dim Existe As Boolean
Dim ClUtil As New Utilitario
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
            If Trim(Mt75(a).Codigo) = Trim(Registro.Codigo) Then
               Existe = True
               Exit For
            End If
        Next
    Else
      Existe = False
    End If

End If

If Not Existe Then
    
    AliquotaIcms = ClUtil.AcertaNumero(CStr(Registro.Aliquota_Icms), 2)
    AliquotaIcms = Replace(AliquotaIcms, ".", "")
    AliquotaIcms = Replace(AliquotaIcms, ",", "")
    AliquotaIcms = Right("0000" & AliquotaIcms, 4)
    
    AliquotaIpi = ClUtil.AcertaNumero(CStr(Registro.Aliquota_Ipi), 2)
    AliquotaIpi = Replace(AliquotaIpi, ".", "")
    AliquotaIpi = Replace(AliquotaIpi, ",", "")
    AliquotaIpi = Right("00000" & AliquotaIpi, 5)
    
    ReDim Preserve Mt75(a)
    Mt75(a).Codigo = Registro.Codigo
    Mt75(a).Nome = Registro.Nome
    Mt75(a).Aliquota_Icms = AliquotaIcms
    Mt75(a).Aliquota_Ipi = AliquotaIpi
    Mt75(a).Base_Icms_subst = Registro.Base_Icms_subst
    Mt75(a).Reducao_Base = Registro.Reducao_Base '"00000"
    Mt75(a).Unidade = Registro.Unidade
End If
CriarDados75 = True

Exit Function
ErroCria75:
CriarDados75 = False
End Function
Public Function Processar_Dados_complementares(Registro As DadosComplementares) As Boolean
Dim a               As Long
Dim TotalRegistros  As Integer

TotalRegistros = UBound(MtIcms) + 1
'For a = 0 To UBound(MtIcms)
    If Registro.Valor_Compl > 0 Then
       Registro.Valor_Complementar = Registro.Valor_Compl / TotalRegistros
       Registro.Codigo_Complementar = 997
       Criar_Dados_complementares Registro
    End If
    
    If Registro.Valor_Despesas > 0 Then
       Registro.Valor_Complementar = Registro.Valor_Despesas / TotalRegistros
       Registro.Codigo_Complementar = 999
       Criar_Dados_complementares Registro
    
    End If
    If Registro.Valor_Frete > 0 Then
       Registro.Valor_Complementar = Registro.Valor_Frete / TotalRegistros
       Registro.Codigo_Complementar = 991
       Criar_Dados_complementares Registro
    End If
    If Registro.Valor_pis > 0 Then
       Registro.Valor_Complementar = Registro.Valor_pis / TotalRegistros
       Registro.Codigo_Complementar = 993
       Criar_Dados_complementares Registro
    End If
    
    If Registro.Valor_Seguro > 0 Then
       Registro.Valor_Complementar = Registro.Valor_Seguro / TotalRegistros
       Registro.Codigo_Complementar = 992
       Criar_Dados_complementares Registro
    End If
    If Registro.Valor_Servicos > 0 Then
       Registro.Valor_Complementar = Registro.Valor_Servicos / TotalRegistros
       Registro.Codigo_Complementar = 998
       Criar_Dados_complementares Registro
    
    End If
'Next
End Function
Public Function Criar_Dados_complementares(Registro As DadosComplementares) As Boolean
On Error GoTo erroCriacomple

Dim a           As Long
Dim StrSql      As String
Dim BaseIcms    As String
Dim ClUtil      As New Utilitario
On Error Resume Next
err.Number = 0
a = UBound(Mt54)
If err.Number <> 0 Then a = 0 Else a = a + 1
Resume 0
ReDim Preserve Mt54(a)
Mt54(a).Cnpj = Registro.Cnpj '2
Mt54(a).Modelo = Registro.Modelo '3
Mt54(a).Serie = Registro.Serie '4
Mt54(a).Numero_Nf = Registro.Numero_Nf '5
Mt54(a).CFOP = Registro.CFOP '6
Mt54(a).cst = "   "  '7
Mt54(a).Numero_Item = Registro.Codigo_Complementar ' CodigoComplemento   '8
Mt54(a).Codigo_Produto = Right("                ", 14) '9
Mt54(a).Quantidade = "00000000000"  '10
Mt54(a).Valor_Produto = "000000000000"  '11
                            
BaseIcms = ClUtil.AcertaNumero(CStr(Registro.Valor_Complementar), 2)
BaseIcms = Replace(BaseIcms, ".", "")
BaseIcms = Replace(BaseIcms, ",", "")
              
Mt54(a).Valor_Desconto = Right("000000000000" & BaseIcms, 12) '12
Mt54(a).Base_Calculo_Icms = "000000000000" '13
Mt54(a).Base_Calculo_subs_Trib = "000000000000" '14
Mt54(a).Valor_Ipi = "000000000000"  '15
Mt54(a).Aliquota_Icms = "0000" '16
Criar_Dados_complementares = True
Exit Function
erroCriacomple:
Criar_Dados_complementares = False
End Function


