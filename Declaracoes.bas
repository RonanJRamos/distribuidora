Attribute VB_Name = "Declaracoes"
Declare Function ConsisteInscricaoEstadual Lib "DllInscE32.dll" (ByVal Insc As String, ByVal UF As String) As Integer
Public Type DadosResumo
    CFOP As String
    ValorTotal As Double
    baseCalculo As Double
    valorIcms As Double
    ValorIsentas As Double
    ValorOutras As Double
End Type
Public Type dados60a
    valor As Double
    Base As Double
    icms As Double
End Type

Public Type Dados54
    Cnpj                    As String
    Modelo                  As String
    Serie                   As String
    Numero_Nf               As String
    CFOP                    As String
    cst                     As String
    Numero_Item             As String
    Codigo_Produto          As String
    Quantidade              As Double
    Valor_Produto           As Double
    Valor_Desconto          As Double
    Base_Calculo_Icms       As Double
    Base_Calculo_subs_Trib  As Double
    Valor_Ipi               As Double
    Aliquota_Icms           As Double
    
End Type
Public Type Dados75
    Codigo          As String
    Nome            As String
    Unidade         As String
    Aliquota_Ipi    As Double
    Aliquota_Icms   As Double
    Reducao_Base    As Double
    Base_Icms_subst As Double
End Type
Public Type DadosComplementares
    Cnpj                    As String
    Modelo                  As String
    Serie                   As String
    Numero_Nf               As String
    CFOP                    As String
    Valor_Frete             As Double
    Valor_Seguro            As Double
    Valor_pis               As Double
    Valor_Compl             As Double
    Valor_Servicos          As Double
    Valor_Despesas          As Double
    Total_Itens             As Long
    Codigo_Complementar     As Long
    Valor_Complementar      As Double
End Type
Public Type Dados50
    Cnpj                    As String
    Inscricao               As String
    Data                    As String
    Estado                  As String
    Modelo                  As String
    Serie                   As String
    Numero_Nf               As String
    CFOP                    As String
    Emitente                As String
    Valor_Total             As Double
    Base_Calculo_Icms       As Double
    Valor_Icms              As Double
    Isenta_Nao_Tributada    As Double
    Outras                  As Double
    Aliquota                As Double
    Situacao                As String
    
End Type

Public Type Dados70
    Cnpj                    As String
    Inscricao               As String
    Data                    As String
    Estado                  As String
    Modelo                  As String
    Serie                   As String
    SubSerie                As String
    Numero_Nf               As String
    CFOP                    As String
    Valor_Total             As Double
    Base_Calculo_Icms       As Double
    Valor_Icms              As Double
    Isenta_Nao_Tributada    As Double
    Outras                  As Double
    CifFob                  As String
    Situacao                As String
End Type
Public Type Dados53
    Cnpj                    As String
    Inscricao               As String
    Data                    As String
    Estado                  As String
    Modelo                  As String
    Serie                   As String
    Numero_Nf               As String
    CFOP                    As String
    Emitente                As String
    Base_Cal_Subst          As Double
    Icms_Retido             As Double
    Despesas_Acessorias     As Double
    Situacao                As String
    Codigo_Antecipacao      As String
End Type
Public Type Dados74
    Data                    As String
    codigoproduto           As String
    Quantidade              As Double
    ValorProduto            As Double
    Codigo_Posse            As String
    Cnpj                    As String
    Inscricao               As String
    Estado                  As String
End Type

Public Type dadosVerificaIcms
   icms     As Double
   valor    As Double
End Type
Public GlComissaoBelclean As Boolean
Public TotalReg50 As Long
Public TotalReg53 As Long
Public TotalReg54 As Long
Public TotalReg60 As Long
Public TotalReg70 As Long
Public TotalReg74 As Long
Public TotalReg75 As Long
Public Total75    As Long
Public Mt75()     As Dados75
Public Mt50()     As Dados50
Public Mt53()     As Dados53
Public Mt54()     As Dados54
Public Mt70()     As Dados70
Public Mt74()     As Dados74
Public ClienteForaEstado As Boolean
Public MtIcms()     As dadosVerificaIcms
Public TemRegistro50 As Boolean
Public TemRegistro53 As Boolean
Public TemRegistro54 As Boolean
Public TemRegistro75 As Boolean
Public TemRegistro70 As Boolean
Public TemRegistro74 As Boolean
