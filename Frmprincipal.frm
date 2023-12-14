VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPrincipal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   8265
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H00800080&
   Icon            =   "Frmprincipal.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   840
      Top             =   2640
   End
   Begin MSComDlg.CommonDialog Abrird 
      Left            =   7200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Serie 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sdsddsdsdddsd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label LbEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vhvhvvbvbbvb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label LbSistema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VIRTUAL DSA LOCADORA  <<<VERSÃO 4.6 >>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   6915
   End
   Begin VB.Menu SistemaPr 
      Caption         =   "&Sistema"
      Begin VB.Menu SAir 
         Caption         =   "&Sair"
      End
      Begin VB.Menu mnUser 
         Caption         =   "&Trocar Usuário"
      End
   End
   Begin VB.Menu Cadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu CadCliente 
         Caption         =   "&Clientes"
         Begin VB.Menu CadCliIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu CadCliAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu CadCliConsulta 
            Caption         =   "&Consulta"
         End
         Begin VB.Menu MnLancfaGrEconomico 
            Caption         =   "Lançar Grupo Economico"
         End
      End
      Begin VB.Menu CadFornecedor 
         Caption         =   "&Fornecedor"
         Begin VB.Menu CadForIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu CadForAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu CadForConsulta 
            Caption         =   "&Consulta"
         End
      End
      Begin VB.Menu CadProduto 
         Caption         =   "&Produto"
         Begin VB.Menu CadProIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu CadProAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu CadProConsulta 
            Caption         =   "&Consulta"
         End
      End
      Begin VB.Menu MnFunc 
         Caption         =   "Fu&ncionarios"
         Begin VB.Menu MnFuncInc 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnFuncAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnFuncConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnGalapao 
         Caption         =   "&Galpão"
         Begin VB.Menu MnGalpaoIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnGalpaoAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnGalpaoConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnCidade 
         Caption         =   "C&idade"
         Begin VB.Menu MnCidadeIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnCidadeAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnCidadeConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnTipo 
         Caption         =   "&Tipo Monetário"
         Begin VB.Menu MnTipoIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnTipoAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnTipoConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnTipoREc 
         Caption         =   "&Tip&o Receitas e Despesas"
         Begin VB.Menu MnIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnTipoRecAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnTipoREcConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnUnidade 
         Caption         =   "&Unidade"
         Begin VB.Menu MnUnidadeIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnUnidadeAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnUnidadeConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnTransportadoras 
         Caption         =   "T&ransportadora"
         Begin VB.Menu MnTraspincluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnTranspAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnTranspConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnCusto 
         Caption         =   "Cu&sto"
         Begin VB.Menu MnIncluirCusto 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnCustoAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnCustoConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MNGrupoEconomico 
         Caption         =   "&Grupo Economico"
      End
      Begin VB.Menu mnEmpresa 
         Caption         =   "Dados da Empresa"
      End
      Begin VB.Menu MnCadCedente 
         Caption         =   "Cadastro Cedente Boleto A4"
      End
      Begin VB.Menu MnLinhaContrato 
         Caption         =   "-"
      End
      Begin VB.Menu MnContrato 
         Caption         =   "Contratos de Fornecimento"
      End
      Begin VB.Menu MnNaturezaNfe 
         Caption         =   "Natureza de Operação NF-e"
      End
   End
   Begin VB.Menu MnFinanceiro 
      Caption         =   "&Financeiro"
      Begin VB.Menu MnFinaReceitas 
         Caption         =   "&Receitas"
         Begin VB.Menu MnFinaRecIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnFinaRecAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnFinaRecConsultar 
            Caption         =   "&Consultar"
         End
         Begin VB.Menu MnFinarecBaixa 
            Caption         =   "&Baixa"
         End
         Begin VB.Menu MNAcertaCredito 
            Caption         =   "Acerta Credito Utilizado Cliente"
         End
      End
      Begin VB.Menu MnFinaDesp 
         Caption         =   "&Despesas"
         Begin VB.Menu MnFinaDespIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnFinaAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnFinaDespConsultar 
            Caption         =   "&Consultar"
         End
         Begin VB.Menu MnFinaDespBaixa 
            Caption         =   "&Baixa"
         End
      End
      Begin VB.Menu MnCheques 
         Caption         =   "&Cheques Recebidos"
         Begin VB.Menu MnChequesIncluir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnChAlterar 
            Caption         =   "&Alterar"
         End
         Begin VB.Menu MnChConsultar 
            Caption         =   "&Consultar"
         End
      End
      Begin VB.Menu MnCaixa 
         Caption         =   "Fechamento de Cai&xa"
      End
      Begin VB.Menu MnComissoes 
         Caption         =   "C&omissões"
      End
      Begin VB.Menu MnCORe 
         Caption         =   "Comissão &Representada"
      End
      Begin VB.Menu MnFluxoCaixa 
         Caption         =   "&Fluxo de Caixa"
         Visible         =   0   'False
      End
      Begin VB.Menu MnCobranca 
         Caption         =   "Cobrança Bancaria"
         Begin VB.Menu MnRemessa 
            Caption         =   "Remessa"
         End
         Begin VB.Menu MnRetorno 
            Caption         =   "Retorno"
         End
      End
   End
   Begin VB.Menu mnestoque 
      Caption         =   "&Estoque"
      Begin VB.Menu MNEntrada 
         Caption         =   "&Entrada"
      End
      Begin VB.Menu MNEntradaXML 
         Caption         =   "Entrada XML"
      End
      Begin VB.Menu MnSaida 
         Caption         =   "&Saida"
         Begin VB.Menu MnIncluiir 
            Caption         =   "&Incluir"
         End
         Begin VB.Menu MnConsultarNF 
            Caption         =   "Consultar"
         End
         Begin VB.Menu MnCancelar 
            Caption         =   "&Cancelar"
         End
      End
      Begin VB.Menu mn1 
         Caption         =   "-"
      End
      Begin VB.Menu MnNFe 
         Caption         =   "Nota Fiscal Eletrônica"
      End
      Begin VB.Menu MnCancelaNFE 
         Caption         =   "Cancelar Nota fiscal Eletronica"
      End
      Begin VB.Menu MnInutilizaNumeros 
         Caption         =   "Inutilização de Numeros"
      End
      Begin VB.Menu mnespaco 
         Caption         =   "-"
      End
      Begin VB.Menu MnVales 
         Caption         =   "&Vales"
      End
      Begin VB.Menu MnCancelaVale 
         Caption         =   "&Excluir Vale"
         Visible         =   0   'False
      End
      Begin VB.Menu mn2 
         Caption         =   "-"
      End
      Begin VB.Menu MnProposta 
         Caption         =   "Pedido de &Vendas"
      End
      Begin VB.Menu MnPedido 
         Caption         =   "S&olicitação de  Vendas"
         Visible         =   0   'False
      End
      Begin VB.Menu MnAlteraPreco 
         Caption         =   "&Alteração de Preços"
      End
      Begin VB.Menu MnOrcamento 
         Caption         =   "O&rçamento e Vendas"
         Visible         =   0   'False
      End
      Begin VB.Menu mncancelaOrcamento 
         Caption         =   "&Cancela Vendas / Orçamento"
         Visible         =   0   'False
      End
      Begin VB.Menu MnPesqComprasCli 
         Caption         =   "&Pesquisa Compras Cliente"
      End
      Begin VB.Menu MnFicha 
         Caption         =   "&Ficha de Estoque"
      End
      Begin VB.Menu MnRomaneioNF 
         Caption         =   "Romaneio NF"
      End
      Begin VB.Menu MnRomaneio 
         Caption         =   "Romaneio Pedido"
      End
      Begin VB.Menu mnestoquefiscal 
         Caption         =   "Verifica estoque fiscal"
      End
   End
   Begin VB.Menu Rel 
      Caption         =   "&Relatórios"
      Begin VB.Menu relcli 
         Caption         =   "&Clientes"
         Begin VB.Menu RelClienteNome 
            Caption         =   "&Nome"
         End
         Begin VB.Menu relclitel 
            Caption         =   "C&lientes/Telemarketing"
         End
         Begin VB.Menu RelCliCondicoes 
            Caption         =   "&Condições Especiais"
         End
         Begin VB.Menu RelCliCompras 
            Caption         =   "&Compras Período"
         End
         Begin VB.Menu RelCliNaoCompraram 
            Caption         =   "&Não Compraram - Período"
         End
         Begin VB.Menu RelCliDevedores 
            Caption         =   "&Devedores"
         End
         Begin VB.Menu FrmCLiCidade 
            Caption         =   "C&idade"
         End
         Begin VB.Menu RelCliBairro 
            Caption         =   "&Bairro"
         End
         Begin VB.Menu RelCliEstado 
            Caption         =   "&Estado"
         End
         Begin VB.Menu RelCliTele 
            Caption         =   "&Telemarketing"
         End
         Begin VB.Menu MnRelCliMala 
            Caption         =   "&Mala Direta"
         End
         Begin VB.Menu mnComprasProdPeriodo 
            Caption         =   "Produtos Comprados Periodo"
         End
      End
      Begin VB.Menu RelFornecedor 
         Caption         =   "&Fornecedores"
         Begin VB.Menu MnNomeFornec 
            Caption         =   "&Nome"
         End
         Begin VB.Menu MnFornecCompras 
            Caption         =   "&Compras Período"
         End
         Begin VB.Menu MnFornecidade 
            Caption         =   "&Cidade"
         End
      End
      Begin VB.Menu MnComissao 
         Caption         =   "C&omissão"
         Begin VB.Menu RelComissaoVendedor 
            Caption         =   "&Vendedor"
         End
         Begin VB.Menu RelComissaoFornec 
            Caption         =   "&Fornecedor"
         End
         Begin VB.Menu RelComissaoVendFornec 
            Caption         =   "V&endedor + Fornecedor"
         End
         Begin VB.Menu MnComRepresentada 
            Caption         =   "&Representada"
         End
      End
      Begin VB.Menu RelProdutos 
         Caption         =   "&Produtos"
         Begin VB.Menu RelProNome 
            Caption         =   "&Nome"
         End
         Begin VB.Menu RelProdSimples 
            Caption         =   "&Simples"
         End
         Begin VB.Menu RelProTabela 
            Caption         =   "&Tabela Preço"
         End
         Begin VB.Menu RelTabMax 
            Caption         =   "Tabela &Valor Max/Min"
         End
         Begin VB.Menu RelProForn 
            Caption         =   "&Fornecedor"
         End
         Begin VB.Menu MnProdCompras 
            Caption         =   "&Compras - Minimo"
         End
         Begin VB.Menu ComprasFornec 
            Caption         =   "Compras - Fornecedor"
         End
         Begin VB.Menu MnGapoos 
            Caption         =   "&Galpões"
         End
         Begin VB.Menu mnPosBalanco 
            Caption         =   "Posição de Balanço"
         End
         Begin VB.Menu MnIventario 
            Caption         =   "Iventario fiscal"
         End
         Begin VB.Menu MnRelFisacalCustoMedio 
            Caption         =   "Inventario Fiscal com Custo Medio"
         End
         Begin VB.Menu MnProdutosNaoComprados 
            Caption         =   "Produtos não comprados"
         End
         Begin VB.Menu MnAcompanhaVendidos 
            Caption         =   "Acompanhamento Vendidos"
         End
      End
      Begin VB.Menu RelEntradaEstoque 
         Caption         =   "&Entrada de Estoque"
         Begin VB.Menu RelEntrEstoque 
            Caption         =   "&Nota Fiscal"
         End
         Begin VB.Menu relnfprod 
            Caption         =   "&NF Produto"
         End
         Begin VB.Menu RelEntradaEstClie 
            Caption         =   "&Fornecedor"
         End
         Begin VB.Menu RelEntradEstPeriodo 
            Caption         =   "&Período"
         End
         Begin VB.Menu MnRelEntrProduto 
            Caption         =   "P&roduto"
         End
         Begin VB.Menu MnPosContabil 
            Caption         =   "Posição Contabil"
         End
      End
      Begin VB.Menu MNCte 
         Caption         =   "&CTe"
      End
      Begin VB.Menu RelSaidaEstoque 
         Caption         =   "&Notas de Saidas"
         Begin VB.Menu RelNotaSaida 
            Caption         =   "&Nota Fiscal"
         End
         Begin VB.Menu RelCliSaida 
            Caption         =   "&Cliente"
         End
         Begin VB.Menu RelPeriodoSaida 
            Caption         =   "&Período"
         End
         Begin VB.Menu MNVendaCFOP 
            Caption         =   "Por CFOP"
         End
         Begin VB.Menu MnDetalhesNota 
            Caption         =   "&Detalhes da Nota"
         End
         Begin VB.Menu mnVSubsEstado 
            Caption         =   "&Vendas com Subs. P/ o Estado"
         End
         Begin VB.Menu MnPosicaoContSaida 
            Caption         =   "&Posição Contabil"
         End
         Begin VB.Menu MnResumoProduto 
            Caption         =   "&Resumo dos Produtos Vendidos"
         End
         Begin VB.Menu MnVendaPorFornecedor 
            Caption         =   "Venda por Fornecedor"
         End
      End
      Begin VB.Menu RelPedido 
         Caption         =   "&Solicitação de Compras"
         Visible         =   0   'False
         Begin VB.Menu RelPedidoCliente 
            Caption         =   "&Cliente"
         End
         Begin VB.Menu RelNumPedido 
            Caption         =   "&Número Pedido"
         End
         Begin VB.Menu RelPedidoPeriodo 
            Caption         =   "&Período"
         End
         Begin VB.Menu RelPedidoFornec 
            Caption         =   "&Fornecedor"
         End
      End
      Begin VB.Menu RelOrcamento 
         Caption         =   "&Orçamento/Venda"
         Begin VB.Menu RelOrcamNumero 
            Caption         =   "&Número"
         End
         Begin VB.Menu RelOrcPeriodo 
            Caption         =   "&Período"
         End
         Begin VB.Menu RelOrcCli 
            Caption         =   "&Cliente"
         End
      End
      Begin VB.Menu MnFuncionarios 
         Caption         =   "&Funcionários"
         Begin VB.Menu MnFuncMala 
            Caption         =   "&Mala Direta"
         End
      End
      Begin VB.Menu MnRelReceita 
         Caption         =   "&Receita"
         Begin VB.Menu RelRecDocumento 
            Caption         =   "&Documento"
         End
         Begin VB.Menu RelReceiPeriodo 
            Caption         =   "&Período"
         End
         Begin VB.Menu RelReceitaCliente 
            Caption         =   "&Cliente"
         End
         Begin VB.Menu mmConfec 
            Caption         =   "C&onferência de Emissão"
         End
         Begin VB.Menu MnConfereRecebimento 
            Caption         =   "Conferencia de Recebimento"
         End
      End
      Begin VB.Menu RelDesp 
         Caption         =   "&Despesas"
         Begin VB.Menu RelDespDoc 
            Caption         =   "&Documento"
         End
         Begin VB.Menu RelDespPeriodo 
            Caption         =   "&Período"
         End
         Begin VB.Menu RelDespFornec 
            Caption         =   "&Fornecedor"
         End
         Begin VB.Menu RelDespTipoMonet 
            Caption         =   "&Tipo Monetário"
         End
         Begin VB.Menu mnconfDesp 
            Caption         =   "C&onferência"
         End
      End
      Begin VB.Menu RelCheq 
         Caption         =   "C&heques"
         Begin VB.Menu RelCheqNumero 
            Caption         =   "&Número"
         End
         Begin VB.Menu RelCheqPedido 
            Caption         =   "&Solicitação de Compras"
            Visible         =   0   'False
         End
         Begin VB.Menu RlCheqCliente 
            Caption         =   "&Cliente"
         End
         Begin VB.Menu RelCheqBanco 
            Caption         =   "&Banco"
         End
         Begin VB.Menu RelCheqDataEntrada 
            Caption         =   "P&rogramados"
         End
         Begin VB.Menu RelDataDeposito 
            Caption         =   "&Data de Depósito"
         End
      End
      Begin VB.Menu MnRelCaixa 
         Caption         =   "Cai&xa"
         Begin VB.Menu MnCaixaDia 
            Caption         =   "&Dia"
         End
         Begin VB.Menu MnCaixaPeriodo 
            Caption         =   "&Periodo"
         End
      End
      Begin VB.Menu MNInvFiscalPeriodo 
         Caption         =   "Inventario Fiscal por periodo"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Utilitarios 
      Caption         =   "&Utilitários"
      Begin VB.Menu mndisquetereceita 
         Caption         =   "&Gerar Disquete Sintegra"
      End
      Begin VB.Menu MnSegunca 
         Caption         =   "Copia de S&egurança"
         Begin VB.Menu MnBackup 
            Caption         =   "&Backup"
         End
         Begin VB.Menu MnRecuperar 
            Caption         =   "&Recuperar"
         End
      End
      Begin VB.Menu Senha 
         Caption         =   "&Senha"
         Begin VB.Menu MnSenhaGrupo 
            Caption         =   "&Grupos"
         End
         Begin VB.Menu MnSenharUser 
            Caption         =   "&Usuario"
         End
      End
      Begin VB.Menu PanoFundo 
         Caption         =   "&Pano de Fundo"
      End
      Begin VB.Menu AlteraData 
         Caption         =   "&Alterar Data do Sistema"
      End
      Begin VB.Menu mReparar 
         Caption         =   "&Reparar Banco de Dados"
      End
      Begin VB.Menu MnOpcoes 
         Caption         =   "&Opções"
      End
      Begin VB.Menu EtMalaDireta 
         Caption         =   "&Configura Etiqueta Mala Direta"
      End
      Begin VB.Menu LcLocalizar 
         Caption         =   "&Localizar Banco de Dados"
      End
      Begin VB.Menu VerExclusao 
         Caption         =   "&Ver Dados Excluídos"
      End
   End
   Begin VB.Menu info 
      Caption         =   "Informações do Sistema"
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Option Explicit

'=== PARA CAPTURAR A PASTA DESEJADA ========================================
'necessário para acionar o browser
Private Type tProcuraInformação
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private a As Integer
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As tProcuraInformação) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lL_Item As Long, ByVal sDir As String) As Long
'===========================================================================

'Private CAB As cCAB
Private CorLetra, CorInicio As Long
Function verificaini()
Dim LcPAth As String
'LcPAth = BuscaDirWin & "\" & App.EXEName & ".ini"

'If Dir(LcPAth, vbArchive) = "" Then
 '  configuracoes.Show , Me
'End If
End Function
Private Sub AlteraData_Click()
On Error Resume Next
frmDataSisema.Show , Me
End Sub



Private Sub CadCliAlterar_Click()
LcTipoDados = 2
FrmCliente.Show , Me

End Sub

Private Sub CadCliConsulta_Click()
LcTipoDados = 3
FrmCliente.Show , Me
End Sub

Private Sub CadCliIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmCliente.Show , Me
End Sub

Private Sub CadForAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmFornecedor.Show , Me
End Sub

Private Sub CadForConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmFornecedor.Show , Me
End Sub

Private Sub CadForIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmFornecedor.Show , Me

End Sub

Private Sub CadFunAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmFuncionarios.Show , Me
End Sub

Private Sub CadFunConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmFuncionarios.Show , Me
End Sub

Private Sub CadFunIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmFuncionarios.Show , Me
End Sub

Private Sub CadProAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmProduto.Show , Me
End Sub

Private Sub CadProConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmProduto.Show , Me
End Sub

Private Sub CadProIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmProduto.Show , Me
End Sub

Private Sub FinaDespAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmDespesas.Show , Me
End Sub

Private Sub FinaDespBaixa_Click()
On Error Resume Next
FrmBaixaDespesa.Show , Me
End Sub

Private Sub FinaDespConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmDespesas.Show , Me
End Sub

Private Sub FinaDespIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmDespesas.Show , Me
End Sub

Private Sub FinaFechaCaixa_Click()
On Error Resume Next
FrmFechaCaixa.Show , Me
End Sub

Private Sub FinaRecAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmContasReceber.Show , Me
End Sub

Private Sub FinaRecBaixa_Click()
On Error Resume Next
FrmBaixaReceita.Show , Me
End Sub

Private Sub FinaRecConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmContasReceber.Show , Me
End Sub

Private Sub FinaRecIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmContasReceber.Show , Me
End Sub

Private Sub Caixa_Click()

End Sub

Private Sub ComprasFornec_Click()
On Error Resume Next
comprasfornecedor.Show , Me
End Sub

Private Sub EtMalaDireta_Click()
On Error Resume Next
ConfiguraEtiqueta.Show , Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
GlSisCarregado = True
LcAlterado = False
GLPesquisa = False
If GlboletoA4 Then
   MnCadCedente.Visible = True
Else
  MnCadCedente.Visible = fasle
End If

If GlUsaEstoqueSeguranca Then
   MnConsultarNF.Visible = True
Else
  MnConsultarNF.Visible = fasle
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 123 Then
'   ConsertaCgc
'End If
If KeyCode = 122 Then Teste.Show , Me
If KeyCode = 121 Then
Unload Me
End If
End Sub
Function ConsertaCgc()
On Error GoTo errcgc
Dim Lc1, Lc2, Lc3, Lc4, lc5, LcCep, LcCgc, LcInsc As String
AbreBase
AbreBanco (Cliente)
LcCap = Me.Caption
Me.Caption = "Processando Dados dos clientes, aguarde..."
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
Me.Caption = LcCap
Exit Function
errcgc:
Exit Function
End Function

Private Sub Form_Load()
On Error Resume Next
If Dir(App.Path & "\fundo.bmp", vbDirectory) <> "" Then
    Me.Picture = LoadPicture(App.Path & "\fundo.bmp")
End If
'cria nova ocorrência da classe...
'Set CAB = New cCAB
'ObtemPath
GlFormInicial = Me.Name
Top = 0
montapainel
Me.Picture = LoadPicture(App.Path & "\fundo.bmp")
'ObtemPath
'VerificaOpcoes
MnCORe.Visible = GlRepresentante
RelComissaoFornec.Visible = GlRepresentante
RelComissaoVendFornec.Visible = GlRepresentante
MnComRepresentada.Visible = GlRepresentante
'ConsertaCgc
'FrmPrincipal.Visible = False
'FrmApresentacao.Show
'frmDataSisema.Show
'frmLogin.SetFocus
verificaini
End Sub
Function montapainel()
On Error Resume Next
StatusBar.Panels(1).Width = 2000
StatusBar.Panels(2).Width = 2000
StatusBar.Panels(3).Width = 8400
'StatusBar.Panels(4).Width = 2000
'StatusBar.Panels(6).Width = 2000

StatusBar.Panels(1).Text = "Usuário:" & GlUsuario
If Len(GlNomeMaquina) = 0 Then
  StatusBar.Panels(2).Text = "Local"
Else
  StatusBar.Panels(2).Text = "Máquina:" & GlNomeMaquina
End If
StatusBar.Panels(3).Text = "Base Atual:" & GLBase
'StatusBar.Panels(5).Text = "Data Atual: " & Format(Date, "dd/mm/yy")
'StatusBar.Panels(6).Text = "Hora Atual: " & Format(Time(), "hh:mm:ss")
End Function
Private Sub LocLoc_Click()
On Error Resume Next
FrmLocacao.Show , Me
End Sub

Private Sub LocProdEst_Click()
GlLocado = 0
GlCap = "Pesquisa Produtos que estão em Estoque"
FrmPesquisaProdutoLocado.Show , Me

End Sub

Private Sub LocPRoLoc_Click()
On Error Resume Next
GlLocado = 1
GlCap = "Pesquisa Produtos que Estão Locados"
FrmPesquisaProdutoLocado.Show , Me

End Sub

Private Sub logoff_Click()
On Error Resume Next
frmLogin.Show
FrmPrincipal.Visible = False
End Sub

Private Sub MnConvAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmConvenio.Show , Me
End Sub

Private Sub MnConvConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmConvenio.Show , Me
End Sub

Private Sub MnConvIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmConvenio.Show , Me
End Sub

Private Sub MnLocAlteracao_Click()
On Error Resume Next
FrmAlteracaoFitas.Show , Me
End Sub

Private Sub MnLocDevolucao_Click()
On Error Resume Next
FrmDevolucao.Show , Me
End Sub

Private Sub mnLocVendas_Click()
On Error Resume Next
FrmVendas.Show , Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Not Cab Is Nothing Then
    Set Cab = Nothing
End If
End
End Sub

Private Sub FrmCLiCidade_Click()
On Error Resume Next
FrmRelClienteCidade.Show , Me
End Sub

Private Sub info_Click()
frmSplash.Show , Me
End Sub

Private Sub LcLocalizar_Click()
On Error Resume Next
'configuracoes.Show , Me
Dim Arqini As String
Arqini = BuscaDirWin
Arqini = App.EXEName & ".ini"
LocalizarBanco Arqini
abreconexao
End Sub

Private Sub mmConfec_Click()
FrmRelReceitaConferencia.Show , Me
End Sub

Private Sub MNAcertaCredito_Click()
FrmAcertaLimiteCliente.Show , Me
End Sub

Private Sub MnAcompanhaVendidos_Click()
FrmProdutosVendidos.Show , Me
End Sub

Private Sub MnAlteraPreco_Click()
FrmReajustaPreco.Show , Me
End Sub

Private Sub MnBackup_Click()
On Error Resume Next
'Set CAB = New cCAB
'frmCab.Show , Me
'Exit Sub
Dim iNum            As Integer
Dim iLinha          As Integer
Dim sArquivo()      As String
Dim LcResposta      As Long
Dim LcCriterio      As String
Dim LcArq           As String
Dim LcCaminho       As String
Dim LCLEtra         As String
Dim LcDestino       As String
Dim LcCap           As String
Dim LcAchou         As Boolean
Dim a               As Integer
Dim LcCamiMak       As String
Dim LcCamiCopia     As String
  

If Not Dbbase Is Nothing Then
   Dbbase.Close
   Set Dbbase = Nothing
End If
LcCap = Me.Caption
Me.Caption = "Aguarde a finalização do Backup...."
DoEvents
LcCamiMak = BuscaDirWin & "\Makecab.exe"
LcCamiCopia = App.Path & "\makecab.exe"
' ===Verifica se existe o Arquivo de compactacao no diretorio Windows
If Dir(LcCamiMak) = "" Then
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
   LcCamiMak = BuscaDirWin & "\Extract.exe"
   LcCamiCopia = App.Path & "\Extract.exe"
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
End If
   
'MONTA LISTA DE ARQUIVOS A COMPRIMIR E CHAMA A ROTINA DE COMPRESSÃO
'se existem arquivos selecionados...
'==== Separa Arquivo e diretorio

For a = Len(GLBase) To 1 Step -1
    LCLEtra = Mid(GLBase, a, 1)
    If LCLEtra = "\" And Not LcAchou Then
       LcAchou = True
    Else
       If Not LcAchou Then
          LcArq = LCLEtra & LcArq
       Else
          LcCaminho = LCLEtra & LcCaminho
       End If
    End If
Next
ReDim Preserve sArquivo(0)
'atualiza o valor deste membro...
LcDestino = "c:\" & LcArq
x = CopyFile(Trim$(GLBase), Trim(LcDestino), False)
If x Then
  'Call Log("Ok")
Else
  MsgBox "Erro Copiando Base de Dados.", 64, "Aviso"
  Exit Sub
          
End If
'FileCopy "d:\dadoscli\leir\lidis.mdb", LcDestino
sArquivo(iLinha) = LcDestino
    
    '======================================
    'ajusta as propriedades, se necessários
    
    'se o drive padrão é A:\ , não precisa ser informado porque é o padrão
    'CAB.BackupDrive = B
    
    'esta propriedade usa por padrao o diretório temp do Windows
    'mais se quiser, pode indicar outro...
    'CAB.PastaTemp = "C:\Temp"
    '=====================================
    
    'chama rotina de compactação com os parâmetros escolhidos...
    Cab.Comprimir True, LcCaminho, sArquivo()
    
'Else
 '   MsgBox "Escolha arquivos os que deseja compactar", vbOKOnly + vbCritical, "Erro"
'End If
'LcCriterio = "Arj u -rva a:BacDSA.arj " & GLBase
Me.Caption = "Aguarde, Apagando Arquivos Temporários...."
Kill LcDestino
'Shell (LcCriterio), 6
Me.Caption = LcCap
End Sub

Private Sub MnCadCedente_Click()
'On Error Resume Next
Dim ClBoleto As New ControlaBoleto
Dim LocalAr As String
Dim ClAcesso As New AcessoAdo.Acessos
LocalAr = App.EXEName & ".ini"
ClBoleto.NomeProjeto = LocalAr

ClBoleto.ConfiguraConta
Set ClBoleto = Nothing

End Sub

Private Sub MnCaixa_Click()
Caixa.Show , Me
End Sub

Private Sub MnCaixaDia_Click()
On Error Resume Next
FrmRelCaixa.Show , Me
End Sub

Private Sub MnCaixaPeriodo_Click()
On Error Resume Next
FrmRelCaixaPerido.Show , Me
End Sub

Private Sub MnCancelaNFE_Click()
'On Error Resume Next
Dim LcCancelaNFe As New Decisao_NFE_FrmCancelaNFe
'Set LcCancelaNFe = New Decisao_NFE.FrmCancelaNFe
'LcCancelaNFe.Load

LcCancelaNFe.Nome_do_Projeto = GlNomeProjeto
LcCancelaNFe.Sistema_Implementado = GlSistemaImplementado

LcCancelaNFe.Show ' , Me
End Sub

Private Sub mncancelaOrcamento_Click()
On Error Resume Next
CancelaOrcamneto.Show , Me
End Sub

Private Sub MnCancelar_Click()
On Error Resume Next
CancelaNota.Show , Me
End Sub

Private Sub MnCancelaVale_Click()
On Error Resume Next
ConfirmaSenha.Show , Me
End Sub

Private Sub MnChAlterar_Click()
On Error Resume Next
LcTipoDados = 2
Frmcheques.Show , Me
End Sub

Private Sub MnChConsultar_Click()
On Error Resume Next
LcTipoDados = 3
Frmcheques.Show , Me
End Sub

Private Sub MnChequesIncluir_Click()
On Error Resume Next
LcTipoDados = 1
Frmcheques.Show , Me
End Sub

Private Sub MnCidadeAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmCidade.Show , Me
End Sub

Private Sub MnCidadeConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmCidade.Show , Me
End Sub

Private Sub MnCidadeIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmCidade.Show , Me
End Sub



Private Sub MnComissoes_Click()
On Error Resume Next
BaixaComissao.Show , Me
End Sub

Private Sub mnComprasProdPeriodo_Click()
On Error Resume Next
RelResumoProdutoCliente.Show , Me
End Sub

Private Sub MnComRepresentada_Click()
On Error Resume Next
FrmRelComissaoRepresent.Show , Me
End Sub

Private Sub mnconfDesp_Click()
On Error Resume Next
FrmRelDespesaConferencia.Show , Me
End Sub

Private Sub MnConfereRecebimento_Click()
On Error Resume Next
FrmConferenciaQuitacao.Show , Me
End Sub

Private Sub MnConsultarNF_Click()
On Error Resume Next
FrmSaidaProdutoAlternativo.Show , Me
End Sub

Private Sub MnContrato_Click()
On Error Resume Next
ContratoFornecimento.Show , Me
End Sub

Private Sub MnCORe_Click()
On Error Resume Next
BaixaComissaoRepresent.Show , Me
End Sub

Private Sub MNCte_Click()
On Error Resume Next
FrmRelatorioCTE.Show , Me
End Sub

Private Sub MnCustoAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmCusto.Show , Me
End Sub

Private Sub MnCustoConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmCusto.Show , Me
End Sub

Private Sub MnDetalhesNota_Click()
On Error Resume Next
FrmRelSaidaDetalhe.Show , Me
End Sub

Private Sub mndisquetereceita_Click()
On Error Resume Next
Dim ClGeraSintegra As New ClSintegra
ClGeraSintegra.AbreFormulario

End Sub

Private Sub mnEmpresa_Click()
On Error Resume Next
ClBoleto.NomeProjeto = ClAcesso.NomedoArquivo
Empresa.Show , Me
End Sub

Private Sub MNEntradaXML_Click()
'On Error Resume Next
Dim LcCaminho As String
Dim Retorno As Double
LcCaminho = App.Path & "\ScaeNet.exe 12 " & GlUsuario & " DOC LIDIS CodProd NomeProd"
Retorno = Shell(LcCaminho, vbNormalFocus)
End Sub

Private Sub mnestoquefiscal_Click()
On Error Resume Next
BaseSintegra.Show , Me
End Sub

Private Sub MnFicha_Click()
fichadeestoque.Show , Me
End Sub

Private Sub MnFinaAlterar_Click()
On Error Resume Next
LcTipoDados = 2
Despesas.Show , Me
End Sub

Private Sub MnFinaDespBaixa_Click()
FrmBaixaDespesas.Show , Me
End Sub

Private Sub MnFinaDespConsultar_Click()
On Error Resume Next
LcTipoDados = 3
Despesas.Show , Me
End Sub

Private Sub MnFinaDespIncluir_Click()
On Error Resume Next
LcTipoDados = 1
Despesas.Show , Me
End Sub

Private Sub MnFinaRecAlterar_Click()
LcTipoDados = 2
alid015.Show , Me
End Sub

Private Sub MnFinarecBaixa_Click()
FrmBaixaReceita.Show , Me
End Sub

Private Sub MnFinaRecConsultar_Click()
LcTipoDados = 3
alid015.Show , Me
End Sub

Private Sub MnFinaRecIncluir_Click()
On Error Resume Next
LcTipoDados = 1
alid015.Show , Me
End Sub

Private Sub MnFluxoCaixa_Click()
On Error Resume Next
FrmDetalhesCaixa.Show , Me
End Sub

Private Sub MnFornecCompras_Click()
On Error Resume Next
FrmRelFornecComprasPeriodo.Show , Me
End Sub

Private Sub MnFornecidade_Click()
On Error Resume Next
FrmRelFornecCidade.Show , Me
End Sub

Private Sub MnFuncMala_Click()
On Error Resume Next
MalaDiretaVendedores.Show , Me
End Sub

Private Sub MnGapoos_Click()
On Error Resume Next
RelProdutoEstoque.Show , Me
End Sub

Private Sub MNGrupoEconomico_Click()
On Error Resume Next
FrmGrupoEconomico.Show , Me

End Sub

Private Sub MnIncluiir_Click()
On Error Resume Next
FrmSaidaProduto.Show , Me
End Sub

Private Sub MnIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmTiporeceita.Show , Me
End Sub

Private Sub MnIncluirCusto_Click()
On Error Resume Next
LcTipoDados = 1
FrmCusto.Show , Me
End Sub



Private Sub MnInutilizaNumeros_Click()
'On Error Resume Next
Dim LcCancelaNFe As New Decisao_NFE_FrmInutilizacao

LcCancelaNFe.Nome_do_Projeto = GlNomeProjeto
LcCancelaNFe.Sistema_Implementado = GlSistemaImplementado

LcCancelaNFe.Show ' , Me
End Sub

Private Sub MnIventario_Click()
'Inventario.Show , Me
On Error Resume Next
RelatorioInventarioFiscal.Show , Me
End Sub

Private Sub MnLancfaGrEconomico_Click()
Dim LcLocal As String
LcLocal = App.Path & "\DistribuirGrupoEconomico.exe"
Shell LcLocal, vbNormalFocus

End Sub

Private Sub MnNaturezaNfe_Click()
Dim LcNota As New Decisao_NFE_FrmNaturezaOperacao
'Set LcNota = New Decisao_NFE_FrmNotaFiscalSaida
LcNota.Nome_do_Projeto = GlNomeProjeto
LcNota.Sistema_Implementado = GlSistemaImplementado
LcNota.Show
End Sub

Private Sub MnNFe_Click()
'Dim LcNota As New Decisao_NFE_FrmNotaFiscalSaida
'Set LcNota = New Decisao_NFE_FrmNotaFiscalSaida
'LcNota.Nome_do_Projeto = GlNomeProjeto
'LcNota.Sistema_Implementado = GlSistemaImplementado
'LcNota.Show
 Dim LcCommando As String
    LcCommando = App.Path & "\"
    'LcCommando = LcCommando & "SCAENET.exe 13 Usuario LIDIS LIDIS"
    LcCommando = LcCommando & "ChamaNFe.exe"
    Shell LcCommando, vbNormalFocus
End Sub

Private Sub MnNomeFornec_Click()
On Error Resume Next
FrmRelFornecedoresNome.Show , Me
End Sub

Private Sub MnOpcoes_Click()
On Error Resume Next
Opcoes.Show , Me
End Sub

Private Sub MnProdutoraAltera_Click()
On Error Resume Next
LcTipoDados = 2
FrmProdutora.Show , Me
End Sub

Private Sub MnProdutoraConsulta_Click()
On Error Resume Next
LcTipoDados = 3
FrmProdutora.Show , Me
End Sub

Private Sub MnProdutoraincluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmProdutora.Show , Me
End Sub


Private Sub mnReajuste_Click()
On Error Resume Next
FrmReajustaPreco.Show , Me
End Sub

Private Sub MnTeclaFuncCad_Click()
On Error Resume Next
ProgramaTeclas.Show , Me
End Sub

Private Sub MnTeclaFuncLoc_Click()
On Error Resume Next
FrmPrgTeclasLocacao.Show , Me
End Sub

Private Sub MnEntrada_Click()
FrmEntradaProduto.Show , Me
End Sub

Private Sub MnFuncAlterar_Click()
LcTipoDados = 2
FrmFuncionario.Show , Me
End Sub

Private Sub MnFuncConsultar_Click()
LcTipoDados = 3
FrmFuncionario.Show , Me
End Sub

Private Sub MnFuncInc_Click()
LcTipoDados = 1
FrmFuncionario.Show , Me
End Sub

Private Sub MnGalpaoAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmGalpao.Show , Me
End Sub

Private Sub MnGalpaoConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmGalpao.Show , Me
End Sub

Private Sub MnGalpaoIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmGalpao.Show , Me
End Sub

Private Sub MnOrcamento_Click()
On Error Resume Next

'FrmVendaOrcam.Show , Me
orcamento.Show , Me
End Sub

Private Sub MnPedido_Click()
FrmPedido.Show , Me
End Sub

Private Sub MnPEsquisaCli_Click()
On Error Resume Next
UltimasComprasClienteLocal.Show , Me
End Sub

Private Sub MnPesqComprasCli_Click()
On Error Resume Next
UltimasComprasClienteLocal.Show , Me
End Sub

Private Sub mnPosBalanco_Click()
On Error Resume Next
PosicaoBalanco.Show , Me
End Sub

Private Sub MnPosContabil_Click()
On Error Resume Next
ContabilEntrada.Show , Me
End Sub

Private Sub MnPosicaoContSaida_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
   ContabilSaida.Show , Me
End If
End Sub

Private Sub MnProdCompras_Click()
On Error Resume Next
RElMinimo.Show , Me
End Sub

Private Sub MnProdutosNaoComprados_Click()
On Error Resume Next
FrmRelComprasPeriodo.Show , Me
End Sub

Private Sub MnProposta_Click()
On Error Resume Next
FrmProposta.Show , Me
End Sub

Private Sub MnRecuperar_Click()
On Error Resume Next
Dim LcResposta      As Long
Dim LcCriterio      As String
Dim LCLEtra         As String
Dim LcArq           As String
Dim LcCaminho       As String
Dim LcAchou         As Boolean
Dim a               As Integer
Dim LcCamiMak       As String
Dim LcCamiCopia     As String
Dim sDiretorio      As String
Dim LcCap           As String
MsgBox "Para Efetuar a Restauração da Base de dados," & Chr(13) & "Todos as Máquinas deverão estar fora do Sistema.", vbInformation, "Aviso"
LcResposta = MsgBox("Esta Operação Irá Sobrescrever o Seu Banco de dados Atual," & Chr(13) & "As Alterações Feitas depois Deste Backup" & Chr(13) & Chr(13) & " SERÃO PERDIDAS." & Chr(13) & Chr(13) & "Confirma a Restauração ?", vbCritical + vbYesNo, "AVISO IMPORTANTE.")
If LcResposta = 7 Then
  MsgBox " Operação Cancelada pelo Usuário.", 64, "Aviso"
  Exit Sub
End If
If Not Dbbase Is Nothing Then
   Set Dbbase = Nothing
End If
'== Tenta Abrir o Banco em Modo Exclusivo Para Ver se existe Alguem Conectado.
Set Dbbase = OpenDatabase(GLBase, True, True)
If Not Dbbase Is Nothing Then
   Dbbase.Close
   'Set Dbbase = Nothing
End If
'LcCamiMak = App.Path & "\bak" & Format(Date, "ddmmyy") & ".mdb"
'X = CopyFile(Trim$(GLBase), Trim(LcCamiMak), False)

LcCap = Me.Caption
Me.Caption = "Aguarde a finalização da Restauração dos dados...."
DoEvents
LcCamiMak = BuscaDirWin & "\Makecab.exe"
LcCamiCopia = App.Path & "\makecab.exe"
' ===Verifica se existe o Arquivo de compactacao no diretorio Windows
If Dir(LcCamiMak) = "" Then
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
   LcCamiMak = BuscaDirWin & "\Extract.exe"
   LcCamiCopia = App.Path & "\Extract.exe"
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
End If

'captura escolha de diretório pelo usuário...
'sDiretorio = sProcuraPorDiretório("Diretório para descompressão de arquivos")
For a = Len(GLBase) To 1 Step -1
    LCLEtra = Mid(GLBase, a, 1)
    If LCLEtra = "\" And Not LcAchou Then
       LcAchou = True
    Else
       If Not LcAchou Then
          LcArq = LCLEtra & LcArq
       Else
          LcCaminho = LCLEtra & LcCaminho
       End If
    End If
Next
sDiretorio = LcCaminho
'se algum foi escolhido...
If sDiretorio <> "" Then

    'verifica contrabarra no caminho...
    sDiretorio = sFormataCaminho(sDiretorio)
    
    'chama rotuina de descompactação...
    Call Cab.Descomprimir(sDiretorio)
End If
Me.Caption = LcCap
Exit Sub
LcResposta = MsgBox("Insira o Primeiro Disco no << DRIVE A >> " & Chr(13) & _
"Os DADOS ATUAIS DO SEU SISTEMA SERÃO SUBSTITUIDOS PELOS DO BACKUP", 65, "Backup")
If LcResposta = 2 Then
   Exit Sub
End If
LcCriterio = "Arj x -rva a:BacDSA.arj " & GLBase

Shell (LcCriterio), 6


End Sub

Private Sub MnRelCliMala_Click()
MalaDiretaClientes.Show , Me
End Sub

Private Sub MnRelEntrProduto_Click()
On Error Resume Next
FrmRelEntradaNfPeriodo.Show , Me
End Sub

Private Sub MnRelFisacalCustoMedio_Click()
FrmInventarioCusto.Show , Me
End Sub

Private Sub MnRemessa_Click()
Dim LcLocal As String
LcLocal = App.Path & "\GerarBoleto.exe 0 GlNomeProjeto GlSistemaImplementado"
Shell LcLocal, vbNormalFocus
End Sub

Private Sub MnResumoProduto_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
   RelNotaSaidaResumoProdutos.Show , Me
End If
End Sub

Private Sub MnRetorno_Click()
Dim LcLocal As String
LcLocal = App.Path & "\GerarBoleto.exe -1 GlNomeProjeto GlSistemaImplementado"
Shell LcLocal, vbNormalFocus
End Sub

Private Sub MnRomaneio_Click()
On Error Resume Next
Romaneio.Show , Me
End Sub

Private Sub MnRomaneioNF_Click()
On Error Resume Next
RomaneioNota.Show , Me
End Sub

Private Sub MnSenhaGrupo_Click()
On Error Resume Next
FrmCadGrupo.Show , Me
End Sub

Private Sub MnSenharUser_Click()
On Error Resume Next
FrmUsuarios.Show , Me
End Sub

Private Sub MnTipoAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmTipoMonetario.Show , Me
End Sub

Private Sub MnTipoConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmTipoMonetario.Show , Me
End Sub

Private Sub MnTipoIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmTipoMonetario.Show , Me
End Sub

Private Sub MnTipoRecAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmTiporeceita.Show , Me
End Sub

Private Sub MnTipoREcConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmTiporeceita.Show , Me
End Sub

Private Sub MnTranspAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmTransportadora.Show , Me
End Sub

Private Sub MnTranspConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmTransportadora.Show , Me
End Sub

Private Sub MnTraspincluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmTransportadora.Show , Me
End Sub

Private Sub MnUnidadeAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmUnidade.Show , Me
End Sub

Private Sub MnUnidadeConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmUnidade.Show , Me
End Sub

Private Sub MnUnidadeIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmUnidade.Show , Me
End Sub

Private Sub mnUser_Click()
On Error Resume Next
frmLogin.Show , Me
End Sub

Private Sub MnVales_Click()
On Error Resume Next
FrmVales.Show , Me
End Sub

Private Sub MNVendaCFOP_Click()

On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
  FrmRelVendaCFOP.Show , Me
End If
End Sub

Private Sub MnVendaPorFornecedor_Click()
FrmVendaPorFornecedor.Show , Me
End Sub

Private Sub mnVSubsEstado_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
   FrmRelVendaEstado.Show , Me
End If
End Sub

Private Sub mReparar_Click()
'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
On Error Resume Next
Dim LcMsg As String

LcMsg = Me.Caption
If Not Dbbase Is Nothing Then
   Set Dbbase = Nothing
End If
'GLBase.Close
Me.Caption = "Aguarde, Reparando a Base de dados..."
Dbbase.Close
DBEngine.RepairDatabase GLBase
MsgBox "Operação Terminada com Sucesso..."
Me.Caption = LcMsg
End Sub

Private Sub MtEspAlterar_Click()
On Error Resume Next
LcTipoDados = 2
FrmEspecie.Show , Me
End Sub

Private Sub MtEspConsultar_Click()
On Error Resume Next
LcTipoDados = 3
FrmEspecie.Show , Me
End Sub

Private Sub MtRspIncluir_Click()
On Error Resume Next
LcTipoDados = 1
FrmEspecie.Show , Me
End Sub

Private Sub PanoFundo_Click()
On Error Resume Next
LocalizaFigura
End Sub

Private Sub RelCategoria_Click()
On Error Resume Next
FrmRelEspecie.Show , Me
End Sub



Private Sub RelConvenio_Click()
On Error Resume Next
FrmRelConvenio.Show , Me
End Sub

Private Sub RelDespesas_Click()
On Error Resume Next
FrmRelDespesas.Show , Me
End Sub

Private Sub RelCheCompensados_Click()
On Error Resume Next
FrmRelChequeCompensado.Show , Me
End Sub

Private Sub RelCheqBanco_Click()
On Error Resume Next
FrmRelChequeBanco.Show , Me
End Sub

Private Sub RelCheqDataEntrada_Click()
On Error Resume Next
FrmRelChequeEntrada.Show , Me
End Sub

Private Sub RelCheqNumero_Click()
On Error Resume Next
FrmRelChequeNumero.Show , Me
End Sub

Private Sub RelCheqPedido_Click()
On Error Resume Next
FrmRelChequePedido.Show , Me
End Sub

Private Sub RelCliBairro_Click()
On Error Resume Next
FrmRelClienteBairro.Show , Me

End Sub

Private Sub RelCliCompras_Click()
On Error Resume Next
FrmRelClienteComprasPeriodo.Show , Me

End Sub

Private Sub RelCliCondicoes_Click()
On Error Resume Next
FrmRelClienteCondicoes.Show , Me
End Sub

Private Sub RelCliDevedores_Click()
On Error Resume Next
FrmRelClienteDevedores.Show , Me
End Sub

Private Sub RelClienteNome_Click()
On Error Resume Next
FrmRelClienteNome.Show , Me
End Sub

Private Sub RelCliEstado_Click()
On Error Resume Next
FrmRelClienteEstado.Show , Me
End Sub

Private Sub RelCliNaoCompraram_Click()
On Error Resume Next
FrmRelClientenaocompraram.Show , Me
End Sub

Private Sub RelCliSaida_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
  FrmRelSaidaProdutoClien.Show , Me
End If
End Sub

Private Sub relclitel_Click()
On Error Resume Next
FrmRelClienteTelem.Show

End Sub

Private Sub RelCliTele_Click()
On Error Resume Next
FrmRelClienteTelemarketing.Show , Me
End Sub

Private Sub RelComissaoFornec_Click()
FrmRelComissaoFornecedor.Show , Me
End Sub

Private Sub RelComissaoVendedor_Click()
FrmRelComissao.Show , Me
End Sub

Private Sub RelComissaoVendFornec_Click()
On Error Resume Next
FrmRelComissaoFornecedorVendedor.Show , Me
End Sub

Private Sub RelDataDeposito_Click()
On Error Resume Next
FrmRelChequeDeposito.Show , Me
End Sub

Private Sub RelDespDoc_Click()
On Error Resume Next
FrmRelDespesaNumero.Show , Me
End Sub

Private Sub RelDespFornec_Click()
On Error Resume Next
FrmRelDespFornec.Show , Me
End Sub

Private Sub RelDespPeriodo_Click()
On Error Resume Next
FrmRelDespesaPeriodo.Show , Me
End Sub

Private Sub RelDespTipoMonet_Click()
On Error Resume Next
FrmRelTipoMonet.Show , Me
End Sub

Private Sub RelEntradaEstClie_Click()
On Error Resume Next
FrmRelEntradaProdutoClien.Show , Me
End Sub

Private Sub RelEntradaEstoque_Click()
'On Error Resume Next
'RelEntradaProdutos.Show , Me
End Sub



Private Sub RelFuncionarios_Click()
On Error Resume Next
FrmRelFuncionarios.Show , Me
End Sub

Private Sub RelLocacao_Click()
On Error Resume Next
FrmRelLocacao.Show , Me
End Sub

Private Sub RelProdutora_Click()
On Error Resume Next
FrmRelProdutora.Show , Me
End Sub



Private Sub RelReceitas_Click()
On Error Resume Next
FrmRelReceitas.Show , Me
End Sub

Private Sub RelReservas_Click()
On Error Resume Next
FrmRelReserva.Show , Me

End Sub

Private Sub ResConsulta_Click()
On Error Resume Next
FrmConsultaReserva.Show , Me

End Sub

Private Sub ResNovo_Click()
On Error Resume Next
FrmReserva.Show , Me
End Sub

Private Sub RelEntradEstPeriodo_Click()
On Error Resume Next
FrmRelEntradEstPeriodo.Show , Me
End Sub

Private Sub RelEntrEstoque_Click()
On Error Resume Next
FrmRelEntradaProduto.Show , Me

End Sub

Private Sub relnfprod_Click()
On Error Resume Next
relentradaProdutoNf.Show , Me

End Sub

Private Sub RelNotaSaida_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
  FrmRelSaidaProduto.Show , Me
End If
End Sub

Private Sub RelNumPedido_Click()
On Error Resume Next
FrmRelPedidoNumero.Show , Me
End Sub

Private Sub RelOrcamNumero_Click()
On Error Resume Next
FrmRelOrcNumero.Show , Me
End Sub

Private Sub RelOrcCli_Click()
On Error Resume Next
FrmRelorcamentoClien.Show , Me
End Sub

Private Sub RelOrcPeriodo_Click()
On Error Resume Next
FrmRelOrcamentoPeriodo.Show , Me
End Sub

Private Sub RelPedidoCliente_Click()
On Error Resume Next
FrmRelPedidoClien.Show , Me

End Sub

Private Sub RelPedidoFornec_Click()
On Error Resume Next
FrmRelPedidoFornec.Show , Me
End Sub

Private Sub RelPedidoPeriodo_Click()
FrmRelPedidoPeriodo.Show , Me

End Sub

Private Sub RelPeriodoSaida_Click()
On Error Resume Next
If UCase(GlUsuario) <> "TELE" Then
  FrmRelSaidaEstPeriodo.Show , Me
End If
End Sub

Private Sub RelProdSimples_Click()
On Error Resume Next
FrmRelProdutoNomeCod.Show , Me
End Sub

Private Sub RelProForn_Click()
On Error Resume Next
FrmRelProdutoFornec.Show , Me

End Sub

Private Sub RelProNome_Click()
On Error Resume Next
FrmRelProdutoNome.Show , Me
End Sub

Private Sub RelProTabela_Click()
On Error Resume Next
FrmRelProdutotabela.Show , Me
End Sub

Private Sub RelRecDocumento_Click()
On Error Resume Next
FrmRelReceitaNumero.Show , Me
End Sub

Private Sub RelReceiPeriodo_Click()
On Error Resume Next
FrmRelReceitaPeriodo.Show , Me
End Sub

Private Sub RelReceitaCliente_Click()
On Error Resume Next
FrmRelReceitaCliente.Show , Me

End Sub

Private Sub RelSaidaEstoque_Click()
'On Error Resume Next
'RelSaidaProdutos.Show , Me

End Sub



Private Sub RelTabMax_Click()
On Error Resume Next
FrmRelProdutotabelaMax.Show , Me

End Sub

Private Sub RlCheqCliente_Click()
On Error Resume Next
FrmRelChequeCliente.Show , Me

End Sub

Private Sub SAir_Click()
On Error Resume Next
End
End Sub

Private Sub SenUsu_Click()
On Error Resume Next
FrmUsuarios.Show , Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

'16777215
'LbSistema.ForeColor = CorLetra
AutoRedraw = True
FontName = "Arial"
'Stop
'fontess
FontSize = 22
Call Etch(Me, "SISTEMA DSA COMERCIAL  <<< VERSÃO " & App.Major & "." & App.Minor & "-" & App.Revision & " >>>", 240, 360, CorLetra)
CorLetra = Int((15 - 0 + 1) * Rnd + 0)
'StatusBar.Panels(6).Text = "Hora Atual: " & Format(Time(), "hh:mm:ss")
Screen.ActiveForm.DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")

Exit Sub

ErroCor:
CorLetra = CorInicio
Exit Sub

End Sub
Function fontess()
Static i    ' Declare variables.
    Dim OldFont
    FontName = Screen.Fonts(i)  ' Change to new font.
    Print Screen.Fonts(i)   ' Print name of font.
    i = i + 1   ' Increment counter.
    If i = FontCount Then
       i = 0 ' Start over.
       FontName = "Times New Roman"  ' Restore original font.
    End If
End Function
Private Sub UtilSegBack_Click()
On Error Resume Next
Dim LcResposta As Long
Dim LcCriterio As String

LcResposta = MsgBox("Insira o Disco no << DRIVE A >> " & Chr(13) & _
"Os Discos Para Backup devem Estar Vazios.", 65, "Backup")
If LcResposta = 2 Then
   Exit Sub
End If
LcCriterio = "Arj a -rva a:BacDSA.arj " & GLBase

Shell (LcCriterio), 6

End Sub

Private Sub UtilSegRestaura_Click()
On Error Resume Next
Dim LcResposta As Long
Dim LcCriterio As String

LcResposta = MsgBox("Insira o Primeiro Disco no << DRIVE A >> " & Chr(13) & _
"Os DADOS ATUAIS DO SEU SISTEMA SERÃO SUBSTITUIDOS PELOS DO BACKUP", 65, "Backup")
If LcResposta = 2 Then
   Exit Sub
End If
LcCriterio = "Arj x -rva a:BacDSA.arj " & GLBase

Shell (LcCriterio), 6

End Sub

Private Sub UtilSenhaGrupo_Click()
 On Error Resume Next
 FrmCadGrupo.Show , Me
End Sub
Function LocalizaFigura()
Dim LcFigura, Destination As String
Destination = App.Path & "\fundo.bmp"
With FrmPrincipal.Abrird
    .InitDir = App.Path
    .FileName = "*.Bmp;*.Gif;*.Jpg;*.Jpeg"
    .ShowOpen
End With
LcFigura = FrmPrincipal.Abrird.FileName
FileCopy LcFigura, Destination
Me.Picture = LoadPicture(App.Path & "\fundo.bmp")
End Function

Private Sub VerExclusao_Click()
Tabelasxcluidas.Show , Me
End Sub
