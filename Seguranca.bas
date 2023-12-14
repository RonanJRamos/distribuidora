Attribute VB_Name = "Seguranca"
Public GlGrupoSenha As String
Function VerificaUsuario(LcUsuario, LcSenha As String) As Integer
On Error Resume Next
Dim RsUser As Recordset
Dim LcCriterio As String

'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
'Set RsUser = Dbbase.OpenRecordset("Usuario", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCriterio = "select * from usuario where Nome='" & LcUsuario & "' and Senha='" & LcSenha & "'"
'RsUser.FindFirst LcCriterio
'abreconexao
Set RsUser = AbreRecordset(LcCriterio)
If RsUser.EOF Then
   MsgBox "Usuário ou Senha Inválido", 48, "Aviso"
   VerificaUsuario = False
Else
   GlGrupoSenha = RsUser!Grupo
   VerificaUsuario = True
   HabilitaMenus
End If
RsUser.Close
'Dbbase.Close
Set RsUser = Nothing
'Set Dbbase = Nothing

End Function

Function HabilitaMenus()
On Error GoTo ErroHab
Dim RsGrupo As ADODB.Recordset
Dim LcCriterio As String
LcCriterio = "Select * From GrpSenhas where Grupo='" & GlGrupo & "'"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
'Set RsGrupo = Dbbase.OpenRecordset(LcCriterio) ', dbOpenDynaset, dbSeeChanges, dbOptimistic)
'abreconexao
Set RsGrupo = AbreRecordset(LcCriterio)
Do Until RsGrupo.EOF
   Select Case RsGrupo!Sistema
   
          Case Is = "Vales"
              FrmPrincipal.MnVales.Enabled = RsGrupo!Incluir
          Case Is = "Excluir Vales"
              FrmPrincipal.MnCancelaVale.Enabled = RsGrupo!Incluir
          Case Is = "Ficha de Estoque"
              FrmPrincipal.MnFicha.Enabled = RsGrupo!Incluir
          Case Is = "Romaneio"
              FrmPrincipal.MnRomaneio.Enabled = RsGrupo!Incluir
          Case Is = "Gerar disquete Receita"
              FrmPrincipal.mndisquetereceita.Enabled = RsGrupo!Incluir
          Case Is = "Localizar banco de dados"
              FrmPrincipal.LcLocalizar.Enabled = RsGrupo!Incluir
          Case Is = "Ver dados Excluidos"
              FrmPrincipal.VerExclusao.Enabled = RsGrupo!Incluir
          Case Is = "Custo"
               FrmPrincipal.MnIncluirCusto.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnCustoAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnCustoConsultar.Enabled = RsGrupo!Consultar
               If Not FrmPrincipal.MnIncluirCusto.Enabled And _
                  Not FrmPrincipal.MnCustoAlterar.Enabled And _
                  Not FrmPrincipal.MnCustoConsultar.Enabled Then FrmPrincipal.MnCusto.Enabled = False
          Case Is = "Comissões Representada"
               FrmPrincipal.MnCORe.Enabled = RsGrupo!Incluir
               
               
          Case Is = "Pedido de Vendas"
               FrmPrincipal.MnProposta.Enabled = RsGrupo!Incluir
                      
          Case Is = "Pesquisa Compras de Cliente"
               FrmPrincipal.MnPesqComprasCli.Enabled = RsGrupo!Incluir
           Case Is = "Copia de Segurança"
               FrmPrincipal.MnBackup.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnRecuperar.Enabled = RsGrupo!Alterar
               If Not FrmPrincipal.MnBackup.Enabled And _
                  Not FrmPrincipal.MnRecuperar.Enabled _
                  Then FrmPrincipal.MnSegunca.Enabled = False

    
          Case Is = "Clientes"
               FrmPrincipal.CadCliIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.CadCliAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.CadCliConsulta.Enabled = RsGrupo!Consultar
               FrmPrincipal.relcli.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.CadCliIncluir.Enabled And _
                  Not FrmPrincipal.CadCliAlterar.Enabled And _
                  Not FrmPrincipal.CadCliConsulta.Enabled Then FrmPrincipal.CadCliente.Enabled = False
          Case Is = "Fornecedores"
               FrmPrincipal.CadForIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.CadForAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.CadForConsulta.Enabled = RsGrupo!Consultar
               FrmPrincipal.RelFornecedor.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.CadForIncluir.Enabled And _
                  Not FrmPrincipal.CadForAlterar.Enabled And _
                  Not FrmPrincipal.CadForConsulta.Enabled Then FrmPrincipal.CadFornecedor.Enabled = False
          
          Case Is = "Transportadora"
               FrmPrincipal.MnTraspincluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnTranspAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnTranspConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelProdutos.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnTraspincluir.Enabled And _
                  Not FrmPrincipal.MnTranspAlterar.Enabled And _
                  Not FrmPrincipal.MnTranspConsultar.Enabled Then FrmPrincipal.MnTransportadoras.Enabled = False

          Case Is = "Funcionarios"
               FrmPrincipal.MnFuncInc.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnFuncAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnFuncConsultar.Enabled = RsGrupo!Consultar
               FrmPrincipal.MnFuncionarios.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnFuncInc.Enabled And _
                  Not FrmPrincipal.MnFuncAlterar.Enabled And _
                  Not FrmPrincipal.MnFuncConsultar.Enabled Then FrmPrincipal.MnFunc.Enabled = False
           
           Case Is = "Cidades"
               FrmPrincipal.MnCidadeIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnCidadeAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnCidadeConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelProdutos.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnCidadeIncluir.Enabled And _
                  Not FrmPrincipal.MnCidadeAlterar.Enabled And _
                  Not FrmPrincipal.MnCidadeConsultar.Enabled Then FrmPrincipal.MnCidade.Enabled = False
 
          Case Is = "Produtos"
               FrmPrincipal.CadProIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.CadProAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.CadProConsulta.Enabled = RsGrupo!Consultar
               FrmPrincipal.RelProdutos.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.CadProIncluir.Enabled And _
                  Not FrmPrincipal.CadProAlterar.Enabled And _
                  Not FrmPrincipal.CadProConsulta.Enabled Then FrmPrincipal.CadProduto.Enabled = False

          Case Is = "Tipo Receitas e Despesas"
               FrmPrincipal.MnIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnTipoRecAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnTipoREcConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelFornecedor.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnIncluir.Enabled And _
                  Not FrmPrincipal.MnTipoRecAlterar.Enabled And _
                  Not FrmPrincipal.MnTipoREcConsultar.Enabled Then FrmPrincipal.MnTipoREc.Enabled = False

          Case Is = "Galpão"
               FrmPrincipal.MnGalpaoIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnGalpaoAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnGalpaoConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelDepartamento.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnGalpaoIncluir.Enabled And _
                  Not FrmPrincipal.MnGalpaoAlterar.Enabled And _
                  Not FrmPrincipal.MnGalpaoConsultar.Enabled Then FrmPrincipal.MnGalapao.Enabled = False

          Case Is = "Tipo Monetário"
               FrmPrincipal.MnTipoIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnTipoAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnTipoConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelCargo.Enabled = RsGrupo!Relatorio
                If Not FrmPrincipal.MnTipoIncluir.Enabled And _
                  Not FrmPrincipal.MnTipoAlterar.Enabled And _
                  Not FrmPrincipal.MnTipoConsultar.Enabled Then FrmPrincipal.MnTipo.Enabled = False

          Case Is = "Unidade"
               FrmPrincipal.MnUnidadeIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnUnidadeAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnUnidadeConsultar.Enabled = RsGrupo!Consultar
               'FrmPrincipal.RelAtas.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnUnidadeIncluir.Enabled And _
                  Not FrmPrincipal.MnUnidadeAlterar.Enabled And _
                  Not FrmPrincipal.MnUnidadeConsultar.Enabled Then FrmPrincipal.MnUnidade.Enabled = False

          Case Is = "Receitas"
               FrmPrincipal.MnFinaRecIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnFinaRecAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnFinaRecConsultar.Enabled = RsGrupo!Consultar
               FrmPrincipal.MnFinarecBaixa.Enabled = RsGrupo!Baixa
               FrmPrincipal.MnRelReceita.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnFinaRecIncluir.Enabled And _
                  Not FrmPrincipal.MnFinaRecAlterar.Enabled And _
                  Not FrmPrincipal.MnFinaRecConsultar.Enabled And _
                  Not FrmPrincipal.MnFinarecBaixa.Enabled Then FrmPrincipal.MnFinaReceitas.Enabled = False

          Case Is = "Despesas"
               FrmPrincipal.MnFinaDespIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnFinaAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnFinaDespConsultar.Enabled = RsGrupo!Consultar
               FrmPrincipal.MnFinaDespBaixa.Enabled = RsGrupo!Baixa
               FrmPrincipal.RelDesp.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnFinaDespIncluir.Enabled And _
                  Not FrmPrincipal.MnFinaAlterar.Enabled And _
                  Not FrmPrincipal.MnFinaDespConsultar.Enabled And _
                  Not FrmPrincipal.MnFinaDespBaixa.Enabled Then FrmPrincipal.MnFinaDesp.Enabled = False

          Case Is = "Cheques"
               FrmPrincipal.MnChequesIncluir.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnChAlterar.Enabled = RsGrupo!Alterar
               FrmPrincipal.MnChConsultar.Enabled = RsGrupo!Consultar
               FrmPrincipal.RelCheq.Enabled = RsGrupo!Relatorio
               If Not FrmPrincipal.MnChequesIncluir.Enabled And _
                  Not FrmPrincipal.MnChAlterar.Enabled And _
                  Not FrmPrincipal.MnChConsultar.Enabled Then FrmPrincipal.MnCheques.Enabled = False

          Case Is = "Entrada de produto"
               FrmPrincipal.MnEntrada.Enabled = RsGrupo!Incluir
               FrmPrincipal.RelEntradaEstoque.Enabled = RsGrupo!Alterar
          
          Case Is = "Caixa"
               FrmPrincipal.MnCaixa.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnRelCaixa.Enabled = RsGrupo!Alterar

          Case Is = "Orçamento e Vendas"
               FrmPrincipal.MnOrcamento.Enabled = RsGrupo!Incluir
               FrmPrincipal.RelOrcamento.Enabled = RsGrupo!Alterar
          Case Is = "Alteração de Preço"
               FrmPrincipal.MnAlteraPreco.Enabled = RsGrupo!Incluir
               
          Case Is = "Pedido de Cliente"
               FrmPrincipal.MnOrcamento.Enabled = RsGrupo!Incluir
               FrmPrincipal.RelOrcamento.Enabled = RsGrupo!Alterar
         
         Case Is = "Cancelar Pedidos"
               FrmPrincipal.mncancelaOrcamento.Enabled = RsGrupo!Incluir
               'FrmPrincipal.RelOrcamento.Enabled = RsGrupo!Alterar
               
         Case Is = "Cancelar Notas"
               FrmPrincipal.MnCancelar.Enabled = RsGrupo!Incluir
              ' FrmPrincipal.RelOrcamento.Enabled = RsGrupo!Alterar
          'Case Is = "Segurança"
          '     FrmPrincipal.UtilSegBack.Enabled = RsGrupo!Incluir
           '    FrmPrincipal.UtilSegRestaura.Enabled = RsGrupo!Alterar
          Case Is = "Notas de Saídas"
               FrmPrincipal.MnSaida.Enabled = RsGrupo!Incluir
               FrmPrincipal.RelSaidaEstoque.Enabled = RsGrupo!Alterar
          Case Is = "Senha"
               FrmPrincipal.MnSenhaGrupo.Enabled = RsGrupo!Incluir
               FrmPrincipal.MnSenharUser.Enabled = RsGrupo!Alterar
               If Not FrmPrincipal.MnSenhaGrupo.Enabled And _
                  Not FrmPrincipal.MnSenharUser.Enabled Then FrmPrincipal.Senha.Enabled = False

          Case Is = "Pano de Fundo"
               FrmPrincipal.PanoFundo.Enabled = RsGrupo!Incluir
           Case Is = "Reparar banco de Dados"
               FrmPrincipal.mReparar.Enabled = RsGrupo!Incluir
           Case Is = "Opções"
                FrmPrincipal.MnOpcoes.Enabled = RsGrupo!Incluir
           Case Is = "Comissões"
                
                FrmPrincipal.MnComissao.Enabled = RsGrupo!Incluir
                FrmPrincipal.MnComissoes.Enabled = RsGrupo!Incluir
           Case Is = "Alterar Data"
                FrmPrincipal.AlteraData.Enabled = RsGrupo!Incluir
           Case Is = "Configura Mala Direta"
               FrmPrincipal.EtMalaDireta.Enabled = RsGrupo!Incluir
  End Select

  RsGrupo.MoveNext
Loop
RsGrupo.Close
Set RsGrupo = Nothing
Exit Function
ErroHab:
MsgBox err.Description & err.Number
'Resume 0
End Function
Function hasbilitatodos()
         On Error Resume Next
               FrmPrincipal.LcLocalizar.Enabled = True
               FrmPrincipal.MnVales.Enabled = True
               FrmPrincipal.MnCancelaVale.Enabled = True
               FrmPrincipal.MnFicha.Enabled = True
               FrmPrincipal.MnRomaneio.Enabled = True
               FrmPrincipal.mndisquetereceita.Enabled = True
               FrmPrincipal.MnSegunca.Enabled = True
               FrmPrincipal.VerExclusao.Enabled = True
               FrmPrincipal.MnCaixa.Enabled = True
               FrmPrincipal.MnRelCaixa.Enabled = True
               FrmPrincipal.SistemaPr.Enabled = True
               FrmPrincipal.Cadastro.Enabled = True
               FrmPrincipal.MnFinanceiro.Enabled = True
               FrmPrincipal.MnEstoque.Enabled = True
               FrmPrincipal.Rel.Enabled = True
               FrmPrincipal.Utilitarios.Enabled = True
               FrmPrincipal.CadCliente.Enabled = True
               FrmPrincipal.CadFornecedor.Enabled = True
               FrmPrincipal.CadProduto.Enabled = True
               FrmPrincipal.MnFunc.Enabled = True
               FrmPrincipal.MnGalapao.Enabled = True
               FrmPrincipal.MnCidade.Enabled = True
               FrmPrincipal.MnTipo.Enabled = True
               FrmPrincipal.MnTipoREc.Enabled = True
               FrmPrincipal.MnUnidade.Enabled = True
               FrmPrincipal.MnFinaReceitas.Enabled = True
               FrmPrincipal.MnFinaDesp.Enabled = True
               FrmPrincipal.MnCheques.Enabled = True
               FrmPrincipal.Senha.Enabled = True
               FrmPrincipal.MnTransportadoras.Enabled = True
               FrmPrincipal.MnTraspincluir.Enabled = True
               FrmPrincipal.MnTranspAlterar.Enabled = True
               FrmPrincipal.MnTranspConsultar.Enabled = True
               FrmPrincipal.MnTransportadoras.Enabled = True

               FrmPrincipal.CadCliIncluir.Enabled = True
               FrmPrincipal.CadCliAlterar.Enabled = True
               FrmPrincipal.CadCliConsulta.Enabled = True
               FrmPrincipal.relcli.Enabled = True
               FrmPrincipal.CadForIncluir.Enabled = True
               FrmPrincipal.CadForAlterar.Enabled = True
               FrmPrincipal.CadForConsulta.Enabled = True
               FrmPrincipal.RelFornecedor.Enabled = True
               FrmPrincipal.MnFuncInc.Enabled = True
               FrmPrincipal.MnFuncAlterar.Enabled = True
               FrmPrincipal.MnFuncConsultar.Enabled = True
               FrmPrincipal.CadProIncluir.Enabled = True
               FrmPrincipal.CadProAlterar.Enabled = True
               FrmPrincipal.CadProConsulta.Enabled = True
               FrmPrincipal.RelProdutos.Enabled = True
               FrmPrincipal.MnIncluir.Enabled = True
               FrmPrincipal.MnTipoRecAlterar.Enabled = True
               FrmPrincipal.MnTipoREcConsultar.Enabled = True
               FrmPrincipal.MnGalpaoIncluir.Enabled = True
               FrmPrincipal.MnGalpaoAlterar.Enabled = True
               FrmPrincipal.MnGalpaoConsultar.Enabled = True
               FrmPrincipal.MnTipoIncluir.Enabled = True
               FrmPrincipal.MnTipoAlterar.Enabled = True
               FrmPrincipal.MnTipoConsultar.Enabled = True
               FrmPrincipal.MnUnidadeIncluir.Enabled = True
               FrmPrincipal.MnUnidadeAlterar.Enabled = True
               FrmPrincipal.MnUnidadeConsultar.Enabled = True
               FrmPrincipal.MnFinaRecIncluir.Enabled = True
               FrmPrincipal.MnFinaRecAlterar.Enabled = True
               FrmPrincipal.MnFinaRecConsultar.Enabled = True
               FrmPrincipal.MnFinarecBaixa.Enabled = True
               FrmPrincipal.MnRelReceita.Enabled = True
               FrmPrincipal.MnFinaDespIncluir.Enabled = True
               FrmPrincipal.MnFinaAlterar.Enabled = True
               FrmPrincipal.MnFinaDespConsultar.Enabled = True
               FrmPrincipal.MnFinaDespBaixa.Enabled = True
               FrmPrincipal.RelDesp.Enabled = True
               FrmPrincipal.MnChequesIncluir.Enabled = True
               FrmPrincipal.MnChAlterar.Enabled = True
               FrmPrincipal.MnChConsultar.Enabled = True
               FrmPrincipal.RelCheq.Enabled = True
               FrmPrincipal.MnEntrada.Enabled = True
               FrmPrincipal.RelEntradaEstoque.Enabled = True
               FrmPrincipal.MnPedido.Enabled = True
               FrmPrincipal.RelPedido.Enabled = True
               FrmPrincipal.MnAlteraPreco.Enabled = True
               FrmPrincipal.MnOrcamento.Enabled = True
               FrmPrincipal.RelOrcamento.Enabled = True
               FrmPrincipal.MnFuncionarios.Enabled = True
               FrmPrincipal.EtMalaDireta.Enabled = True
               FrmPrincipal.MnSenhaGrupo.Enabled = True
               FrmPrincipal.MnSenharUser.Enabled = True
               FrmPrincipal.PanoFundo.Enabled = True
               FrmPrincipal.mReparar.Enabled = True
               FrmPrincipal.MnOpcoes.Enabled = True
               FrmPrincipal.MnComissao.Enabled = True
               FrmPrincipal.MnComissoes.Enabled = True
               FrmPrincipal.MnSaida.Enabled = True
               FrmPrincipal.RelSaidaEstoque.Enabled = True
               FrmPrincipal.MnCidadeIncluir.Enabled = True
               FrmPrincipal.MnCidadeAlterar.Enabled = True
               FrmPrincipal.MnCidadeConsultar.Enabled = True
               FrmPrincipal.AlteraData.Enabled = True
               FrmPrincipal.mncancelaOrcamento.Enabled = True
               FrmPrincipal.MnCancelar.Enabled = True
               FrmPrincipal.MnCusto.Enabled = True
               FrmPrincipal.MnIncluirCusto.Enabled = True
               FrmPrincipal.MnCustoAlterar.Enabled = True
               FrmPrincipal.MnCustoConsultar.Enabled = True
               FrmPrincipal.MnCORe.Enabled = True
               FrmPrincipal.MnProposta.Enabled = True
               FrmPrincipal.MnPesqComprasCli.Enabled = True
               FrmPrincipal.MnSegunca.Enabled = True
               FrmPrincipal.MnBackup.Enabled = True
               FrmPrincipal.MnRecuperar.Enabled = True
End Function
Function Desabilitatodos()
         On Error Resume Next
               'FrmPrincipal.SistemaPr.Enabled = False
               FrmPrincipal.MnVales.Enabled = False
               FrmPrincipal.MnCancelaVale.Enabled = False
               FrmPrincipal.MnFicha.Enabled = False
               FrmPrincipal.MnRomaneio.Enabled = False
               FrmPrincipal.mndisquetereceita.Enabled = False
               FrmPrincipal.MnSegunca.Enabled = False
               FrmPrincipal.VerExclusao.Enabled = False
               FrmPrincipal.MnComissoes.Enabled = True
               FrmPrincipal.LcLocalizar.Enabled = True
               FrmPrincipal.Cadastro.Enabled = False
               FrmPrincipal.MnFinanceiro.Enabled = False
               FrmPrincipal.MnEstoque.Enabled = False
               FrmPrincipal.Rel.Enabled = False
               FrmPrincipal.Utilitarios.Enabled = False
               FrmPrincipal.MnCaixa.Enabled = False
               FrmPrincipal.MnRelCaixa.Enabled = False
               FrmPrincipal.Cadastro.Enabled = False
               FrmPrincipal.MnFinanceiro.Enabled = False
               FrmPrincipal.MnEstoque.Enabled = False
               FrmPrincipal.Rel.Enabled = False
               FrmPrincipal.Utilitarios.Enabled = False
               FrmPrincipal.CadCliente.Enabled = False
               FrmPrincipal.CadFornecedor.Enabled = False
               FrmPrincipal.CadProduto.Enabled = False
               FrmPrincipal.MnFunc.Enabled = False
               FrmPrincipal.MnGalapao.Enabled = False
               FrmPrincipal.MnCidade.Enabled = False
               FrmPrincipal.MnTipo.Enabled = False
               FrmPrincipal.MnTipoREc.Enabled = False
               FrmPrincipal.MnUnidade.Enabled = False
               FrmPrincipal.MnFinaReceitas.Enabled = False
               FrmPrincipal.MnFinaDesp.Enabled = False
               FrmPrincipal.MnCheques.Enabled = False
               FrmPrincipal.Senha.Enabled = False
               FrmPrincipal.MnTransportadoras.Enabled = False
               FrmPrincipal.MnTraspincluir.Enabled = False
               FrmPrincipal.MnTranspAlterar.Enabled = False
               FrmPrincipal.MnTranspConsultar.Enabled = False
               FrmPrincipal.MnTransportadoras.Enabled = False

               FrmPrincipal.CadCliIncluir.Enabled = False
               FrmPrincipal.CadCliAlterar.Enabled = False
               FrmPrincipal.CadCliConsulta.Enabled = False
               FrmPrincipal.relcli.Enabled = False
               FrmPrincipal.CadForIncluir.Enabled = False
               FrmPrincipal.CadForAlterar.Enabled = False
               FrmPrincipal.CadForConsulta.Enabled = False
               FrmPrincipal.RelFornecedor.Enabled = False
               FrmPrincipal.MnFuncInc.Enabled = False
               FrmPrincipal.MnFuncAlterar.Enabled = False
               FrmPrincipal.MnFuncConsultar.Enabled = False
               FrmPrincipal.CadProIncluir.Enabled = False
               FrmPrincipal.CadProAlterar.Enabled = False
               FrmPrincipal.CadProConsulta.Enabled = False
               FrmPrincipal.RelProdutos.Enabled = False
               FrmPrincipal.MnIncluir.Enabled = False
               FrmPrincipal.MnTipoRecAlterar.Enabled = False
               FrmPrincipal.MnTipoREcConsultar.Enabled = False
               FrmPrincipal.MnGalpaoIncluir.Enabled = False
               FrmPrincipal.MnGalpaoAlterar.Enabled = False
               FrmPrincipal.MnGalpaoConsultar.Enabled = False
               FrmPrincipal.MnTipoIncluir.Enabled = False
               FrmPrincipal.MnTipoAlterar.Enabled = False
               FrmPrincipal.MnTipoConsultar.Enabled = False
               FrmPrincipal.MnUnidadeIncluir.Enabled = False
               FrmPrincipal.MnUnidadeAlterar.Enabled = False
               FrmPrincipal.MnUnidadeConsultar.Enabled = False
               FrmPrincipal.MnFinaRecIncluir.Enabled = False
               FrmPrincipal.MnFinaRecAlterar.Enabled = False
               FrmPrincipal.MnFinaRecConsultar.Enabled = False
               FrmPrincipal.MnFinarecBaixa.Enabled = False
               FrmPrincipal.MnRelReceita.Enabled = False
               FrmPrincipal.MnFinaDespIncluir.Enabled = False
               FrmPrincipal.MnFinaAlterar.Enabled = False
               FrmPrincipal.MnFinaDespConsultar.Enabled = False
               FrmPrincipal.MnFinaDespBaixa.Enabled = False
               FrmPrincipal.RelDesp.Enabled = False
               FrmPrincipal.MnChequesIncluir.Enabled = False
               FrmPrincipal.MnChAlterar.Enabled = False
               FrmPrincipal.MnChConsultar.Enabled = False
               FrmPrincipal.RelCheq.Enabled = False
               FrmPrincipal.MnEntrada.Enabled = False
               FrmPrincipal.RelEntradaEstoque.Enabled = False
               FrmPrincipal.MnPedido.Enabled = False
               FrmPrincipal.RelPedido.Enabled = False
               FrmPrincipal.MnAlteraPreco.Enabled = False
               FrmPrincipal.MnOrcamento.Enabled = False
               FrmPrincipal.RelOrcamento.Enabled = False
               FrmPrincipal.MnFuncionarios.Enabled = False
               FrmPrincipal.EtMalaDireta.Enabled = False
               FrmPrincipal.MnSenhaGrupo.Enabled = False
               FrmPrincipal.MnSenharUser.Enabled = False
               FrmPrincipal.PanoFundo.Enabled = False
               FrmPrincipal.mReparar.Enabled = False
               FrmPrincipal.MnOpcoes.Enabled = False
               FrmPrincipal.MnComissao.Enabled = False
               FrmPrincipal.MnSaida.Enabled = False
               FrmPrincipal.RelSaidaEstoque.Enabled = False
               FrmPrincipal.MnCidadeIncluir.Enabled = False
               FrmPrincipal.MnCidadeAlterar.Enabled = False
               FrmPrincipal.MnCidadeConsultar.Enabled = False
               FrmPrincipal.AlteraData.Enabled = False
               FrmPrincipal.mncancelaOrcamento.Enabled = False
               FrmPrincipal.MnCancelar.Enabled = False
               FrmPrincipal.MnCusto.Enabled = False
               FrmPrincipal.MnIncluirCusto.Enabled = False
               FrmPrincipal.MnCustoAlterar.Enabled = False
               FrmPrincipal.MnCustoConsultar.Enabled = False
               FrmPrincipal.MnCORe.Enabled = False
               FrmPrincipal.MnProposta.Enabled = False
               FrmPrincipal.MnPesqComprasCli.Enabled = False
               FrmPrincipal.MnSegunca.Enabled = False
               FrmPrincipal.MnBackup.Enabled = False
               FrmPrincipal.MnRecuperar.Enabled = False

End Function
