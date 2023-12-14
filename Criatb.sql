# Connection: Root
# Host: servidor
# Saved: 2007-01-25 07:48:15
# 
# Host: servidor
# Database: lidis
# Table: 'alid015'
# 
CREATE TABLE `alid015` (
  `NUMLANCTO` varchar(15) default NULL,
  `NF` varchar(25) default NULL,
  `CLIENTE` varchar(20) default NULL,
  `TPMONET` varchar(10) default NULL,
  `VALOR` decimal(24,4) default NULL,
  `DATA` datetime default NULL,
  `DTVENC` datetime default NULL,
  `CONTR` varchar(10) default NULL,
  `DTPAGTO` datetime default NULL,
  `VALPAGO` decimal(24,4) default NULL,
  `TIPORD` varchar(10) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `Acrescimo` decimal(17,2) default '0.00',
  `emitente` varchar(50) default NULL,
  `Obs` longtext,
  `codDesp` varchar(10) default NULL,
  `NomeDesp` varchar(50) default NULL,
  PRIMARY KEY  (`codigo`),
  KEY `CLIENTE` (`CLIENTE`),
  KEY `codigo` (`NF`),
  KEY `DATA` (`DATA`),
  KEY `DTPAGTO` (`DTPAGTO`),
  KEY `DTVENC` (`DTVENC`),
  KEY `nf` (`NF`),
  KEY `NUMLANCTO` (`NUMLANCTO`),
  KEY `TPMONET` (`TPMONET`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'alid050'
# 
CREATE TABLE `alid050` (
  `NUMNF` varchar(20) default NULL,
  `DTEMIS` datetime default NULL,
  `NATUREZA` varchar(15) default NULL,
  `CLIENTE` varchar(20) default NULL,
  `TRANSP` varchar(30) default NULL,
  `TIPOTRANS` varchar(5) default NULL,
  `PLACATRANS` varchar(8) default NULL,
  `UFTRANS` varchar(5) default NULL,
  `CGCCPFTRAN` varchar(20) default NULL,
  `ENDTRANS` varchar(30) default NULL,
  `MUNICTRANS` varchar(20) default NULL,
  `UFMUNIC` varchar(5) default NULL,
  `INSCEST` varchar(20) default NULL,
  `OBS02` varchar(60) default NULL,
  `OBS03` varchar(60) default NULL,
  `OBS04` varchar(60) default NULL,
  `CONTR` varchar(7) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `valorproduto` decimal(17,2) default '0.00',
  `ValorNota` decimal(17,2) default '0.00',
  `Vendedor` varchar(10) default NULL,
  `fonetransp` varchar(50) default NULL,
  `cidade` varchar(40) default NULL,
  `cep` varchar(50) default NULL,
  `formapag` varchar(50) default NULL,
  `dias` varchar(50) default NULL,
  `vencimento1` datetime default NULL,
  `vencimento2` datetime default NULL,
  `vencimento3` datetime default NULL,
  `vencimento4` datetime default NULL,
  `vencimento5` datetime default NULL,
  `CondPag` varchar(50) default NULL,
  `status` varchar(50) default NULL,
  `DESCONTO` decimal(17,2) default '0.00',
  `ICMS` varchar(50) default NULL,
  `CFOP` varchar(50) default NULL,
  `BaseIcms` double(12,2) default '0.00',
  `ValorIcms` double(12,2) default '0.00',
  PRIMARY KEY  (`codigo`),
  KEY `CLIENTE` (`CLIENTE`),
  KEY `codigo` (`codigo`),
  KEY `DTEMIS` (`DTEMIS`),
  KEY `NUMNF` (`NUMNF`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'alid052'
# 
CREATE TABLE `alid052` (
  `NUMNF` varchar(20) default NULL,
  `ITEM` varchar(10) default NULL,
  `GALPAO` varchar(10) default NULL,
  `QTDE` decimal(28,4) default NULL,
  `VALUNIT` decimal(28,4) default NULL,
  `CONTR` varchar(10) default NULL,
  `UNIMED` varchar(150) default NULL,
  `QTDUM` decimal(28,4) default NULL,
  `QTDE01` decimal(28,4) default NULL,
  `QTDE02` decimal(28,4) default NULL,
  `QTDE03` decimal(28,4) default NULL,
  `codigo` int(30) NOT NULL auto_increment,
  `descricao` longtext,
  `codProd` int(50) default NULL,
  `emissao` date default NULL,
  `CST` varchar(4) default NULL,
  `icms` double(8,2) default '0.00',
  PRIMARY KEY  (`codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'entradanf'
# 
CREATE TABLE `entradanf` (
  `NUMLANCTO` varchar(7) default NULL,
  `NF` varchar(10) default NULL,
  `RECDESP` char(1) default NULL,
  `CLICRED` varchar(20) default NULL,
  `VALOR` decimal(20,2) default NULL,
  `DATA` datetime default NULL,
  `CONTR` varchar(7) default NULL,
  `VP` char(1) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `ValorProduto` decimal(20,2) default NULL,
  `BaseIcms` decimal(20,2) default NULL,
  `Icms` decimal(20,2) default NULL,
  `Ipi` decimal(20,2) default NULL,
  `Geral` double(18,2) default NULL,
  `Complementar` double(12,2) default '0.00',
  `CFOP` varchar(10) default NULL,
  `Serie` varchar(4) default NULL,
  `Sintegra` int(1) unsigned default '0',
  `BaseIcmsSubst` double(12,2) default '0.00',
  `IcmsSubst` double(12,2) default '0.00',
  `Frete` double(14,2) default '0.00',
  `Seguro` double(12,2) default '0.00',
  `PIS_COFINS` double(12,2) default '0.00',
  `NaoTributado` double(12,2) default '0.00',
  `DespesasAcessorias` double(12,2) default '0.00',
  `SubSerie` char(2) default NULL,
  `Maquina` varchar(250) default NULL,
  `TipoFrete` int(2) unsigned default NULL,
  `Emissao` datetime default NULL,
  `Desconto` double(12,2) default '0.00',
  PRIMARY KEY  (`codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'estoquefiscal'
# 
CREATE TABLE `estoquefiscal` (
  `codigo` int(10) unsigned NOT NULL auto_increment,
  `data` datetime default NULL,
  `CodigoProduto` int(11) unsigned default NULL,
  `Unidade` varchar(50) default NULL,
  `quantidadeEntradageral` double(12,2) default '0.00',
  `valorcustomediounitario` double(12,2) default NULL,
  `vcustototal` double(12,2) default NULL,
  `saldoGeral` double(14,2) default NULL,
  `QuantidadeSaidaGeral` double(12,2) default '0.00',
  `QTEntradaSanta` double(12,2) default '0.00',
  `QTEntradaSanta1` double(12,2) default '0.00',
  `QTEntradaCalifornia` double(12,2) default '0.00',
  `QTSaidaSanta` double(12,2) default '0.00',
  `QTSaidaSanta1` double(12,2) default '0.00',
  `QTSaidaCalifornia` double(12,2) default '0.00',
  `Saldosanta` double(12,2) default '0.00',
  `SaldoSanta1` double(12,2) default '0.00',
  `SaldoCalifornia` double(12,2) default '0.00',
  `Nome` varchar(50) default NULL,
  `SaldoAnteriorSanta` double(12,2) default '0.00',
  `SaldoAnteriorSanta1` double(12,2) default '0.00',
  `SaldoAnteriorCalifornia` double(12,2) default '0.00',
  `Posicao` int(9) unsigned default '0',
  KEY `codigo` (`codigo`),
  KEY `Data` (`data`),
  KEY `Nome` (`Unidade`)
) TYPE=MyISAM; 

# Host: servidor
# Database: lidis
# Table: 'grpsenhas'
# 
CREATE TABLE `grpsenhas` (
  `Codigo` int(10) NOT NULL auto_increment,
  `Grupo` varchar(15) default NULL,
  `Sistema` varchar(50) default NULL,
  `Incluir` tinyint(1) default NULL,
  `Alterar` tinyint(1) default NULL,
  `Consultar` tinyint(1) default NULL,
  `Baixa` tinyint(1) default NULL,
  `Relatorio` tinyint(1) default NULL,
  PRIMARY KEY  (`Codigo`),
  UNIQUE KEY `Codigo` (`Codigo`),
  KEY `Nome` (`Grupo`),
  KEY `Sistema` (`Sistema`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'historicoproduto'
# 
CREATE TABLE `historicoproduto` (
  `Produto` varchar(20) default NULL,
  `codigo` int(10) unsigned NOT NULL auto_increment,
  `descricao` text,
  `santa` float default NULL,
  `santa2` float default NULL,
  `california` float default NULL,
  `nf` varchar(50) default NULL,
  `data` datetime default NULL,
  `tipo` varchar(50) default NULL,
  `unidade` varchar(50) default NULL,
  `codunid` varchar(50) default NULL,
  `unisanta` float default NULL,
  `unsanta1` float default NULL,
  `Uncalifornia` float default NULL,
  `clienteforn` varchar(250) default NULL,
  PRIMARY KEY  (`codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'itensentradanf'
# 
CREATE TABLE `itensentradanf` (
  `NUMNF` varchar(6) default NULL,
  `ITEM` varchar(20) default NULL,
  `QTDE` float(31,30) default NULL,
  `VALUNIT` decimal(22,2) default NULL,
  `CONTR` varchar(7) default NULL,
  `UNIMED` char(2) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `descricao` varchar(100) default NULL,
  `ValorTotal` decimal(22,2) default NULL,
  `Icms` decimal(22,2) default NULL,
  `Ipi` decimal(22,2) default NULL,
  `QTDUM` varchar(50) default NULL,
  `fornecedor` varchar(150) default NULL,
  `data` date default NULL,
  `OrdemItem` int(4) unsigned default '0',
  `cfop` varchar(50) default NULL,
  `modelo` varchar(50) default NULL,
  `serie` varchar(50) default NULL,
  `CodigoNota` int(10) unsigned default '0',
  PRIMARY KEY  (`codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'numeronota'
# 
CREATE TABLE `numeronota` (
  `numeronota` varchar(50) default NULL,
  `numerosd` varchar(50) default NULL
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'precoatuali'
# 
CREATE TABLE `precoatuali` (
  `CODIGO` varchar(50) default NULL,
  `PRECO` float default NULL,
  `DATAATUALIZACAO` datetime default NULL,
  `PERCENTUAL` float default NULL,
  `VALORACRESCIMO` float default NULL,
  `lucro` float default NULL,
  `miminovenda` float default NULL
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'produtos'
# 
CREATE TABLE `produtos` (
  `CodUsuario` varchar(20) default NULL,
  `NOME` varchar(50) default NULL,
  `CODBAR` varchar(18) default NULL,
  `Custo` decimal(18,2) default NULL,
  `Preco` decimal(18,2) default NULL,
  `MinimoVenda` decimal(18,2) default NULL,
  `MinimoEst` decimal(18,2) default NULL,
  `UnidMedida` char(2) default NULL,
  `QtdMedida` decimal(18,2) default NULL,
  `CST` char(3) default NULL,
  `lucro` decimal(18,2) default NULL,
  `ComissaoFornecedor` decimal(18,2) default NULL,
  `Fornecedor` varchar(50) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `QuantEstoque` decimal(18,2) default NULL,
  `ipi` decimal(18,2) default NULL,
  `percentualcusto` varchar(50) default NULL,
  `maximoEstoque` decimal(18,2) default NULL,
  `custoTotal` decimal(18,2) default NULL,
  `subitens` enum('True','False') default NULL,
  `multiplositens` enum('True','False') default NULL,
  `Santa1` float default NULL,
  `santa2` float default NULL,
  `California` float default NULL,
  `Icms` double default NULL,
  PRIMARY KEY  (`codigo`),
  KEY `Nome` (`NOME`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'sintegra'
# 
CREATE TABLE `sintegra` (
  `Codigo` int(6) NOT NULL auto_increment,
  `Data` date default NULL,
  `Nf` varchar(6) default NULL,
  `Cfop` varchar(4) default NULL,
  `Valor` double(12,2) default NULL,
  `Cliente_Forn` varchar(6) default '0',
  `Origem` char(1) default NULL,
  PRIMARY KEY  (`Codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'sintegra_50'
# 
CREATE TABLE `sintegra_50` (
  `Codigo` int(6) NOT NULL auto_increment,
  `CNPJ` varchar(14) default NULL,
  `Inscricao` varchar(14) default NULL,
  `Data` date default NULL,
  `UF` char(2) default NULL,
  `Modelo` char(2) default NULL,
  `Serie` char(3) default NULL,
  `NF` varchar(6) default NULL,
  `CFOP` varchar(4) default NULL,
  `Emitente` char(1) default NULL,
  `ValorTotal` double(12,2) default '0.00',
  `Base_Calculo_Icms` double(12,2) default '0.00',
  `Valor_Icms` double(12,2) default '0.00',
  `Isenta` double(12,2) default '0.00',
  `Outra` double(12,2) default '0.00',
  `Aliquota` double(4,2) default NULL,
  `Situacao` char(1) default NULL,
  PRIMARY KEY  (`Codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'sintegra_54'
# 
CREATE TABLE `sintegra_54` (
  `Codigo` int(6) NOT NULL auto_increment,
  `CNPJ` varchar(14) default NULL,
  `Modelo` char(2) default NULL,
  `Serie` char(3) default NULL,
  `NF` varchar(6) default NULL,
  `CFOP` varchar(4) default NULL,
  `CST` char(3) default NULL,
  `Item` char(3) default NULL,
  `CodProduto` varchar(5) default NULL,
  `Quantidade` double(12,3) default NULL,
  `Valor_Total_Bruto` double(12,2) default NULL,
  `Valor_Desconto` double(12,2) default NULL,
  `Base_Calculo` double(12,2) default NULL,
  `Base_Calculo_Subst` double(12,2) default NULL,
  `IPI` double(12,2) default NULL,
  `Aliquota_Icms` double(12,2) default NULL,
  `Data` date default NULL,
  PRIMARY KEY  (`Codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'usuario'
# 
CREATE TABLE `usuario` (
  `codigo` int(10) NOT NULL auto_increment,
  `Nome` varchar(50) default NULL,
  `Grupo` varchar(50) default NULL,
  `Senha` varchar(50) default NULL,
  `Expira` datetime default NULL,
  PRIMARY KEY  (`codigo`),
  UNIQUE KEY `codigo` (`codigo`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'vales'
# 
CREATE TABLE `vales` (
  `NUMNF` varchar(6) default NULL,
  `DTEMIS` datetime default NULL,
  `NATUREZA` char(2) default NULL,
  `CLIENTE` varchar(20) default NULL,
  `TRANSP` varchar(30) default NULL,
  `TIPOTRANS` char(1) default NULL,
  `PLACATRANS` varchar(8) default NULL,
  `UFTRANS` char(2) default NULL,
  `CGCCPFTRAN` varchar(20) default NULL,
  `ENDTRANS` varchar(30) default NULL,
  `MUNICTRANS` varchar(20) default NULL,
  `UFMUNIC` char(2) default NULL,
  `INSCEST` varchar(20) default NULL,
  `OBS02` varchar(60) default NULL,
  `OBS03` varchar(60) default NULL,
  `OBS04` varchar(60) default NULL,
  `CONTR` varchar(7) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `valorproduto` decimal(17,2) default '0.00',
  `ValorNota` decimal(17,2) default '0.00',
  `Vendedor` varchar(5) default NULL,
  `fonetransp` varchar(50) default NULL,
  `cidade` varchar(40) default NULL,
  `cep` varchar(50) default NULL,
  `formapag` varchar(50) default NULL,
  `dias` varchar(50) default NULL,
  `vencimento1` datetime default NULL,
  `vencimento2` datetime default NULL,
  `vencimento3` datetime default NULL,
  `vencimento4` datetime default NULL,
  `vencimento5` datetime default NULL,
  `CondPag` varchar(50) default NULL,
  `status` varchar(50) default NULL,
  `DESCONTO` decimal(17,2) default '0.00',
  `ICMS` varchar(50) default NULL,
  `CFOP` varchar(50) default NULL,
  `baixado` int(50) default NULL,
  `marca` varchar(50) default NULL,
  PRIMARY KEY  (`codigo`),
  KEY `CLIENTE` (`CLIENTE`),
  KEY `codigo` (`codigo`),
  KEY `DTEMIS` (`DTEMIS`),
  KEY `NUMNF` (`NUMNF`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'valesprodutos'
# 
CREATE TABLE `valesprodutos` (
  `NUMNF` varchar(6) default NULL,
  `ITEM` varchar(5) default NULL,
  `GALPAO` char(2) default NULL,
  `QTDE` decimal(22,4) default NULL,
  `VALUNIT` decimal(22,4) default NULL,
  `CONTR` varchar(7) default NULL,
  `UNIMED` char(2) default NULL,
  `QTDUM` decimal(22,4) default NULL,
  `QTDE01` decimal(22,4) default NULL,
  `QTDE02` decimal(22,4) default NULL,
  `QTDE03` decimal(22,4) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `descricao` longtext,
  `codProd` varchar(20) default NULL,
  PRIMARY KEY  (`codigo`),
  KEY `codigo` (`codigo`),
  KEY `ITEM` (`ITEM`),
  KEY `NUMNF` (`NUMNF`),
  KEY `UNIMED` (`UNIMED`)
) TYPE=InnoDB; 

# Host: servidor
# Database: lidis
# Table: 'vendasubstestado'
# 
CREATE TABLE `vendasubstestado` (
  `NUMNF` varchar(6) default NULL,
  `DTEMIS` datetime default NULL,
  `NATUREZA` char(2) default NULL,
  `CLIENTE` varchar(20) default NULL,
  `TRANSP` varchar(30) default NULL,
  `TIPOTRANS` char(1) default NULL,
  `PLACATRANS` varchar(8) default NULL,
  `UFTRANS` char(2) default NULL,
  `CGCCPFTRAN` varchar(20) default NULL,
  `ENDTRANS` varchar(30) default NULL,
  `MUNICTRANS` varchar(20) default NULL,
  `UFMUNIC` char(2) default NULL,
  `INSCEST` varchar(20) default NULL,
  `OBS02` varchar(60) default NULL,
  `OBS03` varchar(60) default NULL,
  `OBS04` varchar(60) default NULL,
  `CONTR` varchar(7) default NULL,
  `codigo` int(10) NOT NULL auto_increment,
  `valorproduto` decimal(15,2) default '0.00',
  `ValorNota` decimal(15,2) default '0.00',
  `Vendedor` varchar(5) default NULL,
  `fonetransp` varchar(50) default NULL,
  `cidade` varchar(40) default NULL,
  `cep` varchar(50) default NULL,
  `formapag` varchar(50) default NULL,
  `dias` varchar(50) default NULL,
  `vencimento1` datetime default NULL,
  `vencimento2` datetime default NULL,
  `vencimento3` datetime default NULL,
  `vencimento4` datetime default NULL,
  `vencimento5` datetime default NULL,
  `CondPag` varchar(50) default NULL,
  `status` varchar(50) default NULL,
  `DESCONTO` decimal(15,2) default '0.00',
  `ICMS` varchar(50) default NULL,
  `CFOP` varchar(50) default NULL,
  PRIMARY KEY  (`codigo`),
  KEY `CLIENTE` (`CLIENTE`),
  KEY `codigo` (`codigo`),
  KEY `DTEMIS` (`DTEMIS`),
  KEY `NUMNF` (`NUMNF`)
) TYPE=InnoDB; 

