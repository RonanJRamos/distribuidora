ALTER TABLE itensentradanf ADD COLUMN ALIQ_COFINS decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN VL_COFINS decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN SeguroP decimal(10,2);

ALTER TABLE produtos ADD COLUMN ValorST decimal(10,2);
ALTER TABLE produtos ADD COLUMN ValorIPI decimal(10,2);
ALTER TABLE produtos ADD COLUMN ValorFrete decimal(10,2);
ALTER TABLE produtos ADD COLUMN ValorSeguro decimal(10,2);
ALTER TABLE produtos ADD COLUMN ValorDespesas decimal(10,2);
ALTER TABLE produtos ADD COLUMN ValorDifICMS decimal(10,2);

ALTER TABLE EntradaNf ADD COLUMN Natureza text(150);
ALTER TABLE EntradaNf ADD COLUMN tipoFrete int(11) DEFAULT '0';
ALTER TABLE EntradaNf ADD COLUMN Valor_PIS decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN Valor_COFINS decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN Cancelada int(11) DEFAULT '0';
ALTER TABLE EntradaNf ADD COLUMN pis decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN cofins decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN basecalcpis decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN basecalccofins decimal(10,2);
ALTER TABLE EntradaNf ADD COLUMN referencia varchar(255) DEFAULT NULL;
ALTER TABLE EntradaNf ADD COLUMN estoquenormal int(11) DEFAULT '0';
ALTER TABLE EntradaNf ADD COLUMN estoquefiscal int(11) DEFAULT '0';
ALTER TABLE EntradaNf ADD COLUMN avulsa int(11) DEFAULT '0';
ALTER TABLE EntradaNf ADD COLUMN Finalizado int(11) DEFAULT '0';

ALTER TABLE itensentradanf ADD COLUMN CodProdutoFornecedor varchar(255) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN DescricaoFornecedor varchar(255) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN CST varchar(20) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN NCM varchar(20) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN Icms decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN BC_Icms_Sub decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN Valor_Icms_Sub decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN PercIpi decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN Cst_Pis varchar(20) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN VL_BC_PIS decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN ALIQ_PIS decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN VL_PIS decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN CST_COFINS varchar(20) DEFAULT NULL;
ALTER TABLE itensentradanf ADD COLUMN VL_BC_COFINS decimal(10,2);


ALTER TABLE itensentradanf ADD COLUMN FreteP decimal(10,2);

ALTER TABLE itensentradanf ADD COLUMN OutrasDespP decimal(10,2);
ALTER TABLE itensentradanf ADD COLUMN DescontoP decimal(10,2);

CREATE TABLE produtofornecedor(
	codigo int(10) NOT NULL AUTO_INCREMENT,
	codfornecedor int(11) DEFAULT '0',
	codprodutofornecedor varchar(50) DEFAULT NULL,
	nomeprodutofornecedor varchar(50) DEFAULT NULL,
	codproduto int(11) DEFAULT '0',
PRIMARY KEY (codigo)
);
CREATE TABLE conversaocfop(
	codigo int(10) NOT NULL AUTO_INCREMENT,
	cfopentrada varchar(10) DEFAULT NULL,
	cfopsaida varchar(10) DEFAULT NULL,
PRIMARY KEY (codigo)
);
CREATE TABLE configuracaogeral(
	codigo int(10) NOT NULL AUTO_INCREMENT,
	aproveitamentocredito int(11) DEFAULT '0',
	cstpis varchar(10) DEFAULT NULL,
	cstcofins varchar(10) DEFAULT NULL,
	aliqcofins decimal(10,2) ,
	aliqpis decimal(10,2),
	codigousuario int(11) DEFAULT '0',
	apropriaicms int(11) DEFAULT '0',
	naoapropriast int(11) DEFAULT '0',
	trocacst int(11) DEFAULT '0',
	csttroca varchar(10) DEFAULT NULL,
	csticms varchar(10) DEFAULT NULL,
	rateiafrete int(11) DEFAULT '0',
	rateiaipi int(11) DEFAULT '0',
	rateiaicms int(11) DEFAULT '0',
	rateiaicmsst int(11) DEFAULT '0',
	rateiaseguro int(11) DEFAULT '0',
	rateiaoutrasdesp int(11) DEFAULT '0',
	naomostratelaacertopreco int(11) DEFAULT '0',
	csttroca_cupom varchar(50) DEFAULT NULL,
	naoexigechaveacesso int(11) DEFAULT '0',
	calcularbasest int(11) DEFAULT '0',
	usarconfiguracaoapropriacao int(11) DEFAULT '0',
	SalvarFornecedorMarca int(11) DEFAULT '0',
	caminhoxml varchar(250) DEFAULT NULL,
PRIMARY KEY (codigo)
);

ALTER TABLE ALID015 ADD COLUMN NossoNumero text(50);
ALTER TABLE ALID015 ADD COLUMN Boletogerado YesNo NULL;
ALTER TABLE Produtos ADD COLUMN Desativado YesNo NULL;

ALTER TABLE ALTER TABLE tblNFEProduto ADD COLUMN segurop  double;
ALTER TABLE tblNFEProduto ADD COLUMN outrasdespp  double;
ALTER TABLE tblNFEProduto  ALTER COLUMN Especie Text(150);
ALTER TABLE tblContasPag ALTER COLUMN TipoConta Text(150);
 ADD COLUMN segurop  double;
ALTER TABLE tblNFEProduto ADD COLUMN outrasdespp  double;
ALTER TABLE tblNotaFiscalEntrada ALTER COLUMN Especie Text(150);
ALTER TABLE tblContasPag ALTER COLUMN TipoConta Text(150);
