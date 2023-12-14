ALTER TABLE produtos ADD COLUMN LimiteVenda Decimal(12,2);

CREATE TABLE arquivoxml (
  Codigo int(11) NOT NULL AUTO_INCREMENT,
  Arquivo longblob,
  NF varchar(50) DEFAULT NULL,
  Tipo int(11) DEFAULT 0,
  Evento int(11) DEFAULT 0,
  Data date DEFAULT NULL,
  PRIMARY KEY (Codigo)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

ALTER TABLE Produtos ADD COLUMN ultimaAlteracao date NULL DEFAULT NULL;
ALTER TABLE Produtos ADD COLUMN Desativado int NULL DEFAULT 0;
ALTER TABLE relprodutos ADD COLUMN Desativado int NULL DEFAULT 0;
CREATE TABLE relestoqueperiodo (
  Codigo int(11) NOT NULL AUTO_INCREMENT,
  CodigoProduto int(11) NULL DEFAULT 0,
  Nome varchar(255) NULL DEFAULT NULL,
  Estoque numeric(10,2) NULL DEFAULT 0,
  Custo numeric(10,2) NULL DEFAULT 0,
  QuantUltimaCompra numeric(10,2) NULL DEFAULT 0,
  DataUltimaCompra date NULL DEFAULT NULL,
  NF varchar(255) NULL DEFAULT NULL,
  FORNECEDOR varchar(255) NULL DEFAULT NULL,
  PRIMARY KEY (Codigo)
);
ALTER TABLE produtoimposto ADD COLUMN AliquotaInternaDestino Decimal(12,2) NULL;
ALTER TABLE produtoimposto ADD COLUMN AliquotaInternaOrigem Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN AliquotaInternaDestino Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN AliquotaInternaOrigem Decimal(12,2) NULL;
ALTER TABLE produtoimposto ADD COLUMN  DescricaoCFOP varchar(20);
ALTER TABLE produtoimposto ADD COLUMN pFCPUFDest Decimal(12,2) NULL;
ALTER TABLE produtoimposto ADD COLUMN  pICMSInter varchar(250);
ALTER TABLE alid052 ADD COLUMN pFCPUFDest Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN  pICMSInter varchar(250);
ALTER TABLE alid052 ADD COLUMN PerPartilhaInterestadual Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN VIcmsFCPDestino Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN VIcmsDestino Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN VIcmsRemetente Decimal(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN CodigoImposto long;
ALTER TABLE alid052 ADD COLUMN  CEST varchar(20);
ALTER TABLE alid050 ADD COLUMN ValorTotalFCP Decimal(12,2) NULL;
ALTER TABLE alid050 ADD COLUMN ValorTotalICMSDestino Decimal(12,2) NULL;
ALTER TABLE alid050 ADD COLUMN ValorTotalICMSEmitente Decimal(12,2) NULL;
ALTER TABLE Produtos ADD COLUMN  CEST varchar(20);
ALTER TABLE alid050 ADD COLUMN efornecedor int NULL;
UPDATE ALID050 SET EFORNECEDOR=0;

ALTER TABLE historicoproduto ADD COLUMN Anterior Decimal(12,2) NULL;
ALTER TABLE historicoproduto ADD COLUMN Saldo Decimal(12,2) NULL;
ALTER TABLE alid050 ADD COLUMN NomeVendedorImprimir varchar(255) NULL DEFAULT NULL;

ALTER TABLE itensentradanf ADD COLUMN Santa Decimal(12,2) NULL;
ALTER TABLE itensentradanf ADD COLUMN California Decimal(12,2) NULL;
ALTER TABLE produtos ADD COLUMN SubItem varchar(100) CHARACTER SET utf8 NULL DEFAULT NULL;
ALTER TABLE entradanf ADD COLUMN chave varchar(255) NULL DEFAULT NULL;
ALTER TABLE entradanf ADD COLUMN Protocolo varchar(255) NULL DEFAULT NULL;
ALTER TABLE entradanf ADD COLUMN modelo varchar(255) NULL DEFAULT NULL;
ALTER TABLE naturezaoperacao ADD COLUMN CfopPadrao varchar(50) NULL DEFAULT NULL;
ALTER TABLE relprodutos ADD COLUMN Compra double(12,2) NULL;
ALTER TABLE naturezaoperacao  ADD COLUMN NaoCarregaPedido int(11) NULL DEFAULT 0;
ALTER TABLE alid052 ADD COLUMN OrdemCompra varchar(150) NULL DEFAULT NULL;
ALTER TABLE alid052 ADD COLUMN NumeroOrdemCompra varchar(150) NULL DEFAULT NULL;
ALTER TABLE alid052 ADD COLUMN confins double(12,2) NULL;
ALTER TABLE alid052 ADD COLUMN Pis double(12,2) NULL;