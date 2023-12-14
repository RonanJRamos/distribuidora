/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE */;
/*!40101 SET SQL_MODE='STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES */;
/*!40103 SET SQL_NOTES='ON' */;

CREATE TABLE `NFentrada_Cte` (
  `Codigo` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `CodNota` int(11) DEFAULT '0',
  `Emissao` datetime,
  `NumeroNFCte` varchar(20) DEFAULT NULL,
  `ChaveAcesso` varchar(50) DEFAULT NULL,
  `NumeroNFe` varchar(50) DEFAULT NULL,
  `CNPJ` varchar(20) DEFAULT NULL,
  `Nome` varchar(100) DEFAULT NULL,
  `Valor` decimal(10,2) DEFAULT NULL,
  'Parcelas` int(11) DEFAULT '0',	
  `FormaPag` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`Codigo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
CREATE TABLE `NFentrada_Cte_Vencimentos`(
  `Codigo` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Cod_Cte` int(11) DEFAULT '0',
  `Vencimento` datetime,
  `Valor_Parcela` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`Codigo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;
/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
