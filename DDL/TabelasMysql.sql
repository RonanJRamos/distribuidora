ALTER TABLE alid201 ADD COLUMN LimiteVenda Decimal(12,2);
ALTER TABLE produtos ADD COLUMN LimiteVenda Decimal(12,2);
ALTER TABLE itensentradanf ADD COLUMN FreteP Decimal(12,2);
ALTER TABLE itensentradanf ADD COLUMN OutrasDespP Decimal(12,2);
ALTER TABLE itensentradanf ADD COLUMN DescontoP Decimal(12,2);

ALTER TABLE contasacado ADD COLUMN instrucao3 varchar(60) DEFAULT NULL;

ALTER TABLE contasacado ADD COLUMN instrucao4 varchar(60) DEFAULT NULL;

ALTER TABLE contasacado ADD COLUMN instrucao5 varchar(60) DEFAULT NULL;

ALTER TABLE contasacado ADD COLUMN CodConvenio varchar(50) DEFAULT NULL;

ALTER TABLE contasacado ADD COLUMN ImprimeMora int(11) DEFAULT 0;
ALTER TABLE contasacado ADD COLUMN ImprimeJuros int(11) DEFAULT 0;
ALTER TABLE contasacado ADD COLUMN Especie varchar(50) DEFAULT NULL;

