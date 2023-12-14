ALTER TABLE ALID014 ADD COLUMN AutoNumericoNF int;
update ALID002 SET [CGC]=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE([CGC],'.',''),'/',''),'\',''),'-',''),' ','');
update ALID002 SET [CPF]=REPLACE(REPLACE(REPLACE(REPLACE(REPLACE([CPF],'.',''),'/',''),'\',''),'-',''),' ','');