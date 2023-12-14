***** EMPRESA -----> AHA SISTEMAS
***** ANALISTA ----> ANTONIO DE PADUA MARTINS
***** PROGRAMADOR -> ANTONIO DE PADUA MARTINS
***** SISTEMA -----> SISTEMA LIDS
***** PROGRAMA ----> INDICE - RECUPERACAO DE ARQUIVOS (INDEXACAO) EXTERNA
***** DATA   ------> 27-12-97
**************************************************************

   SET PROCEDURE to PROCEAPM
   DO INICIO

   SELE 99                             &&***  Arq. de Disco
   use ALID999
   wd        = disco
   wmaq      = micro

   SET COLOR TO W+
   @ 00,00 clear to 24,79
   @ 00,01 to 24,79
   @ 01,20 say "         RECUPERACAO ARQUIVOS    "
   SET COLOR TO

   DECLARE VETOR1[25]

   VETOR1[01]  =  " Clientes              ALID001"
   VETOR1[02]  =  " Fornecedores          ALID002"
   VETOR1[03]  =  " CST  ICMS             ALID003"
   VETOR1[04]  =  " Unidades de Medida    ALID004"
   VETOR1[05]  =  " Municipios            ALID005"
   VETOR1[06]  =  " Concorrentes          ALID006"
   VETOR1[07]  =  " Tipos Receita/Desp.   ALID007"
   VETOR1[08]  =  " Tipos Monetarios      ALID008"
   VETOR1[09]  =  " Itens                 ALID009"
   VETOR1[10]  =  " Lancamentos           ALID010"
   VETOR1[11]  =  " ICMS por Estado       ALID011"
   VETOR1[12]  =  " Galpoes               ALID012"
   VETOR1[13]  =  " Estoque               ALID013"
   VETOR1[14]  =  " Contas a Pagar        ALID014"
   VETOR1[15]  =  " Contas a receber      ALID015"
   VETOR1[16]  =  " Caixa                 ALID016"
   VETOR1[17]  =  " Balanco               ALID017"
   VETOR1[18]  =  " Notas Fiscais         ALID050"
   VETOR1[19]  =  " Duplicatas N.F.       ALID051"
   VETOR1[20]  =  " Itens N.F.            ALID052"
   VETOR1[21]  =  " Senhas                ALID900"
   VETOR1[22]  =  " Micros na Rede        ALID996"
   VETOR1[23]  =  " Impressoras na Rede   ALID997"
   VETOR1[24]  =  " Todos Arquivos               "
   VETOR1[25]  =  " Volta                        "

VAR_1      = 0

DO WHILE .T.

   do janelam with 04,42,19,73,1,0
   SET COLOR TO W+
   VAR_1       =  ACHOICE(05,43,18,72,VETOR1)
   SAVE SCREEN TO SALVATELA1

   set color to w+/g
   @ 09,29 clear to 16,60
   @ 09,29 to 16,60
   @ 11,37 SAY "  I N D E X A N D O"
   @ 13,37 SAY "    A G U A R D E  "

   DO CASE
   CASE VAR_1   =  0
        EXIT

   CASE VAR_1   =  01
        @ 15,30 SAY "ALID001 - Clientes      "

        SELE 01
        use &WD\sistemas\lids\ALID001
        index on codigo                                            to &WD\sistemas\lids\ALID001A
        index on substr(fantasia,1,1)+substr(fantasia,2,29)        to &WD\sistemas\lids\ALID001B
        use

   CASE VAR_1   =  02
        @ 15,30 SAY "ALID002 - Fornecedores   "

        SELE 02
        use &WD\sistemas\lids\ALID002
        index on codigo                                            to &WD\sistemas\lids\ALID002A
        index on substr(fantasia,1,1)+substr(fantasia,2,29)        to &WD\sistemas\lids\ALID002B
        use

   CASE VAR_1   =  03
        @ 15,30 SAY "ALID003 - CST  ICMS      "

        SELE 03
        use &WD\sistemas\lids\ALID003
        index on cst                                               to &WD\sistemas\lids\ALID003A
        use

   CASE VAR_1   =  04
        @ 15,30 SAY "ALID004 - U. de Medida  "

        SELE 04
        use &WD\sistemas\lids\ALID004
        index on cod                                               to &WD\sistemas\lids\ALID004A
        index on substr(nome,1,1)+substr(nome,2,14)                to &WD\sistemas\lids\ALID004B
        use

   CASE VAR_1   =  05
        @ 15,30 SAY "ALID005 - Municipios    "

        SELE 05
        use &WD\sistemas\lids\ALID005
        index on cod                                               to &WD\sistemas\lids\ALID005A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID005B
        use

   CASE VAR_1   =  06
        @ 15,30 SAY "ALID006 - Concorrentes  "

        SELE 06
        use &WD\sistemas\lids\ALID006
        index on cod                                               to &WD\sistemas\lids\ALID006A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID006B
        use

   CASE VAR_1   =  07
        @ 15,30 SAY "ALID007 - Receitas/Desp."

        SELE 07
        use &WD\sistemas\lids\ALID007
        index on cod                                               to &WD\sistemas\lids\ALID007A
        index on substr(nome,1,1)+substr(nome,2,19)                to &WD\sistemas\lids\ALID007B
        use

   CASE VAR_1   =  08
        @ 15,30 SAY "ALID008 - Tipo Monetario"

        SELE 08
        use &WD\sistemas\lids\ALID008
        index on tpmonet                                           to &WD\sistemas\lids\ALID008A
        index on substr(xtpmonet,1,1)+substr(xtpmonet,2,14)        to &WD\sistemas\lids\ALID008B
        use

   CASE VAR_1   =  09
        @ 15,30 SAY "ALID009 - Itens         "

        SELE 09
        use &WD\sistemas\lids\ALID009
        index on cod                                               to &WD\sistemas\lids\ALID009A
        index on substr(nome,1,1)+substr(nome,2,1)+substr(nome,3,1)+substr(nome,4,1)+substr(nome,5,1)+substr(nome,6,1)+substr(nome,7,1)+substr(nome,8,1)+substr(nome,9,1)+substr(nome,10,1)+substr(nome,11,20) to &WD\sistemas\lids\ALID009B
        use

   CASE VAR_1   =  10
        @ 15,30 SAY "ALID010 - Lancamentos   "

        SELE 10
        use &WD\sistemas\lids\ALID010
        index on data                                           to &WD\sistemas\lids\ALID010A
        index on numlancto                                      to &WD\sistemas\lids\ALID010B
        index on recdesp+clicred                                to &WD\sistemas\lids\ALID010C
        index on clicred+recdesp                                to &WD\sistemas\lids\ALID010D
        index on contr                                          to &WD\sistemas\lids\ALID010E
        index on recdesp+clicred+nf                             to &WD\sistemas\lids\ALID010F
        use

   CASE VAR_1   =  11
        @ 15,30 SAY "ALID011 - ICMS por Estado"

        SELE 11
        use &WD\sistemas\lids\ALID011
        index on estado                                         to &WD\sistemas\lids\ALID011A
        use

   CASE VAR_1   =  12
        @ 15,30 SAY "ALID012 - Galpoes       "

        SELE 12
        use &WD\sistemas\lids\ALID012
        index on codigo                                            to &WD\sistemas\lids\ALID012A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID012B
        use

   CASE VAR_1   =  13
        @ 15,30 SAY "ALID013 - Estoque         "

        SELE 13
        use &WD\sistemas\lids\ALID013
        index on item+almox                                        to &WD\sistemas\lids\ALID013A
        index on almox+item                                        to &WD\sistemas\lids\ALID013B
        use

   CASE VAR_1   =  14
        @ 15,30 SAY "ALID014 - Contas a Pagar  "

        SELE 14
        use &WD\sistemas\lids\ALID014
        index on data                                           to &WD\sistemas\lids\ALID014A
        index on dtvenc                                         to &WD\sistemas\lids\ALID014B
        index on numlancto                                      to &WD\sistemas\lids\ALID014C
        index on credor+numlancto                               to &WD\sistemas\lids\ALID014D
        index on tpmonet                                        to &WD\sistemas\lids\ALID014E
        index on contr+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)  to &WD\sistemas\lids\ALID014F
        index on nf+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)     to &WD\sistemas\lids\ALID014G
        use

   CASE VAR_1   =  15
        @ 15,30 SAY "ALID015 - Contas Receber"

        SELE 15
        use &WD\sistemas\lids\ALID015
        index on data                                              to &WD\sistemas\lids\ALID015A
        index on dtvenc                                            to &WD\sistemas\lids\ALID015B
        index on numlancto                                         to &WD\sistemas\lids\ALID015C
        index on cliente+numlancto                                 to &WD\sistemas\lids\ALID015D
        index on tpmonet                                           to &WD\sistemas\lids\ALID015E
        index on contr+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)  to &WD\sistemas\lids\ALID015F
        index on nf+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)     to &WD\sistemas\lids\ALID015G
        use

   CASE VAR_1   =  16
        @ 15,30 SAY "ALID016 - Caixa         "

        SELE 16
        use &WD\sistemas\lids\ALID016
        index on data                                           to &WD\sistemas\lids\ALID016A
        index on numlancto                                      to &WD\sistemas\lids\ALID016B
        index on recdesp+clicred                                to &WD\sistemas\lids\ALID016C
        index on tpmonet                                        to &WD\sistemas\lids\ALID016D
        index on contr                                          to &WD\sistemas\lids\ALID016E
        use

   CASE VAR_1   =  17
        @ 15,30 SAY "ALID017 - Balanco       "

        SELE 17
        use &WD\sistemas\lids\ALID017
        index on ano+galpao+item                          to &WD\sistemas\lids\ALID017A
        index on ano+galpao+xitem                         to &WD\sistemas\lids\ALID017B
        use

   CASE VAR_1   =  18
        @ 15,30 SAY "ALID050 - Notas Fiscais "

        SELE 50
        use &WD\sistemas\lids\ALID050
        index on numnf+cliente                                  to &WD\sistemas\lids\ALID050A
        index on cliente                                        to &WD\sistemas\lids\ALID050B
        index on substr(dtoc(dtemis),7,2)+substr(dtoc(dtemis),4,2)+substr(dtoc(dtemis),1,2)+numnf  to &WD\sistemas\lids\ALID050C
        index on contr                                          to &WD\sistemas\lids\ALID050D
        use

   CASE VAR_1   =  19
        @ 15,30 SAY "ALID051 - Duplicatas NF "

        SELE 51
        use &WD\sistemas\lids\ALID051
        index on numnf                                          to &WD\sistemas\lids\ALID051A
        index on contr                                          to &WD\sistemas\lids\ALID051B
        use

   CASE VAR_1   =  20
        @ 15,30 SAY "ALID052 - Itens N.F.    "

        SELE 52
        use &WD\sistemas\lids\ALID052
        index on numnf                                          to &WD\sistemas\lids\ALID052A
        index on item                                           to &WD\sistemas\lids\ALID052B
        index on contr+item+galpao                              to &WD\sistemas\lids\ALID052C
        use

   CASE VAR_1   =  21
        @ 15,30 SAY "ALID900 -  Senhas        "

        SELE 90                                                    &&***  Senhas
        use &WD\sistemas\lids\ALID900
        index on ident+str(senha,10)                               to &WD\sistemas\lids\ALID900A
        index on senha                                             to &WD\sistemas\lids\ALID900B
        use

   CASE VAR_1   =  22
        @ 15,30 SAY "ALID996 -  Micros        "

        SELE 96                                    &&***  Micros
        use &WD\sistemas\lids\ALID996
        index on codigo                                to &WD\sistemas\lids\ALID996A
        index on substr(nome,1,1)+substr(nome,2,07)    to &WD\sistemas\lids\ALID996B
        use

   CASE VAR_1   =  23
        @ 15,30 SAY "ALID997 -  Impressoras   "

        SELE 97                                    &&***  Impressoras
        use &WD\sistemas\lids\ALID997
        index on codigo                                    to &WD\sistemas\lids\ALID997A
        index on micro                                     to &WD\sistemas\lids\ALID997B
        index on substr(nomeimp,1,1)+substr(nomeimp,2,07)  to &WD\sistemas\lids\ALID997C
        use

   CASE VAR_1   =  24
        @ 15,30 SAY "ALID001 - Clientes      "

        SELE 01
        use &WD\sistemas\lids\ALID001
        index on codigo                                            to &WD\sistemas\lids\ALID001A
        index on substr(fantasia,1,1)+substr(fantasia,2,29)        to &WD\sistemas\lids\ALID001B
        use

        @ 15,30 SAY "ALID002 - Fornecedores   "

        SELE 02
        use &WD\sistemas\lids\ALID002
        index on codigo                                            to &WD\sistemas\lids\ALID002A
        index on substr(fantasia,1,1)+substr(fantasia,2,29)        to &WD\sistemas\lids\ALID002B
        use

        @ 15,30 SAY "ALID003 - CST  ICMS      "

        SELE 03
        use &WD\sistemas\lids\ALID003
        index on cst                                               to &WD\sistemas\lids\ALID003A
        use

        @ 15,30 SAY "ALID004 - U. de Medida  "

        SELE 04
        use &WD\sistemas\lids\ALID004
        index on cod                                               to &WD\sistemas\lids\ALID004A
        index on substr(nome,1,1)+substr(nome,2,14)                to &WD\sistemas\lids\ALID004B
        use

        @ 15,30 SAY "ALID005 - Municipios    "

        SELE 05
        use &WD\sistemas\lids\ALID005
        index on cod                                               to &WD\sistemas\lids\ALID005A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID005B
        use

        @ 15,30 SAY "ALID006 - Concorrentes  "

        SELE 06
        use &WD\sistemas\lids\ALID006
        index on cod                                               to &WD\sistemas\lids\ALID006A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID006B
        use

        @ 15,30 SAY "ALID007 - Receitas/Desp."

        SELE 07
        use &WD\sistemas\lids\ALID007
        index on cod                                               to &WD\sistemas\lids\ALID007A
        index on substr(nome,1,1)+substr(nome,2,19)                to &WD\sistemas\lids\ALID007B
        use

        @ 15,30 SAY "ALID008 - Tipo Monetario"

        SELE 08
        use &WD\sistemas\lids\ALID008
        index on tpmonet                                           to &WD\sistemas\lids\ALID008A
        index on substr(xtpmonet,1,1)+substr(xtpmonet,2,14)        to &WD\sistemas\lids\ALID008B
        use

        @ 15,30 SAY "ALID009 - Itens         "

        SELE 09
        use &WD\sistemas\lids\ALID009
        index on cod                                               to &WD\sistemas\lids\ALID009A
        index on substr(nome,1,1)+substr(nome,2,1)+substr(nome,3,1)+substr(nome,4,1)+substr(nome,5,1)+substr(nome,6,1)+substr(nome,7,1)+substr(nome,8,1)+substr(nome,9,1)+substr(nome,10,1)+substr(nome,11,20) to &WD\sistemas\lids\ALID009B
        use

        @ 15,30 SAY "ALID010 - Lancamentos   "

        SELE 10
        use &WD\sistemas\lids\ALID010
        index on data                                           to &WD\sistemas\lids\ALID010A
        index on numlancto                                      to &WD\sistemas\lids\ALID010B
        index on recdesp+clicred                                to &WD\sistemas\lids\ALID010C
        index on clicred+recdesp                                to &WD\sistemas\lids\ALID010D
        index on contr                                          to &WD\sistemas\lids\ALID010E
        index on recdesp+clicred+nf                             to &WD\sistemas\lids\ALID010F
        use

        @ 15,30 SAY "ALID011 - ICMS por Estado"

        SELE 11
        use &WD\sistemas\lids\ALID011
        index on estado                                         to &WD\sistemas\lids\ALID011A
        use

        @ 15,30 SAY "ALID012 - Galpoes       "

        SELE 12
        use &WD\sistemas\lids\ALID012
        index on codigo                                            to &WD\sistemas\lids\ALID012A
        index on substr(nome,1,1)+substr(nome,2,29)                to &WD\sistemas\lids\ALID012B
        use

        @ 15,30 SAY "ALID013 - Estoque         "

        SELE 13
        use &WD\sistemas\lids\ALID013
        index on item+almox                                        to &WD\sistemas\lids\ALID013A
        index on almox+item                                        to &WD\sistemas\lids\ALID013B
        use

        @ 15,30 SAY "ALID014 - Contas a Pagar  "

        SELE 14
        use &WD\sistemas\lids\ALID014
        index on data                                           to &WD\sistemas\lids\ALID014A
        index on dtvenc                                         to &WD\sistemas\lids\ALID014B
        index on numlancto                                      to &WD\sistemas\lids\ALID014C
        index on credor+numlancto                               to &WD\sistemas\lids\ALID014D
        index on tpmonet                                        to &WD\sistemas\lids\ALID014E
        index on contr+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)  to &WD\sistemas\lids\ALID014F
        index on nf+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)     to &WD\sistemas\lids\ALID014G
        use

        @ 15,30 SAY "ALID015 - Contas Receber"

        SELE 15
        use &WD\sistemas\lids\ALID015
        index on data                                              to &WD\sistemas\lids\ALID015A
        index on dtvenc                                            to &WD\sistemas\lids\ALID015B
        index on numlancto                                         to &WD\sistemas\lids\ALID015C
        index on cliente+numlancto                                 to &WD\sistemas\lids\ALID015D
        index on tpmonet                                           to &WD\sistemas\lids\ALID015E
        index on contr+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)  to &WD\sistemas\lids\ALID015F
        index on nf+substr(dtoc(dtvenc),7,2)+substr(dtoc(dtvenc),4,2)+substr(dtoc(dtvenc),1,2)     to &WD\sistemas\lids\ALID015G
        use

        @ 15,30 SAY "ALID016 - Caixa         "

        SELE 16
        use &WD\sistemas\lids\ALID016
        index on data                                           to &WD\sistemas\lids\ALID016A
        index on numlancto                                      to &WD\sistemas\lids\ALID016B
        index on recdesp+clicred                                to &WD\sistemas\lids\ALID016C
        index on tpmonet                                        to &WD\sistemas\lids\ALID016D
        index on contr                                          to &WD\sistemas\lids\ALID016E
        use

        @ 15,30 SAY "ALID017 - Balanco       "

        SELE 17
        use &WD\sistemas\lids\ALID017
        index on ano+galpao+item                          to &WD\sistemas\lids\ALID017A
        index on ano+galpao+xitem                         to &WD\sistemas\lids\ALID017B
        use

        @ 15,30 SAY "ALID050 - Notas Fiscais "

        SELE 50
        use &WD\sistemas\lids\ALID050
        index on numnf+cliente                                  to &WD\sistemas\lids\ALID050A
        index on cliente                                        to &WD\sistemas\lids\ALID050B
        index on substr(dtoc(dtemis),7,2)+substr(dtoc(dtemis),4,2)+substr(dtoc(dtemis),1,2)+numnf  to &WD\sistemas\lids\ALID050C
        index on contr                                          to &WD\sistemas\lids\ALID050D
        use

        @ 15,30 SAY "ALID051 - Duplicatas NF "

        SELE 51
        use &WD\sistemas\lids\ALID051
        index on numnf                                          to &WD\sistemas\lids\ALID051A
        index on contr                                          to &WD\sistemas\lids\ALID051B
        use

        @ 15,30 SAY "ALID052 - Itens N.F.    "

        SELE 52
        use &WD\sistemas\lids\ALID052
        index on numnf                                          to &WD\sistemas\lids\ALID052A
        index on item                                           to &WD\sistemas\lids\ALID052B
        index on contr+item+galpao                              to &WD\sistemas\lids\ALID052C
        use

        @ 15,30 SAY "ALID900 -  Senhas        "

        SELE 90                                                    &&***  Senhas
        use &WD\sistemas\lids\ALID900
        index on ident+str(senha,10)                               to &WD\sistemas\lids\ALID900A
        index on senha                                             to &WD\sistemas\lids\ALID900B
        use

        @ 15,30 SAY "ALID996 -  Micros        "

        SELE 96                                    &&***  Micros
        use &WD\sistemas\lids\ALID996
        index on codigo                                to &WD\sistemas\lids\ALID996A
        index on substr(nome,1,1)+substr(nome,2,07)    to &WD\sistemas\lids\ALID996B
        use

        @ 15,30 SAY "ALID997 -  Impressoras   "

        SELE 97                                    &&***  Impressoras
        use &WD\sistemas\lids\ALID997
        index on codigo                                    to &WD\sistemas\lids\ALID997A
        index on micro                                     to &WD\sistemas\lids\ALID997B
        index on substr(nomeimp,1,1)+substr(nomeimp,2,07)  to &WD\sistemas\lids\ALID997C
        use

   CASE VAR_1   =  25
        EXIT
   ENDCASE

   set color to w+/g
   @ 09,29 clear to 16,60
   @ 09,29 to 16,60
   @ 11,37 SAY "I N D E X A C A O"
   @ 13,37 SAY "    TERMINADA    "
   TONE(300,15)
   TONE(120,03)
   TONE(300,15)
   REST SCRE FROM SALVATELA1

ENDDO

SET COLOR TO W+
@ 00,00 clear to 24,79

RETURN

