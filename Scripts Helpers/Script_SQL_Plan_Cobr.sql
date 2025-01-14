SELECT CONCAT(CAST(VDCLICLI_REGI AS VARCHAR(300)),CAST(VDCLICLI_NUM AS VARCHAR(300))) as COD_CLI, VDCLICLI_RAZAO50, VDCLICLI_CONTATO, VDCLICLI_FONE, VDCLICLI_VEN FROM CADCLI01 LIMIT 0,50;
-- SELECT TABLES specific GIVE CLIENT - COD_REGION, COD_NUMBER, REASON OF CLIENT, NAME OF CONTACT, TELEPHONE, SALLER.

SELECT CAST(VDCLICLI_NUM + VDCLICLI_REGIAS VARCHAR(300)) + VDCLICLI_RAZAO50, VDCLICLI_CONTATO, VDCLICLI_FONE, VDCLICLI_VEN  FROM DBCONTROL3336001.CADCLI01 LIMIT 0,100;
-- SELECT TABLES specific GIVE CLIENT - COD_REGION, COD_NUMBER, REASON OF CLIENT, NAME OF CONTACT, TELEPHONE, SALLER WITH CONCAT.


SELECT CONCAT(CAST(VDCLICLI_REGI AS VARCHAR(300)),CAST(VDCLICLI_NUM AS VARCHAR(300))) FROM CADCLI01


SELECT PGMOVMOV_CGC, PGMOVMOV_NDUPL, PGMOVMOV_DTE, PGMOVMOV_DTEM, PGMOVMOV_DTV, PGMOVMOV_VALOR FROM RECPGM01 LIMIT 0,50;
--SELECT TABLE specific in OPEN TITLES.

SELECT CRMOVMOV_CCLI, CRMOVMOV_NDUPL, CRMOVMOV_CGC, CRMOVMOV_DTE, CRMOVMOV_DTV, CRMOVMOV_VALOR, CRMOVMOV_NPED, CRMOVMOV_DTS, CRMOVMOV_TIME FROM CADMOV01  LIMIT 0,50;
--SELECT TABLE specific OF OPEN TITLE FULL

SELECT * FROM CADMOV01 LIMIT 10;

SELECT column_name(s)
FROM table1
INNER JOIN table2
ON table1.column_name = table2.column_name;

SELECT CRMOVMOV_CCLI, CRMOV MOV_NDUPL, CRMOVMOV_DTE
FROM ((Orders 
INNER JOIN CADBAI01 ON CRMOVMOV_CCLI = CRMOVBAI_CCLI )
INNER JOIN CADBAI01 ON CRMOVMOV_NDUPL = CRMOVBAI_NDUPL)
INNER JOIN CADBAI01 ON CRMOVMOV_DTE = CRMOVBAI_DTE limit 0,5);


SELECT CRMOVMOV_CCLI, CRMOVMOV_NDUPL, CRMOVMOV_DTE, CRMOVMOV_DTV, CRMOVMOV_VALOR
FROM CADMOV01 
INNER JOIN CADBAI01 ON CRMOVMOV_CCLI = CRMOVBAI_CCLI WHERE limit 0,1000;



--ESTRUTURA DA TABELA

SELECT A.CRMOVMOV_CCLI, C.VDCLICLI_RAZAO50, A.CRMOVMOV_NDUPL, A.CRMOVMOV_DTE, A.CRMOVMOV_DTV, D.VDVENVEN_SIGLA, D.VDVENVEN_SUPVD,A.CRMOVMOV_NDUPL, B.CRMOVBAI_NDUPL 
FROM CADMOV01 AS A
INNER JOIN CADBAI01 AS B
ON A.CRMOVMOV_CCLI = B.CRMOVBAI_CCLI
INNER JOIN CADCLI01 AS C
ON A.CRMOVMOV_CGC=C.VDCLICLI_CGC
INNER JOIN CADVEN01 AS D
ON A.CRMOVMOV_VEN=D.VDVENVEN_SIGLA
LIMIT 10

---------------------------------


--ISOLATE SELECT ORDER 
SELECT * FROM NPEDPT01 WHERE VDPEDPPT_NPEDPT = 202112170016;
VDPEDPPT_VEN
---

----SELECT IN CADMOV01
SELECT * FROM CADMOV01 LIMIT 10;

---- OUT

SELECT  DISTINCT A.CRMOVMOV_CCLI, C.VDCLICLI_RAZAO50, A.CRMOVMOV_NDUPL, A.CRMOVMOV_DTE, A.CRMOVMOV_DTV, D.VDVENVEN_SIGLA,
B.VDPEDPPT_TPCOBR, VDFATCFA_VALPRDORI, A.CRMOVMOV_VALOR, D.VDVENVEN_SUPVD, C.VDCLICLI_CONTATO, C.VDCLICLI_FONE
FROM CADMOV01 AS A
INNER JOIN NPEDPT01 AS B
ON A.CRMOVMOV_CCLI=B.VDPEDPPT_CODCLI
INNER JOIN CADCLI01 AS C
ON A.CRMOVMOV_CGC=C.VDCLICLI_CGC
INNER JOIN CADVEN01 AS D
ON A.CRMOVMOV_VEN=D.VDVENVEN_SIGLA
INNER JOIN NC211201 AS E
ON A.CRMOVMOV_NPED=E.VDFATCFA_NPED
WHERE B.VDPEDPPT_TPCOBR > 1

----




--Calculando os dias entre duas datas startdate and enddate
SELECT DAYS_BETWEEN(CAST(LEFT(CAST(CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE), CURDATE()) AS DIAS FROM CADMOV01 LIMIT 10


----- SCRIPT PLANILHA COBRANÇA ------------

SELECT DISTINCT A.CRMOVMOV_NDUPL AS 'NUM.TITULO', A.CRMOVMOV_CCLI AS 'COD. CLIENTE', B.VDCLICLI_RAZAO50 AS 'RAZÃO SOCIAL', CAST(RIGHT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA EMISSAO', CAST(RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA VENCIMENTO', DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) AS DIAS, (CASE WHEN DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) > 0 THEN 'V' ELSE 'AV' END) AS 'V/AV', A.CRMOVMOV_MOD AS 'TP. COBRANÇA', C.VDPEDPPT_VLR_FCEM AS 'VLR. ORIGINAL',A.CRMOVMOV_VALOR AS 'VLR. TITULO', (DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.002 * A.CRMOVMOV_VALOR) AS JUROS, ((DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.002 * A.CRMOVMOV_VALOR) + A.CRMOVMOV_VALOR) AS 'VLR. TOTAL', D.VDVENVEN_CODMOV AS 'VEN', D.VDVENVEN_SUPVD AS 'GA', B.VDCLICLI_CONTATO AS 'CONTATO', B.VDCLICLI_FONE AS 'TELEFONE 1', B.VDCLICLI_FONE2 AS 'TELEFONE 2', B.VDCLICLI_CEL1 AS 'TELEFONE 3', B.VDCLICLI_CEL2 AS 'TELEFONE 4', SUBSTRING(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),1,4) AS 'ANO'

FROM CADMOV01 AS A 
INNER JOIN CADCLI01 AS B ON CAST(A.CRMOVMOV_CCLI AS VARCHAR(300))=CAST(VDCLICLI_REGI AS VARCHAR(300)) || REPEAT('0',4-LENGTH(CAST(VDCLICLI_NUM AS VARCHAR(300)))) || CAST(VDCLICLI_NUM AS VARCHAR(300))
INNER JOIN NPEDPT01 AS C ON A.CRMOVMOV_NPED=C.VDPEDPPT_NPEDPT
INNER JOIN CADVEN01 AS D ON A.CRMOVMOV_VEN=D.VDVENVEN_CODMOV


LIMIT 10


--------------------FIM PLANILHA COBRANÇA ------------

---------PLANILHA PLP - SCRIPT SQL


SELECT A.VDPEDCPE_CODCLI  AS 'COD. CLIENTE',  CAST(SUBSTRING(CAST(A.VDPEDCPE_NPED AS VARCHAR(300)),9,5) AS VARCHAR(300)) AS 'PEDIDO', B.VDCLICLI_RAZAO50 AS 'RAZÃO SOCIAL', A.VDPEDCPE_VEN_ORI_CLI AS 'VD', D.VDVENVEN_SUPVD AS 'SUP', 

CASE 

      WHEN(A.VDPEDCPE_TPCOBR)  >= 4 THEN 'PRAZO'  
      WHEN(A.VDPEDCPE_TPCOBR)  = 1 THEN 'A VISTA' END AS 'PRAZO', CAST(A.VDPEDCPE_CPG AS VARCHAR(300)) || ' - ' || CAST(E.VDCADPAG_DESCR AS VARCHAR(300)) AS 'COND. PG',

CASE 

    WHEN(F.VDPEDIPE_OCOKD) >= 1 THEN 'VENDA'
    WHEN(F.VDPEDIPE_OCOKD) >= 2 THEN 'BONIFICAÇÃO'
    WHEN(F.VDPEDIPE_OCOKD) >= 1 AND F.VDPEDIPE_OCOKD <= 2 THEN 'VENDA/BONI' END AS 'TIPO'

,A.VDPEDCPE_VLR_FCEM AS 'VLR. PEDIDO', C.CRMOVMOV_VALOR AS 'VLR. DO TITULO', 

CASE 

   WHEN(CAST(LEFT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE) < CURDATE()) THEN C.CRMOVMOV_VALOR END AS 'VENCIDOS',
   
   CAST(RIGHT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA VENCIMENTO',

CASE

   WHEN(CAST(LEFT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(C.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE) < CURDATE()) THEN 'NÃO' ELSE 'LIBERAR' END AS 'LIBERAR',

CASE
   WHEN F.VDPEDIPE_CODR = 900090 THEN F.VDPEDIPE_QTDPRD 
END AS '900090',
CASE
   WHEN F.VDPEDIPE_CODR = 901133 THEN F.VDPEDIPE_QTDPRD 
END AS '901133',

CASE
   WHEN F.VDPEDIPE_CODR = 902311 THEN F.VDPEDIPE_QTDPRD 
END AS '902311',

CASE
   WHEN F.VDPEDIPE_CODR = 903061 THEN F.VDPEDIPE_QTDPRD 
END AS '903061',

CASE
   WHEN F.VDPEDIPE_CODR = 903129 THEN F.VDPEDIPE_QTDPRD 
END AS '903129',

CASE
   WHEN F.VDPEDIPE_CODR = 902402 THEN F.VDPEDIPE_QTDPRD 
END AS '902402',

CASE
   WHEN F.VDPEDIPE_CODR = 904213 THEN F.VDPEDIPE_QTDPRD 
END AS '904213',

CASE
   WHEN F.VDPEDIPE_CODR = 903129 THEN F.VDPEDIPE_QTDPRD 
END AS '903129',

CASE
   WHEN F.VDPEDIPE_CODR = 2494 THEN F.VDPEDIPE_QTDPRD 
END AS '2494',

CASE
   WHEN F.VDPEDIPE_CODR = 107381 THEN F.VDPEDIPE_QTDPRD 
END AS '107381',

CASE
   WHEN F.VDPEDIPE_CODR = 1525940 THEN F.VDPEDIPE_QTDPRD 
END AS '1525940',

CAST(RIGHT(CAST(G.VDCEVPEN_DTV AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(G.VDCEVPEN_DTV AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(G.VDCEVPEN_DTV AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'VALIDADE'
 
FROM PEDCP01 AS A 
INNER JOIN CADCLI01 AS B ON CAST(A.VDPEDCPE_CODCLI AS VARCHAR(300))=CAST(VDCLICLI_REGI AS VARCHAR(300)) || REPEAT('0',4-LENGTH(CAST(VDCLICLI_NUM AS VARCHAR(300)))) || CAST(VDCLICLI_NUM AS VARCHAR(300)) 
INNER JOIN CADMOV01 AS C ON A.VDPEDCPE_CODCLI=C.CRMOVMOV_CCLI 
INNER JOIN CADVEN01 AS D ON A.VDPEDCPE_VEN_ORI_CLI=D.VDVENVEN_CODMOV 
INNER JOIN CONDPG01 AS E ON A.VDPEDCPE_CPG=E.VDCADPAG_COD 
INNER JOIN PEDIT01 AS F ON A.VDPEDCPE_NPED=F.VDPEDIPE_NIT 
INNER JOIN CEVPED01 AS G ON A.VDPEDCPE_CODCLI=G.VDCEVPEN_CODCLI 
 
WHERE CAST(LEFT(CAST(VDPEDCPE_NPED AS VARCHAR(300)),4) || SUBSTRING(CAST(VDPEDCPE_NPED AS VARCHAR(300)),5,2) || SUBSTRING(CAST(VDPEDCPE_NPED AS VARCHAR(300)),7,2) AS DATE) = CURDATE()
 
LIMIT 10

----- FIM DO SCRIPT DA PLANILHA PLP --------------
