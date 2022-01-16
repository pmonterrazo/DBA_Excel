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