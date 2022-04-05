SELECT DISTINCT A.CRMOVMOV_NDUPL AS 'NUM.TITULO',

A.CRMOVMOV_CCLI AS 'COD. CLIENTE',

B.VDCLICLI_RAZAO50 AS 'RAZÃO SOCIAL',

CAST(RIGHT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA EMISSAO',

CAST(RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) || '/' || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || '/' || LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) AS VARCHAR(300)) AS 'DATA VENCIMENTO',

DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) AS DIAS,
(CASE WHEN DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) > 0 THEN 'V' ELSE 'AV' END) AS 'V/AV',

A.CRMOVMOV_MOD AS 'TP. COBRANÇA',

(A.CRMOVMOV_VALOR * 1) AS 'VLR. ORIGINAL',

A.CRMOVMOV_VALOR AS 'VLR. TITULO',

(DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.003 * A.CRMOVMOV_VALOR) AS 'JUROS',

((DAYS_BETWEEN(CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE),CURDATE()) * 0.003 * A.CRMOVMOV_VALOR) + A.CRMOVMOV_VALOR) AS 'VLR. TOTAL',

B.VDCLICLI_VEN AS 'VEN',

D.VDVENVEN_SUPVD AS 'GA',

B.VDCLICLI_CONTATO AS 'CONTATO',

B.VDCLICLI_FONE AS 'TELEFONE 1',

B.VDCLICLI_FONE2 AS 'TELEFONE 2',

B.VDCLICLI_CEL1 AS 'TELEFONE 3',

B.VDCLICLI_CEL2 AS 'TELEFONE 4',

SUBSTRING(CAST(A.CRMOVMOV_DTE AS VARCHAR(300)),1,4) AS 'ANO'


FROM DBCONTROL3709006.CADMOV06 AS A
INNER JOIN DBCONTROL3709006.CADCLI01 AS B ON CAST(A.CRMOVMOV_CCLI AS VARCHAR(300))=CAST(B.VDCLICLI_REGI AS VARCHAR(300)) || REPEAT('0',4-LENGTH(CAST(B.VDCLICLI_NUM AS VARCHAR(300)))) || CAST(B.VDCLICLI_NUM AS VARCHAR(300))
INNER JOIN DBCONTROL3709006.CADVEN06 AS D ON A.CRMOVMOV_VEN=D.VDVENVEN_CODMOV