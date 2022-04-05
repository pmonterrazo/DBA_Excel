SELECT A.VDCEVPEN_NRCCEV AS 'NUM. CONTRATO',  A.VDCEVPEN_CODCLI AS 'COD.CLIENTE',

CASE
         WHEN A.VDCEVPEN_PROD = 550011

         THEN A.VDCEVPEN_QTDPRD
END AS '600ML',

CASE
         WHEN A.VDCEVPEN_PROD = 546004
         OR 
         A.VDCEVPEN_PROD = 551002
                  
         THEN A.VDCEVPEN_QTDPRD
   END AS '300ML',

CASE
         WHEN A.VDCEVPEN_PROD = 555001 

         THEN A.VDCEVPEN_QTDPRD
   END AS '1L',

CASE
         WHEN A.VDCEVPEN_PROD = 655744 
         OR 
         A.VDCEVPEN_PROD = 627030 
         OR 
         A.VDCEVPEN_PROD = 620218 
         
         THEN A.VDCEVPEN_QTDPRD
   END AS 'REFR. PEQ',

      CASE
         WHEN A.VDCEVPEN_PROD = 620214 
         OR 
         A.VDCEVPEN_PROD = 622139 
         OR 
         A.VDCEVPEN_PROD = 622140 
         OR 
         A.VDCEVPEN_PROD = 622151 
         OR 
         A.VDCEVPEN_PROD = 620189
         OR 
         A.VDCEVPEN_PROD = 655742

         THEN A.VDCEVPEN_QTDPRD
   END AS 'REFR. GRD',

CASE
         WHEN A.VDCEVPEN_PROD = 999018
         
         THEN A.VDCEVPEN_QTDPRD
   END AS 'MESA PLAST.'


FROM CEVPED01 AS A



--- CONTAGEM DE CONTRATOS POR CLIENTE

SELECT DISTINCT VDCEVPEN_NRCCEV AS 'NUM CONTR.', VDCEVPEN_CODCLI AS 'COD. CLI'  FROM CEVPED01