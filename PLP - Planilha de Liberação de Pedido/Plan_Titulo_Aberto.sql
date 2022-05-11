SELECT A.CRMOVMOV_CCLI AS 'COD. CLIENTE', A.CRMOVMOV_NDUPL AS 'NUM. TITULO', A.CRMOVMOV_VALOR AS 'VLR DO TITULO'

 FROM CADMOV01 AS A 

WHERE CAST(LEFT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),4) || SUBSTRING(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),5,2) || RIGHT(CAST(A.CRMOVMOV_DTV AS VARCHAR(300)),2) AS DATE) < CURDATE()

GROUP BY A.CRMOVMOV_CCLI, A.CRMOVMOV_DTV, A.CRMOVMOV_VALOR, A.CRMOVMOV_NDUPL