Select A.VDPEDIPE_CODCLI AS 'COD.CLIENTE', 

SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 904502 
         OR 
         A.VDPEDIPE_CODR = 903431 
         OR 
         A.VDPEDIPE_CODR = 903061
         OR 
         A.VDPEDIPE_CODR = 900090
         OR 
         A.VDPEDIPE_CODR = 903046
         OR 
         A.VDPEDIPE_CODR = 902411
         
         THEN A.VDPEDIPE_QTDPRD
   END) AS '600ML',

   SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 903482 

         THEN A.VDPEDIPE_QTDPRD
   END) AS 'HEIN.600ML',

   SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 903129 
         OR 
         A.VDPEDIPE_CODR = 902311 
         OR 
         A.VDPEDIPE_CODR = 902451
         
         THEN A.VDPEDIPE_QTDPRD
   END) AS '300ML',

   SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 904213 
         OR 
         A.VDPEDIPE_CODR = 901133 
         OR 
         A.VDPEDIPE_CODR = 902432

         THEN A.VDPEDIPE_QTDPRD
   END) AS '1L',

   SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 100559 
         OR 
         A.VDPEDIPE_CODR = 100579 
         OR 
         A.VDPEDIPE_CODR = 171690

         THEN A.VDPEDIPE_QTDPRD
   END) AS 'REFR. PEQ',

    SUM(
      CASE
         WHEN A.VDPEDIPE_CODR = 198579 
         OR 
         A.VDPEDIPE_CODR = 50473 
         OR 
         A.VDPEDIPE_CODR = 50475 
         OR 
         A.VDPEDIPE_CODR = 171620 
         OR 
         A.VDPEDIPE_CODR = 99900018 
         OR 
         A.VDPEDIPE_CODR = 197061 
         OR 
         A.VDPEDIPE_CODR = 197196 
         OR 
         A.VDPEDIPE_CODR = 33422 
         THEN A.VDPEDIPE_QTDPRD
   END) AS 'REFR. GRANDE'


from pedit01 as A

where A.vdpedipe_nit > 202111010001

Group by A.VDPEDIPE_CODCLI, A.VDPEDIPE_CODR, A.VDPEDIPE_QTDPRD ,A.vdpedipe_nit