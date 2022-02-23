



'GIA NA FTIAXO LINKEDSERVER
'GENERAL
Provider : SQL SERVER NATIVE CLIENT 11.0
Product Name : sqlserver
DATASOURCE : ONOMA_APOMAKRYSMENOY_PC\SQLEXPRESS
'SECURITY
TSEKARV
 Be Made using this security context:
  sa
  p@ssw0rd   π.χ.

πως δουλευω το linked server π.χ. το αρχειο πελατων ειναι [OIKIAS].MERCURY.dbo.PEL
' OIKIA ΕΙΝΑΙ ΤΟ ΟΝΟΜΑ ΤΟΥ LINKED SERVER

USE TEST5
insert into PEL(KOD,EPO) SELECT KOD,EPO FROM [OIKIAS].MERCURY.dbo.PEL