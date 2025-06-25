declare @data_od varchar(25); declare @data_do varchar(25);

SET @data_od = '2025-01-01';
SET @data_do = '2025-01-31';
WITH ListaPracownikow AS
(
	SELECT  distinct(p.PRE_PraId)                                                       AS IdPracownika
	       ,LTRIM(RTRIM(p.PRE_Imie1)) + ' ' + LTRIM(RTRIM(p.PRE_Nazwisko))              AS Pracownik
	       ,p.PRE_Kod                                                                   AS Kod
	       ,CASE WHEN wyp.WPL_NumerPelny not LIKE 'U%' THEN 'etat'  ELSE 'zlecenie' END AS TypZatrudnienia
	       ,wyp.WPL_NumerPelny
	FROM CDN.PracEtaty AS p
	INNER JOIN CDN.Wyplaty AS wyp
	ON wyp.WPL_PraId = p.PRE_PraId
	INNER JOIN CDN.WypElementy AS ele
	ON wyp.WPL_WplId = ele.WPE_WplId
	WHERE p.PRE_DataDo >= CONVERT(datetime, '2999-12-31', 120)
	AND wyp.WPL_DataOd >= @data_od
	AND wyp.WPL_DataOd <= @data_do
	AND UPPER(ele.WPE_Nazwa) not LIKE '%SODIR%'
	AND UPPER(ele.WPE_Nazwa) not LIKE '%ZFRON%'
	AND UPPER(ele.WPE_Nazwa) not LIKE '%PZU%'
	AND UPPER(ele.WPE_Nazwa) not LIKE '%KOMORNICZE%'
	AND UPPER(ele.WPE_Nazwa) not LIKE '%ZASI%' 
)
SELECT  DISTINCT(lp.IdPracownika)
       ,lp.Pracownik
       ,lp.Kod
       ,dp.PRJ_Kod
FROM ListaPracownikow lp
INNER JOIN CDN.PracPlanDni pld
ON pld.PPL_PraId = lp.IdPracownika
INNER JOIN CDN.PracPlanDniGodz pldg
ON pldg.PGL_PplId = pld.PPL_PplId
INNER JOIN CDN.DefProjekty dp
ON pldg.PGL_PrjId = dp.PRJ_PrjId
WHERE pld.PPL_Data >= @data_od
AND pld.PPL_Data <= @data_do
AND pld.PPL_TypDnia = 1
ORDER BY lp.IdPracownika