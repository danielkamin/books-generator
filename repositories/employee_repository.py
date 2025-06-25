from typing import List, Tuple


class EmployeeRepository:
    """Repository for employee-related database operations"""
    
    def __init__(self, db_connection):
        self.db = db_connection
    
    def get_employees_with_project_codes(self, date_from: str, date_to: str) -> List[Tuple]:
        """
        Get list of employees with their project codes for a given date range
        
        Args:
            date_from: Start date in format 'YYYY-MM-DD'
            date_to: End date in format 'YYYY-MM-DD'
            
        Returns:
            List of tuples: (employee_id, employee_name, project_code)
        """
        sql_query = """
        DECLARE @data_od varchar(25) = ?; 
        DECLARE @data_do varchar(25) = ?;
        
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
        """
        
        try:
            return self.db.execute_query(sql_query, (date_from, date_to))
        except Exception as e:
            print(f"Error executing employee query: {e}")
            return []
    
    def get_active_employees(self) -> List[Tuple]:
        """
        Get list of all active employees
        """
        sql_query = """
        SELECT  distinct(p.PRE_PraId)                                                       AS IdPracownika
	       ,p.PRE_Nazwisko + ' ' +p.PRE_Imie1                                           AS Pracownik
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
        ORDER BY p.PRE_PraId 
        """
        
        try:
            return self.db.execute_query(sql_query)
        except Exception as e:
            print(f"Error getting active employees: {e}")
            return []