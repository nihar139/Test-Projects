Imports System.Data.OracleClient
Imports System.IO
'Imports ClosedXML.Excel
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Net.Mail
Imports System.Configuration
Imports System.Globalization

Module Module1

    Dim data_adapt As OracleDataAdapter
    Dim myConn As SqlConnection = New SqlConnection("Server=135.21.21.100;Database=WENCO;Uid=sa;Pwd=Zxc@#$19;")
    Dim ds As DataSet
    ' Dim conn As New ConnectionOracleDB
    Dim xlRange As Excel.Range
    'Dim Str_conn As String = "User ID=RMIS;Password=fin_37km;Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 176.0.0.95)(PORT = 1521))(CONNECT_DATA =(SID =  Noadev)));Persist Security Info=True"
    'Dim Str_conn As String = ConfigurationManager.ConnectionStrings("PRD").ConnectionString
    Dim Str_conn As String = ConfigurationManager.ConnectionStrings("QA").ConnectionString
    'Dim Str_conn As String = ConfigurationManager.ConnectionStrings("PRD").ConnectionString
    Dim conn As New OracleConnection(Str_conn)

#Region "Query"
    Dim qrryrlyScheduleABPDispatchSOomq As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                       "  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And " +
                                       "  ftr_par_cd   IN ('NOADESP_13','NOADESP_01','NOADESP_08','NOADESP_05','NOADESP_19','NOADESP_16','SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML','KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31') " +
                                       "   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION " +
                                       "  Select SUM(FTR_VALUE) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                       "  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And ftr_par_cd   IN " +
                                       "  ('NOADESP_13','NOADESP_01','NOADESP_08','NOADESP_05','NOADESP_19','NOADESP_16','SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML','KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31') " +
                                       "  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL) )"
    'Dim qrrmonthlyScheduleABPDispatchSOomq As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
    '                                    "  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And " +
    '                                    "  ftr_par_cd   IN ('JESOTOTSJ','SIZORTOTML','SIZORTOTSK','SOTOTSIL','NOADESP_01','NOADESP_05','NOADESP_08', " +
    '                                    "  'KBIM_D03','KBIM_D05','KBIM_D07','KBIM_D13','KIM_P57','KIM_P61')  And ftr_year_month='201909'  UNION " +
    '                                    "  Select SUM(FTR_VALUE) As a1,    0                  AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
    '                                    "  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And ftr_par_cd   IN " +
    '                                    "  ('JESOTOTSJ','SIZORTOTML','SIZORTOTSK','SOTOTSIL','NOADESP_01','NOADESP_05','NOADESP_08','KBIM_D03','KBIM_D05','KBIM_D07','KBIM_D13','KIM_P57','KIM_P61') " +
    '                                    "  And ftr_year_month='201909')"
    Dim qryyrlyScheduleABPDispatchFOOMQ As String = " SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And ftr_par_cd   IN ('NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17','KIM_P64','KIM_P66','FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML','KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11') " +
                                                  " And ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And ftr_par_cd   IN ('NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17','KIM_P64','KIM_P66','FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML','KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  )"

    'Dim qrymonthlyScheduleABPDispatchFOOMQ As String = " SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
    '                                      " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And ftr_par_cd   IN ('KIM_P55','KBIM_D09','KIM_P59','KIM_D11','KBIM_D15','FINORTOTML','FINORTOTSK','JEFOTOTSJ','KBFOTOTSJ','NOADESP_03','NOADESP_06','NOADES_09') " +
    '                                      " And ftr_year_month='201909'  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
    '                                      " And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And ftr_par_cd   IN ('KIM_P55','KBIM_D09','KIM_P59','KIM_D11','KBIM_D15','FINORTOTML','FINORTOTSK', " +
    '                                      " 'JEFOTOTSJ','KBFOTOTSJ','NOADESP_03','NOADESP_06','NOADES_09')   And ftr_year_month='201909'  )"
    Dim qrymonthlyactualdispatchSOOMQ As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_13','NOADESP_16','NOADESP_01','NOADESP_08','NOADESP_05','NOADESP_19','SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML','KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qrymonthlyactualdispatchSOOMQ1 As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_13','NOADESP_16','NOADESP_01','NOADESP_08','NOADESP_05','NOADESP_19','SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML','KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qrymonthlyactualdispatchFOOMQ As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17','KIM_P64','KIM_P66','FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML','KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qrymonthlyactualdispatchFOOMQ1 As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17','KIM_P64','KIM_P66','FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML','KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    'Dim qryABPScheduledispatchSOTSJ As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
    '                                    "  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And " +
    '                                    "  ftr_par_cd   IN ('JESOTOTSJ','NOADESP_01' " +
    '                                    "  'KBIM_D13')  And ftr_year_month='Yer-17-18'  UNION " +
    '                                    "  Select SUM(FTR_VALUE) As a1,    0           AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
    '                                    "  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDDESP','KIMDESP')  And ftr_par_cd   IN " +
    '                                    "  ('JESOTOTSJ','NOADESP_01','KBIM_D13') " +
    '                                    "  And ftr_year_month='Yer-17-18')"
    Dim qryABPScheduledispatchSOTSJ As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target WHERE ftr_rec_type ='ABP' And ftr_par_cd   IN ('JESOTOTSJ','NOADESP_01','KBIM_D13','KBIM_D31')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION Select SUM(FTR_VALUE) As a1,    0           AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' And ftr_par_cd   IN ('JESOTOTSJ','NOADESP_01','KBIM_D13','KBIM_D31') And ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qryABPScheduledispatchFOTSJ As String = " SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP' And ftr_par_cd   IN ('KBIM_D15','JEFOTOTSJ','NOADESP_03') " +
                                                  " And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_par_cd   IN ('KBIM_D15', " +
                                                  " 'JEFOTOTSJ','NOADESP_03')   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  )"
    Dim qryABPScheduledispatchSOTSK As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  AND ftr_par_cd    IN ('SIZORTOTSK','NOADESP_08','KBIM_D07')  AND ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)    UNION    SELECT SUM(FTR_VALUE) AS a1,  0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'  AND ftr_fac_cd    IN ('JODDESP','NOADESP','KBDESP','KIMDESP')  AND ftr_par_cd    IN ('SIZORTOTSK','NOADESP_08','KBIM_D07')  AND ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  )"
    Dim qryABPScheduledispatchFOTSK As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  AND ftr_par_cd    IN ('FINORTOTSK','NOADESP_09','KBIM_D09','KIM_P64')  AND ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION    SELECT SUM(FTR_VALUE) AS a1,  0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'  AND ftr_fac_cd    IN ('JODDESP','NOADESP','KBDESP','KIMDESP')  AND ftr_par_cd    IN ('FINORTOTSK','NOADESP_09','KBIM_D09','KIM_P64')  AND ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL) )"
    Dim qryABPScheduledispatchSOTSBSL As String = ""
    Dim qryABPScheduledispatchFOTSBSL As String = ""
    Dim qryABPScheduledispatchSOSis As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  and ftr_par_cd    IN ('SIZORTOTML','SOTOTSIL','NOADESP_19','NOADESP_05','NOADESP_16','KBIM_D03','KBIM_D03','KBIM_D23')  AND ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)   UNION    SELECT SUM(FTR_VALUE) AS a1,  0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'  AND ftr_fac_cd    IN ('JODDESP','NOADESP','KBDESP','KIMDESP')  AND ftr_par_cd    IN ('SIZORTOTML','SOTOTSIL','NOADESP_19','NOADESP_16','NOADESP_05','KBIM_D03','KBIM_D03','KBIM_D23')  AND ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL) )"
    Dim qryABPScheduledispatchFOSis As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP' AND ftr_par_cd    IN ('FINORTOTML','NOADESP_06','KBIM_D11','NOADESP_17')  AND ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)   UNION    SELECT SUM(FTR_VALUE) AS a1,  0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'  AND ftr_fac_cd    IN ('JODDESP','NOADESP','KBDESP','KIMDESP')  AND ftr_par_cd    IN ('FINORTOTML','NOADESP_06','KBIM_D11','NOADESP_17')  AND ftr_year_month= (select to_char(sysdate,'YYYYMM') FROM DUAL)  )"
    Dim qrymonthlyABPScheduleNIMSO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('NOADESP')  And    ftr_par_cd   IN ('NOADESP_08','NOADESP_01','NOADESP_13','NOADESP_19','NOADESP_05')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION     Select SUM(FTR_VALUE) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'    And ftr_fac_cd   In ('NOADESP')  And ftr_par_cd   IN    ('NOADESP_08','NOADESP_01','NOADESP_13','NOADESP_19','NOADESP_05')     And ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qrymonthlyABPScheduleNIMFO As String = " SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('NOADESP')  And ftr_par_cd   IN ('NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17') " +
                                                  " And ftr_year_month= (select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_fac_cd   In ('NOADESP')  And ftr_par_cd   IN ( " +
                                                  " 'NOADESP_09','NOADESP_03','NOADESP_14','NOADESP_06','NOADESP_17')   And ftr_year_month= (select to_char(sysdate,'YYYYMM') FROM DUAL) )"
    Dim qrymonthlyABPScheduleJEIMSO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP')  And    ftr_par_cd   IN ('SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL) UNION     Select SUM(FTR_VALUE) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'    And ftr_fac_cd   In ('JODDESP')  And ftr_par_cd   IN    ('SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML')     And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qrymonthlyABPScheduleJEIMFO As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP')  And ftr_par_cd   IN ('FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML') " +
                                                  " And ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_fac_cd   In ('JODDESP')  And ftr_par_cd   IN ('FINORTOTSK','JEFOTOTSJ','FINTOBSL','FINORTOTML')   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qrymonthlyABPScheduleKIMSO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('KIMDESP')  And    ftr_par_cd   IN ('KIM_P57','KIM_P61')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION     Select SUM(FTR_VALUE) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'    And ftr_fac_cd   In ('KIMDESP')  And ftr_par_cd   IN    ('KIM_P57','KIM_P61')     And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qrymonthlyABPScheduleKIMFO As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('KIMDESP')  And ftr_par_cd   IN ('KIM_P64','KIM_P66','KIM_P55') " +
                                                  " And ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_fac_cd   In ('KIMDESP')  And ftr_par_cd   IN ('KIM_P64','KIM_P66','KIM_P55')   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  )"
    Dim qrymonthlyABPScheduleKBIMSO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('KBDESP')  And    ftr_par_cd   IN ('KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION     Select SUM(FTR_VALUE) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE'    And ftr_fac_cd   In ('KBDESP')  And ftr_par_cd   IN    ('KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')     And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim qrymonthlyABPScheduleKBIMFO As String = "SELECT SUM(a1) AS SCHEDULE,  SUM(b) As ABP FROM  (SELECT 0        AS a1,    SUM(FTR_VALUE) AS b  FROM t_facility_target " +
                                                  " WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('KBDESP')  And ftr_par_cd   IN ('KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11') " +
                                                  " And ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION  Select SUM(FTR_VALUE) As a1,   0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                  " And ftr_fac_cd   In ('KBDESP')  And ftr_par_cd   IN ('KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')   And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  )"
    Dim qrydispatchtoTSJSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('JESOTOTSJ','NOADESP_01','KBIM_D13','KBIM_D31') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoTSJSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('JESOTOTSJ','NOADESP_01','KBIM_D13','KBIM_D31') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchtoTSJFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('JEFOTOTSJ','NOADESP_03','KBIM_D15') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoTSJFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('JEFOTOTSJ','NOADESP_03','KBIM_D15') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE LIKE '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchtoTSKSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('SIZORTOTSK','NOADESP_08','KBIM_D07') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoTSKSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('SIZORTOTSK','NOADESP_08','KBIM_D07') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchtoTSKFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('FINORTOTSK','NOADESP_09','KBIM_D09','KIM_P64') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoTSKFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('FINORTOTSK','NOADESP_09','KBIM_D09','KIM_P64') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchtoTSBSLSO As String = ""
    Dim qrydispatchtoTSBSLFO As String = ""
    Dim qrydispatchtoSisterConcernSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('SIZORTOTML','SOTOTSIL','NOADESP_05','NOADESP_19','NOADESP_16','KBIM_D05','KBIM_D03','KBIM_D23') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoSisterConcernSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('SIZORTOTML','SOTOTSIL','NOADESP_19','NOADESP_05','NOADESP_16','KBIM_D05','KBIM_D03','KBIM_D23') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchtoSisterConcernFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('FINORTOTML','KBIM_D11','NOADESP_06','NOADESP_17') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlydispatchtoSisterConcernFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('FINORTOTML','KBIM_D11','NOADESP_06','NOADESP_17') AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchfromNoamundiSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_08','NOADESP_01','NOADESP_13','NOADESP_19','NOADESP_16','NOADESP_05')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlydispatchfromNoamundiSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_08','NOADESP_01','NOADESP_13','NOADESP_16','NOADESP_19','NOADESP_05')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchfromNoamundiFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_03','NOADESP_06','NOADESP_09','NOADESP_14')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlydispatchfromNoamundiFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_03','NOADESP_06','NOADESP_09','NOADESP_14')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qrydispatchfromJodaSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlydispatchfromJodaSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('SIZORTOTSK','JESOTOTSJ','SOTOBSL','SOTOTSIL','SIZORTOTML')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchfromJodaFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('FINORTOTML','FINORTOTSK', " +
                                                  " 'JEFOTOTSJ','KBFOTOTSJ','FINTOBSL')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlydispatchfromJodaFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('FINORTOTML','FINORTOTSK', " +
                                                  " 'JEFOTOTSJ','KBFOTOTSJ','FINTOBSL')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrydispatchfromKhondbondSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlydispatchfromKhondbondSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KBIM_D07','KBIM_D13','KBIM_D17','KBIM_D21','KBIM_D03','KBIM_D05','KBIM_D23','KBIM_D31')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "

    Dim qrydispatchfromKhondbondFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlydispatchfromKhondbondFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD In ('KBIM_D09','KBIM_D15','KBIM_D19','KBIM_D11')" +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qrydispatchfromkatamatiSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KIM_P57','KIM_P61')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlydispatchfromkatamatiSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KIM_P57','KIM_P61')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qrydispatchfromkatamatiFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KIM_P64','KIM_P66') " +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlydispatchfromkatamatiFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('KIM_P64','KIM_P66') " +
                                               " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qryabpomqRom As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_01','NIM_TNG_04','JEIM_MN_01','JEIM_MN_02','JEIM_MN_04','KIM_P01','KIM_P04','KIM_P011','KBIM_M06','KBIM_M07') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryabpomqOB As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_13','KIM_P11','JEIM_MN_06','KBIM_M10','JEIM_MN_08') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryabptotalexcavation As String = "select NVL(SUM(FTR_VALUE),0) from T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_08','NIM_TNG_09','JEIM_MN_01','JEIM_MN_02','KTM_TNG_01','KTM_TNG_02','KIM_P17','NIM_TNG_12','NIM_TNG_11','NIM_TNG_10','KIM_TNG_05','KIM_TNG_05') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryactualRom As String = "select NVL(SUM(FCP_VALUE),0) from T_facility_perf where fcp_par_cd in ('NIM_TNG_01','NIM_TNG_04','JEIM_MN_01','JEIM_MN_02','JEIM_MN_04','KIM_P01','KIM_P04','KIM_P011','KBIM_M09','KBIM_M07') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualOB As String = "select NVL(SUM(FCP_VALUE),0) from T_facility_perf where fcp_par_cd in ('NIM_TNG_13','KIM_P11','JEIM_MN_06','KBIM_M10','JEIM_MN_08') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualexacavation As String = " Select NVL(SUM(FCP_VALUE),0) from T_FACILITY_perf WHERE Fcp_PAR_CD In ('NIM_TNG_08','NIM_TNG_09','JEIM_MN_01','JEIM_MN_02','KTM_TNG_01','KTM_TNG_02','KIM_P17','NIM_TNG_12','NIM_TNG_11','NIM_TNG_10','KIM_TNG_05','KIM_TNG_05') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryabpmonthlyROM As String = "select SUM(FTR_VALUE) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_01','NIM_TNG_04','JEIM_MN_01','JEIM_MN_02','JEIM_MN_04','KIM_P01','KIM_P04','KIM_P011','KBIM_M06','KBIM_M07') and FTR_YEAR_MONTH=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryabpmonthlyOB As String = " select SUM(FTR_VALUE) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_13','KIM_P11','JEIM_MN_06','KBIM_M10','JEIM_MN_08') and FTR_YEAR_MONTH=(select to_char(sysdate,'YYYYMM') FROM DUAL) "
    Dim qryabpabpmonthlyexcavation As String = "select NVL(SUM(FTR_VALUE),0) from T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NIM_TNG_08','NIM_TNG_09','JEIM_MN_01','JEIM_MN_02','KTM_TNG_01','KTM_TNG_02','KIM_P17','NIM_TNG_12','NIM_TNG_11','NIM_TNG_10','KIM_TNG_05','KIM_TNG_05') and FTR_YEAR_MONTH=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryactualmonthlyROM As String = "select NVL(sum(fcp_value),0) from t_facility_perf where fcp_par_cd in ('NIM_TNG_01','NIM_TNG_04','JEIM_MN_01','KIM_P011','JEIM_MN_02','JEIM_MN_04','KIM_P01','KIM_P04','KBIM_M09','KBIM_M07') and FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim qryactualmonthlyOB As String = "select NVL(sum(fcp_value),0) from t_facility_perf where fcp_par_cd in ('NIM_TNG_13','KIM_P11','JEIM_MN_06','KBIM_M10','JEIM_MN_08') and FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qryactualmonthlyexacavation As String = "Select NVL(sum(fcp_value),0) from T_FACILITY_perf WHERE Fcp_PAR_CD In ('NIM_TNG_08','NIM_TNG_09','JEIM_MN_01','JEIM_MN_02','KTM_TNG_01','KTM_TNG_02','KIM_P17','NIM_TNG_12','NIM_TNG_11','NIM_TNG_10','KIM_TNG_05','KIM_TNG_05') and FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrysoabpproduction As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_32','WPJ_WP_08','WPJ_WP_11','KBIM_P06','KBIM_P07','KBIM_P13') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qryfoabpproduction As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('NDCMP_44','NDCMP_33','KIM_P50','WPJ_WP_09','FO_to_os','PRD_DPC-4','KBIM_P15','KBIM_P08','KBIM_P14') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrysoactualproduction As String = "select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_32','WPJ_WP_08','WPJ_WP_11','KBIM_P06','KBIM_P07','KBIM_P13') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryfoactualprodution As String = "select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_44','NDCMP_33','KIM_P50','WPJ_WP_09','FO_to_os','PRD_DPC-4','KBIM_P15','KBIM_P08','KBIM_P14') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrysomonthlyabpso As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_32','WPJ_WP_08','WPJ_WP_11','KBIM_P06','KBIM_P07','KBIM_P13') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryfomonthlyabpfo As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('NDCMP_44','NDCMP_33','KIM_P50','WPJ_WP_09','FO_to_os','KBIM_P08','KBIM_P14') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrysoactualmonthly As String = "select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_32','WPJ_WP_08','WPJ_WP_11','KBIM_P06','KBIM_P07','KBIM_P13') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qryfoactualmonthly As String = "select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_44','NDCMP_33','KIM_P50','WPJ_WP_09','FO_to_os','PRD_DPC-4','KBIM_P15','KBIM_P08','KBIM_P14') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrynoaabpSOprod As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_32') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrynoaactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_32') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlynoaactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_32') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrynoaabpSOmonthly As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_32') and ftr_rec_type='ABP'"
    Dim qrynoaactualSOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_32') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrynoaabpFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('NDCMP_44','NDCMP_33') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrynoaactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_44','NDCMP_33') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlynoaactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_44','NDCMP_33') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrynoaabpFOmonthly As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('NDCMP_44','NDCMP_33') and ftr_rec_type='ABP'"
    Dim qrynoaactualFOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_44','NDCMP_33')"
    Dim qryjodabpSOprod As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('WPJ_WP_08','WPJ_WP_11') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_08','WPJ_WP_11') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryjodaownactualmonthlyso As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_08','WPJ_WP_11') and fcp_date LIKE '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qryjodabpSOmonthly As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('WPJ_WP_08','WPJ_WP_11')"
    Dim qryjodactualSOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_08','WPJ_WP_11')"
    Dim qryjodabpFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('WPJ_WP_09','FO_to_os','PRD_DPC-4') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_09','FO_to_os','PRD_DPC-4') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlyjodactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_09','FO_to_os','PRD_DPC-4') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qryjodabpFOmonthly As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('WPJ_WP_09','FO_to_os','PRD_DPC-4') and ftr_rec_type='ABP' AND ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodactualFOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_09','FO_to_os','PRD_DPC-4')"
    Dim qrykimabpSOprod As String = "select SUM(FTR_VALUE) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_3211') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykimactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_3211') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qrymonthlyactualkimmactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_3211') and fcp_date LIKE '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrykimabpSOmonthly As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('NDCMP_3211') and ftr_rec_type='ABP' and ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrykimactualSOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('NDCMP_3211')"
    Dim qrykimabpFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KIM_P50') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP' "
    Dim qrykimactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KIM_P50') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim qryactualmonthlykimactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KIM_P50') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrykimabpFOmonthly As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KIM_P50') and ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykimactualFOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KIM_P50')"
    Dim qrykbimabpSOprod As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_P06','KBIM_P07','KBIM_P13') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P06','KBIM_P07','KBIM_P13') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P06','KBIM_P07','KBIM_P13') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrykbimabpSOmonthly As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_P06','KBIM_P07','KBIM_P13') and ftr_rec_type='ABP' and ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrykbimactualSOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P06','KBIM_P07','KBIM_P13')"
    Dim qrykbimabpFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_P08','KBIM_P14','KBIM_P15') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P08','KBIM_P14','KBIM_P15') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P08','KBIM_P14','KBIM_P15') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    Dim qrykbimabpFOmonthly As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_P08','KBIM_P14','KBIM_P15') and ftr_rec_type='ABP' and ftr_year_month=(select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qrykbimactualFOmonthly As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P08','KBIM_P14','KBIM_P15')"
    Dim bhushanscheduleabpdispatchSO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    NVL(SUM(FTR_VALUE),0) AS b  FROM t_facility_target " +
                                                "  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And " +
                                                "  ftr_par_cd   IN ('NOADESP_13','SOTOBSL','KBIM_D17','KBIM_D21')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL)  UNION " +
                                                "  Select NVL(SUM(FTR_VALUE),0) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                "  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And ftr_par_cd   IN " +
                                                "  ('NOADESP_13','KBIM_D17','KBIM_D21') " +
                                                "  And ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim bhushanscheduleabpdispatchFO As String = "SELECT SUM(a1) As SCHEDULE,  SUM(b)       AS ABP FROM  (SELECT 0        AS a1,    NVL(SUM(FTR_VALUE),0) AS b  FROM t_facility_target " +
                                                "  WHERE ftr_rec_type ='ABP'  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And " +
                                                "  ftr_par_cd   IN ('NOADESP_14','KBIM_D19','KIM_P66','FINTOBSL')  And ftr_year_month =(select to_char(sysdate,'YYYYMM') FROM DUAL) UNION " +
                                                "  Select NVL(SUM(FTR_VALUE),0) As a1,    0                   AS b  FROM t_facility_target  WHERE ftr_rec_type ='SCHEDULE' " +
                                                "  And ftr_fac_cd   In ('JODDESP','NOADESP','KBDESP','KIMDESP')  And ftr_par_cd   IN " +
                                                "  ('NOADESP_14','KBIM_D19','KIM_P66','FINTOBSL') " +
                                                "  And ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL))"
    Dim bhushanactualdispatchSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_13','SOTOBSL','KBIM_D17','KBIM_D21')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim bhushanactualdispatchFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_14','KBIM_D19','KIM_P66','FINTOBSL')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI') "
    Dim bhushanactualmonthlydispatchSO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_13','SOTOBSL','KBIM_D17','KBIM_D21')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    Dim bhushanactualmonthlydispatchFO As String = "Select NVL(SUM(FCP_VALUE),0) FROM T_FACILITY_PERF WHERE FCP_PAR_CD IN ('NOADESP_14','KBIM_D19','KIM_P66','FINTOBSL')" +
                                                   " AND FCP_SHIFT IN ('A','B','C') AND FCP_DATE like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%' "
    'Query for Own plant Joda for so
    Dim qryjodaabpownplantso As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('WPJ_WP_08') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodaownactualso As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_08') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryjodaownactulamonthlyso As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_08') and fcp_date LIKE '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Out source plant joda for so
    Dim qryjodaabpoutplantso As String = "select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('WPJ_WP_11') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodaoutactualso As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_11') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryjodaoutactulamonthlyso As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_11') and fcp_date LIKE '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Own plant Joda for fo
    Dim qryjodabpownplantfo As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('WPJ_WP_09','PRD_DPC-4') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodaownctualfo As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_09','PRD_DPC-4') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryjodaownactualmonthlyfo As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('WPJ_WP_09','PRD_DPC-4') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Out source plant Joda for fo
    Dim qryjodaabpoutplantfo As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('FO_to_os') and ftr_rec_type='ABP' and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL)"
    Dim qryjodaoutactualfo As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('FO_to_os') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryjodaoutactulamonthlyfo As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('FO_to_os') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Own plant khondbond for so
    Dim qrykbimabpownSOprod As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_P06','KBIM_P07') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualownSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P06','KBIM_P07') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualownSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P06','KBIM_P07') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Out plant khondbond for so
    Dim qrykbimabpoutSOprod As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_P13') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualoutSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P13') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualoutSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P13') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"


    'Query for Own plant khondbond for fo
    Dim qrykbimabpownFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_P08') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimownactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P08') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimownactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P08') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"


    'Query for Out plant khondbond for fo
    Dim qrykbimabpoutFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_P14','KBIM_P15') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimoutactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P14','KBIM_P15') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimoutactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_P14','KBIM_P15') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    ''' <summary>
    ''' ''new addition Nihar
    ''' </summary>

    'Query for Wet plant khondbond for so
    Dim qrykbimabpWPSOprod As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_WPSO') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualWPSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_WPSO') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualWPSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_WPSO') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"

    'Query for Screening plant khondbond for so
    Dim qrykbimabpSPSOprod As String = " select NVL(SUM(FTR_VALUE),0) FROM T_FACILITY_TARGET WHERE FTR_PAR_CD IN ('KBIM_SPSO') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimactualSPSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_SPSO') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimactualSPSOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_SPSO') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"


    'Query for Wet plant khondbond for fo
    Dim qrykbimabpWPFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_WPFO') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimWPactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_WPFO') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimWPactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_WPFO') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"


    'Query for Screening plant khondbond for fo
    Dim qrykbimabpSPFOprod As String = "select NVL(SUM(FTR_VALUE),0) from T_facility_target where FTR_PAR_CD IN ('KBIM_SPFO') and ftr_year_month = (select to_char(sysdate,'YYYYMM') FROM DUAL) and ftr_rec_type='ABP'"
    Dim qrykbimSPactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_SPFO') and FCP_DATE=TRUNC(SYSDATE-1) and FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI')"
    Dim qryactualmonthlykbimSPactualFOprod As String = "select NVL(SUM(Fcp_VALUE),0) FROM T_FACILITY_perf WHERE Fcp_PAR_CD IN ('KBIM_SPFO') and fcp_date like '%'||(select to_char(sysdate,'MON-YY') FROM DUAL)||'%'"
    ''' <summary>
    ''' end new addition
    ''' </summary>

    'Query to 
    Dim qryRecipientlist As String = "select RecipientList,TypeID from T_AutoMessages where TypeID in ('3','4');"
    Dim datamissingfacility As String = " SELECT DISTINCT(fcp_fac_cd) FROM  (SELECT *   FROM t_facility_perf  WHERE fcp_fac_cd   In ('DPJ','JEIM','KIMPLANT','NOADESP','KBDESP','NDCMP','JODDESP','KIMDESP','NIM','KBIM','KIM','DPK','WPJ') " +
                                         " And FCP_DATE      =TO_CHAR(SYSDATE-1,'dd-mon-yy') " +
                                         " And FCP_CRT_TS<= to_date(to_char(sysdate,'dd-mon-yy')||' 09:31','DD-MON-YY HH:MI'))"

#End Region


    Dim query As String
    Dim month As String = Now.Date.Month
    'Dim month As String = "5"
    Dim year As String = Now.Date.Year
    Dim month_day As Integer = System.DateTime.DaysInMonth(year, month)
    'Dim month_day As Integer = 30
    'Dim month_day As Integer = 31
    ' Dim dttt1 As String = Date.Now.Date
    Dim dttt1 As String = Convert.ToDateTime(Date.Now.Date).ToString("dd-MMM-yyyy").ToString

    Dim dttt2 As String = Convert.ToDateTime(Date.Now.Date.AddDays(-1)).ToString("dd-MMM-yyyy").ToString
    Dim str() As String = dttt1.Split("-")
    Dim till_day As String = str(0).ToString - 1
    ' Dim till_day As String = "31"
    Sub Main()
        ' Dim total_days As String = month_day
        '  month_day = month_day * 1000

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        ' xlApp = New Excel.ApplicationClass
        Try
            Dim total_days As String = month_day
            month_day = month_day * 1000
            Dim reportpath As String = ".\ "
            Dim parentdirectory As String = Directory.GetParent(Directory.GetParent(Directory.GetParent(reportpath).FullName).FullName).FullName
            Dim templatepath As String = parentdirectory + "\Report Formats\Reports1.xlsx"
            Dim newfilepathexcel As String = parentdirectory + "\Report Formats\Daily Report.xlsx"
            'Dim templatepath As String = "D:\Report Formats\Daily Report.xlsx"
            'Dim newfilepathexcel As String = "D:\Report Formats\Daily Report.xlsx"



            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkBook = xlApp.Workbooks.Open(Path.Combine(templatepath), 3, False)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
#Region "Code"
            xlWorkSheet.Range("B2:B36").NumberFormat = "0.0"
            xlWorkSheet.Range("C2:C36").NumberFormat = "0.0"
            xlWorkSheet.Range("D2:D36").NumberFormat = "0.0"
            xlWorkSheet.Range("E2:E36").NumberFormat = "0.0"
            xlWorkSheet.Range("F2:F36").NumberFormat = "0.0"
            xlWorkSheet.Range("G2:G36").NumberFormat = "0.0"
            xlWorkSheet.Range("H2:H36").NumberFormat = "0.0"
            xlWorkSheet.Range("I2:I36").NumberFormat = "0.0"
            xlWorkSheet.Range("J2:J36").NumberFormat = "0.0"
            xlWorkSheet.Range("K2:K36").NumberFormat = "0.0"
            xlWorkSheet.Range("L2:L36").NumberFormat = "0.0"
            xlWorkSheet.Range("M2:M36").NumberFormat = "0.0"
            xlWorkSheet.Range("N2:N36").NumberFormat = "0.0"
            xlWorkSheet.Range("O2:O36").NumberFormat = "0.0"
            xlWorkSheet.Range("P2:P36").NumberFormat = "0.0"
            xlWorkSheet.Range("Q2:Q36").NumberFormat = "0.0"
            xlWorkSheet.Range("R2:R39").NumberFormat = "0.0"
            xlWorkSheet.Range("S2:S39").NumberFormat = "0.0"

            xlWorkSheet.Range("T2:T39").NumberFormat = "0.0"
            xlWorkSheet.Range("U2:U36").NumberFormat = "0.0"
            Dim query As String = qrryrlyScheduleABPDispatchSOomq
            '   Query = Query.Replace(CATUSERID, Session.Item(CUSERID))
            Dim dt As New Data.DataSet
            'DISPATCH SO OMQ
            xlWorkSheet.Cells(1, 23) = total_days
            xlWorkSheet.Cells(1, 22) = till_day
            get_datatable(query, dt)

            If dt.Tables(0).Rows.Count > 0 Then

                xlWorkSheet.Cells(4, 2) = (Integer.Parse((dt.Tables(0).Rows(0)(0).ToString())) / month_day).ToString
                xlWorkSheet.Cells(4, 3) = (Integer.Parse(dt.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                '  xlWorkSheet.Cells(4, 7) = (Integer.Parse(xlWorkSheet.Cells(4, 2)) * till_day)
                xlWorkSheet.Cells(4, 7) = String.Concat("=B4*", till_day)

                xlWorkSheet.Cells(4, 8) = String.Concat("=C4*", till_day)
            Else
            End If
            Dim dt1 As New Data.DataSet
            'ACTUAL DESPATCH SO OMQ
            query = qrymonthlyactualdispatchSOOMQ
            get_datatable(query, dt1)
            If (dt1.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt1.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(4, 4) = str ' (Integer.Parse(dt1.Rows(0)(0).ToString)).ToString
                'xlWorkSheet.Cells(4, 3) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                ''  xlWorkSheet.Cells(4, 7) = (Integer.Parse(xlWorkSheet.Cells(4, 2)) * till_day)
                'xlWorkSheet.Cells(4, 7) = String.Concat("=B4*", till_day)

                'xlWorkSheet.Cells(4, 8) = String.Concat("=C4*", till_day)
                xlWorkSheet.Cells(4, 5) = "=D4/C4%"

                xlWorkSheet.Cells(4, 6) = "=D4/B4%"



            End If
            Dim dt01 As New Data.DataSet
            query = qrymonthlyactualdispatchSOOMQ1
            get_datatable(query, dt01)
            If (dt01.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt01.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(4, 9) = str  '(Integer.Parse(dt3.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(4, 10) = "=I4/G4%"
                xlWorkSheet.Cells(4, 11) = "=I4/H4%"
                xlWorkSheet.Cells(4, 12) = "=((B4*W1)-I4)/(W1-V1)"
                xlWorkSheet.Cells(4, 13) = "= ((C4 * W1) - I4) / (W1 - V1)"
                'xlWorkSheet.Cells(5, 5) = "=D5/C5%"
                'xlWorkSheet.Cells(6, 4) = "=D4+D5"
                'xlWorkSheet.Cells(6, 6) = "=E4+E5"
                'xlWorkSheet.Cells(6, 5) = "=E4+E5"
                'xlWorkSheet.Cells(5, 6) = "=D5/B5%"
                'xlWorkSheet.Cells(6, 6) = "=F4+F5"
                'xlWorkSheet.Cells(5, 12) = "=((B5*V1)-I5)/(V1-W1)"
                'xlWorkSheet.Cells(5, 13) = "=((C5*V1)-I5)/(V1-W1)"
                'xlWorkSheet.Cells(6, 12) = "=L4+L5"
                'xlWorkSheet.Cells(6, 13) = "=M4+M5"

            Else

            End If
            'DISPATCH FO OMQ
            Dim dt2 As New Data.DataSet
            query = qryyrlyScheduleABPDispatchFOOMQ
            get_datatable(query, dt2)
            If (dt2.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(5, 2) = (Integer.Parse(dt2.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(5, 3) = (Integer.Parse(dt2.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(5, 7) = String.Concat("=B5*", till_day)
                xlWorkSheet.Cells(5, 8) = String.Concat("=C5*", till_day)
                'xlWorkSheet.Cells(6, 2) = xlWorkSheet.Cells(4, 2) + xlWorkSheet.Cells(5, 2)
                xlWorkSheet.Cells(6, 2) = "=B4+B5"
                xlWorkSheet.Cells(6, 3) = "=C4+C5"
                xlWorkSheet.Cells(6, 7) = "=G4+G5"
                xlWorkSheet.Cells(6, 8) = "=H4+H5"
                'xlWorkSheet.Cells(6, 3) = xlWorkSheet.Cells(4, 3) + xlWorkSheet.Cells(5, 3)
                'xlWorkSheet.Cells(6, 7) = xlWorkSheet.Cells(4, 7) + xlWorkSheet.Cells(5, 7)
                'xlWorkSheet.Cells(6, 8) = xlWorkSheet.Cells(4, 8) + xlWorkSheet.Cells(5, 8)
                'xlWorkSheet.Cells(5, 10) = "=((B4*V1)-I4)/(V1-W1))"
                'xlWorkSheet.Cells(5, 11) = "=((B4*V1)-I4)/(V1-W1))"

            Else

            End If

            'Actaul Dispatch FO OMQ
            Dim dt3 As New Data.DataSet
            query = qrymonthlyactualdispatchFOOMQ
            get_datatable(query, dt3)
            If (dt3.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt3.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(5, 4) = str / 1000 '(Integer.Parse(dt3.Rows(0)(0).ToString) / month_day).ToString

                xlWorkSheet.Cells(5, 5) = "=D5/C5%"
                xlWorkSheet.Cells(6, 4) = "=D4+D5"
                xlWorkSheet.Cells(6, 6) = "=E4+E5"
                xlWorkSheet.Cells(6, 5) = "=D6/C6%"
                xlWorkSheet.Cells(5, 6) = "=D5/B5%"
                xlWorkSheet.Cells(6, 6) = "=D6/B6%"
                xlWorkSheet.Cells(5, 12) = "=((B5*W1)-I5)/(W1-V1)"
                xlWorkSheet.Cells(5, 13) = "=((C5*W1)-I5)/(W1-V1)"
                xlWorkSheet.Cells(6, 12) = "=((B6*W1)-I6)/(W1-V1)"
                xlWorkSheet.Cells(6, 13) = "=((C6*W1)-I6)/(W1-V1)"

            Else

            End If
            'actual monthly dispatch
            Dim dt_act_so_monthly As New Data.DataSet
            query = qrymonthlyactualdispatchFOOMQ1
            get_datatable(query, dt_act_so_monthly)
            If (dt_act_so_monthly.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt_act_so_monthly.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(5, 9) = str   '(Integer.Parse(dt3.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(6, 9) = "=I4+I5"
                'xlWorkSheet.Cells(5, 5) = "=D5/C5%"
                'xlWorkSheet.Cells(6, 4) = "=D4+D5"
                'xlWorkSheet.Cells(6, 6) = "=E4+E5"
                'xlWorkSheet.Cells(6, 5) = "=E4+E5"
                'xlWorkSheet.Cells(5, 6) = "=D5/B5%"
                'xlWorkSheet.Cells(6, 6) = "=F4+F5"
                'xlWorkSheet.Cells(5, 12) = "=((B5*V1)-I5)/(V1-W1)"
                'xlWorkSheet.Cells(5, 13) = "=((C5*V1)-I5)/(V1-W1)"
                'xlWorkSheet.Cells(6, 12) = "=L4+L5"
                'xlWorkSheet.Cells(6, 13) = "=M4+M5"
                xlWorkSheet.Cells(5, 10) = "=I5/G5%"
                xlWorkSheet.Cells(5, 11) = "=I5/H5%"
                xlWorkSheet.Cells(6, 10) = "=I6/G6%"
                xlWorkSheet.Cells(6, 11) = "=I6/H6%"
            Else

            End If
            'DISPATCH SO/FO TSJ
            Dim dt4 As New Data.DataSet
            query = qryABPScheduledispatchSOTSJ
            get_datatable(query, dt4)
            If (dt4.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(8, 2) = (Integer.Parse(dt4.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(8, 3) = (Integer.Parse(dt4.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(8, 7) = String.Concat("=B8*", till_day)
                xlWorkSheet.Cells(8, 8) = String.Concat("=C8*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString
                xlWorkSheet.Cells(8, 4) = ""

            End If

            'Dispatch SO/FO TSJ
            Dim dt5 As New Data.DataSet
            query = qryABPScheduledispatchFOTSJ
            get_datatable(query, dt5)
            If (dt5.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 2) = (Integer.Parse(dt5.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(9, 3) = (Integer.Parse(dt5.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(9, 7) = String.Concat("=B9*", till_day)
                xlWorkSheet.Cells(9, 8) = String.Concat("=C9*", till_day)


            End If
            'dispatch actual SO/FO TSJ
            Dim dt6 As New Data.DataSet
            query = qrydispatchtoTSJSO
            get_datatable(query, dt6)
            If (dt6.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt6.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(8, 4) = str
                xlWorkSheet.Cells(8, 5) = "=D8/C8%"
                xlWorkSheet.Cells(8, 6) = "=D8/B8%"

                xlWorkSheet.Cells(8, 12) = "=((B8*W1)-I8)/(W1-V1)"
                xlWorkSheet.Cells(8, 13) = "=((C8*W1)-I8)/(W1-V1)"
            End If
            Dim dt_act_tsj_so As New Data.DataSet
            query = qrymonthlydispatchtoTSJSO
            get_datatable(query, dt_act_tsj_so)
            If (dt_act_tsj_so.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt_act_tsj_so.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(8, 9) = str / 1000 '(Integer.Parse(dt_act_tsj_so.Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(9, 5) = "=D9/C9%"
                'xlWorkSheet.Cells(9, 6) = "=D9/B9%"
                'xlWorkSheet.Cells(10, 5) = "=E9+E10"
                'xlWorkSheet.Cells(10, 6) = "=F9+F10"

                'xlWorkSheet.Cells(9, 12) = "=((B9*V1)-I9)/(V1-W1)"
                'xlWorkSheet.Cells(9, 13) = "=((C9*V1)-I9)/(V1-W1)"
                'xlWorkSheet.Cells(10, 12) = "=L8+L9"
                'xlWorkSheet.Cells(10, 13) = "=M8+M9"
                xlWorkSheet.Cells(8, 10) = "=I8/G8%"
                xlWorkSheet.Cells(8, 11) = "=I8/H8%"

            End If
            Dim dt7 As New Data.DataSet
            query = qrydispatchtoTSJFO
            get_datatable(query, dt7)
            If (dt7.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 4) = dt7.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(9, 5) = "=D9/C9%"
                xlWorkSheet.Cells(9, 6) = "=D9/B9%"

                xlWorkSheet.Cells(9, 12) = "=((B9*W1)-I9)/(W1-V1)"
                xlWorkSheet.Cells(9, 13) = "=((C9*W1)-I9)/(W1-V1)"

            End If
            Dim dt02 As New Data.DataSet
            query = qrymonthlydispatchtoTSJFO
            get_datatable(query, dt02)
            If (dt02.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt02.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(9, 9) = str / 1000 '(Integer.Parse(dt02.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(9, 10) = "=I9/G9%"
                xlWorkSheet.Cells(9, 11) = "=I9/H9%"

                'xlWorkSheet.Cells(9, 5) = "=D9/C9%"
                'xlWorkSheet.Cells(9, 6) = "=D9/B9%"


                xlWorkSheet.Cells(10, 2) = "=B8+B9"
                xlWorkSheet.Cells(10, 3) = "=C8+C9"
                xlWorkSheet.Cells(10, 4) = "=D8+D9"

                xlWorkSheet.Cells(10, 5) = "=D10/C10%"
                xlWorkSheet.Cells(10, 6) = "=D10/B10%"
                xlWorkSheet.Cells(10, 7) = "=G8+G9"
                xlWorkSheet.Cells(10, 8) = "=H8+H9"
                'xlWorkSheet.Cells(9, 12) = "=((B9*V1)-I9)/(V1-W1)"
                'xlWorkSheet.Cells(9, 13) = "=((C9*V1)-I9)/(V1-W1)"
                'xlWorkSheet.Cells(10, 12) = "=L8+L9"
                'xlWorkSheet.Cells(10, 13) = "=M8+M9"
                xlWorkSheet.Cells(10, 9) = "=I8+I9"
                xlWorkSheet.Cells(10, 10) = "=I10/G10%"
                xlWorkSheet.Cells(10, 11) = "=I10/H10%"
                xlWorkSheet.Cells(10, 12) = "=((B10*W1)-I10)/(W1-V1)"
                xlWorkSheet.Cells(10, 13) = "=((C10 * W1) - I10) / (W1 - V1)"
            End If
            'Dispatch SO/FO TSK
            Dim dt8 As New Data.DataSet
            query = qryABPScheduledispatchSOTSK
            get_datatable(query, dt8)
            If (dt8.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(12, 2) = (Integer.Parse(dt8.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(12, 3) = (Integer.Parse(dt8.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(12, 7) = String.Concat("=B12*", till_day)
                xlWorkSheet.Cells(12, 8) = String.Concat("=C12*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString



            End If
            Dim dt9 As New Data.DataSet
            query = qryABPScheduledispatchFOTSK
            get_datatable(query, dt9)
            If (dt9.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(13, 2) = (Integer.Parse(dt9.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(13, 3) = (Integer.Parse(dt9.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(13, 7) = String.Concat("=B13*", till_day)
                xlWorkSheet.Cells(13, 8) = String.Concat("=C13*", till_day)
                xlWorkSheet.Cells(14, 2) = "=B12+B13"
                xlWorkSheet.Cells(14, 3) = "=C12+C13"


            End If
            'dispatch actual SO/FO TSK
            Dim dt10 As New Data.DataSet
            query = qrydispatchtoTSKSO
            get_datatable(query, dt10)
            If (dt10.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(12, 4) = dt10.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(12, 5) = "=D12/C12%"
                xlWorkSheet.Cells(12, 6) = "=D12/B12%"
                xlWorkSheet.Cells(12, 12) = "=((B12*W1)-I12)/(W1-V1)"
                xlWorkSheet.Cells(12, 13) = "=((C12*W1)-I12)/(W1-V1)"

            End If
            Dim dt03 As New Data.DataSet
            query = qrymonthlydispatchtoTSKSO
            get_datatable(query, dt03)
            If (dt03.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt03.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(12, 9) = str / 1000 '(Integer.Parse(dt03.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(12, 10) = "=I12/G12%"
                xlWorkSheet.Cells(12, 11) = "=I12/H12%"
                'xlWorkSheet.Cells(12, 5) = "=D12/C12%"
                'xlWorkSheet.Cells(12, 6) = "=D12/B12%"
                'xlWorkSheet.Cells(12, 12) = "=((B12*V1)-I12)/(V1-W1)"
                'xlWorkSheet.Cells(12, 13) = "=((C12*V1)-I12)/(V1-W1)"

            End If
            Dim dt11 As New Data.DataSet
            query = qrydispatchtoTSKFO
            get_datatable(query, dt11)
            If (dt11.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(13, 4) = dt11.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(13, 5) = "=D13/C13%"
                xlWorkSheet.Cells(13, 6) = "=D13/B13%"
                xlWorkSheet.Cells(13, 12) = "=((B13*W1)-I13)/(W1-V1)"
                xlWorkSheet.Cells(13, 13) = "=((C13*W1)-I13)/(W1-V1)"

            End If
            Dim dt04 As New Data.DataSet
            query = qrymonthlydispatchtoTSKFO
            get_datatable(query, dt04)
            If (dt04.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt04.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(13, 9) = str '(Integer.Parse(dt04.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(13, 10) = "=I13/G13%"
                xlWorkSheet.Cells(13, 11) = "=I13/H13%"

                'xlWorkSheet.Cells(13, 5) = "=D13/C13%"
                'xlWorkSheet.Cells(13, 6) = "=D13/B13%"
                'xlWorkSheet.Cells(13, 12) = "=((B13*V1)-I13)/(V1-W1)"
                'xlWorkSheet.Cells(13, 13) = "=((C13*V1)-I13)/(V1-W1)"
                'xlWorkSheet.Cells(14, 12) = "=L12+L13"
                'xlWorkSheet.Cells(14, 13) = "=M12+M13"
                xlWorkSheet.Cells(14, 4) = "=D12+D13"
                xlWorkSheet.Cells(14, 5) = "=D14/C14%"
                xlWorkSheet.Cells(14, 6) = "=D14/B14%"


                xlWorkSheet.Cells(14, 7) = "=G12+G13"
                xlWorkSheet.Cells(14, 8) = "=H12+H13"


                xlWorkSheet.Cells(14, 9) = "=I12+I13"
                xlWorkSheet.Cells(14, 10) = "=I14/G14%"
                xlWorkSheet.Cells(14, 11) = "=I14/H14%"
                xlWorkSheet.Cells(14, 12) = "=((B14*W1)-I14)/(W1-V1)"
                xlWorkSheet.Cells(14, 13) = "=((C14*W1)-I14)/(W1-V1)"
            End If
            'Dispatch SO/FO Sister Concern
            Dim dt12 As New Data.DataSet
            query = qryABPScheduledispatchSOSis
            get_datatable(query, dt12)
            If (dt12.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(20, 2) = (Integer.Parse(dt12.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(20, 3) = (Integer.Parse(dt12.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(20, 7) = String.Concat("=B20*", till_day)
                xlWorkSheet.Cells(20, 8) = String.Concat("=C20*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString

            End If
            Dim dt13 As New Data.DataSet
            query = qryABPScheduledispatchFOSis
            get_datatable(query, dt13)
            If (dt13.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(21, 2) = (Integer.Parse(dt13.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(21, 3) = (Integer.Parse(dt13.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(21, 7) = String.Concat("=B21*", till_day)
                xlWorkSheet.Cells(21, 8) = String.Concat("=C21*", till_day)
                xlWorkSheet.Cells(22, 2) = "=B20+B21"
                xlWorkSheet.Cells(22, 3) = "=C20+C21"
                xlWorkSheet.Cells(22, 7) = "=G20+G21"
                xlWorkSheet.Cells(22, 8) = "=H20+H21"

            End If
            Dim dt14 As New Data.DataSet
            query = qrydispatchtoSisterConcernSO
            get_datatable(query, dt14)
            If (dt14.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(20, 4) = dt14.Tables(0).Rows(0)(0).ToString / 1000
                'xlWorkSheet.Cells(21, 3) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(21, 7) = String.Concat("=B21*", till_day)
                'xlWorkSheet.Cells(21, 8) = String.Concat("=C21*", till_day)
                'xlWorkSheet.Cells(22, 2) = "=B20+B21"
                'xlWorkSheet.Cells(22, 3) = "=C20+C21"
                'xlWorkSheet.Cells(22, 7) = "=G20+G21"
                'xlWorkSheet.Cells(22, 8) = "=H20+H21"
                xlWorkSheet.Cells(20, 12) = "=((B20*W1)-I20)/(W1-V1)"

                xlWorkSheet.Cells(20, 13) = "=((C20*W1)-I20)/(W1-V1)"
            End If
            Dim dt05 As New Data.DataSet
            query = qrymonthlydispatchtoSisterConcernSO
            get_datatable(query, dt05)
            If (dt05.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt05.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(20, 9) = str '(Integer.Parse(dt05.Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(20, 10) = "=I20/G20%"
                xlWorkSheet.Cells(20, 11) = "=I20/H20%"
                'xlWorkSheet.Cells(21, 3) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(21, 7) = String.Concat("=B21*", till_day)
                'xlWorkSheet.Cells(21, 8) = String.Concat("=C21*", till_day)
                'xlWorkSheet.Cells(22, 2) = "=B20+B21"
                'xlWorkSheet.Cells(22, 3) = "=C20+C21"
                'xlWorkSheet.Cells(22, 7) = "=G20+G21"
                'xlWorkSheet.Cells(22, 8) = "=H20+H21"
                'xlWorkSheet.Cells(20, 12) = "=((B20*V1)-I20)/(V1-W1)"

                'xlWorkSheet.Cells(20, 13) = "=((C20*V1)-I20)/(V1-W1)"
            End If
            Dim dt15 As New Data.DataSet
            query = qrydispatchtoSisterConcernFO
            get_datatable(query, dt15)
            If (dt15.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(21, 4) = dt15.Tables(0).Rows(0)(0).ToString / 1000
                'xlWorkSheet.Cells(21, 3) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(21, 7) = String.Concat("=B21*", till_day)
                'xlWorkSheet.Cells(21, 8) = String.Concat("=C21*", till_day)
                'xlWorkSheet.Cells(22, 2) = "=B20+B21"
                'xlWorkSheet.Cells(22, 3) = "=C20+C21"
                'xlWorkSheet.Cells(22, 7) = "=G20+G21"
                'xlWorkSheet.Cells(22, 8) = "=H20+H21"
                xlWorkSheet.Cells(22, 4) = "=D20+D21"
                xlWorkSheet.Cells(20, 5) = "=D20/C20%"
                xlWorkSheet.Cells(21, 5) = "=D21/C21%"

                xlWorkSheet.Cells(20, 6) = "=D20/B20%"
                xlWorkSheet.Cells(21, 6) = "=D21/B21%"
                xlWorkSheet.Cells(22, 5) = "=D22/C22%"
                xlWorkSheet.Cells(22, 6) = "=D22/B22%"
                xlWorkSheet.Cells(21, 12) = "=((B21*W1)-I21)/(W1-V1)"
                xlWorkSheet.Cells(21, 13) = "=((C21*W1)-I21)/(W1-V1)"


            End If
            Dim dt06 As New Data.DataSet
            query = qrymonthlydispatchtoSisterConcernFO
            get_datatable(query, dt06)
            If (dt06.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt06.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(21, 9) = str '(Integer.Parse(dt06.Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(21, 3) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(21, 7) = String.Concat("=B21*", till_day)
                'xlWorkSheet.Cells(21, 8) = String.Concat("=C21*", till_day)
                'xlWorkSheet.Cells(22, 2) = "=B20+B21"
                'xlWorkSheet.Cells(22, 3) = "=C20+C21"
                'xlWorkSheet.Cells(22, 7) = "=G20+G21"
                'xlWorkSheet.Cells(22, 8) = "=H20+H21"
                xlWorkSheet.Cells(22, 9) = "=I20+I21"
                xlWorkSheet.Cells(21, 10) = "=I21/G21%"
                xlWorkSheet.Cells(21, 11) = "=I21/H21%"
                xlWorkSheet.Cells(22, 10) = "=I22/G22%"
                xlWorkSheet.Cells(22, 11) = "=I22/H22%"
                xlWorkSheet.Cells(22, 12) = "=((B22*W1)-I22)/(W1-V1)"
                xlWorkSheet.Cells(22, 13) = "=((C22*W1)-I22)/(W1-V1)"
            End If
            'dispatch abpschedule noamundi so
            Dim dt16 As New Data.DataSet
            query = qrymonthlyABPScheduleNIMSO
            get_datatable(query, dt16)
            If (dt16.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(25, 3) = (Integer.Parse(dt16.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(25, 4) = (Integer.Parse(dt16.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                ' xlWorkSheet.Cells(26, 7) = String.Concat("=B20*", till_day)
                'xlWorkSheet.Cells(26, 8) = String.Concat("=C20*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString


            End If
            'dispatch abpschedule noamundi fo
            Dim dt17 As New Data.DataSet
            query = qrymonthlyABPScheduleNIMFO
            get_datatable(query, dt17)
            If (dt17.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(26, 3) = (Integer.Parse(dt17.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(26, 4) = (Integer.Parse(dt17.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(27, 3) = String.Concat("=C25+C26")
                xlWorkSheet.Cells(27, 4) = String.Concat("=D25+D26")



            End If
            'dispatch actualdispatch from Noamundi SO
            Dim dt18 As New Data.DataSet
            query = qrydispatchfromNoamundiSO
            get_datatable(query, dt18)
            If (dt18.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt18.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(25, 5) = str
                'xlWorkSheet.Cells(25, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                ' xlWorkSheet.Cells(26, 7) = String.Concat("=B20*", till_day)
                'xlWorkSheet.Cells(26, 8) = String.Concat("=C20*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString
                xlWorkSheet.Cells(25, 6) = "=E25/C25%"
                xlWorkSheet.Cells(25, 7) = "=E25/D25%"


            End If
            'actualdispatch from Noamundi FO
            Dim dt19 As New Data.DataSet
            query = qrydispatchfromNoamundiFO
            get_datatable(query, dt19)
            If (dt19.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt19.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(26, 5) = str
                'xlWorkSheet.Cells(25, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                ' xlWorkSheet.Cells(26, 7) = String.Concat("=B20*", till_day)
                'xlWorkSheet.Cells(26, 8) = String.Concat("=C20*", till_day)
                ' xlWorkSheet.Cells(8, 8) = (Integer.Parse(xlWorkSheet.Cells(8, 3)) * 20).ToString
                xlWorkSheet.Cells(27, 5) = "=E25+E26"
                xlWorkSheet.Cells(26, 6) = "=E26/C26%"
                xlWorkSheet.Cells(26, 7) = "=E26/D26%"
                xlWorkSheet.Cells(27, 6) = "=E27/C27%"
                xlWorkSheet.Cells(27, 7) = "=E27/D27%"


            End If
            Dim temp As String = 0
            Dim dt07 As New Data.DataSet
            query = qryactualmonthlydispatchfromNoamundiSO
            get_datatable(query, dt07)
            If (dt07.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt07.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(25, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C25*V1)%"))
                xlWorkSheet.Cells(25, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D25*V1)%"))
                xlWorkSheet.Cells(90, 8) = actual_monthly.ToString
                xlWorkSheet.Cells(25, 10) = String.Concat(String.Concat("=((C25*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(25, 11) = String.Concat(String.Concat("=((D25*W1)-", actual_monthly), ")/(W1-V1)")

            End If
            Dim dt08 As New Data.DataSet

            query = qryactualmonthlydispatchfromNoamundiFO
            get_datatable(query, dt08)
            If (dt08.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt08.Tables(0).Rows(0)(0).ToString / 1000

                xlWorkSheet.Cells(26, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C26*V1)%"))
                xlWorkSheet.Cells(26, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D26*V1)%"))
                xlWorkSheet.Cells(26, 10) = String.Concat(String.Concat("=((C26*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(26, 11) = String.Concat(String.Concat("=((D26*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 9) = actual_monthly.ToString

                xlWorkSheet.Cells(27, 8) = "=(H90+I90)/(C27*V1)%"
                xlWorkSheet.Cells(27, 9) = "=(H90+I90)/(D27*V1)%"

                xlWorkSheet.Range("I70").EntireRow.Hidden = True
                xlWorkSheet.Cells(27, 10) = "=((C27 * W1) -(H90+I90))/(W1 - V1)"
                xlWorkSheet.Cells(27, 11) = "=((D27*W1) - (H90+I90))/(W1-V1)"

            End If
            'dispatch abpschedule katamati so
            Dim dt20 As New Data.DataSet
            query = qrymonthlyABPScheduleKIMSO
            get_datatable(query, dt20)
            If (dt20.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(28, 3) = (Integer.Parse(dt20.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(28, 4) = (Integer.Parse(dt20.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(27, 2) = String.Concat("=C26+C27")
                'xlWorkSheet.Cells(27, 3) = String.Concat("=D26+D27")



            End If
            'dispatch abpschedule katamati fo
            Dim dt21 As New Data.DataSet
            query = qrymonthlyABPScheduleKIMFO
            get_datatable(query, dt21)
            If (dt.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(29, 3) = (Integer.Parse(dt21.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt21.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                xlWorkSheet.Cells(30, 4) = String.Concat("=D28+D29")



            End If
            'actual dispatch from katamati SO
            Dim dt22 As New Data.DataSet
            query = qrydispatchfromkatamatiSO
            get_datatable(query, dt22)
            If (dt22.Tables(0).Rows.Count > 0) Then
                ' xlWorkSheet.Cells(28, 5) = (Integer.Parse(dt22.Tables(0).Rows(0)(0).ToString) / 1000).ToString
                xlWorkSheet.Cells(28, 5) = 0
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                'xlWorkSheet.Cells(30, 4) = String.Concat("=D28+D29")
                ' xlWorkSheet.Cells(28, 6) = "=E28/C28%"
                xlWorkSheet.Cells(28, 6) = 0
                '  xlWorkSheet.Cells(28, 7) = "=E28/D28%"
                xlWorkSheet.Cells(28, 7) = 0

            End If
            'actual dispatch from katamati FO
            Dim dt23 As New Data.DataSet
            query = qrydispatchfromkatamatiFO
            get_datatable(query, dt23)
            If (dt23.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt23.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(29, 5) = str  '(Integer.Parse(dt23.Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                xlWorkSheet.Cells(30, 5) = String.Concat("=E28+E29")
                xlWorkSheet.Cells(29, 6) = "=E29/C29%"
                xlWorkSheet.Cells(29, 7) = "=E29/D29%"
                xlWorkSheet.Cells(30, 6) = "=E30/C30%"
                xlWorkSheet.Cells(30, 7) = "=E30/D30%"


            End If
            Dim dt09 As New Data.DataSet
            query = qryactualmonthlydispatchfromkatamatiSO
            get_datatable(query, dt09)
            If (dt09.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt09.Tables(0).Rows(0)(0).ToString / 1000

                'xlWorkSheet.Cells(28, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C28*V1)%"))
                'xlWorkSheet.Cells(28, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D28*V1)%"))
                xlWorkSheet.Cells(28, 8) = 0
                xlWorkSheet.Cells(28, 9) = 0

                xlWorkSheet.Cells(28, 10) = String.Concat(String.Concat("=((C28*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(28, 11) = String.Concat(String.Concat("=((D28*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 10) = actual_monthly.ToString
            End If
            Dim dt010 As New Data.DataSet
            query = qryactualmonthlydispatchfromkatamatiFO
            get_datatable(query, dt010)
            If (dt010.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt010.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(29, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C29*V1)%"))
                xlWorkSheet.Cells(29, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D29*V1)%"))
                xlWorkSheet.Cells(29, 10) = String.Concat(String.Concat("=((C29*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(29, 11) = String.Concat(String.Concat("=((D29*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 11) = actual_monthly.ToString

                xlWorkSheet.Cells(30, 8) = "=(J90+K90)/(C30*V1)%"
                xlWorkSheet.Cells(30, 9) = "=(J90+K90)/(D30*V1)%"

                'xlWorkSheet.Range("I38").EntireRow.Hidden = True
                xlWorkSheet.Cells(30, 10) = "=((C30*W1) -(J90+K90))/(W1-V1)"
                xlWorkSheet.Cells(30, 11) = "=((D30*W1) -(J90+K90))/(W1-V1)"
            End If
            'dispatch abpschedule jodaeast SO
            Dim dt24 As New Data.DataSet
            query = qrymonthlyABPScheduleJEIMSO
            get_datatable(query, dt24)
            If (dt.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(31, 3) = (Integer.Parse(dt24.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(31, 4) = (Integer.Parse(dt24.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 2) = String.Concat("=C28+C29")
                'xlWorkSheet.Cells(30, 3) = String.Concat("=D28+D29")



            End If
            'dispatch abpschedule jodaeast FO
            Dim dt25 As New Data.DataSet
            query = qrymonthlyABPScheduleJEIMFO
            get_datatable(query, dt25)
            If (dt25.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(32, 3) = (Integer.Parse(dt25.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(32, 4) = (Integer.Parse(dt25.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(33, 3) = String.Concat("=C31+C32")
                xlWorkSheet.Cells(33, 4) = String.Concat("=D31+D32")



            End If
            'actual dispatch from jODA SO
            Dim dt26 As New Data.DataSet
            query = qrydispatchfromJodaSO
            get_datatable(query, dt26)
            If (dt26.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(31, 5) = dt26.Tables(0).Rows(0)(0).ToString / 1000
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                'xlWorkSheet.Cells(30, 4) = String.Concat("=D28+D29")

                xlWorkSheet.Cells(31, 6) = "=E31/C31%"
                xlWorkSheet.Cells(31, 7) = "=E31/D31%"


            End If
            'actual dispatch from joda FO
            Dim dt27 As New Data.DataSet
            query = qrydispatchfromJodaFO
            get_datatable(query, dt27)
            If (dt27.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(32, 5) = dt27.Tables(0).Rows(0)(0).ToString / 1000
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                xlWorkSheet.Cells(33, 5) = String.Concat("=E31+E32")
                xlWorkSheet.Cells(32, 6) = "=E32/C32%"
                xlWorkSheet.Cells(32, 7) = "=E32/D32%"
                xlWorkSheet.Cells(33, 6) = "=E33/C33%"
                xlWorkSheet.Cells(33, 7) = "=E33/D33%"


            End If
            Dim dt011 As New Data.DataSet
            query = qryactualmonthlydispatchfromJodaSO
            get_datatable(query, dt011)
            If (dt011.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt011.Tables(0).Rows(0)(0).ToString / 1000

                xlWorkSheet.Cells(31, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C31*V1)%"))
                xlWorkSheet.Cells(31, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D31*V1)%"))
                xlWorkSheet.Cells(31, 10) = String.Concat(String.Concat("=((C31*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(31, 11) = String.Concat(String.Concat("=((D31*W1)-", actual_monthly), ")/(W1-V1)")

                xlWorkSheet.Cells(90, 12) = actual_monthly.ToString
            End If
            Dim dt012 As New Data.DataSet
            query = qryactualmonthlydispatchfromJodaFO
            get_datatable(query, dt012)
            If (dt012.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt012.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(32, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C32*V1)%"))
                xlWorkSheet.Cells(32, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D32*V1)%"))
                xlWorkSheet.Cells(32, 10) = String.Concat(String.Concat("=((C32*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(32, 11) = String.Concat(String.Concat("=((D32*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 13) = actual_monthly.ToString

                xlWorkSheet.Cells(33, 8) = "=(L90+M90)/(C33*V1)%"
                xlWorkSheet.Cells(33, 9) = "=(L90+M90)/(D33*V1)%"

                'xlWorkSheet.Range("I38").EntireRow.Hidden = True
                xlWorkSheet.Cells(33, 10) = "=((C33*W1) -(L90+M90))/(W1-V1)"
                xlWorkSheet.Cells(33, 11) = "=((D33*W1) -(L90+M90))/(W1-V1)"
            End If
            'dispatch abpschedule khondbondSO
            Dim dt28 As New Data.DataSet
            query = qrymonthlyABPScheduleKBIMSO
            get_datatable(query, dt28)
            If (dt28.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(34, 3) = (Integer.Parse(dt28.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(34, 4) = (Integer.Parse(dt28.Tables(0).Rows(0)(1).ToString) / month_day).ToString

            End If
            'dispatch abpschedule from khondbond fo
            Dim dt29 As New Data.DataSet
            query = qrymonthlyABPScheduleKBIMFO
            get_datatable(query, dt29)
            If (dt29.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(35, 3) = (Integer.Parse(dt29.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(35, 4) = (Integer.Parse(dt29.Tables(0).Rows(0)(1).ToString) / month_day).ToString
                xlWorkSheet.Cells(36, 3) = String.Concat("=C34+C35")
                xlWorkSheet.Cells(36, 4) = String.Concat("=D34+D35")

            End If
            'actual dispatch from khondbond SO
            Dim dt30 As New Data.DataSet
            query = qrydispatchfromKhondbondSO
            get_datatable(query, dt30)
            If (dt30.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(34, 5) = dt30.Tables(0).Rows(0)(0).ToString / 1000
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                'xlWorkSheet.Cells(30, 4) = String.Concat("=D28+D29")
                xlWorkSheet.Cells(34, 6) = "=E34/C34%"
                xlWorkSheet.Cells(34, 7) = "=E34/D34%"


            End If
            'actual dispatch from khondbond FO
            Dim dt31 As New Data.DataSet
            query = qrydispatchfromKhondbondFO
            get_datatable(query, dt31)
            If (dt31.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt31.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(35, 5) = str  '(Integer.Parse(dt31.Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                xlWorkSheet.Cells(36, 5) = String.Concat("=E34+E35")
                xlWorkSheet.Cells(35, 6) = "=E35/C35%"
                xlWorkSheet.Cells(35, 7) = "=E35/D35%"
                xlWorkSheet.Cells(36, 6) = "=F34+F35"
                xlWorkSheet.Cells(36, 7) = "=G34+G35"


            End If
            Dim dt013 As New Data.DataSet
            query = qryactualmonthlydispatchfromKhondbondSO
            get_datatable(query, dt013)
            If (dt013.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt013.Tables(0).Rows(0)(0).ToString / 1000

                xlWorkSheet.Cells(34, 8) = String.Concat("=", String.Concat(actual_monthly, "/(D35*V1)%"))
                xlWorkSheet.Cells(34, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D34*V1)%"))
                xlWorkSheet.Cells(34, 10) = String.Concat(String.Concat("=((C34*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(34, 11) = String.Concat(String.Concat("=((D34*W1)-", actual_monthly), ")/(W1-V1)")

                xlWorkSheet.Cells(90, 14) = actual_monthly.ToString
            End If
            Dim dt014 As New Data.DataSet
            query = qryactualmonthlydispatchfromKhondbondFO
            get_datatable(query, dt014)
            If (dt014.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt014.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(35, 8) = String.Concat("=", String.Concat(actual_monthly, "/(C35*V1)%"))
                xlWorkSheet.Cells(35, 9) = String.Concat("=", String.Concat(actual_monthly, "/(D35*V1)%"))
                xlWorkSheet.Cells(35, 10) = String.Concat(String.Concat("=((C35*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(35, 11) = String.Concat(String.Concat("=((D35*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 15) = actual_monthly.ToString

                xlWorkSheet.Cells(36, 8) = "=(N90+O90)/(C36*V1)%"
                xlWorkSheet.Cells(36, 9) = "=(N90+O90)/(D36*V1)%"

                'xlWorkSheet.Range("I38").EntireRow.Hidden = True
                xlWorkSheet.Cells(36, 10) = "=((C36*W1) -(N90+O90))/(W1-V1)"
                xlWorkSheet.Cells(36, 11) = "=((D36*W1) -(N90+O90))/(W1-V1)"
            End If
            'add some text

            'Rom and OB and Excavation
            Dim dt32 As New Data.DataSet
            query = qryabpomqRom
            get_datatable(query, dt32)
            If (dt32.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(4, 15) = (Integer.Parse(dt32.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                '  xlWorkSheet.Cells(36, 5) = String.Concat("=E34+E35")



            End If
            Dim dt33 As New Data.DataSet
            query = qryabpomqOB
            get_datatable(query, dt33)
            If (dt33.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(5, 15) = (Integer.Parse(dt33.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                '  xlWorkSheet.Cells(36, 5) = String.Concat("=E34+E35")



            End If
            Dim dt34 As New Data.DataSet
            query = qryabptotalexcavation
            get_datatable(query, dt34)
            If (dt34.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(6, 15) = "=O4+O5" '(Integer.Parse(dt34.Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt35 As New Data.DataSet
            query = qryactualRom
            get_datatable(query, dt35)
            If (dt35.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(4, 16) = dt35.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(4, 17) = "=P4/O4%"
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                '  xlWorkSheet.Cells(36, 5) = String.Concat("=E34+E35")




            End If
            Dim dt36 As New Data.DataSet
            query = qryactualOB
            get_datatable(query, dt36)
            If (dt36.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(5, 16) = dt36.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(5, 17) = "=P5/O5%"
                'xlWorkSheet.Cells(29, 4) = (Integer.Parse(dt.Rows(0)(1).ToString) / month_day).ToString
                'xlWorkSheet.Cells(30, 3) = String.Concat("=C28+C29")
                '  xlWorkSheet.Cells(36, 5) = String.Concat("=E34+E35")



            End If
            Dim dt37 As New Data.DataSet
            query = qryactualexacavation
            get_datatable(query, dt37)
            If (dt37.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(6, 16) = "=P4+P5"
                xlWorkSheet.Cells(6, 17) = "=P6/O6%"
            End If
            Dim dt38 As New Data.DataSet
            query = qryabpmonthlyROM
            get_datatable(query, dt38)
            If (dt38.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(4, 18) = "=O4*V1" '(Integer.Parse(dt38.Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt39 As New Data.DataSet
            query = qryabpmonthlyOB
            get_datatable(query, dt39)
            If (dt39.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(5, 18) = "=O5*V1" '(Integer.Parse(dt39.Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt40 As New Data.DataSet
            query = qryabpabpmonthlyexcavation
            get_datatable(query, dt40)
            If (dt40.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(6, 18) = "=R4+R5" '=O6*V1" '(Integer.Parse(dt40.Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt41 As New Data.DataSet
            query = qryactualmonthlyROM
            get_datatable(query, dt41)
            If (dt41.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt41.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(4, 19) = str
                xlWorkSheet.Cells(4, 20) = "=S4/R4%"
                xlWorkSheet.Cells(4, 21) = "=((O4*W1)-S4)/(W1-V1)"
            End If
            Dim dt42 As New Data.DataSet
            query = qryactualmonthlyOB
            get_datatable(query, dt42)
            If (dt42.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(5, 19) = dt42.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(5, 20) = "=S5/R5%"
                xlWorkSheet.Cells(5, 21) = "=((O5*W1)-S5)/(W1-V1)"
            End If
            Dim dt43 As New Data.DataSet
            query = qryactualmonthlyexacavation
            get_datatable(query, dt43)
            If (dt43.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(6, 19) = "=S4+S5"
                xlWorkSheet.Cells(6, 20) = "=S6/R6%"
                xlWorkSheet.Cells(6, 21) = "=((O6*W1)-S6)/(W1-V1)"
            End If
            'end rom ob
            'start production section
            'so abp
            Dim dt44 As New Data.DataSet
            query = qrysoabpproduction
            get_datatable(query, dt44)
            If (dt44.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(8, 15) = (Integer.Parse(dt44.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            'so actual
            Dim dt45 As New Data.DataSet
            query = qrysoactualproduction
            get_datatable(query, dt45)
            If (dt45.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(8, 16) = dt45.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(8, 17) = "=P8/O8%"
            End If
            'so monthly abp
            Dim dt46 As New Data.DataSet
            query = qrysomonthlyabpso
            get_datatable(query, dt46)
            If (dt46.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(8, 18) = "=O8*V1"
            End If
            'so monthly actual
            Dim dt47 As New Data.DataSet
            query = qrysoactualmonthly
            get_datatable(query, dt47)
            If (dt47.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(8, 19) = dt47.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(8, 20) = "=S8/R8%"
                xlWorkSheet.Cells(8, 21) = "=((O8*W1)-S8)/(W1-V1)"
            End If
            'fo abp
            Dim dt48 As New Data.DataSet
            query = qryfoabpproduction
            get_datatable(query, dt48)
            If (dt48.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 15) = (Integer.Parse(dt48.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            'fo actual
            Dim dt49 As New Data.DataSet
            query = qryfoactualprodution
            get_datatable(query, dt49)
            If (dt49.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 16) = dt49.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(9, 17) = "=P9/O9%"
            End If
            'fo monthly abp
            Dim dt50 As New Data.DataSet
            query = qryfomonthlyabpfo
            get_datatable(query, dt50)
            If (dt50.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 18) = "=O9*V1"
            End If
            'fo monthly actual
            Dim dt51 As New Data.DataSet
            query = qryfoactualmonthly
            get_datatable(query, dt51)
            If (dt51.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(9, 19) = dt51.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(9, 20) = "=S9/R9%"
                xlWorkSheet.Cells(9, 21) = "=((O9*W1)-S9)/(W1-V1)"
                xlWorkSheet.Cells(10, 15) = "=O8+O9"
                xlWorkSheet.Cells(10, 16) = "=P8+P9"
                xlWorkSheet.Cells(10, 18) = "=R8+R9"
                xlWorkSheet.Cells(10, 19) = "=S8+S9"
                xlWorkSheet.Cells(10, 17) = "=P10/O10%"
                xlWorkSheet.Cells(10, 20) = "=S10/R10%"
                xlWorkSheet.Cells(10, 21) = "=((O10*W1)-S10)/(W1-V1)"
            End If
            'Location Wise Plant data SO/FO Noamundi
            Dim dt52 As New Data.DataSet
            query = qrynoaabpSOprod
            get_datatable(query, dt52)
            If (dt52.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(12, 16) = (Integer.Parse(dt52.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt53 As New Data.DataSet
            query = qrynoaactualSOprod
            get_datatable(query, dt53)
            If (dt53.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(12, 17) = dt53.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(12, 18) = "=Q12/P12%"

            End If
            Dim nim_actual_so As String = ""
            Dim dt015 As New Data.DataSet
            query = qryactualmonthlynoaactualSOprod
            get_datatable(query, dt015)
            If (dt015.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt015.Tables(0).Rows(0)(0).ToString / 1000
                nim_actual_so = actual_monthly
                xlWorkSheet.Cells(90, 17) = actual_monthly
                xlWorkSheet.Cells(12, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P12*V1)%"))
                xlWorkSheet.Cells(12, 20) = String.Concat(String.Concat("=((P12*W1)-", actual_monthly), ")/(W1-V1)")
            End If

            'Query = qrynoaabpSOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            'Query = qrynoaactualSOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            Dim dt54 As New Data.DataSet
            query = qrynoaabpFOprod
            get_datatable(query, dt54)
            If (dt54.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(13, 16) = (Integer.Parse(dt54.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(14, 16) = "=P12+P13"
            End If
            Dim dt55 As New Data.DataSet
            query = qrynoaactualFOprod
            get_datatable(query, dt55)
            If (dt55.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(13, 17) = dt55.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(13, 18) = "=Q13/P13%"
                xlWorkSheet.Cells(14, 17) = "=Q12+Q13"
                xlWorkSheet.Cells(14, 18) = "=Q14/P14%"
            End If
            Dim dt016 As New Data.DataSet
            Dim nim_actual_fo As String = ""
            query = qryactualmonthlynoaactualFOprod
            get_datatable(query, dt016)
            If (dt016.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt016.Tables(0).Rows(0)(0).ToString / 1000
                nim_actual_fo = actual_monthly
                xlWorkSheet.Cells(13, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P13*V1)%"))
                xlWorkSheet.Cells(13, 20) = String.Concat(String.Concat("=((P13*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 18) = actual_monthly
                xlWorkSheet.Cells(14, 19) = "=(Q90+R90)/(P14*V1)%"
                xlWorkSheet.Cells(14, 20) = "=((P14 * W1)-(Q90+R90))/(W1-V1)"
            End If
            'Query = qrynoaabpFOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            'Query = qrynoaactualFOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If

            'Katamati SO/FO
            Dim dt56 As New Data.DataSet
            query = qrykimabpSOprod
            get_datatable(query, dt56)
            If (dt56.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(15, 16) = "0" '(Integer.Parse(dt56.Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt57 As New Data.DataSet
            query = qrykimactualSOprod
            get_datatable(query, dt57)
            If (dt57.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(15, 17) = "0" 'dt57.Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(15, 18) = 0 '"=Q15/P15%"
            End If
            Dim dt017 As New Data.DataSet
            Dim kim_actual_so As String = ""
            query = qrymonthlyactualkimmactualSOprod
            get_datatable(query, dt017)
            If (dt017.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt017.Tables(0).Rows(0)(0).ToString / 1000
                kim_actual_so = actual_monthly
                xlWorkSheet.Cells(90, 19) = actual_monthly

                '  xlWorkSheet.Cells(15, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P15*V1)%"))
                xlWorkSheet.Cells(15, 19) = "0"
                xlWorkSheet.Cells(15, 20) = String.Concat(String.Concat("=((P15*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            'Query = qrynoaabpSOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            'Query = qrynoaactualSOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            Dim dt58 As New Data.DataSet
            query = qrykimabpFOprod
            get_datatable(query, dt58)
            If (dt58.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(16, 16) = (Integer.Parse(dt58.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(17, 16) = "=P15+P16"

            End If
            Dim dt59 As New Data.DataSet
            query = qrykimactualFOprod
            get_datatable(query, dt59)
            If (dt59.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(16, 17) = dt59.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(16, 18) = "=Q16/P16%"
                xlWorkSheet.Cells(17, 17) = "=Q15+Q16"
                ' xlWorkSheet.Cells(17, 17) = "=P15+P16"
                xlWorkSheet.Cells(17, 18) = "=Q17/P17%"

            End If
            Dim dt018 As New Data.DataSet
            Dim kim_actual_fo As String = ""
            query = qryactualmonthlykimactualFOprod
            get_datatable(query, dt018)
            If (dt018.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt018.Tables(0).Rows(0)(0).ToString / 1000
                kim_actual_fo = actual_monthly
                xlWorkSheet.Cells(16, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P16*V1)%"))
                xlWorkSheet.Cells(16, 20) = String.Concat(String.Concat("=((P16*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 20) = actual_monthly
                xlWorkSheet.Cells(17, 19) = "=(S90+T90)/(P17*V1)%"
                xlWorkSheet.Cells(17, 20) = "=((P17 * W1)-(S90+T90))/(W1 - V1)"
            End If
            'Query = qrynoaabpFOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If
            'Query = qrynoaactualFOmonthly
            'get_datatable(Query, dt)
            'If (dt.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(9, 19) = (Integer.Parse(dt.Rows(0)(0).ToString) / month_day).ToString
            'End If

            'NOamundi and Katamati Together Performance
            xlWorkSheet.Cells(18, 16) = "=P12+P15"
            xlWorkSheet.Cells(18, 17) = "=Q12+Q15"
            xlWorkSheet.Cells(19, 16) = "=P13+P16"
            xlWorkSheet.Cells(19, 17) = "=Q13+Q16"
            xlWorkSheet.Cells(18, 18) = "=Q18/P18%"
            xlWorkSheet.Cells(18, 19) = String.Concat("=", String.Concat((nim_actual_so + kim_actual_so), "/(P18*V1)%"))
            xlWorkSheet.Cells(19, 18) = "=Q19/P19%"
            xlWorkSheet.Cells(50, 4) = nim_actual_so
            xlWorkSheet.Cells(50, 5) = kim_actual_so
            xlWorkSheet.Cells(50, 2) = nim_actual_fo
            xlWorkSheet.Cells(50, 3) = kim_actual_fo
            xlWorkSheet.Cells(19, 19) = "=(B50+C50)/(P19*V1)%"
            xlWorkSheet.Cells(18, 20) = String.Concat(String.Concat("=((P18*W1)-", (nim_actual_so + kim_actual_so)), ")/(W1-V1)")
            xlWorkSheet.Cells(19, 20) = "=((P19*W1)-(B50+C50))/(W1-V1)"
            xlWorkSheet.Cells(20, 16) = "=P18+P19"
            xlWorkSheet.Cells(20, 17) = "=Q18+Q19"
            xlWorkSheet.Cells(20, 18) = "=Q20/P20%"
            xlWorkSheet.Cells(20, 19) = "=(B50+C50+D50+E50)/(P20*V1)%"
            xlWorkSheet.Cells(20, 20) = "=((P20*W1)-(D50+E50+B50+C50))/(W1-V1)"
            'Joda SO/FO
            'Dim dt60 As New Data.DataSet
            'query = qryjodabpSOprod
            'get_datatable(query, dt60)
            'If (dt60.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(21, 16) = (Integer.Parse(dt60.Rows(0)(0).ToString) / month_day).ToString
            'End If
            'Dim dt61 As New Data.DataSet
            'query = qryjodactualSOprod
            'get_datatable(query, dt61)
            'If (dt61.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(21, 17) = dt61.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(21, 18) = "=Q21/P21%"
            'End If
            'Dim dt019 As New Data.DataSet
            'query = qryjodaownactualmonthlyso
            'get_datatable(query, dt019)
            'If (dt019.Rows.Count > 0) Then
            '    Dim actual_monthly = dt019.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(38, 21) = actual_monthly
            '    xlWorkSheet.Cells(21, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P21*V1)%"))
            '    xlWorkSheet.Cells(21, 20) = String.Concat(String.Concat("=(P21*W1)-", actual_monthly), "/(W1-V1)")
            'End If
            'Dim dt62 As New Data.DataSet
            'query = qryjodabpFOprod
            'get_datatable(query, dt62)
            'If (dt62.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(22, 16) = (Integer.Parse(dt62.Rows(0)(0).ToString) / month_day).ToString
            '    xlWorkSheet.Cells(23, 16) = "=P21+P22"
            'End If
            'Dim dt63 As New Data.DataSet
            'query = qryjodactualFOprod
            'get_datatable(query, dt63)
            'If (dt63.Rows.Count > 0) Then
            '    Dim str As String = dt63.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(22, 17) = str
            '    xlWorkSheet.Cells(22, 18) = "=Q22/P22%"
            '    xlWorkSheet.Cells(23, 17) = "=Q21+Q22"
            '    '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
            '    xlWorkSheet.Cells(23, 18) = "=Q23/P23%"

            'End If
            'Dim dt020 As New Data.DataSet
            'query = qryactualmonthlyjodactualFOprod
            'get_datatable(query, dt020)
            'If (dt020.Rows.Count > 0) Then
            '    Dim actual_monthly = dt020.Rows(0)(0).ToString / 1000

            '    xlWorkSheet.Cells(22, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P22*V1)%"))
            '    xlWorkSheet.Cells(22, 20) = String.Concat(String.Concat("=(P22*W1)-", actual_monthly), "/(W1-V1)")
            '    xlWorkSheet.Cells(38, 22) = actual_monthly
            '    xlWorkSheet.Cells(23, 19) = "=(U38+V38)/(P23*V1)%"
            '    xlWorkSheet.Cells(23, 20) = "=((P23 * W1)-(U38+V38))/(W1 - V1)"

            'End If
            ''khondbond so/fo
            'Dim dt64 As New Data.DataSet
            'query = qrykbimabpSOprod
            'get_datatable(query, dt64)
            'If (dt64.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(24, 16) = (Integer.Parse(dt64.Rows(0)(0).ToString) / month_day).ToString

            'End If
            'Dim dt65 As New Data.DataSet
            'query = qrykbimactualSOprod
            'get_datatable(query, dt65)
            'If (dt65.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(24, 17) = dt65.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(24, 18) = "=Q24/P24%"
            'End If
            'Dim dt021 As New Data.DataSet
            'query = qryactualmonthlykbimactualSOprod
            'get_datatable(query, dt021)
            'If (dt021.Rows.Count > 0) Then
            '    Dim actual_monthly = dt021.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(38, 23) = actual_monthly

            '    xlWorkSheet.Cells(24, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P24*V1)%"))
            '    xlWorkSheet.Cells(24, 20) = String.Concat(String.Concat("=(P24*W1)-", actual_monthly), "/(W1-V1)")

            'End If
            'Dim dt66 As New Data.DataSet
            'query = qrykbimabpFOprod
            'get_datatable(query, dt66)
            'If (dt66.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(25, 16) = (Integer.Parse(dt66.Rows(0)(0).ToString) / month_day).ToString
            '    xlWorkSheet.Cells(26, 16) = "=P24+P25"
            'End If
            'Dim dt67 As New Data.DataSet
            'query = qrykbimactualFOprod
            'get_datatable(query, dt67)
            'If (dt67.Rows.Count > 0) Then
            '    xlWorkSheet.Cells(25, 17) = dt67.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(25, 18) = "=Q25/P25%"
            '    xlWorkSheet.Cells(26, 17) = "=Q25+Q25"
            '    xlWorkSheet.Cells(26, 18) = "=Q26/P26%"
            'End If
            'Dim dt022 As New Data.DataSet
            'query = qryactualmonthlykbimactualFOprod
            'get_datatable(query, dt022)
            'If (dt022.Rows.Count > 0) Then
            '    Dim actual_monthly = dt022.Rows(0)(0).ToString / 1000
            '    xlWorkSheet.Cells(25, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P25*V1)%"))
            '    xlWorkSheet.Cells(25, 20) = String.Concat(String.Concat("=(P25*W1)-", actual_monthly), "/(W1-V1)")
            '    xlWorkSheet.Cells(38, 24) = actual_monthly
            '    xlWorkSheet.Cells(26, 19) = "=(W38+X38)/(P26*V1)%"

            '    xlWorkSheet.Cells(26, 20) = "=((P26 * W1)-(W38+X38))/(W1 - V1)"
            'End If


            'OWN Plant Joda...........................
            'Joda SO/FO
            Dim dt60 As New Data.DataSet
            query = qryjodaabpownplantso
            get_datatable(query, dt60)
            If (dt60.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(21, 16) = (Integer.Parse(dt60.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt61 As New Data.DataSet
            query = qryjodaownactualso
            get_datatable(query, dt61)
            If (dt61.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(21, 17) = dt61.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(21, 18) = "=Q21/P21%"
            End If
            Dim dt019 As New Data.DataSet
            Dim jodaown_actual_so As String = ""
            query = qryjodaownactualmonthlyso
            get_datatable(query, dt019)
            If (dt019.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt019.Tables(0).Rows(0)(0).ToString / 1000
                jodaown_actual_so = actual_monthly
                xlWorkSheet.Cells(91, 21) = actual_monthly
                xlWorkSheet.Cells(21, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P21*V1)%"))
                xlWorkSheet.Cells(21, 20) = String.Concat(String.Concat("=((P21*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            Dim dt62 As New Data.DataSet
            query = qryjodabpownplantfo
            get_datatable(query, dt62)
            If (dt62.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(22, 16) = (Integer.Parse(dt62.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(23, 16) = "=P21+P22"
            End If
            Dim dt63 As New Data.DataSet
            query = qryjodaownctualfo
            get_datatable(query, dt63)
            If (dt63.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt63.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(22, 17) = str
                xlWorkSheet.Cells(22, 18) = "=Q22/P22%"
                xlWorkSheet.Cells(23, 17) = "=Q21+Q22"
                '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
                xlWorkSheet.Cells(23, 18) = "=Q23/P23%"

            End If
            Dim dt020 As New Data.DataSet
            Dim jodaown_actual_fo As String = ""
            query = qryjodaownactualmonthlyfo
            get_datatable(query, dt020)
            If (dt020.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt020.Tables(0).Rows(0)(0).ToString / 1000
                jodaown_actual_fo = actual_monthly
                xlWorkSheet.Cells(22, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P22*V1)%"))
                xlWorkSheet.Cells(22, 20) = String.Concat(String.Concat("=((P22*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(91, 22) = actual_monthly
                xlWorkSheet.Cells(23, 19) = "=(U91+V91)/(P23*V1)%"
                xlWorkSheet.Cells(23, 20) = "=((P23 * W1)-(U91+V91))/(W1 - V1)"

            End If

            'Close of Own Plant Joda


            'Outsource Plant Joda...........................
            'Joda SO/FO
            Dim dt006 As New Data.DataSet
            query = qryjodaabpoutplantso
            get_datatable(query, dt006)
            If (dt006.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(24, 16) = (Integer.Parse(dt006.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt007 As New Data.DataSet
            query = qryjodaoutactualso
            get_datatable(query, dt007)
            If (dt007.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(24, 17) = dt007.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(24, 18) = "=Q24/P24%"
            End If
            Dim dt008 As New Data.DataSet
            Dim jodaout_actual_so As String = ""
            query = qryjodaoutactulamonthlyso
            get_datatable(query, dt008)
            If (dt008.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt008.Tables(0).Rows(0)(0).ToString / 1000
                jodaout_actual_so = actual_monthly
                xlWorkSheet.Cells(92, 21) = actual_monthly
                xlWorkSheet.Cells(24, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P24*V1)%"))
                xlWorkSheet.Cells(24, 20) = String.Concat(String.Concat("=((P24*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            Dim dt009 As New Data.DataSet
            query = qryjodaabpoutplantfo
            get_datatable(query, dt009)
            If (dt009.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(25, 16) = (Integer.Parse(dt009.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(26, 16) = "=P24+P25"
            End If
            Dim dt0010 As New Data.DataSet
            query = qryjodaoutactualfo
            get_datatable(query, dt0010)
            If (dt0010.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt0010.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(25, 17) = str
                xlWorkSheet.Cells(25, 18) = "=Q25/P25%"
                xlWorkSheet.Cells(26, 17) = "=Q24+Q25"
                '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
                xlWorkSheet.Cells(26, 18) = "=Q26/P26%"

            End If
            Dim dt0011 As New Data.DataSet
            Dim jodaout_actual_fo As String = ""
            query = qryjodaoutactulamonthlyfo
            get_datatable(query, dt0011)
            If (dt0011.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt0011.Tables(0).Rows(0)(0).ToString / 1000
                jodaout_actual_fo = actual_monthly
                xlWorkSheet.Cells(25, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P25*V1)%"))
                xlWorkSheet.Cells(25, 20) = String.Concat(String.Concat("=((P25*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(92, 22) = actual_monthly
                xlWorkSheet.Cells(26, 19) = "=(U92+V92)/(P26*V1)%"
                xlWorkSheet.Cells(26, 20) = "=((P26 * W1)-(U92+V92))/(W1 - V1)"

            End If

            'Close of Out Plant Joda

            'Joda Own and out Together Performance
            xlWorkSheet.Cells(27, 16) = "=P21+P24"
            xlWorkSheet.Cells(27, 17) = "=Q21+Q24"
            xlWorkSheet.Cells(28, 16) = "=P22+P25"
            xlWorkSheet.Cells(28, 17) = "=Q22+Q25"
            xlWorkSheet.Cells(27, 18) = "=Q27/P27%"
            xlWorkSheet.Cells(88, 2) = jodaown_actual_so
            xlWorkSheet.Cells(88, 3) = jodaout_actual_so
            xlWorkSheet.Cells(27, 19) = "=(B88+C88)/(P27*V1)%"
            xlWorkSheet.Cells(28, 18) = "=Q28/P28%"
            xlWorkSheet.Cells(91, 2) = jodaown_actual_fo
            xlWorkSheet.Cells(91, 3) = jodaout_actual_fo
            xlWorkSheet.Cells(28, 19) = "=(B91+C91)/(P28*V1)%"
            xlWorkSheet.Cells(27, 20) = "=((P27*W1)-(B88+C88))/(W1-V1)"
            xlWorkSheet.Cells(28, 20) = "=((P28*W1)-(B91+C91))/(W1-V1)"
            xlWorkSheet.Cells(29, 16) = "=P27+P28"
            xlWorkSheet.Cells(29, 17) = "=Q27+Q28"
            xlWorkSheet.Cells(29, 18) = "=Q29/P29%"
            xlWorkSheet.Cells(29, 19) = "=(B88+C88+B91+C91)/(P29*V1)%"
            xlWorkSheet.Cells(29, 20) = "=((P29*W1)-(B88+C88+B91+C91))/(W1-V1)"

            'OWN Plant Khondbond...........................
            'Khondbond SO/FO

            Dim dt64 As New Data.DataSet
            query = qrykbimabpownSOprod
            get_datatable(query, dt64)
            If (dt64.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(30, 16) = (Integer.Parse(dt64.Tables(0).Rows(0)(0).ToString) / month_day).ToString

            End If
            Dim dt65 As New Data.DataSet
            query = qrykbimactualownSOprod
            get_datatable(query, dt65)
            If (dt65.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(30, 17) = dt65.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(30, 18) = "=Q30/P30%"
            End If
            Dim dt021 As New Data.DataSet
            Dim kbimown_actual_so As String = ""
            query = qryactualmonthlykbimactualownSOprod
            get_datatable(query, dt021)
            If (dt021.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly As String = dt021.Tables(0).Rows(0)(0).ToString / 1000
                kbimown_actual_so = actual_monthly
                xlWorkSheet.Cells(90, 23) = actual_monthly

                xlWorkSheet.Cells(30, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P30*V1)%"))
                xlWorkSheet.Cells(30, 20) = String.Concat(String.Concat("=((P30*W1)-", actual_monthly), ")/(W1-V1)")

            End If
            Dim dt66 As New Data.DataSet
            query = qrykbimabpownFOprod
            get_datatable(query, dt66)
            If (dt66.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(31, 16) = (Integer.Parse(dt66.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(32, 16) = "=P30+P31"
            End If
            Dim dt67 As New Data.DataSet
            query = qrykbimownactualFOprod
            get_datatable(query, dt67)
            If (dt67.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(31, 17) = dt67.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(31, 18) = "=Q31/P31%"
                xlWorkSheet.Cells(32, 17) = "=Q30+Q31"
                xlWorkSheet.Cells(32, 18) = "=Q32/P32%"
            End If
            Dim dt022 As New Data.DataSet
            query = qryactualmonthlykbimownactualFOprod
            Dim kbimown_actual_fo As String = ""
            get_datatable(query, dt022)
            If (dt022.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt022.Tables(0).Rows(0)(0).ToString / 1000
                kbimown_actual_fo = actual_monthly
                xlWorkSheet.Cells(31, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P31*V1)%"))
                xlWorkSheet.Cells(31, 20) = String.Concat(String.Concat("=((P31*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(90, 24) = actual_monthly
                xlWorkSheet.Cells(32, 19) = "=(W90+X90)/(P32*V1)%"

                xlWorkSheet.Cells(32, 20) = "=((P32 * W1)-(W90+X90))/(W1 - V1)"
            End If

            'Close of Own Plant Khondbond


            'Outsource Plant Khondbond...........................
            'Khondbond SO/FO
            Dim dt0012 As New Data.DataSet
            query = qrykbimabpoutSOprod
            get_datatable(query, dt0012)
            If (dt0012.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(33, 16) = (Integer.Parse(dt0012.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt0013 As New Data.DataSet
            query = qrykbimactualoutSOprod
            get_datatable(query, dt0013)
            If (dt0013.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(33, 17) = dt0013.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(33, 18) = "=Q33/P33%"
            End If
            Dim dt0014 As New Data.DataSet
            Dim kbimout_actual_so As String = ""
            query = qryactualmonthlykbimactualoutSOprod
            get_datatable(query, dt0014)
            If (dt0014.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0014.Tables(0).Rows(0)(0).ToString / 1000
                kbimout_actual_so = actual_monthly
                xlWorkSheet.Cells(93, 21) = actual_monthly
                xlWorkSheet.Cells(33, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P33*V1)%"))
                xlWorkSheet.Cells(33, 20) = String.Concat(String.Concat("=((P33*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            Dim dt0015 As New Data.DataSet
            query = qrykbimabpoutFOprod
            get_datatable(query, dt0015)
            If (dt0015.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(34, 16) = (Integer.Parse(dt0015.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(35, 16) = "=P33+P34"
            End If
            Dim dt0016 As New Data.DataSet
            query = qrykbimoutactualFOprod
            get_datatable(query, dt0016)
            If (dt0016.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt0016.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(34, 17) = str
                xlWorkSheet.Cells(34, 18) = "=Q34/P34%"
                xlWorkSheet.Cells(35, 17) = "=Q33+Q34"
                '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
                xlWorkSheet.Cells(35, 18) = "=Q35/P35%"

            End If
            Dim dt0017 As New Data.DataSet
            query = qryactualmonthlykbimoutactualFOprod
            Dim kbimout_actual_fo As String = ""
            get_datatable(query, dt0017)
            If (dt0017.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0017.Tables(0).Rows(0)(0).ToString / 1000
                kbimout_actual_fo = actual_monthly
                xlWorkSheet.Cells(34, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P34*V1)%"))
                xlWorkSheet.Cells(34, 20) = String.Concat(String.Concat("=((P34*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(93, 22) = actual_monthly
                xlWorkSheet.Cells(35, 19) = "=(U93+V93)/(P35*V1)%"
                xlWorkSheet.Cells(35, 20) = "=((P35 * W1)-(U93+V93))/(W1 - V1)"

            End If

            '''added by nihar for SO/FO of Wet Plant'''
            '''

            Dim dt0056 As New Data.DataSet
            query = qrykbimabpWPSOprod
            get_datatable(query, dt0056)
            If (dt0056.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(36, 16) = (Integer.Parse(dt0056.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt0057 As New Data.DataSet
            query = qrykbimactualWPSOprod
            get_datatable(query, dt0057)
            If (dt0057.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(36, 17) = dt0057.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(36, 18) = "=Q36/P36%"
            End If
            Dim dt0058 As New Data.DataSet
            Dim kbimWP_actual_so As String = ""
            query = qryactualmonthlykbimactualWPSOprod
            get_datatable(query, dt0058)
            If (dt0058.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0058.Tables(0).Rows(0)(0).ToString / 1000
                kbimWP_actual_so = actual_monthly
                xlWorkSheet.Cells(96, 21) = actual_monthly
                xlWorkSheet.Cells(36, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P36*V1)%"))
                xlWorkSheet.Cells(36, 20) = String.Concat(String.Concat("=((P36*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            Dim dt0059 As New Data.DataSet
            query = qrykbimabpWPFOprod
            get_datatable(query, dt0059)
            If (dt0059.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(37, 16) = (Integer.Parse(dt0059.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(38, 16) = "=P36+P37"
            End If
            Dim dt0060 As New Data.DataSet
            query = qrykbimWPactualFOprod
            get_datatable(query, dt0060)
            If (dt0060.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt0060.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(37, 17) = str
                xlWorkSheet.Cells(37, 18) = "=Q37/P37%"
                xlWorkSheet.Cells(38, 17) = "=Q36+Q37"
                '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
                xlWorkSheet.Cells(38, 18) = "=Q38/P38%"

            End If
            Dim dt0061 As New Data.DataSet
            query = qryactualmonthlykbimWPactualFOprod
            Dim kbimWP_actual_fo As String = ""
            get_datatable(query, dt0061)
            If (dt0061.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0061.Tables(0).Rows(0)(0).ToString / 1000
                kbimWP_actual_fo = actual_monthly
                xlWorkSheet.Cells(37, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P37*V1)%"))
                xlWorkSheet.Cells(37, 20) = String.Concat(String.Concat("=((P37*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(96, 22) = actual_monthly
                xlWorkSheet.Cells(38, 19) = "=(U96+V96)/(P38*V1)%"
                xlWorkSheet.Cells(38, 20) = "=((P38 * W1)-(U96+V96))/(W1 - V1)"

            End If

            Dim dt0062 As New Data.DataSet
            query = qrykbimabpSPSOprod
            get_datatable(query, dt0062)
            If (dt0062.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(39, 16) = (Integer.Parse(dt0062.Tables(0).Rows(0)(0).ToString) / month_day).ToString
            End If
            Dim dt0063 As New Data.DataSet
            query = qrykbimactualSPSOprod
            get_datatable(query, dt0063)
            If (dt0063.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(39, 17) = dt0063.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(39, 18) = "=Q39/P39%"
            End If
            Dim dt0064 As New Data.DataSet
            Dim kbimSP_actual_so As String = ""
            query = qryactualmonthlykbimactualSPSOprod
            get_datatable(query, dt0064)
            If (dt0064.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0064.Tables(0).Rows(0)(0).ToString / 1000
                kbimSP_actual_so = actual_monthly
                xlWorkSheet.Cells(99, 21) = actual_monthly
                xlWorkSheet.Cells(39, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P39*V1)%"))
                xlWorkSheet.Cells(39, 20) = String.Concat(String.Concat("=((P39*W1)-", actual_monthly), ")/(W1-V1)")
            End If
            Dim dt0065 As New Data.DataSet
            query = qrykbimabpSPFOprod
            get_datatable(query, dt0065)
            If (dt0065.Tables(0).Rows.Count > 0) Then
                xlWorkSheet.Cells(40, 16) = (Integer.Parse(dt0065.Tables(0).Rows(0)(0).ToString) / month_day).ToString
                xlWorkSheet.Cells(41, 16) = "=P39+P40"
            End If
            Dim dt0066 As New Data.DataSet
            query = qrykbimSPactualFOprod
            get_datatable(query, dt0066)
            If (dt0066.Tables(0).Rows.Count > 0) Then
                Dim str As String = dt0066.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(40, 17) = str
                xlWorkSheet.Cells(40, 18) = "=Q40/P40%"
                xlWorkSheet.Cells(41, 17) = "=Q39+Q40"
                '  xlWorkSheet.Cells(20, 1) = "=Q12+Q13"
                xlWorkSheet.Cells(41, 18) = "=Q41/P41%"

            End If
            Dim dt0067 As New Data.DataSet
            query = qryactualmonthlykbimSPactualFOprod
            Dim kbimSP_actual_fo As String = ""
            get_datatable(query, dt0067)
            If (dt0067.Tables(0).Rows.Count > 0) Then
                Dim actual_monthly = dt0067.Tables(0).Rows(0)(0).ToString / 1000
                kbimSP_actual_fo = actual_monthly
                xlWorkSheet.Cells(40, 19) = String.Concat("=", String.Concat(actual_monthly, "/(P40*V1)%"))
                xlWorkSheet.Cells(40, 20) = String.Concat(String.Concat("=((P40*W1)-", actual_monthly), ")/(W1-V1)")
                xlWorkSheet.Cells(99, 22) = actual_monthly
                xlWorkSheet.Cells(41, 19) = "=(U99+V99)/(P41*V1)%"
                xlWorkSheet.Cells(41, 20) = "=((P41 * W1)-(U99+V99))/(W1 - V1)"

            End If



            '''




            'Close of Out Plant Khondbond

            'Khondbond Own and out Together Performance
            xlWorkSheet.Cells(42, 16) = "=P30+P33+P36+P39"
            xlWorkSheet.Cells(42, 17) = "=Q30+Q33+Q36+Q39"
            xlWorkSheet.Cells(43, 16) = "=P31+P34+P37+P40"
            xlWorkSheet.Cells(43, 17) = "=Q31+Q34+Q37+Q40"
            xlWorkSheet.Cells(42, 18) = "=Q42/P42%"
            xlWorkSheet.Cells(89, 2) = kbimown_actual_so
            xlWorkSheet.Cells(89, 3) = kbimout_actual_so
            xlWorkSheet.Cells(89, 4) = kbimWP_actual_so
            xlWorkSheet.Cells(89, 5) = kbimSP_actual_so
            xlWorkSheet.Cells(42, 19) = "=(B89+C89+D89+E89)/(P42*V1)%"
            xlWorkSheet.Cells(43, 18) = "=Q43/P43%"
            xlWorkSheet.Cells(92, 2) = kbimown_actual_fo
            xlWorkSheet.Cells(92, 3) = kbimout_actual_fo
            xlWorkSheet.Cells(92, 4) = kbimWP_actual_fo
            xlWorkSheet.Cells(92, 5) = kbimSP_actual_fo
            xlWorkSheet.Cells(43, 19) = "=(B92+C92+D92+E92)/(P43*V1)%"
            xlWorkSheet.Cells(42, 20) = "=((P42*W1)-(B89+C89+D89+E89))/(W1-V1)"
            xlWorkSheet.Cells(43, 20) = "=((P43*W1)-(B92+C92+D92+E92))/(W1-V1)"
            xlWorkSheet.Cells(44, 16) = "=P42+P43"
            xlWorkSheet.Cells(44, 17) = "=Q43+Q43"
            xlWorkSheet.Cells(44, 18) = "=Q44/P44%"
            xlWorkSheet.Cells(44, 19) = "=(B89+C89+D89+E89+B92+C92+D92+E92)/(P44*V1)%"
            xlWorkSheet.Cells(44, 20) = "=((P44*W1)-(B89+C89+D89+E89+B92+C92+D92+E92))/(W1-V1)"
            'Bhushan Steel Dispatch
            Dim dt023 As New Data.DataSet
            query = bhushanscheduleabpdispatchSO
            get_datatable(query, dt023)
            If (dt023.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt023.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(16, 2) = str / month_day
                Dim str1 As String = dt023.Tables(0).Rows(0)(1).ToString
                xlWorkSheet.Cells(16, 3) = str1 / month_day

            End If
            Dim dt024 As New Data.DataSet
            query = bhushanactualdispatchSO
            get_datatable(query, dt024)
            If (dt024.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt024.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(16, 4) = str
                xlWorkSheet.Cells(16, 5) = "=D16/C16%"
                xlWorkSheet.Cells(16, 6) = "=D16/B16%"
                xlWorkSheet.Cells(16, 7) = "=B16*V1"
                xlWorkSheet.Cells(16, 8) = "=C16*V1"

            End If
            Dim dt028 As New Data.DataSet
            query = bhushanactualmonthlydispatchSO
            get_datatable(query, dt028)
            If (dt028.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt028.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(16, 9) = str
                xlWorkSheet.Cells(16, 10) = "=I16/G16%"
                xlWorkSheet.Cells(16, 11) = "=I16/H16%"
                xlWorkSheet.Cells(16, 12) = "=((B16*W1)-I16)/(W1-V1)"
                xlWorkSheet.Cells(16, 13) = "=((C16*W1)-I16)/(W1-V1)"
            End If
            Dim dt025 As New Data.DataSet
            query = bhushanscheduleabpdispatchFO
            get_datatable(query, dt025)
            If (dt025.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt025.Tables(0).Rows(0)(0).ToString
                xlWorkSheet.Cells(17, 2) = str / month_day
                Dim str1 As String = dt025.Tables(0).Rows(0)(1).ToString
                xlWorkSheet.Cells(17, 3) = str1 / month_day
                xlWorkSheet.Cells(18, 2) = "=B16+B17"
                xlWorkSheet.Cells(18, 3) = "=C16+C17"
            End If
            Dim dt026 As New Data.DataSet
            query = bhushanactualdispatchFO
            get_datatable(query, dt026)
            If (dt026.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt026.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(17, 4) = str

                xlWorkSheet.Cells(18, 4) = "=D16+D17"
                xlWorkSheet.Cells(18, 5) = "=D18/C18%"
                xlWorkSheet.Cells(18, 6) = "=D18/B18%"
                xlWorkSheet.Cells(17, 5) = "=D17/C17%"
                xlWorkSheet.Cells(17, 6) = "=D17/B17%"
                xlWorkSheet.Cells(17, 7) = "=B17*V1"
                xlWorkSheet.Cells(17, 8) = "=C17*V1"
            End If
            Dim dt029 As New Data.DataSet
            query = bhushanactualmonthlydispatchFO
            get_datatable(query, dt029)
            If (dt029.Tables(0).Rows.Count > 0) Then

                Dim str As String = dt029.Tables(0).Rows(0)(0).ToString / 1000
                xlWorkSheet.Cells(17, 9) = str
                xlWorkSheet.Cells(18, 7) = "=G16+G17"
                xlWorkSheet.Cells(18, 8) = "=H16+H17"
                xlWorkSheet.Cells(18, 9) = "=I16+I17"
                xlWorkSheet.Cells(17, 10) = "=I17/G17%"
                xlWorkSheet.Cells(17, 11) = "=I17/H17%"

                xlWorkSheet.Cells(17, 12) = "=((B17*W1)-I17)/(W1-V1)"
                xlWorkSheet.Cells(17, 13) = "=((C17*W1)-I17)/(W1-V1)"
                xlWorkSheet.Cells(18, 10) = "=I18/G18%"
                xlWorkSheet.Cells(18, 11) = "=I18/H18%"
            End If

            If (xlWorkSheet.Range("E4").Value >= 100) Then
                xlWorkSheet.Range("E4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E4").Value < 100 And xlWorkSheet.Range("E4").Value >= 95) Then
                xlWorkSheet.Range("E4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F4").Value >= 100) Then
                xlWorkSheet.Range("F4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F4").Value < 100 And xlWorkSheet.Range("F4").Value >= 95) Then
                xlWorkSheet.Range("F4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E5").Value >= 100) Then
                xlWorkSheet.Range("E5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E5").Value < 100 And xlWorkSheet.Range("E5").Value >= 95) Then
                xlWorkSheet.Range("E5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F5").Value >= 100) Then
                xlWorkSheet.Range("F5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F5").Value < 100 And xlWorkSheet.Range("F5").Value >= 95) Then
                xlWorkSheet.Range("F5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E6").Value >= 100) Then
                xlWorkSheet.Range("E6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E6").Value < 100 And xlWorkSheet.Range("E6").Value >= 95) Then
                xlWorkSheet.Range("E6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F6").Value >= 100) Then
                xlWorkSheet.Range("F6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F6").Value < 100 And xlWorkSheet.Range("F6").Value >= 95) Then
                xlWorkSheet.Range("F6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J4").Value >= 100) Then
                xlWorkSheet.Range("J4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J4").Value < 100 And xlWorkSheet.Range("J4").Value >= 95) Then
                xlWorkSheet.Range("J4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K4").Value >= 100) Then
                xlWorkSheet.Range("K4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K4").Value < 100 And xlWorkSheet.Range("K4").Value >= 95) Then
                xlWorkSheet.Range("K4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J5").Value >= 100) Then
                xlWorkSheet.Range("J5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J5").Value < 100 And xlWorkSheet.Range("J5").Value >= 95) Then
                xlWorkSheet.Range("J5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K5").Value >= 100) Then
                xlWorkSheet.Range("K5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K5").Value < 100 And xlWorkSheet.Range("K5").Value >= 95) Then
                xlWorkSheet.Range("K5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J6").Value >= 100) Then
                xlWorkSheet.Range("J6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J6").Value < 100 And xlWorkSheet.Range("J6").Value >= 95) Then
                xlWorkSheet.Range("J6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K6").Value >= 100) Then
                xlWorkSheet.Range("K6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K6").Value < 100 And xlWorkSheet.Range("K6").Value >= 95) Then
                xlWorkSheet.Range("K6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E8").Value >= 100) Then
                xlWorkSheet.Range("E8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E8").Value < 100 And xlWorkSheet.Range("E8").Value >= 95) Then
                xlWorkSheet.Range("E8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E9").Value >= 100) Then
                xlWorkSheet.Range("E9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E9").Value < 100 And xlWorkSheet.Range("E9").Value >= 95) Then
                xlWorkSheet.Range("E9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E10").Value >= 100) Then
                xlWorkSheet.Range("E10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E10").Value < 100 And xlWorkSheet.Range("E10").Value >= 95) Then
                xlWorkSheet.Range("E10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F8").Value >= 100) Then
                xlWorkSheet.Range("F8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F8").Value < 100 And xlWorkSheet.Range("F8").Value >= 95) Then
                xlWorkSheet.Range("F8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F9").Value >= 100) Then
                xlWorkSheet.Range("F9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F9").Value < 100 And xlWorkSheet.Range("F9").Value >= 95) Then
                xlWorkSheet.Range("F9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F10").Value >= 100) Then
                xlWorkSheet.Range("F10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F10").Value < 100 And xlWorkSheet.Range("F10").Value >= 95) Then
                xlWorkSheet.Range("F10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J8").Value >= 100) Then
                xlWorkSheet.Range("J8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J8").Value < 100 And xlWorkSheet.Range("J8").Value >= 95) Then
                xlWorkSheet.Range("J8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J9").Value >= 100) Then
                xlWorkSheet.Range("J9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J9").Value < 100 And xlWorkSheet.Range("J9").Value >= 95) Then
                xlWorkSheet.Range("J9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J10").Value >= 100) Then
                xlWorkSheet.Range("J10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J10").Value < 100 And xlWorkSheet.Range("J10").Value >= 95) Then
                xlWorkSheet.Range("J10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K8").Value >= 100) Then
                xlWorkSheet.Range("K8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K8").Value < 100 And xlWorkSheet.Range("K8").Value >= 95) Then
                xlWorkSheet.Range("K8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K9").Value >= 100) Then
                xlWorkSheet.Range("K9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K9").Value < 100 And xlWorkSheet.Range("K9").Value >= 95) Then
                xlWorkSheet.Range("K9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K10").Value >= 100) Then
                xlWorkSheet.Range("K10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K10").Value < 100 And xlWorkSheet.Range("K10").Value >= 95) Then
                xlWorkSheet.Range("K10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E12").Value >= 100) Then
                xlWorkSheet.Range("E12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E12").Value < 100 And xlWorkSheet.Range("E12").Value >= 95) Then
                xlWorkSheet.Range("E12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E13").Value >= 100) Then
                xlWorkSheet.Range("E13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E13").Value < 100 And xlWorkSheet.Range("E13").Value >= 95) Then
                xlWorkSheet.Range("E13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E14").Value >= 100) Then
                xlWorkSheet.Range("E14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E14").Value < 100 And xlWorkSheet.Range("E14").Value >= 95) Then
                xlWorkSheet.Range("E14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F12").Value >= 100) Then
                xlWorkSheet.Range("F12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F12").Value < 100 And xlWorkSheet.Range("F12").Value >= 95) Then
                xlWorkSheet.Range("F12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F13").Value >= 100) Then
                xlWorkSheet.Range("F13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F13").Value < 100 And xlWorkSheet.Range("F13").Value >= 95) Then
                xlWorkSheet.Range("F13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F14").Value >= 100) Then
                xlWorkSheet.Range("F14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F14").Value < 100 And xlWorkSheet.Range("F14").Value >= 95) Then
                xlWorkSheet.Range("F14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J12").Value >= 100) Then
                xlWorkSheet.Range("J12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J12").Value < 100 And xlWorkSheet.Range("J12").Value >= 95) Then
                xlWorkSheet.Range("J12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J13").Value >= 100) Then
                xlWorkSheet.Range("J13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J13").Value < 100 And xlWorkSheet.Range("J13").Value >= 95) Then
                xlWorkSheet.Range("J13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J14").Value >= 100) Then
                xlWorkSheet.Range("J14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J14").Value < 100 And xlWorkSheet.Range("J14").Value >= 95) Then
                xlWorkSheet.Range("J14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K12").Value >= 100) Then
                xlWorkSheet.Range("K12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K12").Value < 100 And xlWorkSheet.Range("K12").Value >= 95) Then
                xlWorkSheet.Range("K12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K13").Value >= 100) Then
                xlWorkSheet.Range("K13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K13").Value < 100 And xlWorkSheet.Range("K13").Value >= 95) Then
                xlWorkSheet.Range("K13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K14").Value >= 100) Then
                xlWorkSheet.Range("K14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K14").Value < 100 And xlWorkSheet.Range("K14").Value >= 95) Then
                xlWorkSheet.Range("K14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E16").Value >= 100) Then
                xlWorkSheet.Range("E16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E16").Value < 100 And xlWorkSheet.Range("E16").Value >= 95) Then
                xlWorkSheet.Range("E16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E17").Value >= 100) Then
                xlWorkSheet.Range("E17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E17").Value < 100 And xlWorkSheet.Range("E17").Value >= 95) Then
                xlWorkSheet.Range("E17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E18").Value >= 100) Then
                xlWorkSheet.Range("E18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E18").Value < 100 And xlWorkSheet.Range("E18").Value >= 95) Then
                xlWorkSheet.Range("E18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F16").Value >= 100) Then
                xlWorkSheet.Range("F16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F16").Value < 100 And xlWorkSheet.Range("F16").Value >= 95) Then
                xlWorkSheet.Range("F16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F17").Value >= 100) Then
                xlWorkSheet.Range("F17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F17").Value < 100 And xlWorkSheet.Range("F17").Value >= 95) Then
                xlWorkSheet.Range("F17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F18").Value >= 100) Then
                xlWorkSheet.Range("F18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F18").Value < 100 And xlWorkSheet.Range("F18").Value >= 95) Then
                xlWorkSheet.Range("F18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J16").Value >= 100) Then
                xlWorkSheet.Range("J16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J16").Value < 100 And xlWorkSheet.Range("J16").Value >= 95) Then
                xlWorkSheet.Range("J16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J17").Value >= 100) Then
                xlWorkSheet.Range("J17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J17").Value < 100 And xlWorkSheet.Range("J17").Value >= 95) Then
                xlWorkSheet.Range("J17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J18").Value >= 100) Then
                xlWorkSheet.Range("J18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J18").Value < 100 And xlWorkSheet.Range("J18").Value >= 95) Then
                xlWorkSheet.Range("J18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K16").Value >= 100) Then
                xlWorkSheet.Range("K16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K16").Value < 100 And xlWorkSheet.Range("K16").Value >= 95) Then
                xlWorkSheet.Range("K16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K17").Value >= 100) Then
                xlWorkSheet.Range("K17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K17").Value < 100 And xlWorkSheet.Range("K17").Value >= 95) Then
                xlWorkSheet.Range("K17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K18").Value >= 100) Then
                xlWorkSheet.Range("K18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K18").Value < 100 And xlWorkSheet.Range("K18").Value >= 95) Then
                xlWorkSheet.Range("K18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E20").Value >= 100) Then
                xlWorkSheet.Range("E20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E20").Value < 100 And xlWorkSheet.Range("E20").Value >= 95) Then
                xlWorkSheet.Range("E20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E21").Value >= 100) Then
                xlWorkSheet.Range("E21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E21").Value < 100 And xlWorkSheet.Range("E21").Value >= 95) Then
                xlWorkSheet.Range("E21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("E22").Value >= 100) Then
                xlWorkSheet.Range("E22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("E22").Value < 100 And xlWorkSheet.Range("E22").Value >= 95) Then
                xlWorkSheet.Range("E22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("E22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F20").Value >= 100) Then
                xlWorkSheet.Range("F20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F20").Value < 100 And xlWorkSheet.Range("F20").Value >= 95) Then
                xlWorkSheet.Range("F20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F21").Value >= 100) Then
                xlWorkSheet.Range("F21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F21").Value < 100 And xlWorkSheet.Range("F21").Value >= 95) Then
                xlWorkSheet.Range("F21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F22").Value >= 100) Then
                xlWorkSheet.Range("F22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F22").Value < 100 And xlWorkSheet.Range("F22").Value >= 95) Then
                xlWorkSheet.Range("F22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J20").Value >= 100) Then
                xlWorkSheet.Range("J20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J20").Value < 100 And xlWorkSheet.Range("J20").Value >= 95) Then
                xlWorkSheet.Range("J20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J21").Value >= 100) Then
                xlWorkSheet.Range("J21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J21").Value < 100 And xlWorkSheet.Range("J21").Value >= 95) Then
                xlWorkSheet.Range("J21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("J22").Value >= 100) Then
                xlWorkSheet.Range("J22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("J22").Value < 100 And xlWorkSheet.Range("J22").Value >= 95) Then
                xlWorkSheet.Range("J22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("J22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K20").Value >= 100) Then
                xlWorkSheet.Range("K20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K20").Value < 100 And xlWorkSheet.Range("K20").Value >= 95) Then
                xlWorkSheet.Range("K20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K21").Value >= 100) Then
                xlWorkSheet.Range("K21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K21").Value < 100 And xlWorkSheet.Range("K21").Value >= 95) Then
                xlWorkSheet.Range("K21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("K22").Value >= 100) Then
                xlWorkSheet.Range("K22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("K22").Value < 100 And xlWorkSheet.Range("K22").Value >= 95) Then
                xlWorkSheet.Range("K22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("K22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F25").Value >= 100) Then
                xlWorkSheet.Range("F25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F25").Value < 100 And xlWorkSheet.Range("F25").Value >= 95) Then
                xlWorkSheet.Range("F25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F26").Value >= 100) Then
                xlWorkSheet.Range("F26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F26").Value < 100 And xlWorkSheet.Range("F26").Value >= 95) Then
                xlWorkSheet.Range("F26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F27").Value >= 100) Then
                xlWorkSheet.Range("F27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F27").Value < 100 And xlWorkSheet.Range("F27").Value >= 95) Then
                xlWorkSheet.Range("F27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G25").Value >= 100) Then
                xlWorkSheet.Range("G25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G25").Value < 100 And xlWorkSheet.Range("G25").Value >= 95) Then
                xlWorkSheet.Range("G25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G26").Value >= 100) Then
                xlWorkSheet.Range("G26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G26").Value < 100 And xlWorkSheet.Range("G26").Value >= 95) Then
                xlWorkSheet.Range("G26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G27").Value >= 100) Then
                xlWorkSheet.Range("G27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G27").Value < 100 And xlWorkSheet.Range("G27").Value >= 95) Then
                xlWorkSheet.Range("G27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H25").Value >= 100) Then
                xlWorkSheet.Range("H25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H25").Value < 100 And xlWorkSheet.Range("H25").Value >= 95) Then
                xlWorkSheet.Range("H25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H26").Value >= 100) Then
                xlWorkSheet.Range("H26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H26").Value < 100 And xlWorkSheet.Range("H26").Value >= 95) Then
                xlWorkSheet.Range("H26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H27").Value >= 100) Then
                xlWorkSheet.Range("H27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H27").Value < 100 And xlWorkSheet.Range("H27").Value >= 95) Then
                xlWorkSheet.Range("H27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I25").Value >= 100) Then
                xlWorkSheet.Range("I25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I25").Value < 100 And xlWorkSheet.Range("I25").Value >= 95) Then
                xlWorkSheet.Range("I25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I26").Value >= 100) Then
                xlWorkSheet.Range("I26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I26").Value < 100 And xlWorkSheet.Range("I26").Value >= 95) Then
                xlWorkSheet.Range("I26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I27").Value >= 100) Then
                xlWorkSheet.Range("I27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I27").Value < 100 And xlWorkSheet.Range("I27").Value >= 95) Then
                xlWorkSheet.Range("I27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F28").Value >= 100) Then
                xlWorkSheet.Range("F28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F28").Value < 100 And xlWorkSheet.Range("F28").Value >= 95) Then
                xlWorkSheet.Range("F28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F29").Value >= 100) Then
                xlWorkSheet.Range("F29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F29").Value < 100 And xlWorkSheet.Range("F29").Value >= 95) Then
                xlWorkSheet.Range("F29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F30").Value >= 100) Then
                xlWorkSheet.Range("F30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F30").Value < 100 And xlWorkSheet.Range("F30").Value >= 95) Then
                xlWorkSheet.Range("F30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G28").Value >= 100) Then
                xlWorkSheet.Range("G28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G28").Value < 100 And xlWorkSheet.Range("G28").Value >= 95) Then
                xlWorkSheet.Range("G28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G29").Value >= 100) Then
                xlWorkSheet.Range("G29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G29").Value < 100 And xlWorkSheet.Range("G29").Value >= 95) Then
                xlWorkSheet.Range("G29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G30").Value >= 100) Then
                xlWorkSheet.Range("G30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G30").Value < 100 And xlWorkSheet.Range("G30").Value >= 95) Then
                xlWorkSheet.Range("G30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H28").Value >= 100) Then
                xlWorkSheet.Range("H28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H28").Value < 100 And xlWorkSheet.Range("H28").Value >= 95) Then
                xlWorkSheet.Range("H28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H29").Value >= 100) Then
                xlWorkSheet.Range("H29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H29").Value < 100 And xlWorkSheet.Range("H29").Value >= 95) Then
                xlWorkSheet.Range("H29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H30").Value >= 100) Then
                xlWorkSheet.Range("H30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H30").Value < 100 And xlWorkSheet.Range("H30").Value >= 95) Then
                xlWorkSheet.Range("H30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I28").Value >= 100) Then
                xlWorkSheet.Range("I28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I28").Value < 100 And xlWorkSheet.Range("I28").Value >= 95) Then
                xlWorkSheet.Range("I28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I29").Value >= 100) Then
                xlWorkSheet.Range("I29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I29").Value < 100 And xlWorkSheet.Range("I29").Value >= 95) Then
                xlWorkSheet.Range("I29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I30").Value >= 100) Then
                xlWorkSheet.Range("I30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I30").Value < 100 And xlWorkSheet.Range("I30").Value >= 95) Then
                xlWorkSheet.Range("I30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F31").Value >= 100) Then
                xlWorkSheet.Range("F31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F31").Value < 100 And xlWorkSheet.Range("F31").Value >= 95) Then
                xlWorkSheet.Range("F31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F32").Value >= 100) Then
                xlWorkSheet.Range("F32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F32").Value < 100 And xlWorkSheet.Range("F32").Value >= 95) Then
                xlWorkSheet.Range("F32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F33").Value >= 100) Then
                xlWorkSheet.Range("F33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F33").Value < 100 And xlWorkSheet.Range("F33").Value >= 95) Then
                xlWorkSheet.Range("F33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F33").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G31").Value >= 100) Then
                xlWorkSheet.Range("G31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G31").Value < 100 And xlWorkSheet.Range("G31").Value >= 95) Then
                xlWorkSheet.Range("G31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G32").Value >= 100) Then
                xlWorkSheet.Range("G32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G32").Value < 100 And xlWorkSheet.Range("G32").Value >= 95) Then
                xlWorkSheet.Range("G32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G33").Value >= 100) Then
                xlWorkSheet.Range("G33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G33").Value < 100 And xlWorkSheet.Range("G33").Value >= 95) Then
                xlWorkSheet.Range("G33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G33").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H31").Value >= 100) Then
                xlWorkSheet.Range("H31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H31").Value < 100 And xlWorkSheet.Range("H31").Value >= 95) Then
                xlWorkSheet.Range("H31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H32").Value >= 100) Then
                xlWorkSheet.Range("H32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H32").Value < 100 And xlWorkSheet.Range("H32").Value >= 95) Then
                xlWorkSheet.Range("H32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H33").Value >= 100) Then
                xlWorkSheet.Range("H33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H33").Value < 100 And xlWorkSheet.Range("H33").Value >= 95) Then
                xlWorkSheet.Range("H33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H33").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I31").Value >= 100) Then
                xlWorkSheet.Range("I31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I31").Value < 100 And xlWorkSheet.Range("I31").Value >= 95) Then
                xlWorkSheet.Range("I31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I32").Value >= 100) Then
                xlWorkSheet.Range("I32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I32").Value < 100 And xlWorkSheet.Range("I32").Value >= 95) Then
                xlWorkSheet.Range("I32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I33").Value >= 100) Then
                xlWorkSheet.Range("I33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I33").Value < 100 And xlWorkSheet.Range("I33").Value >= 95) Then
                xlWorkSheet.Range("I33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I33").Interior.Color = Color.OrangeRed
            End If

            If (xlWorkSheet.Range("F34").Value >= 100) Then
                xlWorkSheet.Range("F34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F34").Value < 100 And xlWorkSheet.Range("F34").Value >= 95) Then
                xlWorkSheet.Range("F34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F34").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F35").Value >= 100) Then
                xlWorkSheet.Range("F35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F35").Value < 100 And xlWorkSheet.Range("F35").Value >= 95) Then
                xlWorkSheet.Range("F35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("F36").Value >= 100) Then
                xlWorkSheet.Range("F36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("F36").Value < 100 And xlWorkSheet.Range("F36").Value >= 95) Then
                xlWorkSheet.Range("F36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("F36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G34").Value >= 100) Then
                xlWorkSheet.Range("G34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G34").Value < 100 And xlWorkSheet.Range("G34").Value >= 95) Then
                xlWorkSheet.Range("G34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G34").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G35").Value >= 100) Then
                xlWorkSheet.Range("G35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G35").Value < 100 And xlWorkSheet.Range("G35").Value >= 95) Then
                xlWorkSheet.Range("G35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("G36").Value >= 100) Then
                xlWorkSheet.Range("G36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("G36").Value < 100 And xlWorkSheet.Range("G36").Value >= 95) Then
                xlWorkSheet.Range("G36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("G36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H34").Value >= 100) Then
                xlWorkSheet.Range("H34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H34").Value < 100 And xlWorkSheet.Range("H34").Value >= 95) Then
                xlWorkSheet.Range("H34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H34").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H35").Value >= 100) Then
                xlWorkSheet.Range("H35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H35").Value < 100 And xlWorkSheet.Range("H35").Value >= 95) Then
                xlWorkSheet.Range("H35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("H36").Value >= 100) Then
                xlWorkSheet.Range("H36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("H36").Value < 100 And xlWorkSheet.Range("H36").Value >= 95) Then
                xlWorkSheet.Range("H36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("H36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I34").Value >= 100) Then
                xlWorkSheet.Range("I34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I34").Value < 100 And xlWorkSheet.Range("I34").Value >= 95) Then
                xlWorkSheet.Range("I34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I34").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I35").Value >= 100) Then
                xlWorkSheet.Range("I35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I35").Value < 100 And xlWorkSheet.Range("I35").Value >= 95) Then
                xlWorkSheet.Range("I35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("I36").Value >= 100) Then
                xlWorkSheet.Range("I36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("I36").Value < 100 And xlWorkSheet.Range("I36").Value >= 95) Then
                xlWorkSheet.Range("I36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("I36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q4").Value >= 100) Then
                xlWorkSheet.Range("Q4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q4").Value < 100 And xlWorkSheet.Range("Q4").Value >= 95) Then
                xlWorkSheet.Range("Q4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q5").Value >= 100) Then
                xlWorkSheet.Range("Q5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q5").Value < 100 And xlWorkSheet.Range("Q5").Value >= 95) Then
                xlWorkSheet.Range("Q5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q6").Value >= 100) Then
                xlWorkSheet.Range("Q6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q6").Value < 100 And xlWorkSheet.Range("Q6").Value >= 95) Then
                xlWorkSheet.Range("Q6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q8").Value >= 100) Then
                xlWorkSheet.Range("Q8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q8").Value < 100 And xlWorkSheet.Range("Q8").Value >= 95) Then
                xlWorkSheet.Range("Q8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q9").Value >= 100) Then
                xlWorkSheet.Range("Q9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q9").Value < 100 And xlWorkSheet.Range("Q9").Value >= 95) Then
                xlWorkSheet.Range("Q9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("Q10").Value >= 100) Then
                xlWorkSheet.Range("Q10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("Q10").Value < 100 And xlWorkSheet.Range("Q10").Value >= 95) Then
                xlWorkSheet.Range("Q10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("Q10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R12").Value >= 100) Then
                xlWorkSheet.Range("R12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R12").Value < 100 And xlWorkSheet.Range("R12").Value >= 95) Then
                xlWorkSheet.Range("R12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R13").Value >= 100) Then
                xlWorkSheet.Range("R13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R13").Value < 100 And xlWorkSheet.Range("R13").Value >= 95) Then
                xlWorkSheet.Range("R13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R14").Value >= 100) Then
                xlWorkSheet.Range("R14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R14").Value < 100 And xlWorkSheet.Range("R14").Value >= 95) Then
                xlWorkSheet.Range("R14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R15").Value >= 100) Then
                xlWorkSheet.Range("R15").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R15").Value < 100 And xlWorkSheet.Range("R15").Value >= 95) Then
                xlWorkSheet.Range("R15").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R15").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R16").Value >= 100) Then
                xlWorkSheet.Range("R16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R16").Value < 100 And xlWorkSheet.Range("R16").Value >= 95) Then
                xlWorkSheet.Range("R16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R17").Value >= 100) Then
                xlWorkSheet.Range("R17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R17").Value < 100 And xlWorkSheet.Range("R17").Value >= 95) Then
                xlWorkSheet.Range("R17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R18").Value >= 100) Then
                xlWorkSheet.Range("R18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R18").Value < 100 And xlWorkSheet.Range("R18").Value >= 95) Then
                xlWorkSheet.Range("R18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R19").Value >= 100) Then
                xlWorkSheet.Range("R19").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R19").Value < 100 And xlWorkSheet.Range("R19").Value >= 95) Then
                xlWorkSheet.Range("R19").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R19").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R20").Value >= 100) Then
                xlWorkSheet.Range("R20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R20").Value < 100 And xlWorkSheet.Range("R20").Value >= 95) Then
                xlWorkSheet.Range("R20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R21").Value >= 100) Then
                xlWorkSheet.Range("R21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R21").Value < 100 And xlWorkSheet.Range("R21").Value >= 95) Then
                xlWorkSheet.Range("R21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R22").Value >= 100) Then
                xlWorkSheet.Range("R22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R22").Value < 100 And xlWorkSheet.Range("R22").Value >= 95) Then
                xlWorkSheet.Range("R22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R23").Value >= 100) Then
                xlWorkSheet.Range("R23").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R23").Value < 100 And xlWorkSheet.Range("R23").Value >= 95) Then
                xlWorkSheet.Range("R23").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R23").Interior.Color = Color.OrangeRed
            End If
            'JHBJHDBKSBJKBSKBD
            If (xlWorkSheet.Range("S12").Value >= 100) Then
                xlWorkSheet.Range("S12").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S12").Value < 100 And xlWorkSheet.Range("S12").Value >= 95) Then
                xlWorkSheet.Range("S12").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S12").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S13").Value >= 100) Then
                xlWorkSheet.Range("S13").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S13").Value < 100 And xlWorkSheet.Range("S13").Value >= 95) Then
                xlWorkSheet.Range("S13").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S13").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S14").Value >= 100) Then
                xlWorkSheet.Range("S14").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S14").Value < 100 And xlWorkSheet.Range("S14").Value >= 95) Then
                xlWorkSheet.Range("S14").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S14").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S15").Value >= 100) Then
                xlWorkSheet.Range("S15").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S15").Value < 100 And xlWorkSheet.Range("S15").Value >= 95) Then
                xlWorkSheet.Range("S15").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S15").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S16").Value >= 100) Then
                xlWorkSheet.Range("S16").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S16").Value < 100 And xlWorkSheet.Range("S16").Value >= 95) Then
                xlWorkSheet.Range("S16").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S16").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S17").Value >= 100) Then
                xlWorkSheet.Range("S17").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S17").Value < 100 And xlWorkSheet.Range("S17").Value >= 95) Then
                xlWorkSheet.Range("S17").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S17").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S18").Value >= 100) Then
                xlWorkSheet.Range("S18").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S18").Value < 100 And xlWorkSheet.Range("S18").Value >= 95) Then
                xlWorkSheet.Range("S18").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S18").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S19").Value >= 100) Then
                xlWorkSheet.Range("S19").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S19").Value < 100 And xlWorkSheet.Range("S19").Value >= 95) Then
                xlWorkSheet.Range("S19").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S19").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S20").Value >= 100) Then
                xlWorkSheet.Range("S20").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S20").Value < 100 And xlWorkSheet.Range("S20").Value >= 95) Then
                xlWorkSheet.Range("S20").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S20").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S21").Value >= 100) Then
                xlWorkSheet.Range("S21").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S21").Value < 100 And xlWorkSheet.Range("S21").Value >= 95) Then
                xlWorkSheet.Range("S21").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S21").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S22").Value >= 100) Then
                xlWorkSheet.Range("S22").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S22").Value < 100 And xlWorkSheet.Range("S22").Value >= 95) Then
                xlWorkSheet.Range("S22").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S22").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S23").Value >= 100) Then
                xlWorkSheet.Range("S23").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S23").Value < 100 And xlWorkSheet.Range("S23").Value >= 95) Then
                xlWorkSheet.Range("S23").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S23").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T4").Value >= 100) Then
                xlWorkSheet.Range("T4").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T4").Value < 100 And xlWorkSheet.Range("T4").Value >= 95) Then
                xlWorkSheet.Range("T4").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T4").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T5").Value >= 100) Then
                xlWorkSheet.Range("T5").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T5").Value < 100 And xlWorkSheet.Range("T5").Value >= 95) Then
                xlWorkSheet.Range("T5").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T5").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T6").Value >= 100) Then
                xlWorkSheet.Range("T6").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T6").Value < 100 And xlWorkSheet.Range("T6").Value >= 95) Then
                xlWorkSheet.Range("T6").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T6").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T7").Value >= 100) Then
                xlWorkSheet.Range("T7").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T7").Value < 100 And xlWorkSheet.Range("T7").Value >= 95) Then
                xlWorkSheet.Range("T7").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T7").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T8").Value >= 100) Then
                xlWorkSheet.Range("T8").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T8").Value < 100 And xlWorkSheet.Range("T8").Value >= 95) Then
                xlWorkSheet.Range("T8").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T8").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T9").Value >= 100) Then
                xlWorkSheet.Range("T9").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T9").Value < 100 And xlWorkSheet.Range("T9").Value >= 95) Then
                xlWorkSheet.Range("T9").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T9").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("T10").Value >= 100) Then
                xlWorkSheet.Range("T10").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("T10").Value < 100 And xlWorkSheet.Range("T10").Value >= 95) Then
                xlWorkSheet.Range("T10").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("T10").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R24").Value >= 100) Then
                xlWorkSheet.Range("R24").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R24").Value < 100 And xlWorkSheet.Range("R24").Value >= 95) Then
                xlWorkSheet.Range("R24").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R24").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R25").Value >= 100) Then
                xlWorkSheet.Range("R25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R25").Value < 100 And xlWorkSheet.Range("R25").Value >= 95) Then
                xlWorkSheet.Range("R25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R26").Value >= 100) Then
                xlWorkSheet.Range("R26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R26").Value < 100 And xlWorkSheet.Range("R26").Value >= 95) Then
                xlWorkSheet.Range("R26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R27").Value >= 100) Then
                xlWorkSheet.Range("R27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R27").Value < 100 And xlWorkSheet.Range("R27").Value >= 95) Then
                xlWorkSheet.Range("R27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R28").Value >= 100) Then
                xlWorkSheet.Range("R28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R28").Value < 100 And xlWorkSheet.Range("R28").Value >= 95) Then
                xlWorkSheet.Range("R28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R29").Value >= 100) Then
                xlWorkSheet.Range("R29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R29").Value < 100 And xlWorkSheet.Range("R29").Value >= 95) Then
                xlWorkSheet.Range("R29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R30").Value >= 100) Then
                xlWorkSheet.Range("R30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R30").Value < 100 And xlWorkSheet.Range("R30").Value >= 95) Then
                xlWorkSheet.Range("R30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R31").Value >= 100) Then
                xlWorkSheet.Range("R31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R31").Value < 100 And xlWorkSheet.Range("R31").Value >= 95) Then
                xlWorkSheet.Range("R31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R32").Value >= 100) Then
                xlWorkSheet.Range("R32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R32").Value < 100 And xlWorkSheet.Range("R32").Value >= 95) Then
                xlWorkSheet.Range("R32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R33").Value >= 100) Then
                xlWorkSheet.Range("R33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R33").Value < 100 And xlWorkSheet.Range("R33").Value >= 95) Then
                xlWorkSheet.Range("R33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R33").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R34").Value >= 100) Then
                xlWorkSheet.Range("R34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R34").Value < 100 And xlWorkSheet.Range("R34").Value >= 95) Then
                xlWorkSheet.Range("R34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R34").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S24").Value >= 100) Then
                xlWorkSheet.Range("S24").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S24").Value < 100 And xlWorkSheet.Range("S24").Value >= 95) Then
                xlWorkSheet.Range("S24").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S24").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S25").Value >= 100) Then
                xlWorkSheet.Range("S25").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S25").Value < 100 And xlWorkSheet.Range("S25").Value >= 95) Then
                xlWorkSheet.Range("S25").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S25").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S26").Value >= 100) Then
                xlWorkSheet.Range("S26").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S26").Value < 100 And xlWorkSheet.Range("S26").Value >= 95) Then
                xlWorkSheet.Range("S26").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S26").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S27").Value >= 100) Then
                xlWorkSheet.Range("S27").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S27").Value < 100 And xlWorkSheet.Range("S27").Value >= 95) Then
                xlWorkSheet.Range("S27").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S27").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S28").Value >= 100) Then
                xlWorkSheet.Range("S28").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S28").Value < 100 And xlWorkSheet.Range("S28").Value >= 95) Then
                xlWorkSheet.Range("S28").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S28").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S29").Value >= 100) Then
                xlWorkSheet.Range("S29").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S29").Value < 100 And xlWorkSheet.Range("S29").Value >= 95) Then
                xlWorkSheet.Range("S29").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S29").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S30").Value >= 100) Then
                xlWorkSheet.Range("S30").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S30").Value < 100 And xlWorkSheet.Range("S30").Value >= 95) Then
                xlWorkSheet.Range("S30").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S30").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S31").Value >= 100) Then
                xlWorkSheet.Range("S31").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S31").Value < 100 And xlWorkSheet.Range("S31").Value >= 95) Then
                xlWorkSheet.Range("S31").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S31").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S32").Value >= 100) Then
                xlWorkSheet.Range("S32").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S32").Value < 100 And xlWorkSheet.Range("S32").Value >= 95) Then
                xlWorkSheet.Range("S32").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S32").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S33").Value >= 100) Then
                xlWorkSheet.Range("S33").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S33").Value < 100 And xlWorkSheet.Range("S33").Value >= 95) Then
                xlWorkSheet.Range("S33").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S33").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S34").Value >= 100) Then
                xlWorkSheet.Range("S34").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S34").Value < 100 And xlWorkSheet.Range("S34").Value >= 95) Then
                xlWorkSheet.Range("S34").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S34").Interior.Color = Color.OrangeRed
            End If

            If (xlWorkSheet.Range("S35").Value >= 100) Then
                xlWorkSheet.Range("S35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S35").Value < 100 And xlWorkSheet.Range("S35").Value >= 95) Then
                xlWorkSheet.Range("S35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S36").Value >= 100) Then
                xlWorkSheet.Range("S36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S36").Value < 100 And xlWorkSheet.Range("S36").Value >= 95) Then
                xlWorkSheet.Range("S36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S37").Value >= 100) Then
                xlWorkSheet.Range("S37").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S37").Value < 100 And xlWorkSheet.Range("S37").Value >= 95) Then
                xlWorkSheet.Range("S37").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S37").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("S38").Value >= 100) Then
                xlWorkSheet.Range("S38").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("S38").Value < 100 And xlWorkSheet.Range("S38").Value >= 95) Then
                xlWorkSheet.Range("S38").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("S38").Interior.Color = Color.OrangeRed
            End If

            If (xlWorkSheet.Range("R35").Value >= 100) Then
                xlWorkSheet.Range("R35").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R35").Value < 100 And xlWorkSheet.Range("R35").Value >= 95) Then
                xlWorkSheet.Range("R35").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R35").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R36").Value >= 100) Then
                xlWorkSheet.Range("R36").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R36").Value < 100 And xlWorkSheet.Range("R36").Value >= 95) Then
                xlWorkSheet.Range("R36").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R36").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R37").Value >= 100) Then
                xlWorkSheet.Range("R37").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R37").Value < 100 And xlWorkSheet.Range("R37").Value >= 95) Then
                xlWorkSheet.Range("R37").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R37").Interior.Color = Color.OrangeRed
            End If
            If (xlWorkSheet.Range("R38").Value >= 100) Then
                xlWorkSheet.Range("R38").Interior.Color = Color.LightGreen
            ElseIf (xlWorkSheet.Range("R38").Value < 100 And xlWorkSheet.Range("R38").Value >= 95) Then
                xlWorkSheet.Range("R38").Interior.Color = Color.Yellow
            Else
                xlWorkSheet.Range("R38").Interior.Color = Color.OrangeRed
            End If
            'xlApp.DisplayAlerts = False
            ' xlWorkSheet.Range("N2:N36").VerticalAlignment = Excel.Constants.xlCenter
            'xlWorkSheet.Range("A25:N34").VerticalAlignment = Excel.Constants.xlCenter

            'xlWorkSheet.Range("A1", "A3000").Replace("1", "#DIV/0!")
            xlWorkSheet.Cells(2, 1) = "Parameter"
            xlWorkSheet.Cells(2, 2) = "OSP(On date)"
            xlWorkSheet.Cells(2, 3) = "ABP(On Date)"
            xlWorkSheet.Cells(2, 4) = "Actual(On Date)"
            xlWorkSheet.Cells(2, 5) = "ABP Comp(On Date)"
            xlWorkSheet.Cells(2, 6) = "OSP Comp(On Date)"
            xlWorkSheet.Cells(2, 7) = "OSP(To Date)"
            xlWorkSheet.Cells(2, 8) = "ABP(To Date)"
            xlWorkSheet.Cells(2, 9) = "Actual(To Date)"
            xlWorkSheet.Cells(2, 10) = "OSP Comp(On Date)"
            xlWorkSheet.Cells(2, 11) = "ABP Comp(To Date)"
            xlWorkSheet.Cells(2, 12) = "Asking Rate(OSP)"
            xlWorkSheet.Cells(2, 13) = "Asking Rate(ABP)"
            xlWorkSheet.Cells(2, 14) = "Paramater"
            xlWorkSheet.Cells(2, 15) = "ABP"
            xlWorkSheet.Cells(2, 16) = "Actual"
            xlWorkSheet.Cells(2, 17) = "Comp(On Date)"
            xlWorkSheet.Cells(2, 18) = "ABP(To Date)"
            xlWorkSheet.Cells(2, 19) = "Actual(To Date)"
            xlWorkSheet.Cells(2, 20) = "Comp(To Date)"
            xlWorkSheet.Cells(2, 21) = "Asking Rate"
            xlWorkSheet.Range("A3:M3").MergeCells = True
            xlWorkSheet.Range("A7:M7").MergeCells = True
            xlWorkSheet.Range("A11:M11").MergeCells = True
            xlWorkSheet.Range("A15:M15").MergeCells = True
            xlWorkSheet.Range("A19:M19").MergeCells = True

            xlWorkSheet.Cells(3, 1) = "Dispatch(OMQ)"
            xlWorkSheet.Cells(4, 1) = "SO"
            xlWorkSheet.Cells(5, 1) = "FO"
            xlWorkSheet.Cells(6, 1) = "Total"
            xlWorkSheet.Range("A4").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A5").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A6").Interior.Color = Color.Yellow

            xlWorkSheet.Cells(7, 1) = "Dispatch To TSJ"
            xlWorkSheet.Range("A8").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A9").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A10").Interior.Color = Color.Yellow
            xlWorkSheet.Cells(8, 1) = "SO"
            xlWorkSheet.Cells(9, 1) = "FO"
            xlWorkSheet.Cells(10, 1) = "Total"
            xlWorkSheet.Cells(11, 1) = "Dispatch To TSK"
            xlWorkSheet.Range("A12").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A13").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A14").Interior.Color = Color.Yellow
            xlWorkSheet.Cells(12, 1) = "SO"
            xlWorkSheet.Cells(13, 1) = "FO"
            xlWorkSheet.Cells(14, 1) = "Total"
            xlWorkSheet.Cells(15, 1) = "Dispatch To TSBSL"
            xlWorkSheet.Range("A16").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A17").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A18").Interior.Color = Color.Yellow
            xlWorkSheet.Cells(16, 1) = "SO"
            xlWorkSheet.Cells(17, 1) = "FO"
            xlWorkSheet.Cells(18, 1) = "Total"
            xlWorkSheet.Cells(19, 1) = "Dispatch To Sister Concern"
            xlWorkSheet.Range("A20").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A21").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A22").Interior.Color = Color.Yellow
            xlWorkSheet.Cells(20, 1) = "SO"
            xlWorkSheet.Cells(21, 1) = "FO"
            xlWorkSheet.Cells(22, 1) = "Total"
            xlWorkSheet.Range("A3").Interior.Color = Color.Gold
            xlWorkSheet.Range("A7").Interior.Color = Color.Gold
            xlWorkSheet.Range("A11").Interior.Color = Color.Gold
            xlWorkSheet.Range("A15").Interior.Color = Color.Gold
            xlWorkSheet.Range("A19").Interior.Color = Color.Gold
            'xlWorkSheet.Range("A3:M3").VerticalAlignment = Excel.Constants.xlCenter
            xlWorkSheet.Range("N3").Interior.Color = Color.Gold
            xlWorkSheet.Range("A2:V48").Cells.Borders.LineStyle = 1

            xlWorkSheet.Range("A2:M2").Interior.Color = Color.LightBlue
            xlWorkSheet.Range("N2:U2").Interior.Color = Color.LightBlue
            xlWorkSheet.Range("A24:K24").Interior.Color = Color.LightBlue
            xlWorkSheet.Cells(24, 1) = "Mines"
            xlWorkSheet.Cells(24, 3) = "OSP On Date"
            xlWorkSheet.Cells(24, 4) = "ABP(On Date)"
            xlWorkSheet.Cells(24, 5) = "Actual(On Date)"
            xlWorkSheet.Cells(24, 6) = "Comp(OSP On Date)"
            xlWorkSheet.Cells(24, 7) = "Comp(ABP On Date)"
            xlWorkSheet.Cells(24, 8) = "Comp(OSP To Date)"
            xlWorkSheet.Cells(24, 9) = "Comp(ABP To Date)"
            xlWorkSheet.Cells(24, 10) = "Asking Rate(OSP)"
            ' xlWorkSheet.Cells(24, 11) = "Comp (Monthly W/o ABP)"
            xlWorkSheet.Cells(24, 11) = "Asking Rate(ABP)"
            xlWorkSheet.Cells(25, 2) = "SO"
            xlWorkSheet.Cells(26, 2) = "FO"
            xlWorkSheet.Cells(27, 2) = "Total"
            xlWorkSheet.Cells(28, 2) = "SO"
            xlWorkSheet.Cells(29, 2) = "FO"
            xlWorkSheet.Cells(30, 2) = "Total"
            xlWorkSheet.Cells(31, 2) = "SO"
            xlWorkSheet.Cells(32, 2) = "FO"
            xlWorkSheet.Cells(33, 2) = "Total"
            xlWorkSheet.Cells(34, 2) = "SO"
            xlWorkSheet.Cells(35, 2) = "FO"
            xlWorkSheet.Cells(36, 2) = "Total"
            xlWorkSheet.Range("A25:A27").MergeCells = True
            xlWorkSheet.Range("A28:A30").MergeCells = True
            xlWorkSheet.Range("A31:A33").MergeCells = True
            xlWorkSheet.Range("A34:A36").MergeCells = True
            xlWorkSheet.Cells(25, 1) = "NIM"
            xlWorkSheet.Cells(28, 1) = "KTM"
            xlWorkSheet.Cells(31, 1) = "JEIM"
            xlWorkSheet.Cells(34, 1) = "KIM"
            xlWorkSheet.Range("A25").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A28").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A31").Interior.Color = Color.Yellow
            xlWorkSheet.Range("A34").Interior.Color = Color.Yellow

            '2nd part

            xlWorkSheet.Cells(3, 14) = "Mining(OMQ)"
            xlWorkSheet.Range("N3:U3").MergeCells = True
            xlWorkSheet.Range("N3").Interior.Color = Color.Gold
            xlWorkSheet.Cells(4, 14) = "ROM"
            xlWorkSheet.Cells(5, 14) = "OB"
            xlWorkSheet.Cells(6, 14) = "Excavation"
            xlWorkSheet.Range("N4").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N5").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N6").Interior.Color = Color.Yellow

            xlWorkSheet.Cells(7, 14) = "Production(OMQ)"
            xlWorkSheet.Range("N7:U7").MergeCells = True
            xlWorkSheet.Range("N7").Interior.Color = Color.Gold
            xlWorkSheet.Cells(8, 14) = "SO"
            xlWorkSheet.Cells(9, 14) = "FO"
            xlWorkSheet.Cells(10, 14) = "Total"
            xlWorkSheet.Range("N8").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N9").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N10").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N11:T11").Interior.Color = Color.LightBlue
            xlWorkSheet.Cells(11, 14) = "Plant"
            xlWorkSheet.Cells(11, 15) = " "
            xlWorkSheet.Cells(11, 16) = "ABP(On Date)"
            xlWorkSheet.Cells(11, 17) = "Actual(On Date)"
            xlWorkSheet.Cells(11, 18) = "Comp(On date)"
            xlWorkSheet.Cells(11, 19) = "Comp(To Date)"
            xlWorkSheet.Cells(11, 20) = " Asking Rate(ABP)"
            xlWorkSheet.Range("N12:N14").MergeCells = True
            xlWorkSheet.Range("N15:N17").MergeCells = True
            xlWorkSheet.Range("N18:N20").MergeCells = True
            xlWorkSheet.Range("N21:N23").MergeCells = True
            xlWorkSheet.Range("N24:N26").MergeCells = True
            xlWorkSheet.Range("N27:N29").MergeCells = True
            xlWorkSheet.Range("N30:N32").MergeCells = True
            xlWorkSheet.Range("N33:N35").MergeCells = True
            xlWorkSheet.Range("N36:N38").MergeCells = True
            '2 lines added by nihar 31/08/21
            xlWorkSheet.Range("N39:N41").MergeCells = True
            xlWorkSheet.Range("N42:N44").MergeCells = True
            'end
            xlWorkSheet.Cells(12, 14) = "NIM"
            xlWorkSheet.Cells(15, 14) = "KTM"
            xlWorkSheet.Cells(18, 14) = "NIM+KTM"
            xlWorkSheet.Cells(21, 14) = "JEIM(Own Plant)"
            xlWorkSheet.Cells(24, 14) = "JEIM(O/S Plant)"
            xlWorkSheet.Cells(27, 14) = "JEIM(Total)"
            xlWorkSheet.Cells(30, 14) = "KIM(Own Plant)"
            xlWorkSheet.Cells(33, 14) = "KIM(O/S Plant)"
            '2 lines added by nihar 31/08/21
            xlWorkSheet.Cells(36, 14) = "KIM(Wet Plant)"
            xlWorkSheet.Cells(39, 14) = "KIM(Screening Plant)"
            'end
            xlWorkSheet.Cells(42, 14) = "KIM(Total)"
            xlWorkSheet.Range("N12").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N15").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N18").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N21").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N24").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N27").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N30").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N33").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N36").Interior.Color = Color.Yellow
            '2 lines added by nihar 31/08/21
            xlWorkSheet.Range("N39").Interior.Color = Color.Yellow
            xlWorkSheet.Range("N42").Interior.Color = Color.Yellow
            'end
            xlWorkSheet.Range("A2:U2").WrapText = True
            xlWorkSheet.Range("N12:N42").WrapText = True
            'xlWorkSheet.Range("A24:K24").WrapText = True
            ' xlWorkSheet.Range("O24:P24").MergeCells = True
            '  xlWorkSheet.Range("N24:T24").Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 14) = "Mines"

            '''xlWorkSheet.Range(24, 15).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 15) = "Sch(Till Date)"
            '''xlWorkSheet.Range(24, 17).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 17) = "Actual"
            '''xlWorkSheet.Range(24, 18).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 18) = "Total Monthly"
            '''xlWorkSheet.Range(24, 19).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 19) = "Comp(Mon)"
            '''xlWorkSheet.Range(24, 20).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 20) = "Required(Rake/Day)"
            '''xlWorkSheet.Range(24, 21).Interior.Color = Color.LightBlue
            ''xlWorkSheet.Cells(24, 21) = "Total(Mon)"
            ''xlWorkSheet.Range("N25:N33").Interior.Color = Color.Yellow
            ''xlWorkSheet.Range("N25:N26").MergeCells = True
            ''xlWorkSheet.Range("N27:N28").MergeCells = True
            ''xlWorkSheet.Range("N30:N31").MergeCells = True
            ''xlWorkSheet.Range("N32:N33").MergeCells = True
            ''xlWorkSheet.Cells(25, 14) = "TSK"
            ''xlWorkSheet.Cells(27, 14) = "TSJ"
            ''xlWorkSheet.Cells(29, 14) = "TSIL"
            ''xlWorkSheet.Cells(30, 14) = "TMIL"
            ''xlWorkSheet.Cells(32, 14) = "TSBSL"
            ''xlWorkSheet.Cells(25, 15) = "SO"
            ''xlWorkSheet.Cells(26, 15) = "FO"
            ''xlWorkSheet.Cells(27, 15) = "SO"
            ''xlWorkSheet.Cells(28, 15) = "FO"
            ''xlWorkSheet.Cells(29, 15) = "SO"
            ''xlWorkSheet.Cells(30, 15) = "SO"
            ''xlWorkSheet.Cells(31, 15) = "FO"
            ''xlWorkSheet.Cells(32, 15) = "SO"
            ''xlWorkSheet.Cells(33, 15) = "FO"
            xlWorkSheet.Cells(12, 15) = "SO"
            xlWorkSheet.Cells(13, 15) = "FO"
            xlWorkSheet.Cells(14, 15) = "Total"
            xlWorkSheet.Cells(15, 15) = "SO"
            xlWorkSheet.Cells(16, 15) = "FO"
            xlWorkSheet.Cells(17, 15) = "Total"
            xlWorkSheet.Cells(18, 15) = "SO"
            xlWorkSheet.Cells(19, 15) = "FO"
            xlWorkSheet.Cells(20, 15) = "Total"
            xlWorkSheet.Cells(21, 15) = "SO"
            xlWorkSheet.Cells(22, 15) = "FO"
            xlWorkSheet.Cells(23, 15) = "Total"
            xlWorkSheet.Cells(24, 15) = "SO"
            xlWorkSheet.Cells(25, 15) = "FO"
            xlWorkSheet.Cells(26, 15) = "Total"
            xlWorkSheet.Cells(27, 15) = "SO"
            xlWorkSheet.Cells(28, 15) = "FO"
            xlWorkSheet.Cells(29, 15) = "Total"
            xlWorkSheet.Cells(30, 15) = "SO"
            xlWorkSheet.Cells(31, 15) = "FO"
            xlWorkSheet.Cells(32, 15) = "Total"
            xlWorkSheet.Cells(33, 15) = "SO"
            xlWorkSheet.Cells(34, 15) = "FO"
            xlWorkSheet.Cells(35, 15) = "Total"
            'added by nihar 31/8/21
            xlWorkSheet.Cells(36, 15) = "SO"
            xlWorkSheet.Cells(37, 15) = "FO"
            xlWorkSheet.Cells(38, 15) = "Total"
            xlWorkSheet.Cells(39, 15) = "SO"
            xlWorkSheet.Cells(40, 15) = "FO"
            xlWorkSheet.Cells(41, 15) = "Total"

            xlWorkSheet.Cells(42, 15) = "SO"
            xlWorkSheet.Cells(43, 15) = "FO"
            xlWorkSheet.Cells(44, 15) = "Total"

            xlWorkSheet.Range("A1:U1").MergeCells = True

#End Region
            ' xlWorkSheet.Range("N24:N26").MergeCells = True
            ' xlWorkSheet.Shapes.AddPicture("C:\xl_pic.JPG",
            '   Microsoft.Office.Core.MsoTriState.msoFalse,
            ' Microsoft.Office.Core.MsoTriState.msoCTrue, 50, 50, 300, 45)
            Dim dt_today As Date = Date.Today
            dt_today = Date.Now.ToString("dd-MMM-yy")

            xlWorkSheet.Cells(1, 1) = dt_today
            ' xlWorkSheet.SaveAs("G:\TeamProjects\rmis_final\rmis\file.xls")
            ' Dim FileName2 As String = Path.Combine("D:\Report Formats\Daily Report.xlsx")
            ' Dim FileNamePdf As String = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, ".\Report Formats\Daily Report  " + DDLFACILITY.SelectedItem.Text + "  Iron Mines_" + TextBox1.Text.ToString.Substring(0, 2) + ".pdf")

            If System.IO.File.Exists(newfilepathexcel) = True Then
                System.IO.File.Delete(newfilepathexcel)
            End If
            Dim excelFlePath As String = String.Empty

            xlWorkBook.SaveAs(newfilepathexcel)
            excelFlePath = newfilepathexcel
            xlWorkBook.Close()
            xlApp.Quit()
            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)
            Dim sa As SqlDataAdapter = New SqlDataAdapter(qryRecipientlist, myConn)
            'For checking the missing facility 
            Dim message As String = "Dear All,Please find the OMQ daily report for '" + dttt2 + "'"
            Dim ds_missing_facility As New Data.DataSet
            Dim facility() As String = {"DPJ", "JEIM", "KIMPLANT", "NOADESP", "KBDESP", "NDCMP", "JODDESP", "KIMDESP", "NIM", "KBIM", "KIM", "DPK", "WPJ", "0"}
            query = datamissingfacility
            Dim flag As Integer = 0
            Dim mailbody As New StringBuilder
            mailbody.AppendLine("Dear All,")


            mailbody.AppendLine("Please find the OMQ daily report for " + dttt2)
            Console.WriteLine("Please find the OMQ daily report for " + dttt2)
            get_datatable(query, ds_missing_facility)
            Dim temp10 As Integer = 0
            Dim cnt As Integer = 0
            If (ds_missing_facility.Tables(0).Rows.Count > 0) Then

                For i = 0 To 12
                    flag = 0
                    For j = 0 To ds_missing_facility.Tables(0).Rows.Count - 1
                        If facility(i).ToString = ds_missing_facility.Tables(0).Rows(j)(0).ToString Then
                            flag = 1
                            Exit For
                        End If

                    Next
                    If flag = 0 Then
                        temp10 = 1
                        If cnt = 0 Then
                            Console.WriteLine()
                            mailbody.AppendLine("Data not entered for –")
                            Console.WriteLine("Data not entered for –")
                        End If
                        cnt = 1
                        'mailbody.AppendLine(facility(i).ToString + " data not entered")
                        If facility(i).ToString = "NIM" Then
                            Console.WriteLine("Noamundi Mining")
                            mailbody.AppendLine("Noamundi Mining")

                        End If
                        If facility(i).ToString = "NDCMP" Then
                            Console.WriteLine("Noamundi Plant")
                            mailbody.AppendLine("Noamundi Plant")
                        End If
                        If facility(i).ToString = "NOADESP" Then
                            Console.WriteLine("Noamundi Despatch")
                            mailbody.AppendLine("Noamundi Despatch")
                        End If
                        If facility(i).ToString = "KIMDESP" Then
                            Console.WriteLine("Katamati Despatch")
                            mailbody.AppendLine("Katamati Despatch")
                        End If
                        If facility(i).ToString = "KIM" Then
                            Console.WriteLine("Katamati Mining")
                            mailbody.AppendLine("Katamati Mining")
                        End If
                        If facility(i).ToString = "KIMPLANT" Then
                            Console.WriteLine("Katamati Plant")
                            mailbody.AppendLine("Katamati Plant")
                        End If
                        If facility(i).ToString = "DPJ" Then
                            Console.WriteLine("Dry Plant Joda")
                            mailbody.AppendLine("Dry Plant Joda")
                        End If
                        If facility(i).ToString = "WPJ" Then
                            Console.WriteLine("Wet Plant Joda")
                            mailbody.AppendLine("Wet Plant Joda")
                        End If
                        If facility(i).ToString = "JEIM" Then
                            Console.WriteLine("Joda Mining")
                            mailbody.AppendLine("Joda Mining")
                        End If
                        If facility(i).ToString = "JODDESP" Then
                            Console.WriteLine("Joda Despatch")
                            mailbody.AppendLine("Joda Despatch")
                        End If

                        If facility(i).ToString = "KBDESP" Then
                            Console.WriteLine("Khondbond Despatch")
                            mailbody.AppendLine("Khondbond Despatch")
                        End If
                        If facility(i).ToString = "KBIM" Then
                            Console.WriteLine("Khondbond Mining")
                            mailbody.AppendLine("Khondbond Mining")
                        End If
                        If facility(i).ToString = "DPK" Then
                            Console.WriteLine("Khondbond Plant")
                            mailbody.AppendLine("Khondbond Plant")
                        End If

                    End If
                Next

            Else
                mailbody.AppendLine("Data Not Entered for any location")
                Console.WriteLine("Data Not Entered for any location")
                Console.WriteLine()
                Console.WriteLine("------------This is M/C generated Mail please do not reply to this-------------")
            End If
            If temp10 = 0 Then
                Console.WriteLine("Data Entered for all the location")
                Console.WriteLine()
                Console.WriteLine("------------This is M/C generated Mail please do not reply to this-------------")
            Else
                Console.WriteLine()
                Console.WriteLine("------------This is M/C generated Mail please do not reply to this-------------")
            End If

            Dim mail As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
            Dim ma As System.Net.Mail.MailAddress = New System.Net.Mail.MailAddress("neeraj.thakur@tatasteel.com", "Neeraj Kumar Thakur")
            mail.From = ma
            mail.Subject = "OMQ Daily Report"
            ' mail.To.Add("OMQDailyReport@tslin.onmicrosoft.com")
            mail.To.Add("nihar.nanda@tatasteel.com")
            ' mail.CC.Add("its.operation@tatasteel.com")
            ' mail.CC.Add("neeraj.thakur@tatasteel.com")
            mail.CC.Add("nihar.nanda@tatasteel.com")
            mail.Body = mailbody.ToString

            mail.Attachments.Add(New Attachment(newfilepathexcel))
            Dim smtpClient As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient("144.0.11.253")
            smtpClient.Send(mail)

            'System.Web.HttpContext.Current.Response.Clear()
            'System.Web.HttpContext.Current.Response.ContentType = "application/ms-excel"
            'System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment; filename=" & System.IO.Path.GetFileName(excelFlePath))
            'System.Web.HttpContext.Current.Response.TransmitFile(excelFlePath)

            'System.Web.HttpContext.Current.Response.End()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
            Console.WriteLine(ex.StackTrace)
            Console.ReadLine()
            Console.WriteLine(ex.InnerException)
            Console.ReadLine()
        End Try


    End Sub
    Public Sub get_datatable(ByVal query As String, ByRef dt As DataSet)
        Try
            ' conn = New OracleConnection(Str_conn)

            data_adapt = New OracleDataAdapter(query, conn)
            data_adapt.Fill(dt)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
            Console.WriteLine(ex.StackTrace)
            Console.ReadLine()
            Console.WriteLine(ex.InnerException)
            Console.ReadLine()
            Throw
        Finally
            data_adapt.Dispose()
            Disconnectconn()
        End Try
    End Sub
    Public Sub Disconnectconn()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
    End Sub
    Public Sub get_dataset(ByVal query As String, ByRef ds As DataSet)
        Try
            data_adapt = New OracleDataAdapter(query, conn)
            data_adapt.Fill(ds)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
            Console.WriteLine(ex.StackTrace)
            Console.ReadLine()
            Console.WriteLine(ex.InnerException)
            Console.ReadLine()

            Throw
        Finally
            data_adapt.Dispose()
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
            Console.WriteLine(ex.StackTrace)
            Console.ReadLine()
            Console.WriteLine(ex.InnerException)
            Console.ReadLine()
        Finally
            o = Nothing
        End Try
    End Sub


End Module
