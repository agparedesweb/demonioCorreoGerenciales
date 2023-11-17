Imports Sistema.Comunes.Comun
Imports System.Windows.Forms
Imports System.Drawing.Design
Imports System.Drawing
Imports Sistema
Imports System.Web
Imports System.IO
Imports Sistema.Comunes.Catalogos
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports Microsoft.VisualBasic
Imports Microsoft.Office.Core


Namespace Comunes.Comun
    Public Class ClsInformes



        Public Shared Function fgInformeEnviado(ByVal prmReporte As String, ByVal prmFecha As Date) As Boolean

            Dim DAO As New DataAccessCls
            Dim vcSQL As String = ""
            Dim DT As New DataTable

            Try

            Catch ex As Exception

            End Try

            DAO = DataAccessCls.DevuelveInstancia


            vcSQL = "SELECT TOP 1 * FROM ENVIACORREOS(NOLOCK)"
            vcSQL += vbCrLf & "WHERE CNOMBRE = '" & prmReporte & "' AND CONVERT(VARCHAR(20),DFECHA,112) = '" & Format(prmFecha, "yyyyMMdd") & "'"


            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then Return True

            Return False


        End Function

        Public Shared Function fgGeneralDeEmpaque(ByVal prmFecha As Date, ByVal prmTemporada As String) As Boolean


            Dim vcSQL As String

            Try
                vcSQL = ""
                vcSQL = vcSQL & "SELECT   ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "         ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "         ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "         ntipo, " & vbCrLf
                vcSQL = vcSQL & "         cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "         cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "         cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nempaquerango)     AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nempaqueacum)      AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nnacionalrango)    AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nnacionalacum)     AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nexportacionrango) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "         Sum(nexportacionacum)  AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "FROM     (SELECT   emp.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)     AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(emp.nbultos)         AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_empaque emp(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = emp.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = emp.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = emp.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    emp.cstatus NOT IN ('C','N') " & vbCrLf
                vcSQL = vcSQL & "                   AND emp.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),emp.dfecha,112) BETWEEN '" & Format(prmFecha, "yyyyMMdd") & "' AND '" & Format(prmFecha, "yyyyMMdd") & "'" & vbCrLf


                vcSQL = vcSQL & "                   AND emp.ccve_agricultor = '0001'" & vbCrLf


                vcSQL = vcSQL & "          GROUP BY emp.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre " & vbCrLf
                vcSQL = vcSQL & "          UNION  " & vbCrLf
                vcSQL = vcSQL & "          SELECT   emp.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)     AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(emp.nbultos)         AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_empaque emp(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = emp.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = emp.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = emp.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    emp.cstatus NOT IN ('C','N') " & vbCrLf
                vcSQL = vcSQL & "                   AND emp.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),emp.dfecha,112) <= '" & Format(prmFecha, "yyyyMMdd") & "' " & vbCrLf

                vcSQL = vcSQL & "                   AND emp.ccve_agricultor = '0001'" & vbCrLf

                vcSQL = vcSQL & "          GROUP BY emp.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   emp.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre " & vbCrLf
                vcSQL = vcSQL & "          UNION  " & vbCrLf
                vcSQL = vcSQL & "          SELECT   ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)     AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(ed.nbultos)          AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_detembarques ed(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_encembarques een(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON een.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND een.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND een.cfolio_manif = ed.cfolio_manif " & vbCrLf
                vcSQL = vcSQL & "                        AND ed.cmercado = een.cmercado " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_empaque ee(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON ee.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.cfolio_palet = ed.cfolio_palet " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = ed.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = ed.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = ed.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    een.cstatus = 'A' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),een.dfecha_trabajo,112) BETWEEN '" & Format(prmFecha, "yyyyMMdd") & "' AND '" & Format(prmFecha, "yyyyMMdd") & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.cmercado = 'N' " & vbCrLf

                vcSQL = vcSQL & "                   AND ee.ccve_agricultor = '0001'" & vbCrLf

                vcSQL = vcSQL & "          GROUP BY ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre " & vbCrLf
                vcSQL = vcSQL & "          UNION  " & vbCrLf
                vcSQL = vcSQL & "          SELECT   ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)        AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(ed.nbultos)          AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_detembarques ed(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_encembarques een(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON een.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND een.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND een.cfolio_manif = ed.cfolio_manif " & vbCrLf
                vcSQL = vcSQL & "                        AND ed.cmercado = een.cmercado " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_empaque ee(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON ee.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.cfolio_palet = ed.cfolio_palet " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = ed.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = ed.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = ed.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    een.cstatus = 'A' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),een.dfecha_trabajo,112) <= '" & Format(prmFecha, "yyyyMMdd") & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.cmercado = 'N' " & vbCrLf

                vcSQL = vcSQL & "                   AND ee.ccve_agricultor = '0001'" & vbCrLf

                vcSQL = vcSQL & "          GROUP BY ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre " & vbCrLf
                vcSQL = vcSQL & "          UNION  " & vbCrLf
                vcSQL = vcSQL & "          SELECT   ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)        AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(ed.nbultos)          AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_detembarques ed(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_encembarques een(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON een.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND een.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND een.cfolio_manif = ed.cfolio_manif " & vbCrLf
                vcSQL = vcSQL & "                        AND ed.cmercado = een.cmercado " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_empaque ee(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON ee.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.cfolio_palet = ed.cfolio_palet " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = ed.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = ed.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = ed.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    een.cstatus = 'A' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),een.dfecha_trabajo,112) BETWEEN '" & Format(prmFecha, "yyyyMMdd") & "' AND '" & Format(prmFecha, "yyyyMMdd") & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.cmercado = 'E' " & vbCrLf


                vcSQL = vcSQL & "                   AND ee.ccve_agricultor = '0001'" & vbCrLf

                vcSQL = vcSQL & "          GROUP BY ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre " & vbCrLf
                vcSQL = vcSQL & "          UNION  " & vbCrLf
                vcSQL = vcSQL & "          SELECT   ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   left(eti.cnombre,20)     AS cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre              AS cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre              AS cdescenvase, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaquerango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nempaqueacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalrango, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nnacionalacum, " & vbCrLf
                vcSQL = vcSQL & "                   Convert(NUMERIC(18,0),0) AS nexportacionrango, " & vbCrLf
                vcSQL = vcSQL & "                   Sum(ed.nbultos)          AS nexportacionacum " & vbCrLf
                vcSQL = vcSQL & "          FROM     eye_detembarques ed(nolock) " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_encembarques een(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON een.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND een.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND een.cfolio_manif = ed.cfolio_manif " & vbCrLf
                vcSQL = vcSQL & "                        AND ed.cmercado = een.cmercado " & vbCrLf
                vcSQL = vcSQL & "                   JOIN eye_empaque ee(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON ee.ccve_temporada = ed.ccve_temporada " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.ccve_empaque = ed.ccve_empaque " & vbCrLf
                vcSQL = vcSQL & "                        AND ee.cfolio_palet = ed.cfolio_palet " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_etiquetas eti(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON eti.ccve_etiqueta = ed.ccve_etiqueta " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_cultivos cul(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON cul.ccve_cultivo = ed.ccve_cultivo " & vbCrLf
                vcSQL = vcSQL & "                   JOIN ctl_envases env(nolock) " & vbCrLf
                vcSQL = vcSQL & "                     ON env.ccve_envase = ed.ccve_envase " & vbCrLf
                vcSQL = vcSQL & "          WHERE    een.cstatus = 'A' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.ccve_temporada = '" & prmTemporada & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND Convert(CHAR(10),een.dfecha_trabajo,112) <= '" & Format(prmFecha, "yyyyMMdd") & "' " & vbCrLf
                vcSQL = vcSQL & "                   AND een.cmercado = 'E' " & vbCrLf

                vcSQL = vcSQL & "                   AND ee.ccve_agricultor = '0001'" & vbCrLf


                vcSQL = vcSQL & "          GROUP BY ed.ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "                   ed.ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "                   eti.ntipo, " & vbCrLf
                vcSQL = vcSQL & "                   eti.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   cul.cnombre, " & vbCrLf
                vcSQL = vcSQL & "                   env.cnombre) AS resultado " & vbCrLf
                vcSQL = vcSQL & "WHERE    ccve_cultivo not in (" & ClsTools.fgGetParametro("CULEXCREP") & ")" & vbCrLf
                vcSQL = vcSQL & "GROUP BY ccve_cultivo, " & vbCrLf
                vcSQL = vcSQL & "         ccve_etiqueta, " & vbCrLf
                vcSQL = vcSQL & "         ccve_envase, " & vbCrLf
                vcSQL = vcSQL & "         ntipo, " & vbCrLf
                vcSQL = vcSQL & "         cdescetiqueta, " & vbCrLf
                vcSQL = vcSQL & "         cdesccultivo, " & vbCrLf
                vcSQL = vcSQL & "         cdescenvase"



                If ClsTools.fgCreaVista(vcSQL, "VWEMPAQUEGENERAL") Then



                    vcSQL = ""
                    vcSQL = vcSQL & "SELECT Max(chm)                  AS chm, " & vbCrLf
                    vcSQL = vcSQL & "       Max(cmar)                 AS cmar, " & vbCrLf
                    vcSQL = vcSQL & "       Max(cdvm)                 AS cdvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaquerangohm)      AS nempaquerangohm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaqueacumhm)       AS nempaqueacumhm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaquerangomar)     AS nempaquerangomar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaqueacummar)      AS nempaqueacummar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaquerangodvm)     AS nempaquerangodvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nempaqueacumdvm)      AS nempaqueacumdvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalrangohm)     AS nnacionalrangohm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalacumhm)      AS nnacionalacumhm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalrangomar)    AS nnacionalrangomar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalacummar)     AS nnacionalacummar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalrangodvm)    AS nnacionalrangodvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nnacionalacumdvm)     AS nnacionalacumdvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionrangohm)  AS nexportacionrangohm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionacumhm)   AS nexportacionacumhm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionrangomar) AS nexportacionrangomar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionacummar)  AS nexportacionacummar, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionrangodvm) AS nexportacionrangodvm, " & vbCrLf
                    vcSQL = vcSQL & "       Sum(nexportacionacumdvm)  AS nexportacionacumdvm " & vbCrLf
                    vcSQL = vcSQL & "FROM   (SELECT   CASE  " & vbCrLf
                    vcSQL = vcSQL & "                   WHEN ntipo = 1 THEN 'HM' " & vbCrLf
                    vcSQL = vcSQL & "                   ELSE '' " & vbCrLf
                    vcSQL = vcSQL & "                 END AS chm, " & vbCrLf
                    vcSQL = vcSQL & "                 CASE  " & vbCrLf
                    vcSQL = vcSQL & "                   WHEN ntipo = 2 THEN 'MARENGO' " & vbCrLf
                    vcSQL = vcSQL & "                   ELSE '' " & vbCrLf
                    vcSQL = vcSQL & "                 END AS cmar, " & vbCrLf
                    vcSQL = vcSQL & "                 CASE  " & vbCrLf
                    vcSQL = vcSQL & "                   WHEN ntipo = 3 THEN 'DVM' " & vbCrLf
                    vcSQL = vcSQL & "                   ELSE '' " & vbCrLf
                    vcSQL = vcSQL & "                 END AS cdvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nempaquerango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaquerangohm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nempaquerango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaquerangomar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nempaquerango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaquerangodvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nempaqueacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaqueacumhm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nempaqueacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaqueacummar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nempaqueacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nempaqueacumdvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nnacionalrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalrangohm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nnacionalrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalrangomar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nnacionalrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalrangodvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nnacionalacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalacumhm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nnacionalacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalacummar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nnacionalacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nnacionalacumdvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nexportacionrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionrangohm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nexportacionrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionrangomar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nexportacionrango " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionrangodvm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 1 THEN nexportacionacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionacumhm, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 2 THEN nexportacionacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionacummar, " & vbCrLf
                    vcSQL = vcSQL & "                 Sum(CASE  " & vbCrLf
                    vcSQL = vcSQL & "                       WHEN ntipo = 3 THEN nexportacionacum " & vbCrLf
                    vcSQL = vcSQL & "                       ELSE 0 " & vbCrLf
                    vcSQL = vcSQL & "                     END) AS nexportacionacumdvm " & vbCrLf
                    vcSQL = vcSQL & "        FROM     vwempaquegeneral (NOLOCK) " & vbCrLf
                    vcSQL = vcSQL & "        GROUP BY ntipo) AS resultado"



                    If ClsTools.fgCreaVista(vcSQL, "VWEMPAQUEGENERALTOTALES") Then
                        Return True
                    End If
                Else
                    Return False
                End If
            Catch ex As Exception
                Return False
            End Try



        End Function




    End Class
End Namespace

