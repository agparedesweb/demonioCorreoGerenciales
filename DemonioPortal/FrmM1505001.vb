Imports Sistema.Comunes.Comun.ClsTools
Imports Sistema.Comunes.Catalogos
Imports Sistema.Comunes.Registros
Imports Sistema.Comunes.Registros.EscribanoRegistros
Imports Sistema.Comunes.Registros.FabricaRegistros
Imports Sistema.Comunes.Comun
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Globalization
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports Sistema.DataAccessCls
Imports System.Data.OleDb
Imports OpenPop.Pop3
Imports OpenPop.Mime
Imports OpenPop.Mime.Header
Imports OpenPop.Pop3.Exceptions
Imports OpenPop.Common.Logging
Imports Message = OpenPop.Mime.Message
Imports Microsoft.Office.Interop
Imports System.Net
Imports Newtonsoft.Json
Imports Sistema.Comunes.Comun.WS
Imports System.Diagnostics.Contracts

'Imports Herramientas


Public Class FrmM1505001

    Dim vcOC As String = ""
    Dim vcVerficationcode As String = Guid.NewGuid().ToString
    Dim username As String = "joser2203@gmail.com"
    Dim password As String = "trruuorcpgjfdugq"

#Region "Declaraciones"

    Dim aplicacionExcel As Microsoft.Office.Interop.Excel.Application = Nothing
    Dim objHojaExcel As Object
    Dim gcServidorSMTP As String = ""
    Private gnPuertoSMTP As Integer
    Private gbUsaSSL As Boolean

    Private vbBerenjena1ra As Boolean = True
    Private vbBerenjena2da As Boolean = True

    Private vbVerdes1ra As Boolean = True
    Private vbVerdes2da As Boolean = True

    Private vbRojos1ra As Boolean = True
    Private vbRojos2da As Boolean = True

    Private vbAmarillo1ra As Boolean = True
    Private vbAmarillo2da As Boolean = True

    Private vbNaranja1ra As Boolean = True
    Private vbNaranja2da As Boolean = True

    Private vbBolas1ra As Boolean = True
    Private vbBolas2da As Boolean = True

    Private vbSaladette1ra As Boolean = True
    Private vbSaladette2da As Boolean = True


    Dim vcCorreosEnfriados As String = "<arturopp@aparedes.com.mx>,<carlosmm@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<raultb@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosVentas As String = "<arturopp@aparedes.com.mx>,<sergiopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosVentasSemanales As String = "<arturopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosVentasDist As String = "<arturopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>,<carlosmm@aparedes.com.mx>"
    Dim vcCorreosDisponibleExportacion As String = "<arturopp@aparedes.com.mx>,<carlosmm@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"


    Dim vcCorreosAvanceLabores As String = "<arturopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosPorcEmpaqueCalidad As String = "<arturopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<carlosmm@aparedes.com.mx>,<guillermoba@aparedes.com.mx>,<antoniogv@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosRentaTractores As String = "<sergiopp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<antoniogv@aparedes.com.mx>,<yuleidy.higuera@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosPorcentajeEmpaque As String = "<arturopp@aparedes.com.mx>,<carlosmm@aparedes.com.mx>,<guillermoba@aparedes.com.mx>,<antoniogv@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"


    Dim vcCorreosChequesCln As String = "<arturopp@aparedes.com.mx>,<enriquerm@aparedes.com.mx>,<kareli@aparedes.com.mx>,<gisela.iribe@aparedes.com.mx>,<anamaria.leal@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosChequesJal As String = "<arturopp@aparedes.com.mx>,<enriquerm@aparedes.com.mx>,<ruben.elias@aparedes.com.mx>,<yara.moreno@aparedes.com.mx>,<vanessa.moreno@aparedes.com.mx>,<jesusramonpp@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreosComplementoNomina As String = "<carlosmm@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"
    Dim vcCorreoElegiblesTargetas As String = "<antoniogv@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"

    Dim vcCorreosAsistenciaTSPV As String = "<joel@aparedes.com.mx>,<enriqueca@aparedes.com.mx>,<alfredor@aparedes.com.mx>"


    'Dim vcCorreos As String = "<enriqueca@aparedes.com.mx>"

    'Dim gcArchivoBitacora As String = "C:\rpt\BitacoraCorreos.txt"
    Dim gcArchivoBitacora As String = "C:\CROP\BitacoraCorreos.txt"

#End Region

#Region "Procedimientos y Funciones"

    Private Sub EscribeEnBitacora(ByVal prmMensajeEscribir As String)

        'If Not File.Exists(gcArchivoBitacora) Then
        '    File.Create(gcArchivoBitacora)
        'End If

        'If File.Exists(gcArchivoBitacora) Then
        '    Dim oSW As New StreamWriter(gcArchivoBitacora, True)
        '    Dim Linea As String = Format(Date.Now, "dd-MM-yyyy HH:mm:ss") & " - " & prmMensajeEscribir
        '    oSW.WriteLine(Linea)
        '    oSW.Flush()
        '    oSW.Close()
        '    oSW.Dispose()
        '    oSW = Nothing
        'End If


    End Sub


    Private Sub plObtenEnfriados()
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        EscribeEnBitacora("Obteniendo informacion de Preenfriado")

        Dim ds As DataSet = fgTraeEnfriados(vcTemporada, vdFecha)
        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then


            Try
                oRpt.Load("C:\CROP\RPT_PALLETSPREENFRIADO.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                AgregarParametro("@PRMTEPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMFECHA", vdFecha, oRpt)

                Dim oStrem As New System.IO.MemoryStream

                oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\CROP\PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
                ArchivoPDF.Flush()
                ArchivoPDF.Close()

                EscribeEnBitacora("Se creo PDF de Preenfriado")

                If File.Exists(vcArchivo) Then
                    Dim vcAdjuntos As New ArrayList()

                    vcAdjuntos.Add(vcArchivo)

                    EscribeEnBitacora("Se enviara coreo de PDF de Preenfriado")

                    ' Enviamos correo
                    'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                    flEnviarMail(vcCorreosEnfriados, vcAdjuntos, "INFORME DE PREENFRIADO AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    DAO.EjecutaComandoSQL("INSERT EYE_ENVIOPREENFRIADOS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','001'")

                    EscribeEnBitacora("Se inserta en tabla de preenfriado")
                End If

                'If File.Exists(vcArchivo) Then
                '    File.Delete(vcArchivo)
                'End If

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenEnfriados", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para preenfriado")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOPREENFRIADOS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','001'")
            EscribeEnBitacora("Se inserta en tabla de preenfriado")
            Exit Sub
        End If

        Proc.Dispose()
        'Proc.Kill()
        oRpt = Nothing

    End Sub

    Private Sub plObtenDatosAvanceLabores()
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        Console.Write("### Obteniendo información de Avance de Labores")

        Dim ds As DataSet = fgTraeVanceLabores(vcTemporada, "01", vdFecha)
        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then


            Try
                oRpt.Load("C:\CROP\RPT_AVANCELABORES.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMNOMINA", "01", oRpt)
                AgregarParametro("@PRMFECHA", vdFecha, oRpt)

                Dim oStrem As New System.IO.MemoryStream

                oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\CROP\AVANCE DE LABORES DEL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
                ArchivoPDF.Flush()
                ArchivoPDF.Close()

                EscribeEnBitacora("Se creo PDF de Avance de Labores")

                If File.Exists(vcArchivo) Then
                    Dim vcAdjuntos As New ArrayList()

                    vcAdjuntos.Add(vcArchivo)

                    EscribeEnBitacora("Se enviara correo de PDF de Avance de Labores")

                    ' Enviamos correo
                    'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    flEnviarMail(vcCorreosAvanceLabores, vcAdjuntos, "INFORME DE AVANCE DE LABORES DEL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    DAO.EjecutaComandoSQL("INSERT NOM_ENVIOAVANCELABORES SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                    EscribeEnBitacora("Se inserta en tabla de Avance de Labores")
                End If

                'If File.Exists(vcArchivo) Then
                '    File.Delete(vcArchivo)
                'End If

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenDatosAvanceLabores", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para Avance de Labores")
            DAO.EjecutaComandoSQL("INSERT NOM_ENVIOAVANCELABORES SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            Console.WriteLine("### Se inserta en tabla NOM_ENVIOAVANCELABORES")
            Exit Sub
        End If

        Proc.Dispose()
        'Proc.Kill()
        oRpt = Nothing

    End Sub

    Private Sub plObtenDatosRentaTractores()
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        EscribeEnBitacora("Obteniendo informacion de Renta de Tractores")

        Dim ds As DataSet = fgTraeInfoTractoresRenta(vcTemporada, "01", vdFecha)
        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then


            Try
                oRpt.Load("C:\CROP\RPT_LABORTRACTORES.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMNOMINA", "01", oRpt)
                AgregarParametro("@PRMFECHA", vdFecha, oRpt)

                Dim oStrem As New System.IO.MemoryStream

                oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = Replace("C:\CROP\INFORME DE TRACTORISTAS DEL " & UCase(Format(vdFecha, "dd-MM-yy")), ".", "") & ".pdf"

                If File.Exists(vcArchivo) Then
                    File.Delete(vcArchivo)
                End If

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
                ArchivoPDF.Flush()
                ArchivoPDF.Close()

                EscribeEnBitacora("Se creo PDF de Renta de Tractores")

                If File.Exists(vcArchivo) Then
                    Dim vcAdjuntos As New ArrayList()

                    vcAdjuntos.Add(vcArchivo)

                    EscribeEnBitacora("Se enviara correo de PDF de Renta de Tractores")

                    ' Enviamos correo
                    'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    flEnviarMail(vcCorreosRentaTractores, vcAdjuntos, "INFORME DE TRACTORISTAS DEL " & UCase(Format(vdFecha, "dd-MM-yy")), "SE ANEXA ARCHIVO")
                    'plEnviarMail("<enriqueca@aparedes.com.mx>", vcAdjuntos, "INFORME DE TRACTORISTAS DEL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    DAO.EjecutaComandoSQL("INSERT NOM_ENVIORENTATRACTORES SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                    EscribeEnBitacora("Se inserta en tabla de Avance de Labores")
                End If

                'If File.Exists(vcArchivo) Then
                '    File.Delete(vcArchivo)
                'End If

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosRentaTractores", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para Renta de Tractores")
            DAO.EjecutaComandoSQL("INSERT NOM_ENVIORENTATRACTORES SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de Renta de Tractores")
            Exit Sub
        End If

        Proc.Dispose()
        'Proc.Kill()
        oRpt = Nothing

    End Sub
    Class ResponseApi
        Public Property vendor As Integer
        Public Property growerDealNumber As String
        Public Property postDate As Date
        Public Property reference As String
        Public Property product As String
        Public Property lotNo As String
        Public Property palletID As Integer
        Public Property commodity As String
        Public Property variety As String
        Public Property packaging As String
        Public Property container As String
        Public Property size As String
        Public Property grade As String
        Public Property qty As Integer
        Public Property grossSales As Double
        Public Property netSales As Double
        Public Property totalAdjustments As Double
    End Class
    Private Sub plObtenDatosVentasDiarias()

        Dim Parametro1 As Integer
        ' optima companyid
        Dim Parametro2 As Integer
        ' company id
        Dim Parametro3 As Integer
        ' yearid
        Dim Parametro4 As Integer
        ' batchid
        Dim Parametro5 As Integer
        ' packer id
        Dim Parametro6 As Integer
        ' grower id
        Parametro1 = 2
        Parametro2 = 1
        'Parametro3 = 34
        'Parametro3 = 39
        'Parametro3 = 40 temp 20-21
        'Parametro3 = 41 ' temp 21-22
        Parametro3 = 42 ' temp 22-23
        Parametro5 = 125
        Parametro6 = 0

        Dim DTResultado As DataTable
        Dim DSResultado As DataSet

        Dim vcLiq As String = ""

        DTResultado = New DataTable()
        DSResultado = New DataSet()

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        Dim wsservice As New net.optimaproduce.webservices.Service()

        'Try
        wsservice.Url = "http://webservices.optimaproduce.net/GrowerService/service.asmx"

        '' Obteniendo Inventarios
        Parametro6 = 0

        EscribeEnBitacora("Obteniendo Ventas Diarias")

        Parametro4 = 125

        'Dim vnCiclo As Integer = -87

        'For vnCiclo = 87 To 1 Step -1

        '    Dim vcFechaNueva As String = Format(DateAdd(DateInterval.Day, vnCiclo * -1, DAO.RegresaFechaDelSistema.Date), "dd/MM/yyyy")

        '    DTResultado = Nothing

        '    DTResultado = wsservice.GetTodayInvoices(Parametro1, Parametro2, Parametro3, Parametro5, Parametro6, vcFechaNueva, 2)


        '    If Not fgGrabaVentasDiarias(DTResultado, fgObtenParametroEMB("TEMPORADA"), vcFechaNueva) Then
        '        EscribeEnBitacora("Ocurrio un error al Obtener las Ventas")
        '        Exit Sub
        '    End If

        '    Dim vdFechaNueva As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)


        'Next vnCiclo

        Dim vcFecha As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "dd/MM/yyyy")
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        DTResultado = Nothing
        DTResultado = wsservice.GetTodayInvoices(Parametro1, Parametro2, Parametro3, Parametro5, Parametro6, vcFecha, 2)

        If Not fgGrabaVentasDiarias(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha) Then
            EscribeEnBitacora("Ocurrio un error al Obtener las Ventas")
            Exit Sub
        End If


        Dim ds As DataSet = fgTraeVentasDiarias(vcTemporada, vdFecha)
        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

            'Try
            oRpt.Load("C:\CROP\RPT_VENTASDIARIASHMCLN.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFecha, oRpt)

            Dim oStrem As New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS HM AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            'Si lo deseamos escribimos el pdf a disco.
            Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            Dim vcAdjuntos As New ArrayList()

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            EscribeEnBitacora("Se creo PDF de Ventas Diarias de HM")

            Proc.Dispose()
            oRpt = Nothing

            'plObtenDatosVentasDiariasMarengo()
            Do While True
                If flObtenVentasWSMarengo() Then 'WS Ventas marengo 
                    Exit Do
                End If
            Loop

            vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            If vcAdjuntos.Count > 0 Then

                EscribeEnBitacora("Se enviara correo de PDF de Ventas Diarias")

                ' Enviamos correo
                flEnviarMail(vcCorreosVentas, vcAdjuntos, "CLN INFORME DE VENTAS DIARIAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")
                'flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "INFORME DE VENTAS DIARIAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOVENTASDIARIASDISTRIBUIDORAS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','" & sucursal & "'")

                EscribeEnBitacora("Se inserta en tabla de VENTAS DIARIAS")

            End If

            'If File.Exists(vcArchivo) Then
            '    File.Delete(vcArchivo)
            'End If

            'Catch ex As Exception
            'EscribeEnBitacora(ex.Message)
            'End Try
        Else

            Do While True
                If flObtenVentasWSMarengoV2() Then 'Consumo de api Marengo GrowerNetSales'
                    Exit Do
                End If
            Loop



            Dim vcAdjuntos As New ArrayList()
            Console.WriteLine("### Generación C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf")
            vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            If vcAdjuntos.Count > 0 Then

                EscribeEnBitacora("Se enviara correo de PDF de Ventas Diarias")

                ' Enviamos correo
                flEnviarMail(vcCorreosVentas, vcAdjuntos, "CLN INFORME DE VENTAS DIARIAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")
                'flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "INFORME DE VENTAS DIARIAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOVENTASDIARIASDISTRIBUIDORAS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','" & sucursal & "'")

                EscribeEnBitacora("Se inserta en tabla de VENTAS DIARIAS")

            End If


            EscribeEnBitacora("No hubo datos para ventas diarias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOVENTASDIARIASDISTRIBUIDORAS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','" & sucursal & "'")
            EscribeEnBitacora("Se inserta en tabla de ventas diarias")
            Exit Sub
        End If

        Proc.Dispose()
        oRpt = Nothing



        'Catch ex As Exception

        'EscribeEnBitacora(ex.Message)

        'End Try

        'DAO.CierraTransaccion()

    End Sub

    Private Sub plObtenDatosVentasDiariasMarengoAlterna()

        Try


            Dim xls_cn As New System.Data.OleDb.OleDbConnection
            Dim xls_cmd As New System.Data.OleDb.OleDbCommand
            Dim xls_reader As New System.Data.OleDb.OleDbDataAdapter

            Dim strExtension As String = ""
            Dim nombreXls As String
            Dim m_Excel As Microsoft.Office.Interop.Excel.Application

            Dim DTResultadoMarengo As New DataTable

            DTResultadoMarengo.Columns.Add("Customer", GetType(String))
            DTResultadoMarengo.Columns.Add("CommodityName", GetType(String))
            DTResultadoMarengo.Columns.Add("PackStyle", GetType(String))
            DTResultadoMarengo.Columns.Add("Label", GetType(String))
            DTResultadoMarengo.Columns.Add("Size", GetType(String))
            DTResultadoMarengo.Columns.Add("UoM", GetType(String))
            DTResultadoMarengo.Columns.Add("Qty", GetType(Double))
            DTResultadoMarengo.Columns.Add("Gross", GetType(Double))
            DTResultadoMarengo.Columns.Add("Adj", GetType(Double))
            DTResultadoMarengo.Columns.Add("Net", GetType(Double))
            DTResultadoMarengo.Columns.Add("UnitPrice", GetType(Double))
            DTResultadoMarengo.Columns.Add("SalesType", GetType(String))
            DTResultadoMarengo.Columns.Add("ShipDate", GetType(Date))


            Dim vcCustomer As String = ""
            Dim vcCommodityName As String = ""
            Dim vcPackStyle As String = ""
            Dim vcLabel As String = ""
            Dim vcSize As String = ""
            Dim vcUoM As String = ""
            Dim vnQty As Double = 0
            Dim vnGross As Double = 0
            Dim vnAdj As Double = 0
            Dim vnNet As Double = 0
            Dim vnUnitPrice As Double = 0
            Dim vcSalesType As String = ""
            Dim vnShipDate As Date = Now
            Dim vcSplit() As String
            Dim vcSplit2() As String
            Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)


            Dim lblArchivo As String = "C:\CROP\00029676.csv"

            nombreXls = Path.GetFileName(lblArchivo)
            strExtension = Path.GetExtension(lblArchivo)
            nombreXls = Strings.Replace(nombreXls, strExtension, "")

            If strExtension = ".csv" Then
                'MsgBox("es un archivo excel")
                If (File.Exists(lblArchivo)) Then
                    xls_cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + lblArchivo + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'"
                    'xls_cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.12.0;Data Source=" + xlsx + ";Extended Properties='Excel 12.0;HDR=YES'"
                    Using xls_cn

                        Dim dt As New DataTable("Datos")

                        m_Excel = CreateObject("Excel.Application")
                        m_Excel.Workbooks.Open(lblArchivo)
                        xls_cn.Open()
                        xls_cmd.CommandText = "SELECT * FROM [" & nombreXls & "$]"
                        xls_cmd.Connection = xls_cn
                        xls_reader.SelectCommand = xls_cmd

                        Dim da As New System.Data.OleDb.OleDbDataAdapter(xls_cmd)
                        da.Fill(dt)

                        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

                            Dim vnRenglones As Integer = dt.Rows.Count
                            Dim vnRenglon As Integer = 0

                            For Each vRowCiclo As DataRow In dt.Rows

                                vnRenglon += 1

                                If Not vnRenglon = vnRenglones Then
                                    Dim vcProducto As String = ""
                                    Dim vcTipoProducto As String = ""
                                    Dim vcTamaño As String = ""
                                    Dim vcEtiqueta As String = ""
                                    Dim vcTipoCarton As String = ""
                                    Dim vcCarton As String = ""

                                    vcSplit = Strings.Split(vRowCiclo("PRODUCT DESCRIPTION"), "-")

                                    vcProducto = vcSplit(0)
                                    vcTipoProducto = vcSplit(1)
                                    vcTamaño = vcSplit(2)
                                    vcEtiqueta = vcSplit(3)
                                    vcSplit2 = Strings.Split(vcSplit(4), "      ")
                                    vcTipoCarton = vcSplit2(0)
                                    vcCarton = vcSplit2(1)


                                    vcCustomer = ""
                                    vcCommodityName = fgObtenCommoditieName(vcProducto + "-" + vcTipoProducto)
                                    vcPackStyle = vcCarton
                                    vcLabel = vcEtiqueta
                                    vcSize = vcTamaño
                                    vcUoM = vcTipoCarton
                                    vnQty = vRowCiclo("QTY")
                                    vnGross = vRowCiclo("GROSS")
                                    vnAdj = vRowCiclo("ADJ")
                                    vnNet = vRowCiclo("NET")
                                    vnUnitPrice = vRowCiclo("AVG PRICE1")
                                    vcSalesType = ""
                                    vnShipDate = CDate(Strings.Mid(vRowCiclo("DATE"), 4, 2) & "/" & Strings.Left(vRowCiclo("DATE"), 2) & "/20" & Strings.Right(vRowCiclo("DATE"), 2))

                                    vdFecha = CDate(Strings.Mid(vRowCiclo("DATE"), 4, 2) & "/" & Strings.Left(vRowCiclo("DATE"), 2) & "/20" & Strings.Right(vRowCiclo("DATE"), 2))

                                    Dim vRow As DataRow

                                    vRow = DTResultadoMarengo.NewRow

                                    vRow("Customer") = vcCustomer
                                    vRow("CommodityName") = vcCommodityName
                                    vRow("PackStyle") = vcPackStyle
                                    vRow("Label") = vcLabel
                                    vRow("Size") = vcSize
                                    vRow("UoM") = vcUoM
                                    vRow("Qty") = vnQty
                                    vRow("Gross") = vnGross
                                    vRow("Adj") = vnAdj
                                    vRow("Net") = vnNet
                                    vRow("UnitPrice") = vnUnitPrice
                                    vRow("SalesType") = vcSalesType
                                    vRow("ShipDate") = vnShipDate

                                    DTResultadoMarengo.Rows.Add(vRow)

                                End If


                            Next

                            If Not DTResultadoMarengo Is Nothing AndAlso DTResultadoMarengo.Rows.Count > 0 Then

                                Dim vcResultado As String = ""
                                'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
                                Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
                                vcResultado = fgGrabaVentasDiariasMarengo(DTResultadoMarengo, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

                            End If
                        End If

                        m_Excel.Quit()
                        m_Excel = Nothing

                    End Using



                End If
            End If

        Catch ex As Exception
            'MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlterna", ex.Message)
        End Try


    End Sub


    Private Sub plObtenDatosVentasDiariasMarengoAlternaConexion()


        Dim DTResultado As DataTable
        Dim DSResultado As DataSet

        Dim vcLiq As String = ""

        DTResultado = New DataTable()
        DSResultado = New DataSet()
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String


        Dim miConexion As SqlConnection
        Dim miComando As SqlCommand
        Dim miAdapter As SqlDataAdapter
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -30, DAO.RegresaFechaDelSistema.Date)
        'vdFecha = "20/10/2017" 
        'vdFecha = "2017-10-20"
        Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        EscribeEnBitacora("Intenta conectarse al server de Marengo")

        'En caso de haber transacción abierta, se abre una nueva conexión solo para traernos los parámetros del Procedimiento almacenado.
        miConexion = New SqlConnection("SERVER= db.marengoserver.dyndns.org; Initial Catalog= PD; User= PDuser; Pwd= bd&CRYL3%mb#3E5%")
        miConexion.Open()

        EscribeEnBitacora("Se conecta al server de Marengo")

        miComando = New SqlCommand("pGrowerNetSalesPeriod", miConexion)
        miComando.CommandType = CommandType.StoredProcedure
        'miComando.CommandTimeout = 30
        miComando.CommandTimeout = 0

        miAdapter = New SqlDataAdapter(miComando)

        Try
            SqlCommandBuilder.DeriveParameters(miComando)
        Catch ex As Exception
            'MsgBox(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message)
        End Try

        miComando.Connection = miConexion

        Dim vParam(3) As Object
        Dim DSVentas As New DataSet

        vParam(0) = "00000035"
        vParam(1) = vdFecha
        vParam(2) = vdFechaFin
        vParam(3) = "D"

        CargaParametrosProcedimientoAlmacenado(miComando, vParam)

        Try

            EscribeEnBitacora("Se piden los datos al servidor de Marengo")

            miAdapter.Fill(DSVentas)
        Catch ex As Exception
            'MsgBox(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message)
        End Try

        If miConexion.State = System.Data.ConnectionState.Open Then
            miConexion.Close()
        End If

        miConexion.Dispose()

        Dim dt As New DataTable

        dt = DSVentas.Tables(0)


        'DAO.RegresaConsultaSQL("SELECT * FROM VTAMAR", dt)

        Dim DTResultadoMarengo As New DataTable

        DTResultadoMarengo.Columns.Add("Customer", GetType(String))
        DTResultadoMarengo.Columns.Add("CommodityName", GetType(String))
        DTResultadoMarengo.Columns.Add("PackStyle", GetType(String))
        DTResultadoMarengo.Columns.Add("Label", GetType(String))
        DTResultadoMarengo.Columns.Add("Size", GetType(String))
        DTResultadoMarengo.Columns.Add("UoM", GetType(String))
        DTResultadoMarengo.Columns.Add("Qty", GetType(Double))
        DTResultadoMarengo.Columns.Add("Gross", GetType(Double))
        DTResultadoMarengo.Columns.Add("Adj", GetType(Double))
        DTResultadoMarengo.Columns.Add("Net", GetType(Double))
        DTResultadoMarengo.Columns.Add("UnitPrice", GetType(Double))
        DTResultadoMarengo.Columns.Add("SalesType", GetType(String))
        DTResultadoMarengo.Columns.Add("ShipDate", GetType(Date))


        Dim vRowsError As DataRow()

        vRowsError = dt.Select("[NET SALES] IS NULL")

        vRowsError = dt.Select("QTY IS NULL")

        vRowsError = dt.Select("TOTAL_ADJUSTMENTS IS NULL")

        Dim vcCustomer As String = ""
        Dim vcCommodityName As String = ""
        Dim vcPackStyle As String = ""
        Dim vcLabel As String = ""
        Dim vcSize As String = ""
        Dim vcUoM As String = ""
        Dim vnQty As Double = 0
        Dim vnGross As Double = 0
        Dim vnAdj As Double = 0
        Dim vnNet As Double = 0
        Dim vnUnitPrice As Double = 0
        Dim vcSalesType As String = ""
        Dim vnShipDate As Date = Now

        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

            Dim vnRenglones As Integer = dt.Rows.Count
            Dim vnRenglon As Integer = 0

            For Each vRowCiclo As DataRow In dt.Rows

                Dim vcProducto As String = ""
                Dim vcTipoProducto As String = ""
                Dim vcTamaño As String = ""
                Dim vcEtiqueta As String = ""
                Dim vcTipoCarton As String = ""
                Dim vcCarton As String = ""

                If Not vRowCiclo("NET SALES") Is DBNull.Value Then
                    Try
                        Try
                            vcCustomer = ""
                            vcCommodityName = IIf(vRowCiclo("COMMODITY") Is DBNull.Value, "", vRowCiclo("COMMODITY").ToString.Trim) & " - " & IIf(vRowCiclo("VARIETY") Is DBNull.Value, "", vRowCiclo("VARIETY").ToString.Trim)
                            vcPackStyle = Strings.Left(Strings.Replace(vRowCiclo("PACKAGING"), " ", ""), 5).ToString.Trim
                            vcLabel = IIf(vRowCiclo("GRADE").ToString.Trim = "GRADE #1", 1, 2)
                            vcSize = Strings.Replace(vRowCiclo("SIZE"), " ", "")
                            vcUoM = IIf(vRowCiclo("CONTAINER").ToString.Trim = "CARTON", "CTN", "RPC")
                            vnQty = IIf(vRowCiclo("QTY") Is DBNull.Value, 0, vRowCiclo("QTY"))
                            vnGross = IIf(vRowCiclo("NET SALES") Is DBNull.Value, 0, vRowCiclo("NET SALES"))
                            vnAdj = IIf(vRowCiclo("TOTAL_ADJUSTMENTS") Is DBNull.Value, 0, vRowCiclo("TOTAL_ADJUSTMENTS"))
                            vnNet = IIf(vRowCiclo("NET SALES") Is DBNull.Value, 0, vRowCiclo("NET SALES")) - IIf(vRowCiclo("TOTAL_ADJUSTMENTS") Is DBNull.Value, 0, vRowCiclo("TOTAL_ADJUSTMENTS"))
                            vnUnitPrice = 0
                            If IIf(vRowCiclo("QTY") Is DBNull.Value, 0, vRowCiclo("QTY")) > 0 Then
                                vnUnitPrice = (IIf(vRowCiclo("NET SALES") Is DBNull.Value, 0, vRowCiclo("NET SALES")) - IIf(vRowCiclo("TOTAL_ADJUSTMENTS") Is DBNull.Value, 0, vRowCiclo("TOTAL_ADJUSTMENTS"))) / IIf(vRowCiclo("QTY") Is DBNull.Value, 0, vRowCiclo("QTY"))
                            End If
                            vcSalesType = ""
                            vnShipDate = Strings.Left(vRowCiclo("POST_DATE"), 4) & "-" & Strings.Mid(vRowCiclo("POST_DATE"), 5, 2) & "-" & Strings.Right(vRowCiclo("POST_DATE"), 2)
                            'vnShipDate = CDate(Strings.Right(vRowCiclo("POST_DATE"), 2) & "/" & Strings.Mid(vRowCiclo("POST_DATE"), 5, 2) & "/" & Strings.Left(vRowCiclo("POST_DATE"), 4))
                            vdFecha = vnShipDate

                            Dim vRow As DataRow

                            vRow = DTResultadoMarengo.NewRow

                            vRow("Customer") = vcCustomer
                            vRow("CommodityName") = vcCommodityName
                            vRow("PackStyle") = vcPackStyle
                            vRow("Label") = vcLabel
                            vRow("Size") = vcSize
                            vRow("UoM") = vcUoM
                            vRow("Qty") = vnQty
                            vRow("Gross") = vnGross
                            vRow("Adj") = vnAdj
                            vRow("Net") = vnNet
                            vRow("UnitPrice") = vnUnitPrice
                            vRow("SalesType") = vcSalesType
                            vRow("ShipDate") = vnShipDate

                            DTResultadoMarengo.Rows.Add(vRow)

                        Catch ex As Exception
                            'MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
                            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message)
                        End Try


                    Catch ex As Exception
                        'MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
                        flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message)
                    End Try

                End If

            Next


            If Not DTResultadoMarengo Is Nothing AndAlso DTResultadoMarengo.Rows.Count > 0 Then

                Dim vcResultado As String = ""
                vdFecha = DateAdd(DateInterval.Day, -30, DAO.RegresaFechaDelSistema.Date)

                vcResultado = fgGrabaVentasDiariasMarengo(DTResultadoMarengo, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

            End If


            EscribeEnBitacora("Se inserto la informacion de Ventas Marengo con Exito")

        End If


        Dim dsMarengo As DataSet = fgTraeVentasDiariasMarengo(vcTemporada, vdFechaFin)

        If Not dsMarengo Is Nothing AndAlso dsMarengo.Tables.Count > 0 AndAlso dsMarengo.Tables(0).Rows.Count > 0 Then
            oRpt = New ReportDocument
            oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGO.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFechaFin, oRpt)

            Dim oStremMarengo As New System.IO.MemoryStream

            Try
                oStremMarengo = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\ARCHIVOS\VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & ".pdf"

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDFMarengo As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDFMarengo.Write(oStremMarengo.ToArray, 0, oStremMarengo.ToArray.Length)
                ArchivoPDFMarengo.Flush()
                ArchivoPDFMarengo.Close()
                ArchivoPDFMarengo.Dispose()
                ArchivoPDFMarengo = Nothing

                Proc.Dispose()
                oRpt = Nothing

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message)
            End Try
        End If


    End Sub

    'Private Sub plObtenDatosVentasDiariasMarengo()



    '    Dim DTResultado As DataTable
    '    Dim DSResultado As DataSet

    '    Dim vcLiq As String = ""

    '    DTResultado = New DataTable()
    '    DSResultado = New DataSet()

    '    Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", "001")
    '    Dim oRpt = New ReportDocument
    '    Dim Proc As New System.Diagnostics.Process
    '    Dim vcAdjuntosPDF As New ArrayList
    '    Dim vcArchivo As String

    '    Try

    '        EscribeEnBitacora("Obteniendo Ventas Diarias Marengo")

    '        'For vnCiclo As Integer = -180 To 0


    '        Dim vcResultado As String = ""
    '        Dim DTResultadoMarengo As New DataTable
    '        Dim wsserviceMarengo As New wsMarengo.GrowerService
    '        Dim Parametro1Marengo As String = ""
    '        Dim Parametro2Marengo As String = ""

    '        DTResultadoMarengo.Columns.Add("Grower", GetType(String))
    '        DTResultadoMarengo.Columns.Add("Customer", GetType(String))
    '        DTResultadoMarengo.Columns.Add("CommodityName", GetType(String))
    '        DTResultadoMarengo.Columns.Add("PackStyle", GetType(String))
    '        DTResultadoMarengo.Columns.Add("Label", GetType(String))
    '        DTResultadoMarengo.Columns.Add("Size", GetType(String))
    '        DTResultadoMarengo.Columns.Add("UoM", GetType(String))
    '        DTResultadoMarengo.Columns.Add("Qty", GetType(Double))
    '        DTResultadoMarengo.Columns.Add("Gross", GetType(Double))
    '        DTResultadoMarengo.Columns.Add("Adj", GetType(Double))
    '        DTResultadoMarengo.Columns.Add("Net", GetType(Double))
    '        DTResultadoMarengo.Columns.Add("UnitPrice", GetType(Double))
    '        DTResultadoMarengo.Columns.Add("SalesType", GetType(String))
    '        DTResultadoMarengo.Columns.Add("ShipDate", GetType(Date))

    '        Try
    '            wsserviceMarengo.Url = "https://www.marengosite.com/marengowebservices/growerservice.asmx"

    '            '' Obteniendo Inventarios
    '            Parametro1Marengo = "001"
    '            Parametro2Marengo = "p@r3d3s"


    '            Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

    '            Dim vObjMarengo() As wsMarengo.SalesByDate

    '            vObjMarengo = wsserviceMarengo.SalesByDate(Parametro1Marengo, Parametro2Marengo, vdFecha)

    '            For i As Integer = 0 To vObjMarengo.Length - 1
    '                Dim vRow As DataRow

    '                vRow = DTResultadoMarengo.NewRow

    '                vRow("Grower") = vObjMarengo(i).Grower
    '                vRow("Customer") = vObjMarengo(i).Customer
    '                vRow("CommodityName") = vObjMarengo(i).CommodityName
    '                vRow("PackStyle") = vObjMarengo(i).PackStyle
    '                vRow("Label") = vObjMarengo(i).Label
    '                vRow("Size") = vObjMarengo(i).Size
    '                vRow("UoM") = vObjMarengo(i).UoM
    '                vRow("Qty") = vObjMarengo(i).Qty
    '                vRow("Gross") = vObjMarengo(i).Gross
    '                vRow("Adj") = vObjMarengo(i).Adj
    '                vRow("Net") = vObjMarengo(i).Net
    '                vRow("UnitPrice") = vObjMarengo(i).UnitPrice
    '                vRow("SalesType") = vObjMarengo(i).SalesType
    '                vRow("ShipDate") = vObjMarengo(i).ShipDate

    '                DTResultadoMarengo.Rows.Add(vRow)

    '            Next


    '            EscribeEnBitacora("Se obtiene la información de existencias del webservice")

    '            vcResultado = fgGrabaVentasDiariasMarengo(DTResultadoMarengo, fgObtenParametroEMB("TEMPORADA", "001"), vdFecha)

    '            If vcResultado <> "" Then
    '                EscribeEnBitacora(vcResultado)
    '                Exit Sub
    '            End If

    '            EscribeEnBitacora("Se inserto la informacion de Ventas Marengo con Exito")

    '            Dim dsMarengo As DataSet = fgTraeVentasDiariasMarengo(vcTemporada, vdFecha)

    '            If Not dsMarengo Is Nothing AndAlso dsMarengo.Tables.Count > 0 AndAlso dsMarengo.Tables(0).Rows.Count > 0 Then
    '                oRpt.Load("C:\RPT_VENTASDIARIASMARENGO.rpt")

    '                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
    '                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
    '                AgregarParametro("@PRMFECHA", vdFecha, oRpt)

    '                Dim oStremMarengo As New System.IO.MemoryStream

    '                Try
    '                    oStremMarengo = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

    '                    vcArchivo = "C:\ARCHIVOS\VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

    '                    'Si lo deseamos escribimos el pdf a disco.
    '                    Dim ArchivoPDFMarengo As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
    '                    ArchivoPDFMarengo.Write(oStremMarengo.ToArray, 0, oStremMarengo.ToArray.Length)
    '                    ArchivoPDFMarengo.Flush()
    '                    ArchivoPDFMarengo.Close()
    '                    ArchivoPDFMarengo.Dispose()
    '                    ArchivoPDFMarengo = Nothing

    '                    Proc.Dispose()
    '                    oRpt = Nothing

    '                Catch ex As Exception
    '                    EscribeEnBitacora(ex.Message)
    '                End Try



    '            End If



    '        Catch ex As Exception
    '            EscribeEnBitacora(ex.Message)
    '        End Try

    '        'Next




    '        Proc.Dispose()
    '        oRpt = Nothing



    '    Catch ex As Exception

    '        EscribeEnBitacora(ex.Message)

    '    End Try

    'End Sub

    Private Sub plObtenDatosExistenciasPiso()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As New System.Text.StringBuilder()


        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        Dim vldirec As String = "C:\CROP\CLN EXISTENCIANACIONAL_" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            xlLibro = xlApp.Workbooks.Open("C:\CROP\NACIONAL DISPONIBLE V2.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")

            'xlApp.Visible = True

            With xlHoja


                '' BOLA CHELITA 3X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '006'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(5, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 3X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '006'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")


                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(5, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BOLA CHELITA 4X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '007'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(6, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 4X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '007'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(6, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA CHELITA 4X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '008'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(7, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 4X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '008'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(7, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA CHELITA 5X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '009'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(8, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 5X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '009'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(8, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA CHELITA 5X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '010'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(9, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 5X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '010'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(9, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BOLA CHELITA 6X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '011'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(10, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 6X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '011'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(10, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA CHELITA 6X7
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '012'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(11, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA CHELITA 6X7
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '012'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(11, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA MEZQUITILLO 3X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '006'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(16, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA MEZQUITILLO 3X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '006'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(16, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA MEZQUITILLO 4X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '007'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(17, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA MEZQUITILLO 4X4
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '007'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(17, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA MEZQUITILLO 4X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '008'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(18, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' EMPAQUE BOLA MEZQUITILLO 4X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '008'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(18, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BOLA MEZQUITILLO 5X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '009'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(19, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA MEZQUITILLO 5X5
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '009'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(19, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' BOLA MEZQUITILLO 5X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '010'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(20, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' EMPAQUE BOLA MEZQUITILLO 5X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '010'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(20, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BOLA MEZQUITILLO 6X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '011'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(21, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BOLA MEZQUITILLO 6X6
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '011'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(21, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' BOLA MEZQUITILLO 6X7
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '012'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(22, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' EMPAQUE BOLA MEZQUITILLO 6X7
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '002' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '012'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(22, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' ROMA CHELITA SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(28, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA CHELITA SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(28, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA CHELITA JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(29, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA CHELITA JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(29, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable




                '' ROMA CHELITA XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(30, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA CHELITA XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(30, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' ROMA CHELITA LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(31, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA CHELITA LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(31, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA CHELITA MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(32, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA CHELITA MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(32, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' ROMA CHELITA SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(33, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' EMPAQUE ROMA CHELITA SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '002'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(33, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' ROMA SPV SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(38, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA SPV SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(38, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA SPV JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(39, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA SPV JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(39, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA SPV XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(40, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA SPV XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(40, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA SPV LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(41, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE SPV LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(41, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA SPV MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(42, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA SPV MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(42, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' ROMA SPV SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(43, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA SPV SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(43, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable





                '' ROMA MEZQUITILLO SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(49, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA MEZQUITILLO SJMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '033'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(49, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA MEZQUITILLO JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(50, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA MEZQUITILLO JMB
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '032'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(50, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA MEZQUITILLO XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(51, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA MEZQUITILLO XLG
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '002'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(51, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA MEZQUITILLO LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(52, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE MEZQUITILLO LGE
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '003'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(52, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' ROMA MEZQUITILLO MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(53, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA MEZQUITILLO MED
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '004'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(53, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable



                '' ROMA MEZQUITILLO SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(54, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE ROMA MEZQUITILLO SML
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '004' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '009'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '005'")
                vcSQL.AppendLine("AND E.CCVE_ENVASE = '016'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(54, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BERENJENA SPV 18's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '025'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(62, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA SPV 18's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '025'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(62, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BERENJENA SPV 24's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '026'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(63, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA SPV 24's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '026'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(63, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BERENJENA SPV 32's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '027'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(64, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA SPV 32's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '027'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(64, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BERENJENA SPV 40's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '029'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(65, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA SPV 40's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '001'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '029'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(65, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BERENJENA PARIS 18's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '025'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(70, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA PARIS 18's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '025'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(70, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BERENJENA PARIS 24's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '026'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(71, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' EMPAQUE BERENJENA PARIS 24's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '026'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(71, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' BERENJENA PARIS 32's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '027'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(72, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA PARIS 32's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '027'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(72, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable


                '' BERENJENA PARIS 40's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS = 'A' AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '029'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(73, "A").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

                '' EMPAQUE BERENJENA PARIS 40's
                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT SUM(COALESCE(E.NPALLETS,0)) AS NPALLETS")
                vcSQL.AppendLine("FROM EYE_EMPAQUE E(NOLOCK) WHERE E.CCVE_CULTIVO = '005' AND E.CCVE_TEMPORADA = '" & vcTemporada & "' AND E.CSTATUS IN ('A','E') AND E.CCVE_ETIQUETA = '003'")
                vcSQL.AppendLine("AND E.CCVE_TAMAÑO = '029'")
                vcSQL.AppendLine("AND CONVERT(VARCHAR(20),E.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "'")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then
                    .Cells(73, "C").Value = Val(IIf(DTCabecero.Rows(0)("NPALLETS") Is DBNull.Value, 0, DTCabecero.Rows(0)("NPALLETS")))
                End If

                DTCabecero = New DataTable

            End With

            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosExistenciasPiso", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Piso Nacional")

                ' Enviamos correo
                flEnviarMail(vcCorreosAvanceLabores, vcAdjuntos, "CLN NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOEXISTENCIASPISO SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Existencias Piso")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosExistenciasPiso", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para existencias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOEXISTENCIASPISO SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

    End Sub

    Private Sub plObtenEmpleadosMayoresDeEdad()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As New System.Text.StringBuilder()

        Dim vnRenglon As Integer

        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        Dim vldirec As String = "C:\CROP\ELIGIBLESTARJETA" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            xlLibro = xlApp.Workbooks.Open("C:\CROP\ELEGIBLESTARJETA.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")

            'xlApp.Visible = True

            With xlHoja

                vnRenglon = 8

                Dim DTEmpleado As DataSet = fgTraeEmpleadosMayoresDeEdad(vcTemporada)

                If Not DTEmpleado Is Nothing AndAlso DTEmpleado.Tables.Count > 0 AndAlso DTEmpleado.Tables(0).Rows.Count > 0 Then
                    .Cells(5, "B").Value = vdFecha
                    For Each vRow As DataRow In DTEmpleado.Tables(0).Rows
                        .Cells(vnRenglon, "B").Value = vRow("CCLAVE")
                        .Cells(vnRenglon, "C").Value = vRow("CNOMBRE")
                        .Cells(vnRenglon, "D").Value = vRow("NEDAD")
                        .Cells(vnRenglon, "E").Value = vRow("DD/MM/YY")

                        vnRenglon += 1
                    Next
                End If

                DTCabecero = New DataTable

            End With
            If File.Exists(vldirec) Then
                File.Delete(vldirec)
            End If
            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenEmpleadosMayoresDeEdad", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Piso Nacional")

                ' Enviamos correo
                flEnviarMail(vcCorreoElegiblesTargetas, vcAdjuntos, "CLN EMPLEADOS ELEGIBLES PARA TARJETA" & UCase(Format(vdFecha, "dd-MMM-yy")), "")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT NOM_ELEGIBLESTARJETA SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Existencias ELIGIBLESTARJETA")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenEmpleadosMayoresDeEdad", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para existencias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOEXISTENCIASPISO SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

    End Sub

    Private Sub plObtenPorcentajeCalidad()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As String = ""
        Dim vnRenglon As Integer = 0
        Dim vnRenglonCultivoInicio As Integer = 0
        Dim vnRenglonCultivoFin As Integer = 0
        Dim vcCultivo As String = ""


        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        Dim vldirec As String = "C:\CROP\CLN PORCENTAJECALIDADES_" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            xlLibro = xlApp.Workbooks.Open("C:\CROP\PORCENTAJESCALIDAD.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")

            'xlApp.Visible = True

            With xlHoja


                Dim vParam(8) As Object
                Dim DS As New DataSet

                vParam(0) = vcTemporada
                vParam(1) = "001"
                vParam(2) = ""
                vParam(3) = vdFecha
                vParam(4) = vdFecha

                vcSQL = "SPPORCENTAJESEMPAQUECULTIVOSLOTES"

                Try
                    If Not DAO.RegresaConsultaSQL(vcSQL, DS, vParam) Then
                        Exit Sub
                    End If
                Catch ex As Exception
                    flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenPorcentajeCalidad", ex.Message)
                    Exit Sub
                End Try

                If DS Is Nothing OrElse DS.Tables.Count = 0 Then
                    Exit Sub
                End If

                DTCabecero = DS.Tables(0)

                vnRenglon = 4

                .Cells(vnRenglon, "A").Value = "CULTIVO: TODOS LOS CULTIVOS AL: " & Format(vdFecha, "dd/MMM/yy")

                vnRenglon = 7


                For Each vRow As DataRow In DTCabecero.Rows

                    If vcCultivo = "" Then
                        vcCultivo = vRow("CULTIVO")
                        '.Cells(vnRenglon, "A").Value = vRow("CULTIVO")
                        plLlenaValoresCeldaCalidad(vnRenglon, 1, xlHoja, vRow("CULTIVO"), False)

                        vnRenglonCultivoInicio = vnRenglon
                    Else
                        If vcCultivo <> vRow("CULTIVO") Then

                            '' Aqui debo de Aplicar los efectos

                            vnRenglonCultivoFin = vnRenglon - 1


                            Dim objRango = xlHoja.Range(flLetraExcel(1) & vnRenglonCultivoInicio & ":" & flLetraExcel(1) & vnRenglonCultivoFin)

                            objRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                            With objRango.Cells
                                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                                .WrapText = True
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = Fix(Excel.Constants.xlLTR)
                                .MergeCells = False
                                .Merge()
                            End With

                            plBordesGruesos(vnRenglonCultivoInicio, vnRenglonCultivoFin, 1, 1, xlHoja)
                            plBordesGruesos(vnRenglonCultivoInicio, vnRenglonCultivoFin, 2, 7, xlHoja)


                            vcCultivo = vRow("CULTIVO")
                            .Cells(vnRenglon, "A").Value = vRow("CULTIVO")

                            vnRenglonCultivoInicio = vnRenglon
                        End If
                    End If

                    plLlenaValoresCeldaCalidad(vnRenglon, 2, xlHoja, vRow("LOTE"), False)
                    '.Cells(vnRenglon, "B").Value = vRow("LOTE")

                    plLlenaValoresCeldaCalidad(vnRenglon, 3, xlHoja, vRow("EMPAQUE"))
                    plLlenaValoresCeldaCalidad(vnRenglon, 4, xlHoja, vRow("PRIMERAS"))
                    plLlenaValoresCeldaCalidad(vnRenglon, 5, xlHoja, vRow("SEGUNDAS"))
                    plLlenaValoresCeldaCalidad(vnRenglon, 6, xlHoja, vRow("TERCERAS"))
                    plLlenaValoresCeldaCalidad(vnRenglon, 7, xlHoja, vRow("REZAGA"))

                    '.Cells(vnRenglon, "C").Value = vRow("EMPAQUE")
                    '.Cells(vnRenglon, "D").Value = vRow("PRIMERAS")
                    '.Cells(vnRenglon, "E").Value = vRow("SEGUNDAS")
                    '.Cells(vnRenglon, "F").Value = vRow("TERCERAS")
                    '.Cells(vnRenglon, "G").Value = vRow("REZAGA")

                    vnRenglon += 1

                Next

                vnRenglonCultivoFin = vnRenglon - 1

                Dim objRangoFinal = xlHoja.Range(flLetraExcel(1) & vnRenglonCultivoInicio & ":" & flLetraExcel(1) & vnRenglonCultivoFin)

                objRangoFinal.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                With objRangoFinal.Cells
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Fix(Excel.Constants.xlLTR)
                    .MergeCells = False
                    .Merge()
                End With


                plBordesGruesos(vnRenglonCultivoInicio, vnRenglonCultivoFin, 1, 1, xlHoja)
                plBordesGruesos(vnRenglonCultivoInicio, vnRenglonCultivoFin, 2, 7, xlHoja)


                DTCabecero = Nothing
                DS = Nothing

            End With

            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenPorcentajeCalidad", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Porcentajes de Calidad por Lote")

                ' Enviamos correo
                flEnviarMail(vcCorreosPorcEmpaqueCalidad, vcAdjuntos, "CLN PORCENTAJES CALIDAD POR LOTE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOPORCCALIDAD SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Existencias Piso")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenPorcentajeCalidad", ex.Message)
            End Try

        Else
            EscribeEnBitacora("No hubo datos para existencias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOPORCCALIDAD SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

    End Sub


    Private Sub plObtenChequesSinFacturasCln()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As New System.Text.StringBuilder()
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vnRenglon As Integer

        Dim vldirec As String = "C:\CROP\CLN CHEQUESSINFACTURA_" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            xlLibro = xlApp.Workbooks.Open("C:\CROP\RELACION DE CHEQUES.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")

            'xlApp.Visible = True

            With xlHoja


                vnRenglon = 5

                .Cells(vnRenglon, "B").Value = "AL " & Format(vdFecha, "dd-MM-yyyy")

                vnRenglon = 8

                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT * FROM SERVIDOR.PAREDES0708.DBO.VWCHEQUESSINFACTURACLN ORDER BY DFECHA,CCHEQUE")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then

                    For Each vRow As DataRow In DTCabecero.Rows

                        .Cells(vnRenglon, "B").Value = vRow("CPROVEEDOR")
                        .Cells(vnRenglon, "C").Value = vRow("CDESCRIP")
                        .Cells(vnRenglon, "D").Value = vRow("CCHEQUE")
                        .Cells(vnRenglon, "E").Value = vRow("DFECHA")
                        .Cells(vnRenglon, "F").Value = vRow("NIMPORTECHEQUE")
                        .Cells(vnRenglon, "G").Value = vRow("NIMPORTEFACT")
                        .Cells(vnRenglon, "H").Value = vRow("NFACTURAS")
                        .Cells(vnRenglon, "I").Value = vRow("COBSERVACIONES")
                        .Cells(vnRenglon, "J").Value = vRow("CDESCCUENTA")
                        .Cells(vnRenglon, "K").Value = vRow("DFECHA_COBRADO")

                        vnRenglon += 1
                    Next

                End If

                DTCabecero = New DataTable

            End With

            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenChequesSinFacturasCln", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Cheques sin Facturas asignadas")

                ' Enviamos correo
                flEnviarMail(vcCorreosChequesCln, vcAdjuntos, "CLN CHEQUES SIN FACTURAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'flEnviarMail("<enriqueca@aparedes.com.mx>", vcAdjuntos, "CLN CHEQUES SIN FACTURAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOCHEQUESCLN SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Cheques Cln Piso")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenChequesSinFacturasCln", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para existencias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOCHEQUESCLN SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

    End Sub

    Private Sub plObtenChequesSinFacturasJal()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As New System.Text.StringBuilder()
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vnRenglon As Integer

        Dim vldirec As String = "C:\CROP\JAL CHEQUESSINFACTURA_" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            xlLibro = xlApp.Workbooks.Open("C:\CROP\RELACION DE CHEQUES.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")

            'xlApp.Visible = True

            With xlHoja


                vnRenglon = 5

                .Cells(vnRenglon, "B").Value = "AL " & Format(vdFecha, "dd-MM-yyyy")

                vnRenglon = 8

                vcSQL = New System.Text.StringBuilder()

                vcSQL.AppendLine("SELECT * FROM SERVIDOR.PAREDES0708.DBO.VWCHEQUESSINFACTURAJAL ORDER BY DFECHA,CCHEQUE")

                DAO.RegresaConsultaSQL(vcSQL.ToString, DTCabecero)

                If Not DTCabecero Is Nothing AndAlso DTCabecero.Rows.Count > 0 Then

                    For Each vRow As DataRow In DTCabecero.Rows

                        .Cells(vnRenglon, "B").Value = vRow("CPROVEEDOR")
                        .Cells(vnRenglon, "C").Value = vRow("CDESCRIP")
                        .Cells(vnRenglon, "D").Value = vRow("CCHEQUE")
                        .Cells(vnRenglon, "E").Value = vRow("DFECHA")
                        .Cells(vnRenglon, "F").Value = vRow("NIMPORTECHEQUE")
                        .Cells(vnRenglon, "G").Value = vRow("NIMPORTEFACT")
                        .Cells(vnRenglon, "H").Value = vRow("NFACTURAS")
                        .Cells(vnRenglon, "I").Value = vRow("COBSERVACIONES")
                        .Cells(vnRenglon, "J").Value = vRow("CDESCCUENTA")
                        .Cells(vnRenglon, "K").Value = vRow("DFECHA_COBRADO")

                        vnRenglon += 1
                    Next

                End If

                DTCabecero = New DataTable

            End With

            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenChequesSinFacturasJal", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Cheques sin Facturas asignadas")

                ' Enviamos correo
                flEnviarMail(vcCorreosChequesJal, vcAdjuntos, "JAL CHEQUES SIN FACTURAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'flEnviarMail("<enriqueca@aparedes.com.mx>", vcAdjuntos, "CLN CHEQUES SIN FACTURAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOCHEQUESJAL SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Cheques Jal")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenChequesSinFacturasJal", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para existencias")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOCHEQUESJAL SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

    End Sub

    Private Sub plObtenDatosRotacionDistribuidoras()

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet

        Dim vcSQL As New System.Text.StringBuilder()
        Dim vnCiclo As Integer = 0
        Dim vnCicloNormal As Integer = 0
        Dim vnRenglonInicial As Integer = 5
        Dim vnRenglon As Integer = 0
        Dim DT As New DataTable

        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -8, DAO.RegresaFechaDelSistema.Date)

        Dim vldirec As String = "C:\CROP\CLN ROTACIONDISTRIBUIDORAS_" & Format(vdFecha, "dd-MM-yyyy") & ".xlsx"
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        Dim DTCabecero As New DataTable

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Try

            If File.Exists(vldirec) Then
                Kill(vldirec)
            End If


            xlLibro = xlApp.Workbooks.Open("C:\CROP\ROTACION DE INVENTARIO DIARIO.xlsx")
            xlHoja = xlLibro.Worksheets.Application.Sheets("Sheet1")

            'xlApp.Visible = True

            With xlHoja


                For vnCiclo = 0 To 14 Step 2
                    .Cells(vnRenglonInicial + vnCiclo, "A") = DAO.RegresaDatoSQL("SELECT UPPER(DATENAME(dw,'" & Format(DateAdd(DateInterval.Day, vnCicloNormal, vdFechaFin), "yyyyMMdd") & "'))")
                    .Cells(vnRenglonInicial + vnCiclo, "B") = DateAdd(DateInterval.Day, vnCicloNormal, vdFechaFin)
                    vnCicloNormal += 1
                Next

                '' BERENJENAS

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "005", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "D").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "D").Value = vRows(0)("NPORCMAR") / 100


                        End If

                    Next

                End If


                '' CH VERDE

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "003", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "E").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "E").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If


                '' CH ROJO

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "008", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "F").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "F").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If


                '' ROMAS

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "004", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "G").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "G").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If

                '' BOLAS

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "002", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "H").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "H").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If

                '' CH AMARILLO

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "061", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "I").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "I").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If


                '' CH NARANJA

                DT = fgTraeRotacionDistribuidoras(vcTemporada, "010", vdFechaFin, vdFecha)

                If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then


                    For vnCiclo = 0 To 14 Step 2

                        Dim vRows() As DataRow

                        vnRenglon = vnRenglonInicial

                        vRows = DT.Select("CFECHA = " & Format(xlHoja.Cells(vnRenglon + vnCiclo, "B").Value(), "yyyyMMdd"))

                        If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                            .Cells(vnRenglon + vnCiclo, "J").Value = vRows(0)("NPORCHM") / 100
                            vnRenglon += 1
                            .Cells(vnRenglon + vnCiclo, "J").Value = vRows(0)("NPORCMAR") / 100

                        End If

                    Next

                End If


                DT = New DataTable

            End With

            xlHoja.SaveAs(vldirec)
            xlLibro.Close()
            xlApp.Quit()

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosRotacionDistribuidoras", ex.Message)
        End Try

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing

        If File.Exists(vldirec) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vldirec)

                EscribeEnBitacora("Se enviara correo de Excel de Rotacion de Distribuidoras")

                ' Enviamos correo
                flEnviarMail(vcCorreosAvanceLabores, vcAdjuntos, "CLN ROTACION DE PRODUCTOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOROTACIONPRODUCTOS SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Rotacion de Distribuidoras")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosRotacionDistribuidoras", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para Rotacion de Productos")
            DAO.EjecutaComandoSQL("INSERT EYE_ENVIOROTACIONPRODUCTOS SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")
            EscribeEnBitacora("Se inserta en tabla de Rotacion de Productos")
            Exit Sub
        End If

    End Sub

    Private Sub plObtenDatos()
        Console.WriteLine(vbCrLf & "## Inicia plObtenDatos() ## " & vbCrLf)
        Dim Parametro1 As Integer
        ' optima companyid
        Dim Parametro2 As Integer
        ' company id
        Dim Parametro3 As Integer
        ' yearid
        Dim Parametro4 As Integer
        ' batchid
        Dim Parametro5 As Integer
        ' packer id
        Dim Parametro6 As Integer
        ' grower id
        Parametro1 = 2
        Parametro2 = 1
        'Parametro3 = 36
        Parametro3 = 42
        'Parametro5 = 125
        Parametro5 = 125
        Parametro6 = 0

        Dim DTResultado As DataTable
        Dim DSResultado As DataSet

        Dim vcLiq As String = ""

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-MX")
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String


        DTResultado = New DataTable()
        DSResultado = New DataSet()

        Dim vcResultado As String = ""

        Dim wsservice As New net.optimaproduce.webservices.Service()
        'Try
        wsservice.Url = "http://webservices.optimaproduce.net/GrowerService/service.asmx" '' fgTraeExistenciasDistribuidoras() HM

        '' Obteniendo Inventarios
        Parametro6 = 0

        Parametro4 = 125

        Dim vnDiaSemana As Integer = DAO.RegresaDatoSQL("SELECT DATEPART(dw,GETDATE()) ")

        Dim vcFecha As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "dd/MM/yyyy")
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        DTResultado = Nothing

        '' Por mientras se tomara siempre datos del webservice esto aplica solamente para la distribuidora HM
        'vnDiaSemana = 4
        If vnDiaSemana <> 1 Then
            DTResultado = wsservice.GetInventoryActivity(Parametro1, Parametro2, vcFecha, Parametro4, Parametro5, Parametro6, "E")

            EscribeEnBitacora("Se obtiene la información de existencias del webservice")

            vcResultado = fgGrabainventarios(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha) 'Inserta inventario para distribuidora HM

            If vcResultado <> "" Then
                EscribeEnBitacora(vcResultado)
                Exit Sub
            End If

            EscribeEnBitacora("Se inserto la informacion de Inventarios con Exito")
        Else
            EscribeEnBitacora("Se tomaran las existencias del dia sabado anterior al cierre")

            Dim vdFechaAnt As Date = DateAdd(DateInterval.Day, -2, DAO.RegresaFechaDelSistema.Date)
            Dim vdFechaActual As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

            Dim vDs As New DataSet
            Dim vParamDetalle(2) As Object
            Dim vlSQL As String

            vlSQL = "SPACTUALIZAEXISTENCIASHM"

            vParamDetalle(0) = vcTemporada
            vParamDetalle(1) = vdFechaAnt
            vParamDetalle(2) = vdFechaActual

            If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamDetalle) Then
                EscribeEnBitacora("Ocurrio un error al actualizar los inventarios de HM")
            End If

        End If

        ' Inicia Proceso de Inventarios y Ventas Diarias de Api Marengo 
        Do While True
            If flObtenExistenciaMarengoWS() Then 'Flujo Transactions
                Exit Do
            End If
        Loop


        Dim ds As DataSet = fgTraeExistenciasDistribuidoras(vcTemporada, vdFecha)
        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Try
                oRpt = New ReportDocument
                Dim CLN As String = "001"
                Dim oStrem As New System.IO.MemoryStream

                If sucursal = CLN Then 'CULIACAN
                    oRpt.Load("C:\CROP\RPT_INVDISTRIBUIDORASCLN.rpt")

                    LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                    AgregarParametro("@SUCURSAL", "CULIACAN, SIN.", oRpt)
                    AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                    AgregarParametro("@PRMFECHA", vdFecha, oRpt)

                    vcArchivo = "C:\ARCHIVOS\CLN EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"
                    oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)
                Else 'JALISCO
                    oRpt.Load("C:\CROP\RPT_INVDISTRIBUIDORASJAL.rpt")

                    LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                    AgregarParametro("@SUCURSAL", "JALISCO.", oRpt)
                    AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                    AgregarParametro("@PRMFECHA", vdFecha, oRpt)

                    vcArchivo = "C:\ARCHIVOS\JAL EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"
                    oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)
                End If

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
                ArchivoPDF.Flush()
                ArchivoPDF.Close()

                EscribeEnBitacora("Se creo PDF de Existencias")

                If File.Exists(vcArchivo) Then
                    Dim vcAdjuntos As New ArrayList()

                    vcAdjuntos.Add(vcArchivo)

                    EscribeEnBitacora("Se enviara correo de PDF de Existencias")

                    ' Enviamos correo
                    'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                    'plEnviarMail(vcCorreosVentas, vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    'Dim vbResultado As Boolean = False

                    ' Do While True

                    'vbResultado = flEnviarMail(vcCorreosVentas, vcAdjuntos, "CLN EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                    'vbResultado = flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    'If vbResultado Then
                    'Exit Do
                    '  End If

                    ' Loop

                    DAO.EjecutaComandoSQL("INSERT EYE_ENVIOEXISTENCIASDISTRIBUIDORAS SELECT '" & Format(vdFecha, "yyyyMMdd") & "','" & sucursal & "'")

                    EscribeEnBitacora("Se inserta en tabla de EYE_ENVIOEXISTENCIASDISTRIBUIDORAS")
                End If

                'If File.Exists(vcArchivo) Then
                '    File.Delete(vcArchivo)
                'End If

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatos", ex.Message)
            End Try

        Else
            EscribeEnBitacora("No hubo datos para existencias")
            EscribeEnBitacora("Se inserta en tabla de envio de existencias")
            Exit Sub
        End If

        Proc.Dispose()
        oRpt = Nothing


        'GeneraExcelExistencias()

        'Catch ex As Exception
        'EscribeEnBitacora(ex.Message)
        'End Try

    End Sub

    Private Sub CargaParametrosProcedimientoAlmacenado(ByRef prmComando As SqlCommand, ByVal prmParametros() As Object)
        For i As Int32 = 1 To prmComando.Parameters.Count - 1
            Dim miParametro As SqlParameter = prmComando.Parameters(i)
            If i > prmParametros.Length Then
                miParametro.Value = DBNull.Value
            Else
                miParametro.Value = prmParametros(i - 1)
            End If
        Next
    End Sub


    'Private Sub plObtenExistenciaMarengo(ByVal prmFecha As Date)

    '    Dim vcResultado As String = ""
    '    Dim DTResultado As New DataTable
    '    Dim wsservice As New wsMarengo.GrowerService
    '    Dim Parametro1 As String = ""
    '    Dim Parametro2 As String = ""

    '    DTResultado.Columns.Add("Grower", GetType(String))
    '    DTResultado.Columns.Add("Branch", GetType(String))
    '    DTResultado.Columns.Add("CommodityName", GetType(String))
    '    DTResultado.Columns.Add("PackStyle", GetType(String))
    '    DTResultado.Columns.Add("Label", GetType(String))
    '    DTResultado.Columns.Add("Size", GetType(String))
    '    DTResultado.Columns.Add("UoM", GetType(String))
    '    DTResultado.Columns.Add("Inventory", GetType(Double))
    '    DTResultado.Columns.Add("Ins", GetType(Double))
    '    DTResultado.Columns.Add("Outs", GetType(Double))

    '    Try
    '        wsservice.Url = "https://www.marengosite.com/marengowebservices/growerservice.asmx"

    '        '' Obteniendo Inventarios
    '        Parametro1 = "001"
    '        Parametro2 = "p@r3d3s"

    '        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

    '        'vcFecha = "24/01/2015"

    '        Dim vObjMarengo() As wsMarengo.GrowerInOuts

    '        vObjMarengo = wsservice.InOuts(Parametro1, Parametro2, vdFecha)

    '        For i As Integer = 0 To vObjMarengo.Length - 1
    '            Dim vRow As DataRow

    '            vRow = DTResultado.NewRow

    '            vRow("Grower") = vObjMarengo(i).Grower
    '            vRow("Branch") = vObjMarengo(i).Branch
    '            vRow("CommodityName") = vObjMarengo(i).CommodityName
    '            vRow("PackStyle") = vObjMarengo(i).PackStyle
    '            vRow("Label") = vObjMarengo(i).Label
    '            vRow("Size") = vObjMarengo(i).Size
    '            vRow("UoM") = vObjMarengo(i).UoM
    '            vRow("Inventory") = vObjMarengo(i).Inventory
    '            vRow("Ins") = vObjMarengo(i).Ins
    '            vRow("Outs") = vObjMarengo(i).Outs

    '            DTResultado.Rows.Add(vRow)

    '        Next


    '        EscribeEnBitacora("Se obtiene la información de existencias del webservice")

    '        vcResultado = fgGrabainventariosMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", "001"), vdFecha)

    '        If vcResultado <> "" Then
    '            EscribeEnBitacora(vcResultado)
    '            Exit Sub
    '        End If

    '        EscribeEnBitacora("Se inserto la informacion de Inventarios Marengo con Exito")

    '    Catch ex As Exception
    '        EscribeEnBitacora(ex.Message)
    '    End Try

    '    'Dim ConnectionString As String = "Provider=SQLOLEDB.1;Password=marengo123;Persist Security Info=True;User ID=MarengoRO;Initial Catalog=MarengoFoods;Data Source=db.marengoserver.dyndns.org;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    '    'Dim vcFecha As String = Format(prmFecha, "MM/dd/yyyy")

    '    'Using connection As New OleDb.OleDbConnection(ConnectionString)

    '    '    EscribeEnBitacora("Se conecta al servidor de Marengo")

    '    '    connection.Open()

    '    '    Dim dAdpter As New OleDb.OleDbDataAdapter("EXEC sExcelGrowerInOuts 2,'" & vcFecha & "'", connection)
    '    '    Dim DT As New DataTable

    '    '    EscribeEnBitacora("Obtiene informacion de Existencia")
    '    '    dAdpter.Fill(DT)

    '    '    If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

    '    '        EscribeEnBitacora("Graba informacion de Inventarios de Marengo")
    '    '        vcResultado = fgGrabainventariosMarengo(DT, fgObtenParametroEMB("TEMPORADA"), prmFecha)

    '    '    End If

    '    'End Using


    'End Sub

    Private Sub plObtenExistenciaMarengoAlterna(ByVal prmFecha As Date)

        Dim pop As New OpenPop.Pop3.Pop3Client
        Dim vnMensajes As Integer
        Dim vnCiclo As Integer

        Try
            '  Set the POP3 server's hostname   '  Set the POP3 login/password.
            pop.Connect("mail.aparedes.com.mx", 110, False)
            pop.Authenticate("vmarengo@aparedes.com.mx", "Paredes@123:)", AuthenticationMethod.UsernameAndPassword)

            vnMensajes = 0

            EscribeEnBitacora("Obteniendo Correos Pendientes")

            vnMensajes = pop.GetMessageCount

            If Not vnMensajes > 0 Then
                EscribeEnBitacora("No existieron correos pendientes para procesar")
                Exit Sub
            End If

            EscribeEnBitacora("Se procesaran " & vnMensajes & " correos")

            Dim vcAdjuntos As New ArrayList()
            For vnCiclo = 1 To vnMensajes

                Try
                    Dim message As Message = pop.GetMessage(vnCiclo)
                    'Dim messageheader As MessageHeader = pop.GetMessageHeaders(vnCiclo)

                    Dim attachments As List(Of MessagePart) = message.FindAllAttachments()

                    If attachments.Count > 0 Then

                        Dim oMsg As MailMessage = New MailMessage()
                        Dim vObjReceptor As New System.Net.Mail.MailAddress("enriqueca@aparedes.com.mx")

                        Dim vObjEmisor As New System.Net.Mail.MailAddress("vmarengo@aparedes.com.mx")
                        Dim Servidor As New System.Net.Mail.SmtpClient
                        Dim vAdjuntos As New ArrayList

                        oMsg.From = vObjEmisor
                        oMsg.To.Add(vObjReceptor)
                        oMsg.Subject = message.Headers.Subject
                        oMsg.IsBodyHtml = True
                        oMsg.Body = "<HTML><BODY><B>" & "Archivo" & "</B></BODY></HTML>"


                        For Each attachment As MessagePart In attachments

                            Dim vcRutaNombreDescarga As String

                            vcRutaNombreDescarga = "C:\CROP" & "\" & IIf(InStr(attachment.FileName, ":") > 0, Replace(attachment.FileName, ":", "_"), attachment.FileName)

                            vcRutaNombreDescarga = IIf(InStr(vcRutaNombreDescarga, "/") > 0, Replace(vcRutaNombreDescarga, "/", "_"), vcRutaNombreDescarga)

                            Dim file As New FileInfo(vcRutaNombreDescarga)

                            If file.Exists Then
                                file.Delete()
                            End If

                            If file.Extension.ToUpper = ".CSV" Then
                                Try
                                    attachment.Save(file)
                                    vAdjuntos.Add(vcRutaNombreDescarga)
                                Catch ex As Exception
                                    'EscribeEnBitacora(ex.Message + "No se encontro el Archivo " + file.FullName)
                                    flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenDatosVentasDiariasMarengoAlternaConexion", ex.Message + "No se encontro el Archivo " + file.FullName)
                                End Try
                            End If

                        Next

                        For vnPosicion As Integer = 0 To vAdjuntos.Count - 1
                            Dim vcAdjunto As String = ""
                            vcAdjunto = vAdjuntos.Item(vnPosicion)
                            oMsg.Attachments.Add(New System.Net.Mail.Attachment(vcAdjunto))
                        Next

                        Servidor.Host = "smtp.aparedes.com.mx"
                        Servidor.Port = 587
                        Servidor.EnableSsl = False
                        Servidor.Credentials = New System.Net.NetworkCredential("vmarengo@aparedes.com.mx", "Paredes@123:)")
                        Servidor.Send(oMsg)

                        oMsg.Dispose()
                        oMsg = Nothing

                    End If

                    pop.DeleteMessage(vnCiclo)
                    message = Nothing

                Catch ex As Exception
                    EscribeEnBitacora(ex.Message)
                    flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlterna", ex.Message)
                End Try

            Next
            pop.DeleteAllMessages()

            pop.Disconnect()
            pop.Dispose()
            pop = Nothing


        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlterna", ex.Message)
        End Try

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        EscribeEnBitacora("Intento leer los archivos de excel")

        For Each Archivo As String In Directory.GetFiles("C:\CROP", "*.CSV", SearchOption.AllDirectories)

            Dim vcResultado As String = ""
            Dim DTResultado As New DataTable
            Dim vcSQL As String = ""

            EscribeEnBitacora("Proceso el archivo " & Archivo)


            Dim xls_cn As New System.Data.OleDb.OleDbConnection
            Dim xls_cmd As New System.Data.OleDb.OleDbCommand
            Dim xls_reader As New System.Data.OleDb.OleDbDataAdapter

            Dim strExtension As String = ""
            Dim nombreXls As String
            Dim m_Excel As Microsoft.Office.Interop.Excel.Application

            Dim vcSplit() As String
            'Dim vcSplit2() As String
            Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

            DTResultado.Columns.Add("Grower", GetType(String))
            DTResultado.Columns.Add("Branch", GetType(String))
            DTResultado.Columns.Add("CommodityName", GetType(String))
            DTResultado.Columns.Add("PackStyle", GetType(String))
            DTResultado.Columns.Add("Label", GetType(String))
            DTResultado.Columns.Add("Size", GetType(String))
            DTResultado.Columns.Add("UoM", GetType(String))
            DTResultado.Columns.Add("Inventory", GetType(Double))
            DTResultado.Columns.Add("Ins", GetType(Double))
            DTResultado.Columns.Add("Outs", GetType(Double))
            DTResultado.Columns.Add("Floor", GetType(Double))

            Try

                Dim lblArchivo As String = Archivo

                nombreXls = Path.GetFileName(lblArchivo)
                strExtension = Path.GetExtension(lblArchivo)
                nombreXls = Strings.Replace(nombreXls, strExtension, "")

                If strExtension = ".csv" Then
                    'MsgBox("es un archivo excel")
                    If (File.Exists(lblArchivo)) Then
                        xls_cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + lblArchivo + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'"
                        'xls_cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.12.0;Data Source=" + xlsx + ";Extended Properties='Excel 12.0;HDR=YES'"
                        Using xls_cn

                            Dim dt As New DataTable("Datos")

                            m_Excel = CreateObject("Excel.Application")
                            m_Excel.Workbooks.Open(lblArchivo)
                            xls_cn.Open()
                            'xls_cmd.CommandText = "SELECT * FROM [" & nombreXls & "$]"

                            vcSQL = "SELECT DESCRIPTION,PACK,SUM([RECEIVE QNTY]) AS [RECEIVE QNTY],SUM([ON HND QNTY]) AS [ON HND QNTY],SUM([AVAIL QNTY]) AS [AVAIL QNTY]"
                            vcSQL = vcSQL & vbCrLf & "FROM [" & nombreXls & "$]"
                            vcSQL = vcSQL & vbCrLf & "GROUP BY DESCRIPTION,PACK"

                            xls_cmd.CommandText = vcSQL
                            xls_cmd.Connection = xls_cn
                            xls_reader.SelectCommand = xls_cmd

                            Dim da As New System.Data.OleDb.OleDbDataAdapter(xls_cmd)
                            da.Fill(dt)

                            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

                                'Dim vnRenglones As Integer = dt.Rows.Count
                                'Dim vnRenglon As Integer = 0

                                For Each vRowCiclo As DataRow In dt.Rows

                                    Dim vnInventarioAnterior As Double = 0
                                    Dim vnEntradas As Double = 0
                                    Dim vnSalidas As Double = 0
                                    Dim vnExistenciaActual As Double = 0

                                    If vRowCiclo("DESCRIPTION").ToString.Trim <> "" Then
                                        Dim vcProducto As String = ""
                                        Dim vcTipoProducto As String = ""
                                        Dim vcTamaño As String = ""
                                        Dim vcEtiqueta As String = ""
                                        Dim vcTipoCarton As String = ""
                                        Dim vcCarton As String = ""


                                        vcSplit = Strings.Split(vRowCiclo("DESCRIPTION"), "-")

                                        vcProducto = vcSplit(0)
                                        vcTipoProducto = vcSplit(1)
                                        vcTamaño = vcSplit(2)
                                        If vcSplit.Length = 4 Then
                                            vcEtiqueta = 1
                                            vcTipoCarton = vcSplit(3)
                                        Else
                                            vcEtiqueta = vcSplit(3)
                                            vcTipoCarton = vcSplit(4)
                                        End If
                                        vcCarton = vRowCiclo("PACK")


                                        Dim vRow As DataRow

                                        vRow = DTResultado.NewRow

                                        vRow("Grower") = ""
                                        vRow("Branch") = ""
                                        vRow("CommodityName") = fgObtenCommoditieName(vcProducto + "-" + vcTipoProducto)
                                        If vRow("CommodityName") = "" Then
                                            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenExistenciaMarengoAlterna", "CommodityName=''")
                                            'MsgBox("parate aki")
                                        End If
                                        vRow("PackStyle") = vcCarton
                                        vRow("Label") = vcEtiqueta
                                        vRow("Size") = vcTamaño
                                        vRow("UoM") = vcTipoCarton

                                        vnInventarioAnterior = fgObtenerInventario(vcTemporada, vRow("CommodityName"), vcCarton, vcEtiqueta, vcTamaño, vcTipoCarton)
                                        vnEntradas = vRowCiclo("RECEIVE QNTY")
                                        vnExistenciaActual = vRowCiclo("AVAIL QNTY")
                                        vnSalidas = (vnInventarioAnterior + vnEntradas) - vnExistenciaActual

                                        vRow("Inventory") = vnInventarioAnterior
                                        vRow("Ins") = vnEntradas
                                        vRow("Outs") = vnSalidas
                                        vRow("Floor") = vnExistenciaActual

                                        DTResultado.Rows.Add(vRow)

                                    End If


                                Next

                            End If

                            m_Excel.Quit()
                            m_Excel = Nothing

                        End Using

                    End If
                End If

                EscribeEnBitacora("Se obtiene la información de existencias del webservice")

                vcResultado = fgGrabainventariosMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

                If vcResultado <> "" Then
                    EscribeEnBitacora(vcResultado)
                    Exit Sub
                End If


                Try
                    vcSQL = "INSERT EYE_INVENTARIOSMARENGO"
                    vcSQL = vcSQL & vbCrLf & "SELECT O.CCVE_TEMPORADA,'" & Format(vdFecha, "yyyyMMdd") & "' AS DFECHA,O.Branch,O.CommodityName,O.PackStyle,O.Label,O.Size,O.UoM,"
                    vcSQL = vcSQL & vbCrLf & "O.Floor AS Inventory,0 AS INS,0 AS OUT,O.Floor,"
                    vcSQL = vcSQL & vbCrLf & "O.CCVE_CULTIVO,O.CCVE_ETIQUETA,O.CCVE_TAMAÑO,O.CCVE_ENVASE"
                    vcSQL = vcSQL & vbCrLf & "FROM EYE_INVENTARIOSMARENGO O(NOLOCK)"
                    vcSQL = vcSQL & vbCrLf & "WHERE O.CCVE_TEMPORADA = '" & vcTemporada & "' AND CONVERT(VARCHAR(20),O.DFECHA,112) = '" & Format(vdFecha.AddDays(-1), "yyyyMMdd") & "'"
                    vcSQL = vcSQL & vbCrLf & "AND NOT EXISTS(SELECT TOP 1 * FROM EYE_INVENTARIOSMARENGO S(NOLOCK) WHERE S.CCVE_TEMPORADA = O.CCVE_TEMPORADA AND S.CommodityName = O.CommodityName"
                    vcSQL = vcSQL & vbCrLf & "																	AND S.PackStyle = O.PackStyle AND S.Label = O.Label AND S.Size = O.Size AND S.UoM = O.UoM"
                    vcSQL = vcSQL & vbCrLf & "																	AND CONVERT(VARCHAR(20),S.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "')"

                    DAO.EjecutaComandoSQL(vcSQL)

                Catch ex As Exception
                    EscribeEnBitacora(ex.Message)
                    flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlterna", ex.Message)
                End Try

                EscribeEnBitacora("Se inserto la informacion de Inventarios Marengo con Exito")

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlterna", ex.Message)
            End Try


            If File.Exists(Archivo) Then
                Kill(Archivo)
            End If


        Next


    End Sub


    Private Sub plObtenExistenciaMarengoAlternaConexion()


        Dim miConexion As SqlConnection
        Dim miComando As SqlCommand
        Dim miAdapter As SqlDataAdapter
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        'Dim vdFecha As Date = DAO.RegresaFechaDelSistema.Date

        EscribeEnBitacora("Intenta conectarse al server de Marengo")

        'En caso de haber transacción abierta, se abre una nueva conexión solo para traernos los parámetros del Procedimiento almacenado.
        miConexion = New SqlConnection("SERVER= db.marengoserver.dyndns.org; Initial Catalog= PD; User= PDuser; Pwd= bd&CRYL3%mb#3E5%")
        miConexion.Open()

        EscribeEnBitacora("Se conecta al server de Marengo")

        miComando = New SqlCommand("sGwrTrans", miConexion)
        miComando.CommandType = CommandType.StoredProcedure
        'miComando.CommandTimeout = 30
        miComando.CommandTimeout = 0

        miAdapter = New SqlDataAdapter(miComando)

        Try
            SqlCommandBuilder.DeriveParameters(miComando)
        Catch ex As Exception
            'MsgBox(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlternaConexion", ex.Message)
        End Try

        miComando.Connection = miConexion

        Dim vParam(1) As Object
        Dim DSVentas As New DataSet

        vParam(0) = "00000035"
        vParam(1) = vdFecha

        CargaParametrosProcedimientoAlmacenado(miComando, vParam)

        Try

            EscribeEnBitacora("Se piden los datos al servidor de Marengo")

            miAdapter.Fill(DSVentas)
        Catch ex As Exception
            'MsgBox(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlternaConexion", ex.Message)
        End Try

        If miConexion.State = System.Data.ConnectionState.Open Then
            miConexion.Close()
        End If

        miConexion.Dispose()

        Dim dt As New DataTable

        dt = DSVentas.Tables(0)

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        EscribeEnBitacora("Intento leer los archivos de excel")


        Dim vcResultado As String = ""
        Dim DTResultado As New DataTable
        Dim vcSQL As String = ""

        'Dim vcSplit2() As String
        'Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        DTResultado.Columns.Add("Grower", GetType(String))
        DTResultado.Columns.Add("Branch", GetType(String))
        DTResultado.Columns.Add("CommodityName", GetType(String))
        DTResultado.Columns.Add("PackStyle", GetType(String))
        DTResultado.Columns.Add("Label", GetType(String))
        DTResultado.Columns.Add("Size", GetType(String))
        DTResultado.Columns.Add("UoM", GetType(String))
        DTResultado.Columns.Add("Inventory", GetType(Double))
        DTResultado.Columns.Add("Ins", GetType(Double))
        DTResultado.Columns.Add("Outs", GetType(Double))
        DTResultado.Columns.Add("Floor", GetType(Double))

        Try

            For Each vRowCiclo As DataRow In dt.Rows

                Dim vnInventarioAnterior As Double = 0
                Dim vnEntradas As Double = 0
                Dim vnSalidas As Double = 0
                Dim vnExistenciaActual As Double = 0

                Dim vcProducto As String = ""
                Dim vcTipoProducto As String = ""
                Dim vcTamaño As String = ""
                Dim vcEtiqueta As String = ""
                Dim vcTipoCarton As String = ""
                Dim vcCarton As String = ""



                vcProducto = vRowCiclo("COMMODITY")
                vcTipoProducto = vRowCiclo("VARIETY")
                vcTamaño = vRowCiclo("SIZE")
                vcEtiqueta = vRowCiclo("GRADE_CODE")
                vcTipoCarton = vRowCiclo("CONT_CODE")
                vcCarton = vRowCiclo("PACK")


                Dim vRow As DataRow

                vRow = DTResultado.NewRow

                vRow("Grower") = ""
                vRow("Branch") = ""
                vRow("CommodityName") = vcProducto + " - " + vcTipoProducto
                If vRow("CommodityName") = "" Then
                    'MsgBox("parate aki " & vcProducto + " - " + vcTipoProducto)
                    flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenExistenciaMarengoAlternaConexion", "CommodityName=''" & vcProducto + " - " + vcTipoProducto)
                End If
                vRow("PackStyle") = vcCarton
                vRow("Label") = vcEtiqueta.ToString.Trim
                vRow("Size") = vcTamaño
                vRow("UoM") = vcTipoCarton

                vnInventarioAnterior = vRowCiclo("FLOOR")
                vnEntradas = vRowCiclo("RECEIVED")
                vnSalidas = vRowCiclo("SHIPPED")
                vnExistenciaActual = vRowCiclo("FLOOR") + vRowCiclo("RECEIVED") - vRowCiclo("UNPACK") + vRowCiclo("REPACK") - vRowCiclo("SHIPPED")

                'If vnExistenciaActual < 0 Then
                '    MsgBox("parate aki")
                'End If


                vRow("Inventory") = vnInventarioAnterior
                vRow("Ins") = vnEntradas
                vRow("Outs") = vnSalidas
                vRow("Floor") = vnExistenciaActual

                DTResultado.Rows.Add(vRow)


            Next

            EscribeEnBitacora("Se graba la informacion en tabla")

            vcResultado = fgGrabainventariosMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

            If vcResultado <> "" Then
                EscribeEnBitacora(vcResultado)
                Exit Sub
            End If

            'Try
            '    vcSQL = "INSERT EYE_INVENTARIOSMARENGO"
            '    vcSQL = vcSQL & vbCrLf & "SELECT O.CCVE_TEMPORADA,'" & Format(vdFecha, "yyyyMMdd") & "' AS DFECHA,O.Branch,O.CommodityName,O.PackStyle,O.Label,O.Size,O.UoM,"
            '    vcSQL = vcSQL & vbCrLf & "O.Floor AS Inventory,0 AS INS,0 AS OUT,O.Floor,"
            '    vcSQL = vcSQL & vbCrLf & "O.CCVE_CULTIVO,O.CCVE_ETIQUETA,O.CCVE_TAMAÑO,O.CCVE_ENVASE"
            '    vcSQL = vcSQL & vbCrLf & "FROM EYE_INVENTARIOSMARENGO O(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE O.CCVE_TEMPORADA = '" & vcTemporada & "' AND CONVERT(VARCHAR(20),O.DFECHA,112) = '" & Format(vdFecha.AddDays(-1), "yyyyMMdd") & "'"
            '    vcSQL = vcSQL & vbCrLf & "AND NOT EXISTS(SELECT TOP 1 * FROM EYE_INVENTARIOSMARENGO S(NOLOCK) WHERE S.CCVE_TEMPORADA = O.CCVE_TEMPORADA AND S.CommodityName = O.CommodityName"
            '    vcSQL = vcSQL & vbCrLf & "																	AND S.PackStyle = O.PackStyle AND S.Label = O.Label AND S.Size = O.Size AND S.UoM = O.UoM"
            '    vcSQL = vcSQL & vbCrLf & "																	AND CONVERT(VARCHAR(20),S.DFECHA,112) = '" & Format(vdFecha, "yyyyMMdd") & "')"

            '    DAO.EjecutaComandoSQL(vcSQL)

            'Catch ex As Exception
            '    EscribeEnBitacora(ex.Message)
            'End Try

            EscribeEnBitacora("Se inserto la informacion de Inventarios Marengo con Exito")

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plObtenExistenciaMarengoAlterna", ex.Message)
        End Try


    End Sub

    Private Sub GeneraExcelExistencias()

        'Dim vcFecha As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "dd/MM/yyyy")
        Dim vcFecha As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "yyyyMMdd")
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim vcCultivo As String = ""
        Dim vcCalidad As String = ""
        Dim vcDiaSemana As String = UCase(vdFecha.ToString("dddddddddddd"))
        Dim vnDiaMes As Integer = vdFecha.Day
        Dim formatoFecha As DateTimeFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat
        Dim vcNombreMes As String = UCase(formatoFecha.GetMonthName(vdFecha.Month))
        Dim vcSQL As String = ""


        vbBerenjena1ra = False
        vbBerenjena2da = False

        vbVerdes1ra = False
        vbVerdes2da = False

        vbRojos1ra = False
        vbRojos2da = False

        vbAmarillo1ra = False
        vbAmarillo2da = False

        vbNaranja1ra = False
        vbNaranja2da = False

        vbBolas1ra = False
        vbBolas2da = False

        vbSaladette1ra = False
        vbSaladette2da = False

        Try
            DAO.EjecutaComandoSQL("SET DATEFORMAT YMD")
            DAO.EjecutaComandoSQL("DELETE INF_INVENTARIOSDISTRIBUIDORAS")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVO '" & vcTemporada & "','" & vcFecha & "','005',1")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVO '" & vcTemporada & "','" & vcFecha & "','003',2")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVO '" & vcTemporada & "','" & vcFecha & "','004',3")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASE '" & vcTemporada & "','" & vcFecha & "','002',4")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','008',5,'020',1")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','008',5,'003',2")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','061',6,'020',1")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','061',6,'003',2")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','010',7,'020',1")
            DAO.EjecutaComandoSQL("EXEC SPINSERTAINVENTARIOSDISTPORCULTIVOYENVASEFILTRO '" & vcTemporada & "','" & vcFecha & "','010',7,'003',2")

            vcSQL = "DELETE INF_INVENTARIOSDISTRIBUIDORAS"
            vcSQL = vcSQL & vbCrLf & "WHERE nInicialHM = 0 AND nInicialMarengo = 0 AND nInicialTotal = 0"
            vcSQL = vcSQL & vbCrLf & "AND nLlegadasHM = 0 AND nLlegadasMarengo = 0 AND nLlegadasTotal = 0"
            vcSQL = vcSQL & vbCrLf & "AND nSalidasHM = 0 AND nSalidasMarengo = 0 AND nSalidasTotal = 0"
            vcSQL = vcSQL & vbCrLf & "AND nFinalHM = 0 AND nFinalMarengo = 0 AND nFinalTotal = 0"

            DAO.EjecutaComandoSQL(vcSQL)

            DAO.EjecutaComandoSQL("UPDATE INF_INVENTARIOSDISTRIBUIDORAS SET CNOMBREGRUPOSECUNDARIO = 'CALIDAD #1' WHERE CNOMBREGRUPOSECUNDARIO = 'PRIMERAS'")

            DAO.EjecutaComandoSQL("UPDATE INF_INVENTARIOSDISTRIBUIDORAS SET CNOMBREGRUPOSECUNDARIO = 'CALIDAD #2' WHERE CNOMBREGRUPOSECUNDARIO = 'SEGUNDAS'")


        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en GeneraExcelExistencias", ex.Message)
            Exit Sub
        End Try

        Try

            If (aplicacionExcel Is Nothing) Then
                aplicacionExcel = New Microsoft.Office.Interop.Excel.Application
            End If

            'aplicacionExcel = New Microsoft.Office.Interop.Excel.Application

            aplicacionExcel.Workbooks.Open("C:\EXISTENCIASDISTRIBUIDORAS.xlsx")
            'aplicacionExcel.Workbooks.Open("C:\ENRIQUE\MIS DOCUMENTOS\Documentos Paredes\Formatos Existencias Distribuidoras\EXISTENCIASDISTRIBUIDORAS.xlsx")

            objHojaExcel = aplicacionExcel.Worksheets(1)

            objHojaExcel.Range("M1:M1").Value = fgObtenParametrosSemanaEnBaseFecha(vcTemporada, vdFecha)
            objHojaExcel.Range("M2:M2").Value = vcDiaSemana & " " & vcNombreMes & " " & vnDiaMes & ", " & vdFecha.Year


            '' Obtengo los datos de Berenjena
            '' CCVE_CULTIVO = '005'
            vcCultivo = "005"
            vcCalidad = "001"

            '' 18'S PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '025'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "025", 7)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "025", 7, vbBerenjena1ra)

            '' 24'S PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '026'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "026", 8)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "026", 8, vbBerenjena1ra)

            '' 32'S PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '027'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "027", 9)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "027", 9, vbBerenjena1ra)



            '' 18'S SEGUNDA CALIDAD
            '' CCVE_TAMAÑO = '025'
            '' CCVE_CALIDAD = '002'
            vcCalidad = "002"

            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "025", 14)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "025", 14, vbBerenjena2da)

            '' 24'S SEGUNDA CALIDAD
            '' CCVE_TAMAÑO = '026'
            '' CCVE_CALIDAD = '002'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "026", 15)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "026", 15, vbBerenjena2da)



            If vbBerenjena1ra = False And vbBerenjena2da = False Then

                objHojaExcel.Range("B3:I3").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B4:I4").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B5:I5").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B11:I11").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B12:I12").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B17:I17").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B18:I18").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B19:I19").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("B20:I20").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

            End If


            '' Obtengo los datos de Chile Verde
            '' CCVE_CULTIVO = '003'
            vcCultivo = "003"


            vcCalidad = "001"

            '' JMB PRIMERA(CALIDAD)
            '' CCVE_TAMAÑO = '043'
            '' CCVE_CALIDAD = '001'
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "043", 24, vbVerdes1ra)

            '' XLG PRIMERA(CALIDAD)
            '' CCVE_TAMAÑO = '025'
            '' CCVE_CALIDAD = '001'


            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "002", 25)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "002", 25, vbVerdes1ra)

            '' LAR PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '003'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "003", 26)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "003", 26, vbVerdes1ra)

            '' MED PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '027'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "004", 27)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "004", 27, vbVerdes1ra)

            '' SML PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '005'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "005", 28)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "005", 28, vbVerdes1ra)

            '' XSMALL PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '042'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "042", 29)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "042", 29, vbVerdes1ra)



            vcCalidad = "002"

            '' JMB SEGUNDA(CALIDAD)
            '' CCVE_TAMAÑO = '043'
            '' CCVE_CALIDAD = '002'


            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "043", 36, vbVerdes2da)

            '' XLG SEGUNDA(CALIDAD)
            '' CCVE_TAMAÑO = '025'
            '' CCVE_CALIDAD = '002'


            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "002", 37)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "002", 37, vbVerdes2da)

            '' LAR SEGUNDA CALIDAD
            '' CCVE_TAMAÑO = '003'
            '' CCVE_CALIDAD = '002'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "003", 38)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "003", 38, vbVerdes2da)

            '' MED SEGUNDA CALIDAD
            '' CCVE_TAMAÑO = '027'
            '' CCVE_CALIDAD = '002'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "004", 39)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "004", 39, vbVerdes2da)

            '' SML SEGUNDA CALIDAD
            '' CCVE_TAMAÑO = '005'
            '' CCVE_CALIDAD = '002'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "005", 40)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "005", 40, vbVerdes2da)


            If vbVerdes1ra = False And vbVerdes2da = False Then

                objHojaExcel.Range("A21").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A22").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A23").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A25").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A26").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A27").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A28").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A29").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A34").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A35").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A37").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A38").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A39").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A40").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A41").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A42").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A43").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A44").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If

            '' Obtengo los datos de ROMAS
            '' CCVE_CULTIVO = '004'
            vcCultivo = "004"

            '' XLG PRIMERA(CALIDAD)
            '' CCVE_TAMAÑO = '002'
            '' CCVE_CALIDAD = '001'
            vcCalidad = "001"

            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "002", 49)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "002", 49, vbSaladette1ra)

            '' LAR PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '003'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "003", 50)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "003", 50, vbSaladette1ra)

            '' MED PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '027'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "004", 51)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "004", 51, vbSaladette1ra)

            '' SML PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '005'
            '' CCVE_CALIDAD = '001'
            'plLlenaDatosExcelTamaño(vcTemporada, vcFecha, vcCalidad, vcCultivo, "005", 52)
            plLlenaDatosExcelTamaño(vcTemporada, vdFecha, vcCalidad, vcCultivo, "005", 52, vbSaladette1ra)

            If vbSaladette1ra = False Then

                objHojaExcel.Range("A45").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A46").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A47").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A49").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A50").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A51").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A52").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A54").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A55").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A64").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A65").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If


            '' Obtengo los datos de BOLAS
            '' CCVE_CULTIVO = '002'
            vcCultivo = "002"

            '' 4X4 2T PRIMERA(CALIDAD)
            '' CCVE_TAMAÑO = '007'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '011'
            vcCalidad = "001"

            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "007", 70, "011")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "007", 70, "011", vbBolas1ra)

            '' 4X5 2T PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '008'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '011'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "008", 71, "011")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "008", 71, "011", vbBolas1ra)

            '' 5X5 2T PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '009'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '011'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "009", 72, "011")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "009", 72, "011", vbBolas1ra)

            '' 5X6 2T PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '010'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '011'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "010", 73, "011")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "010", 73, "011", vbBolas1ra)

            '' 5X5 25# PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '009'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '001'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "009", 76, "001")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "009", 76, "001", vbBolas1ra)

            '' 5X6 25# PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '010'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '001'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "010", 77, "001")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "010", 77, "001", vbBolas1ra)

            '' 6X6 25# PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '011'
            '' CCVE_CALIDAD = '001'
            '' ENVASE = '001'
            'plLlenaDatosExcelTamañoEnvase(vcTemporada, vcFecha, vcCalidad, vcCultivo, "011", 78, "001")
            plLlenaDatosExcelTamañoEnvase(vcTemporada, vdFecha, vcCalidad, vcCultivo, "011", 78, "001", vbBolas1ra)


            If vbBolas1ra = False Then
                objHojaExcel.Range("A66").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A67").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A68").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A70").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A71").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A72").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A73").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A76").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A77").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A78").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A80").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A81").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A95").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A96").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If

            '' Obtengo los datos de CHILES ROJOS
            '' CCVE_CULTIVO = '008'
            vcCultivo = "008"

            '' XLG 11#
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '020'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 101, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 101, "020", vbRojos1ra)

            '' LGE 11#
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 102, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 102, "020", vbRojos1ra)

            '' MED 11#
            '' CCVE_TAMAÑO = '004'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 103, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 103, "020", vbRojos1ra)


            '' XLG 1 1/9
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '003'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 106, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 106, "003", vbRojos1ra)

            '' LGE 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '003'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 107, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 107, "003", vbRojos1ra)

            '' MED 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 108, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 108, "003", vbRojos1ra)

            '' SML 1 1/9
            '' CCVE_TAMAÑO = '005'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "005", 109, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "005", 109, "003", vbRojos1ra)

            If vbRojos1ra = False Then
                objHojaExcel.Range("A97").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A98").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A99").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A101").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A102").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A103").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A104").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A105").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A106").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A107").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A108").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A109").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A111").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A112").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A113").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A114").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If


            '' Obtengo los datos de CHILES AMARILLOS
            '' CCVE_CULTIVO = '061'
            vcCultivo = "061"

            '' XLG 11#
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '020'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 119, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 119, "020", vbAmarillo1ra)

            '' LGE 11#
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 120, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 120, "020", vbAmarillo1ra)

            '' MED 11#
            '' CCVE_TAMAÑO = '004'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 121, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 121, "020", vbAmarillo1ra)


            '' XLG 1 1/9
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '003'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 124, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 124, "003", vbAmarillo1ra)

            '' LGE 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '003'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 125, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 125, "003", vbAmarillo1ra)

            '' MED 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 126, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 126, "003", vbAmarillo1ra)

            '' SML 1 1/9
            '' CCVE_TAMAÑO = '005'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "005", 127, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "005", 127, "003", vbAmarillo1ra)

            If vbAmarillo1ra = False Then
                objHojaExcel.Range("A115").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A116").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A117").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A119").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A120").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A121").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A122").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A123").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A124").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A125").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A126").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A127").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A129").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A130").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A131").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A132").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If


            '' Obtengo los datos de CHILES NARANJA
            '' CCVE_CULTIVO = '010'
            vcCultivo = "010"

            '' XLG 11#
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '020'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 137, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 137, "020", vbNaranja1ra)

            '' LGE 11#
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 138, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 138, "020", vbNaranja1ra)

            '' MED 11#
            '' CCVE_TAMAÑO = '004'
            '' ENVASE = '020'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 139, "020")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 139, "020", vbNaranja1ra)


            '' XLG 1 1/9
            '' CCVE_TAMAÑO = '002'
            '' ENVASE = '003'

            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "002", 142, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "002", 142, "003", vbNaranja1ra)

            '' LGE 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '003'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "003", 143, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "003", 143, "003", vbNaranja1ra)

            '' MED 1 1/9
            '' CCVE_TAMAÑO = '003'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "004", 144, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "004", 144, "003", vbNaranja1ra)

            '' SML 1 1/9
            '' CCVE_TAMAÑO = '005'
            '' ENVASE = '004'
            'plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vcFecha, vcCultivo, "005", 145, "003")
            plLlenaDatosExcelTamañoEnvaseSinCalidad(vcTemporada, vdFecha, vcCultivo, "005", 145, "003", vbNaranja1ra)

            If vbNaranja1ra = False Then
                objHojaExcel.Range("A133").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A134").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A135").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A137").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A138").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A139").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A140").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A141").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A142").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A143").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A144").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A145").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A147").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A148").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A149").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

                objHojaExcel.Range("A150").Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True
            End If


            With aplicacionExcel.ActiveSheet.PageSetup
                .PrintTitleRows = "$1:$2"
                .PrintTitleColumns = ""
                .BottomMargin = aplicacionExcel.InchesToPoints(1.0)
            End With

            Dim vcNombreArchivo As String = ""


            vcNombreArchivo = "C:\ARCHIVOS\EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".xlsx"

            If File.Exists(vcNombreArchivo) Then
                File.Delete(vcNombreArchivo)
            End If
            EscribeEnBitacora("Se intenta grabar archivo excel de existencias")
            aplicacionExcel.ActiveWorkbook.SaveAs(vcNombreArchivo)
            EscribeEnBitacora("Se grabo archivo excel de existencias")

            'aplicacionExcel.SaveWorkspace(vcNombreArchivo)
            Dim vcAdjuntos As New ArrayList()
            'vcAdjuntos.Add(vcNombreArchivo)

            vcNombreArchivo = "C:\ARCHIVOS\EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"
            EscribeEnBitacora("Se intenta grabar archivo pdf de existencias")

            aplicacionExcel.ActiveSheet.ExportAsFixedFormat(Type:=0, Filename:=vcNombreArchivo, Quality:=1, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)

            EscribeEnBitacora("Se grabo archivo pdf de existencias")

            aplicacionExcel.Quit()

            If File.Exists(vcNombreArchivo) Then


                vcAdjuntos.Add(vcNombreArchivo)

                ' Enviamos correo
                'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                EscribeEnBitacora("Enviando Correo de Existencias...")

                flEnviarMail(vcCorreosVentas, vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOEXISTENCIASDISTRIBUIDORAS SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

            End If

            'vcNombreArchivo = "C:\EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".xlsx"

            'If File.Exists(vcNombreArchivo) Then
            '    File.Delete(vcNombreArchivo)
            'End If

            'vcNombreArchivo = "C:\EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

            'If File.Exists(vcNombreArchivo) Then
            '    File.Delete(vcNombreArchivo)
            'End If

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en GeneraExcelExistencias", ex.Message)
        End Try


    End Sub

    Private Sub checkVerificationCode()
        Dim entercode As String = ""
        If entercode.Equals(vcVerficationcode, StringComparison.OrdinalIgnoreCase) Then
            MsgBox("Email verification succeeded!")
        Else
            MsgBox("Email verification failed!")
        End If
    End Sub

    Private Function flEnviarMail(ByVal prmCorreo As String, ByVal prmAdjuntos As ArrayList, ByVal prmTítulo As String, ByVal prmCuerpo As String) As Boolean

        Dim vcAdjunto As String
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcPass As String
        Try
            gcServidorSMTP = fgObtenParametroEMB("SERVIDORSMTP", sucursal)
            gnPuertoSMTP = fgObtenParametroEMB("PUERTOCORREOEMISOR", sucursal)
            vcPass = fgObtenParametroMail("PASSNOREPLY")

            'gcServidorSMTP = "mail.aparedes.com.mx"
            'gnPuertoSMTP = "465"
            Dim oMsg As MailMessage = New MailMessage()
            oMsg.To.Add(prmCorreo)
            'Dim vObjReceptor As New System.Net.Mail.MailAddress(prmCorreo)
            Dim vObjEmisor As New System.Net.Mail.MailAddress("noreply@aparedes.com.mx")
            Dim Servidor As New System.Net.Mail.SmtpClient

            oMsg.From = vObjEmisor
            'oMsg.To.Add(vObjReceptor)
            oMsg.Subject = prmTítulo
            oMsg.IsBodyHtml = True
            oMsg.Body = "<HTML><BODY><B>" & prmCuerpo & "</B></BODY></HTML>"

            '[ en el caso de querer agregar un archivo adjunto, realizar el siguiente paso ]

            For vnPosicion As Integer = 0 To prmAdjuntos.Count - 1
                vcAdjunto = prmAdjuntos.Item(vnPosicion)
                oMsg.Attachments.Add(New System.Net.Mail.Attachment(vcAdjunto))
            Next

            Servidor.Host = gcServidorSMTP
            Servidor.Port = gnPuertoSMTP
            'Servidor.EnableSsl = True
            Servidor.Credentials = New System.Net.NetworkCredential("noreply@aparedes.com.mx", vcPass)
            Servidor.EnableSsl = False
            Servidor.Send(oMsg)

            oMsg = Nothing

            EscribeEnBitacora("Se envio el correo")

        Catch ex As Exception
            Return flEnviarMail
            EscribeEnBitacora(ex.Message)
        End Try

    End Function

    Private Function flEnviarMail2(ByVal prmCorreo As String, ByVal prmAdjuntos As ArrayList, ByVal prmTítulo As String, ByVal prmCuerpo As String) As Boolean

        Dim vcAdjunto As String
        Dim vcPass As String


        Try
            gcServidorSMTP = "smtp.gmail.com" 'fgObtenParametroEMB("SERVIDORSMTP", "001")
            gnPuertoSMTP = 587 'fgObtenParametroEMB("PUERTOCORREOEMISOR", "001")
            gbUsaSSL = True ' fgObtenParametroEMB("USASSL", "001")
            vcPass = "trruuorcpgjfdugq" 'fgObtenParametroMail("PASSNOREPLY")


            Dim oMsg As MailMessage = New MailMessage()
            oMsg.To.Add(prmCorreo)

            Dim vObjEmisor As New MailAddress("joser2203@gmail.com", "No Reply")
            Dim vObjReceptor As New MailAddress(prmCorreo)

            oMsg.From = vObjEmisor
            oMsg.To.Add(vObjReceptor)
            oMsg.Subject = prmTítulo
            oMsg.IsBodyHtml = True
            oMsg.Body = "<HTML><BODY><B>" & prmCuerpo & "</B></BODY></HTML>"

            '[ en el caso de querer agregar un archivo adjunto, realizar el siguiente paso ]
            If prmAdjuntos IsNot Nothing Then
                For vnPosicion As Integer = 0 To prmAdjuntos.Count - 1
                    vcAdjunto = prmAdjuntos.Item(vnPosicion)
                    oMsg.Attachments.Add(New System.Net.Mail.Attachment(vcAdjunto))
                Next
            End If

            Dim Servidor As New SmtpClient()
            Servidor.Host = gcServidorSMTP
            Servidor.EnableSsl = gbUsaSSL

            Servidor.Credentials = New NetworkCredential(vObjEmisor.ToString(), vcPass)

            Servidor.UseDefaultCredentials = False
            Servidor.DeliveryMethod = SmtpDeliveryMethod.Network
            Servidor.Port = gnPuertoSMTP
            Servidor.Send(oMsg)

            oMsg = Nothing

            EscribeEnBitacora("Se envio el correo")
            Return True

        Catch ex As Exception
            Return flEnviarMail2
            EscribeEnBitacora(ex.Message)
        End Try


    End Function

    Private Function flEnviarMailoAuth(ByVal prmCorreo As String, ByVal prmAdjuntos As ArrayList, ByVal prmTítulo As String, ByVal prmCuerpo As String) As Boolean

        Dim vcAdjunto As String
        Dim vcPass As String


        Try
            gcServidorSMTP = "smtp.gmail.com" 'fgObtenParametroEMB("SERVIDORSMTP", "001")
            gnPuertoSMTP = 587 'fgObtenParametroEMB("PUERTOCORREOEMISOR", "001")
            gbUsaSSL = True ' fgObtenParametroEMB("USASSL", "001")
            vcPass = "trruuorcpgjfdugq" 'fgObtenParametroMail("PASSNOREPLY")


            Dim oMsg As MailMessage = New MailMessage()
            oMsg.To.Add(prmCorreo)

            Dim vObjEmisor As New MailAddress("joser2203@gmail.com", "No Reply")
            Dim vObjReceptor As New MailAddress(prmCorreo)

            oMsg.From = vObjEmisor
            oMsg.To.Add(vObjReceptor)
            oMsg.Subject = prmTítulo
            oMsg.IsBodyHtml = True
            oMsg.Body = "<HTML><BODY><B>" & prmCuerpo & "</B></BODY></HTML>"

            '[ en el caso de querer agregar un archivo adjunto, realizar el siguiente paso ]
            If prmAdjuntos IsNot Nothing Then
                For vnPosicion As Integer = 0 To prmAdjuntos.Count - 1
                    vcAdjunto = prmAdjuntos.Item(vnPosicion)
                    oMsg.Attachments.Add(New System.Net.Mail.Attachment(vcAdjunto))
                Next
            End If

            Dim Servidor As New SmtpClient()
            Servidor.Host = gcServidorSMTP
            Servidor.EnableSsl = gbUsaSSL

            Servidor.Credentials = New NetworkCredential(vObjEmisor.ToString(), vcPass)

            Servidor.UseDefaultCredentials = False
            Servidor.DeliveryMethod = SmtpDeliveryMethod.Network
            Servidor.Port = gnPuertoSMTP
            Servidor.Send(oMsg)

            oMsg = Nothing

            EscribeEnBitacora("Se envio el correo")
            Return True

        Catch ex As Exception
            Return flEnviarMailoAuth
            EscribeEnBitacora(ex.Message)
        End Try


    End Function

    Private Sub plLlenaDatosExcelTamañoEnvaseSinCalidad(ByVal prmTemporada As String, ByVal vcFecha As Date, ByVal prmCultivo As String,
                                  ByVal prmTamaño As String, ByVal prmRenglon As Integer, ByVal prmEnvase As String, ByRef prmCalidadProducto As Boolean)

        Dim vcSQL As String = ""
        Dim vDs As New DataSet
        Dim vParametros(4) As Object


        '' 18'S PRIMERA CALIDAD
        '' CCVE_TAMAÑO = '025'
        '' CCVE_CALIDAD = '001'
        vParametros(0) = prmTemporada
        vParametros(1) = vcFecha
        vParametros(2) = prmCultivo '' CULTIVO
        vParametros(3) = prmTamaño '' TAMAÑO
        vParametros(4) = prmEnvase '' TAMAÑO


        vcSQL = "SPOBTENEXISTENCIASDISTRIBUIDORASPORTAMAÑOYENVASESINCALIDAD"
        If Not DAO.RegresaConsultaSQL(vcSQL, vDs, vParametros) Then
            Exit Sub
        End If

        If Not vDs Is Nothing AndAlso vDs.Tables.Count > 0 Then
            plAlimentaHojaExcel(vDs, prmRenglon, prmCalidadProducto)
        End If

        vDs = Nothing

    End Sub

    Private Sub plLlenaDatosExcelTamañoEnvase(ByVal prmTemporada As String, ByVal vcFecha As Date, ByVal prmCalidad As String, ByVal prmCultivo As String,
                                  ByVal prmTamaño As String, ByVal prmRenglon As Integer, ByVal prmEnvase As String, ByRef prmCalidadProducto As Boolean)

        Dim vcSQL As String = ""
        Dim vDs As New DataSet
        Dim vParametros(5) As Object


        '' 18'S PRIMERA CALIDAD
        '' CCVE_TAMAÑO = '025'
        '' CCVE_CALIDAD = '001'
        vParametros(0) = prmTemporada
        vParametros(1) = vcFecha
        vParametros(2) = prmCalidad '' CALIDAD
        vParametros(3) = prmCultivo '' CULTIVO
        vParametros(4) = prmTamaño '' TAMAÑO
        vParametros(5) = prmEnvase '' TAMAÑO


        vcSQL = "SPOBTENEXISTENCIASDISTRIBUIDORASPORTAMAÑOYENVASE"
        If Not DAO.RegresaConsultaSQL(vcSQL, vDs, vParametros) Then
            Exit Sub
        End If

        If Not vDs Is Nothing AndAlso vDs.Tables.Count > 0 Then
            plAlimentaHojaExcel(vDs, prmRenglon, prmCalidadProducto)
        End If

        vDs = Nothing

    End Sub

    Private Sub plLlenaDatosExcelTamaño(ByVal prmTemporada As String, ByVal vcFecha As Date, ByVal prmCalidad As String, ByVal prmCultivo As String,
                                  ByVal prmTamaño As String, ByVal prmRenglon As Integer, ByRef prmCalidadProducto As Boolean)

        Dim vcSQL As String = ""
        Dim vDs As New DataSet
        Dim vParametros(4) As Object

        Try
            '' 18'S PRIMERA CALIDAD
            '' CCVE_TAMAÑO = '025'
            '' CCVE_CALIDAD = '001'
            vParametros(0) = prmTemporada
            vParametros(1) = vcFecha
            vParametros(2) = prmCalidad '' CALIDAD
            vParametros(3) = prmCultivo '' CULTIVO
            vParametros(4) = prmTamaño '' TAMAÑO

            vcSQL = "SPOBTENEXISTENCIASDISTRIBUIDORASPORTAMAÑO"
            If Not DAO.RegresaConsultaSQL(vcSQL, vDs, vParametros) Then
                Exit Sub
            End If

            If Not vDs Is Nothing AndAlso vDs.Tables.Count > 0 Then
                plAlimentaHojaExcel(vDs, prmRenglon, prmCalidadProducto)
            End If

            vDs = Nothing

        Catch ex As Exception
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plLLenaDatosExcelTamaño", ex.Message)
        End Try


    End Sub

    Private Sub plAlimentaHojaExcel(ByVal vDs As DataSet, ByVal prmRenglon As Integer, ByRef prmCalidadProducto As Boolean)

        Dim DT As New DataTable

        Try
            DT = vDs.Tables(0)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

                If DT(0)(0) = 0 And DT(0)(1) = 0 And DT(0)(2) = 0 And DT(0)(3) = 0 And DT(0)(4) = 0 And DT(0)(5) = 0 Then

                    objHojaExcel.Range("B" & prmRenglon & ":I" & prmRenglon).Select()
                    aplicacionExcel.Selection.EntireRow.Hidden = True

                Else
                    '' INICIAL HM
                    objHojaExcel.Range("B" & prmRenglon & ":B" & prmRenglon).Value = DT(0)(0)
                    '' INICIAL MARENGO
                    objHojaExcel.Range("C" & prmRenglon & ":C" & prmRenglon).Value = DT(0)(1)

                    '' ENTRADA HM
                    objHojaExcel.Range("E" & prmRenglon & ":E" & prmRenglon).Value = DT(0)(2)
                    '' ENTRADA MARENGO
                    objHojaExcel.Range("F" & prmRenglon & ":F" & prmRenglon).Value = DT(0)(3)

                    '' SALIDA HM
                    objHojaExcel.Range("H" & prmRenglon & ":H" & prmRenglon).Value = DT(0)(4)
                    '' SALIDA MARENGO
                    objHojaExcel.Range("I" & prmRenglon & ":I" & prmRenglon).Value = DT(0)(5)

                    prmCalidadProducto = True

                End If
            Else

                objHojaExcel.Range("B" & prmRenglon & ":I" & prmRenglon).Select()
                aplicacionExcel.Selection.EntireRow.Hidden = True

            End If

        Catch ex As Exception
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plAlimentaHojaExcel", ex.Message)
        End Try


        DT = Nothing

    End Sub

    'Private Function flObtenVentasWSMarengo() As Boolean

    '    'Dim vObj As New wsVentasMarengo.GrowerClient
    '    Dim vObj As New wsVentasMarengoV2.GrowerClient


    '    'Dim vObjRes As New wsVentasMarengo.Transaction
    '    'Dim vObjRes As New wsVentasMarengo.TransactionV2
    '    Dim vcProduct As String = ""
    '    Dim DTResultado As New DataTable
    '    'Dim vdFecha As DateTime
    '    Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", "001")
    '    Dim oRpt = New ReportDocument
    '    Dim Proc As New System.Diagnostics.Process
    '    Dim vcAdjuntosPDF As New ArrayList
    '    Dim vcArchivo As String

    '    'Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

    '    DTResultado = New DataTable

    '    DTResultado.Columns.Add("Customer", GetType(String))
    '    DTResultado.Columns.Add("CommodityName", GetType(String))
    '    DTResultado.Columns.Add("PackStyle", GetType(String))
    '    DTResultado.Columns.Add("Label", GetType(String))
    '    DTResultado.Columns.Add("Size", GetType(String))
    '    DTResultado.Columns.Add("UoM", GetType(String))
    '    DTResultado.Columns.Add("Qty", GetType(Double))
    '    DTResultado.Columns.Add("Gross", GetType(Double))
    '    DTResultado.Columns.Add("Adj", GetType(Double))
    '    DTResultado.Columns.Add("net", GetType(Double))
    '    DTResultado.Columns.Add("UnitPrice", GetType(Double))
    '    DTResultado.Columns.Add("SalesType", GetType(String))
    '    DTResultado.Columns.Add("ShipDate", GetType(DateTime))
    '    DTResultado.Columns.Add("Variety", GetType(String))
    '    DTResultado.Columns.Add("PalletID", GetType(String))
    '    DTResultado.Columns.Add("cManif", GetType(String))


    '    'Dim vdFecha As Date = DateAdd(DateInterval.Day, -60, DAO.RegresaFechaDelSistema.Date)
    '    Dim vdFecha As Date = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)
    '    Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)



    '    Dim vObjVentas() As wsVentasMarengoV2.NetSalesV2

    '    vObjVentas = vObj.GetNetSalesV2("ws.paredes", "Wsp111817+", "00000133", vdFecha, vdFechaFin, "D")



    '    'vObjVentas = vObj.GetNetSales("ws.paredes", "Wsp111817+", "00000107", vdFecha, vdFechaFin, "D")



    '    Dim vcCustomer As String = ""
    '    Dim vcCommodityName As String = ""
    '    Dim vcPackStyle As String = ""
    '    Dim vcLabel As String = ""
    '    Dim vcSize As String = ""
    '    Dim vcUoM As String = ""
    '    Dim vnQty As Double = 0
    '    Dim vnGross As Double = 0
    '    Dim vnAdj As Double = 0
    '    Dim vnNet As Double = 0
    '    Dim vnUnitPrice As Double = 0
    '    Dim vcSalesType As String = ""
    '    Dim vnShipDate As Date = Now
    '    Dim vcVariety As String = ""
    '    Dim vcPalletID As String = ""
    '    Dim vcManif As String = ""

    '    For i As Integer = 0 To vObjVentas.Length - 1
    '        Try
    '            If vObjVentas(i).NetSalesMember <> 0 Then
    '                vcCustomer = ""

    '                vcPalletID = vObjVentas(i).Pallet_Ids
    '                vcManif = vObjVentas(i).Lot_No


    '                vcCommodityName = vObjVentas(i).Commodity
    '                vcPackStyle = Strings.Left(Strings.Replace(vObjVentas(i).Packaging, " ", ""), 5).ToString.Trim
    '                vcLabel = IIf(vObjVentas(i).Grade.ToString.Trim = "GRADE #1", 1, 2)
    '                vcSize = Strings.Replace(vObjVentas(i).Size, " ", "")
    '                vcUoM = IIf(vObjVentas(i).Container.ToString.Trim = "CARTON", "CTN", "RPC")
    '                vnQty = IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty)
    '                vnGross = IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember)
    '                vnAdj = IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)
    '                vnNet = IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember) - IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)
    '                vnUnitPrice = 0
    '                If IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty) > 0 Then
    '                    vnUnitPrice = (IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember) - IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)) / IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty)
    '                End If
    '                vcSalesType = ""
    '                vnShipDate = Strings.Left(vObjVentas(i).PostDate, 4) & "-" & Strings.Mid(vObjVentas(i).PostDate, 5, 2) & "-" & Strings.Right(vObjVentas(i).PostDate, 2)
    '                'vnShipDate = CDate(Strings.Right(vRowCiclo("POST_DATE"), 2) & "/" & Strings.Mid(vRowCiclo("POST_DATE"), 5, 2) & "/" & Strings.Left(vRowCiclo("POST_DATE"), 4))
    '                vdFecha = vnShipDate
    '                vcVariety = vObjVentas(i).Variety

    '                Dim vRow As DataRow

    '                vRow = DTResultado.NewRow

    '                vRow("Customer") = vcCustomer
    '                vRow("CommodityName") = vcCommodityName
    '                vRow("PackStyle") = vcPackStyle
    '                vRow("Label") = vcLabel
    '                vRow("Size") = vcSize
    '                vRow("UoM") = vcUoM
    '                vRow("Qty") = vnQty
    '                vRow("Gross") = vnGross
    '                vRow("Adj") = vnAdj
    '                vRow("Net") = vnNet
    '                vRow("UnitPrice") = vnUnitPrice
    '                vRow("SalesType") = vcSalesType
    '                vRow("ShipDate") = vnShipDate
    '                vRow("Variety") = vcVariety

    '                vRow("PalletID") = vcPalletID
    '                vRow("cManif") = vcManif


    '                DTResultado.Rows.Add(vRow)
    '            End If
    '        Catch ex As Exception
    '            MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
    '        End Try

    '    Next

    '    If Not DTResultado Is Nothing AndAlso DTResultado.Rows.Count > 0 Then

    '        Dim vcResultado As String = ""
    '        vdFecha = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)

    '        vcResultado = fgGrabaVentasDiariasMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", "001"), vdFecha)

    '    End If

    '    Dim dsMarengo As DataSet = fgTraeVentasDiariasMarengo(vcTemporada, vdFechaFin)

    '    If Not dsMarengo Is Nothing AndAlso dsMarengo.Tables.Count > 0 AndAlso dsMarengo.Tables(0).Rows.Count > 0 Then
    '        oRpt = New ReportDocument
    '        oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGOCLN.rpt")

    '        LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
    '        AgregarParametro("@SUCURSAL", "CULIACAN", oRpt)
    '        AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
    '        AgregarParametro("@PRMFECHA", vdFechaFin, oRpt)


    '        Dim oStremMarengo As New System.IO.MemoryStream

    '        Try
    '            oStremMarengo = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

    '            vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & ".pdf"

    '            'Si lo deseamos escribimos el pdf a disco.
    '            Dim ArchivoPDFMarengo As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
    '            ArchivoPDFMarengo.Write(oStremMarengo.ToArray, 0, oStremMarengo.ToArray.Length)
    '            ArchivoPDFMarengo.Flush()
    '            ArchivoPDFMarengo.Close()
    '            ArchivoPDFMarengo.Dispose()
    '            ArchivoPDFMarengo = Nothing

    '            Proc.Dispose()
    '            oRpt = Nothing

    '        Catch ex As Exception
    '            EscribeEnBitacora(ex.Message)
    '            Return False
    '        End Try
    '    End If

    '    Return True

    'End Function

    Private Function flObtenVentasWSMarengoV2() As Boolean 'Funcion que consume GrowerNetSales
        Console.WriteLine(vbCrLf & "### flObtenVentasWSMarengoV2() ### API")

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim contractID As String = fgObtenerConexionBD("contract_id")

        Dim DTResultado As New DataTable
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        Dim vdFecha As Date = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)
        Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        'Formateo de fechas para consumo de URL de API'        
        Dim vdFecha_format As String = Format(DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date), "yyyy-MM-dd")
        Dim vdFechaFin_format As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "yyyy-MM-dd")

        Dim api = New ExecuteAPI '//Llamado  a clase de consumo api
        Dim RespItem = New Response_GrowerNetSales '//Response de Api

        'URL y valores enviar
        Dim apiUrl = $"http://marengosite.com/GrowerAPI/GrowerNetSales/v2/{contractID}/{vdFecha_format}/{vdFechaFin_format}"

        Console.WriteLine("URL_API: " + apiUrl)

        'Guardamos respuesta de API
        Dim response = api.MGet(apiUrl)

        ' Deserializamos la respuesta JSON en una lista de objetos Response_GrowerNetSales
        Dim responseObjectList As List(Of Response_GrowerNetSales) = JsonConvert.DeserializeObject(Of List(Of Response_GrowerNetSales))(response)

        DTResultado = New DataTable

        ' Agregar columnas a la DataTable
        DTResultado.Columns.Add("vendor", GetType(Integer))
        DTResultado.Columns.Add("growerDealNumber", GetType(String))
        DTResultado.Columns.Add("postDate", GetType(String))
        DTResultado.Columns.Add("reference", GetType(String))
        DTResultado.Columns.Add("product", GetType(String))
        DTResultado.Columns.Add("lotNo", GetType(String))
        DTResultado.Columns.Add("palletID", GetType(Long))
        DTResultado.Columns.Add("commodity", GetType(String))
        DTResultado.Columns.Add("variety", GetType(String))
        DTResultado.Columns.Add("packaging", GetType(String))
        DTResultado.Columns.Add("container", GetType(String))
        DTResultado.Columns.Add("size", GetType(String))
        DTResultado.Columns.Add("grade", GetType(String))
        DTResultado.Columns.Add("qty", GetType(Integer))
        DTResultado.Columns.Add("grossSales", GetType(Double))
        DTResultado.Columns.Add("netSales", GetType(Double))
        DTResultado.Columns.Add("totalAdjustments", GetType(Double))

        Try
            ' Recorrer la lista de objetos y agregar los datos a la DataTable
            For Each item As Response_GrowerNetSales In responseObjectList
                Dim newRow As DataRow = DTResultado.NewRow()
                newRow("vendor") = item.vendor
                newRow("growerDealNumber") = item.growerDealNumber
                newRow("postDate") = item.postDate
                newRow("reference") = item.reference
                newRow("product") = item.product
                newRow("lotNo") = item.lotNo
                newRow("palletID") = item.palletID
                newRow("commodity") = item.commodity
                newRow("variety") = item.variety
                newRow("packaging") = item.packaging
                newRow("container") = item.container
                newRow("size") = item.size
                newRow("grade") = item.grade
                newRow("qty") = item.qty
                newRow("grossSales") = item.grossSales
                newRow("netSales") = item.netSales
                newRow("totalAdjustments") = item.totalAdjustments
                DTResultado.Rows.Add(newRow)
            Next

        Catch ex As Exception
            MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
            Return False
        End Try

        If Not DTResultado Is Nothing AndAlso DTResultado.Rows.Count > 0 Then

            Dim vcResultado As String = ""
            vdFecha = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)

            vcResultado = fgGrabaVentasDiariasMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

        End If

        Dim dsMarengo As DataSet = fgTraeVentasDiariasMarengo(vcTemporada, vdFechaFin)

        If Not dsMarengo Is Nothing AndAlso dsMarengo.Tables.Count > 0 AndAlso dsMarengo.Tables(0).Rows.Count > 0 Then
            Dim oStremMarengo As New System.IO.MemoryStream
            oRpt = New ReportDocument
            Dim oStrem As New System.IO.MemoryStream
            Console.WriteLine("### Generación de Reporte C:\CROP\RPT_VENTASDIARIASMARENGO ### ")
            If sucursal Is "001" Then 'CULIACAN
                oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGOCLN.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                AgregarParametro("@SUCURSAL", "CULIACAN", oRpt)
                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMFECHA", vdFechaFin, oRpt)

                vcArchivo = "C:\ARCHIVOS\CLN EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

                oStremMarengo = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)
                vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & ".pdf"
            Else 'JALISCO
                oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGOJAL.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)

                AgregarParametro("@SUCURSAL", "JALISCO", oRpt)
                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMFECHA", vdFechaFin, oRpt)

                vcArchivo = "C:\ARCHIVOS\JAL EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"
                oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\ARCHIVOS\JAL VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & ".pdf"
            End If

            Try
                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDFMarengo As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDFMarengo.Write(oStremMarengo.ToArray, 0, oStremMarengo.ToArray.Length)
                ArchivoPDFMarengo.Flush()
                ArchivoPDFMarengo.Close()
                ArchivoPDFMarengo.Dispose()
                ArchivoPDFMarengo = Nothing

                Proc.Dispose()
                oRpt = Nothing

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                Return False
            End Try
        End If

        Return True

    End Function


    Private Function flObtenVentasWSMarengo() As Boolean
        'Private Function flObtenVentasWSMarengo17nov22() As Boolean
        Console.WriteLine("### Ingresa a flObtenVentasWSMarengo()")

        '' RESPALDADA POR ENRIQUE CORRAL 17NVOV22

        Dim vObj As New wsVentasMarengo.GrowerClient
        Dim vObjRes As New wsVentasMarengo.Transaction
        Dim vcProduct As String = ""
        Dim DTResultado As New DataTable
        'Dim vdFecha As DateTime
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String

        'Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        Dim vdFecha As Date = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)
        'Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)


        Dim vObjVentas() As wsVentasMarengo.NetSales

        Try
            vObjVentas = vObj.GetNetSales("ws.paredes", "Wsp111817+", "00000133", vdFecha, vdFechaFin, "D") ' 22-23
            'vObjVentas = vObj.GetNetSales("ws.paredes", "Wsp111817+", "00000095", vdFecha, vdFechaFin, "D") ' 20-21
            'vObjVentas = vObj.GetNetSales("ws.paredes", "Wsp111817+", "00000074", vdFecha, vdFechaFin, "D") ' 19-20
            'vObjVentas = vObj.GetNetSales("ws.paredes", "Wsp111817+", "00000056", vdFecha, vdFechaFin, "D") ' 18-19

        Catch ex As Exception
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenVentasWSMarengo", ex.Message)
            Return False
        End Try

        DTResultado.Columns.Add("Customer", GetType(String))
        DTResultado.Columns.Add("CommodityName", GetType(String))
        DTResultado.Columns.Add("PackStyle", GetType(String))
        DTResultado.Columns.Add("Label", GetType(String))
        DTResultado.Columns.Add("Size", GetType(String))
        DTResultado.Columns.Add("UoM", GetType(String))
        DTResultado.Columns.Add("Qty", GetType(Double))
        DTResultado.Columns.Add("Gross", GetType(Double))
        DTResultado.Columns.Add("Adj", GetType(Double))
        DTResultado.Columns.Add("net", GetType(Double))
        DTResultado.Columns.Add("UnitPrice", GetType(Double))
        DTResultado.Columns.Add("SalesType", GetType(String))
        DTResultado.Columns.Add("ShipDate", GetType(DateTime))
        DTResultado.Columns.Add("Variety", GetType(String))


        Dim vcCustomer As String = ""
        Dim vcCommodityName As String = ""
        Dim vcPackStyle As String = ""
        Dim vcLabel As String = ""
        Dim vcSize As String = ""
        Dim vcUoM As String = ""
        Dim vnQty As Double = 0
        Dim vnGross As Double = 0
        Dim vnAdj As Double = 0
        Dim vnNet As Double = 0
        Dim vnUnitPrice As Double = 0
        Dim vcSalesType As String = ""
        Dim vnShipDate As Date = Now
        Dim vcVariety As String = ""

        For i As Integer = 0 To vObjVentas.Length - 1
            Try
                If vObjVentas(i).NetSalesMember <> 0 Then
                    vcCustomer = ""
                    vcCommodityName = vObjVentas(i).Commodity
                    vcPackStyle = Strings.Left(Strings.Replace(vObjVentas(i).Packaging, " ", ""), 5).ToString.Trim
                    vcLabel = IIf(vObjVentas(i).Grade.ToString.Trim = "GRADE #1", 1, 2)
                    vcSize = Strings.Replace(vObjVentas(i).Size, " ", "")
                    vcUoM = IIf(vObjVentas(i).Container.ToString.Trim = "CARTON", "CTN", "RPC")
                    vnQty = IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty)
                    vnGross = IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember)
                    vnAdj = IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)
                    vnNet = IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember) - IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)
                    vnUnitPrice = 0
                    If IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty) > 0 Then
                        vnUnitPrice = (IIf(vObjVentas(i).NetSalesMember = 0, 0, vObjVentas(i).NetSalesMember) - IIf(vObjVentas(i).TotalAdjustments = 0, 0, vObjVentas(i).TotalAdjustments)) / IIf(vObjVentas(i).Qty = 0, 0, vObjVentas(i).Qty)
                    End If
                    vcSalesType = ""
                    vnShipDate = Strings.Left(vObjVentas(i).PostDate, 4) & "-" & Strings.Mid(vObjVentas(i).PostDate, 5, 2) & "-" & Strings.Right(vObjVentas(i).PostDate, 2)
                    'vnShipDate = CDate(Strings.Right(vRowCiclo("POST_DATE"), 2) & "/" & Strings.Mid(vRowCiclo("POST_DATE"), 5, 2) & "/" & Strings.Left(vRowCiclo("POST_DATE"), 4))
                    vdFecha = vnShipDate
                    vcVariety = vObjVentas(i).Variety

                    Dim vRow As DataRow

                    vRow = DTResultado.NewRow

                    vRow("Customer") = vcCustomer
                    vRow("CommodityName") = vcCommodityName
                    vRow("PackStyle") = vcPackStyle
                    vRow("Label") = vcLabel
                    vRow("Size") = vcSize
                    vRow("UoM") = vcUoM
                    vRow("Qty") = vnQty
                    vRow("Gross") = vnGross
                    vRow("Adj") = vnAdj
                    vRow("Net") = vnNet
                    vRow("UnitPrice") = vnUnitPrice
                    vRow("SalesType") = vcSalesType
                    vRow("ShipDate") = vnShipDate
                    vRow("Variety") = vcVariety

                    DTResultado.Rows.Add(vRow)
                End If
            Catch ex As Exception
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenVentasWSMarengo", ex.Message)
                'MsgBox("Error" & Chr(13) & Chr(13) & ex.Message)
                Return False
            End Try

        Next

        If Not DTResultado Is Nothing AndAlso DTResultado.Rows.Count > 0 Then

            Dim vcResultado As String = ""
            vdFecha = DateAdd(DateInterval.Day, -15, DAO.RegresaFechaDelSistema.Date)

            vcResultado = fgGrabaVentasDiariasMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

        End If

        Dim dsMarengo As DataSet = fgTraeVentasDiariasMarengo(vcTemporada, vdFechaFin)

        If Not dsMarengo Is Nothing AndAlso dsMarengo.Tables.Count > 0 AndAlso dsMarengo.Tables(0).Rows.Count > 0 Then
            oRpt = New ReportDocument
            oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGOCLN.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@SUCURSAL", "CULIACAN", oRpt)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFechaFin, oRpt)


            Dim oStremMarengo As New System.IO.MemoryStream

            Try
                oStremMarengo = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\ARCHIVOS\CLN VENTAS DIARIAS MARENGO AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & ".pdf"

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDFMarengo As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDFMarengo.Write(oStremMarengo.ToArray, 0, oStremMarengo.ToArray.Length)
                ArchivoPDFMarengo.Flush()
                ArchivoPDFMarengo.Close()
                ArchivoPDFMarengo.Dispose()
                ArchivoPDFMarengo = Nothing

                Proc.Dispose()
                oRpt = Nothing

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en flObtenVentasWSMarengo", ex.Message)
                Return False
            End Try
        End If

        Return True

    End Function

    Private Function flObtenExistenciaMarengoWS() As Boolean 'Clase que consume las Transactions 
        Console.WriteLine(vbCrLf & "## Inicia flObtenExistenciaMarengoWS() ## " & vbCrLf & "Funcion que consume las Transactions" & vbCrLf)
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim contractID As String = fgObtenerConexionBD("contract_id")

        Dim vObj As New wsVentasMarengo.GrowerClient
        Dim vObjRes As New wsVentasMarengo.TransactionV2
        Dim vcProduct As String = ""
        Dim DTResultado As New DataTable

        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)

        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

        Dim vcResultado As String = ""

        'Formateo de fechas para consumo de URL de API'
        Dim vdFecha_format As String = Format(DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date), "yyyy-MM-dd") '-1

        Dim api = New ExecuteAPI '//Llamado  a clase de consumo api
        Dim RespItem = New Response_Transactions_v3 '//Response de Api

        'URL y valores enviar
        Dim apiUrl = $"http://marengosite.com/GrowerAPI/transactions/v3/{contractID}/{vdFecha_format}"
        'Dim apiUrl = "https://mocki.io/v1/6b67f82e-c93b-4b43-b12f-644a9d4508b7"
        Console.WriteLine("### URL_API: " + apiUrl)

        'Guardamos respuesta de API
        Dim response = api.MGet(apiUrl)

        ' Deserializamos la respuesta JSON en una lista de objetos Response_Transactions
        Dim responseObjectList As List(Of Response_Transactions_v3) = JsonConvert.DeserializeObject(Of List(Of Response_Transactions_v3))(response)

        ' Agregar columnas a la DataTable
        DTResultado.Columns.Add("branch", GetType(String))
        DTResultado.Columns.Add("reference", GetType(String))
        DTResultado.Columns.Add("product", GetType(String))
        DTResultado.Columns.Add("lotNo", GetType(String))
        DTResultado.Columns.Add("palletIds", GetType(String))
        DTResultado.Columns.Add("tSource", GetType(String))
        DTResultado.Columns.Add("vendorProductCode", GetType(String))
        DTResultado.Columns.Add("commodity", GetType(String))
        DTResultado.Columns.Add("var", GetType(String))
        DTResultado.Columns.Add("variety", GetType(String))
        DTResultado.Columns.Add("pack", GetType(String))
        DTResultado.Columns.Add("pack2", GetType(String))
        DTResultado.Columns.Add("packaging", GetType(String))
        DTResultado.Columns.Add("contCode", GetType(String))
        DTResultado.Columns.Add("container", GetType(String))
        DTResultado.Columns.Add("sizeCode", GetType(String))
        DTResultado.Columns.Add("size", GetType(String))
        DTResultado.Columns.Add("gradeCode", GetType(String))
        DTResultado.Columns.Add("grade", GetType(String))
        DTResultado.Columns.Add("floor", GetType(Integer))
        DTResultado.Columns.Add("received", GetType(Integer))
        DTResultado.Columns.Add("unpack", GetType(Integer))
        DTResultado.Columns.Add("repack", GetType(Integer))
        DTResultado.Columns.Add("shipped", GetType(Integer))
        DTResultado.Columns.Add("shippedret", GetType(Integer))
        DTResultado.Columns.Add("unreceived", GetType(Integer))
        DTResultado.Columns.Add("endOfDay", GetType(Integer))

        Try
            ' Recorrer la lista de objetos de la nueva API y agregar los datos a la DataTable
            For Each item As Response_Transactions_v3 In responseObjectList
                Dim newRow As DataRow = DTResultado.NewRow()
                newRow("branch") = item.branch
                newRow("reference") = item.reference
                newRow("product") = item.product
                newRow("lotNo") = item.lotNo
                newRow("palletIds") = item.palletIds
                newRow("tSource") = item.tSource
                newRow("vendorProductCode") = item.vendorProductCode
                newRow("commodity") = item.commodity
                newRow("var") = item.var
                newRow("variety") = item.variety
                newRow("pack") = item.pack
                newRow("pack2") = item.pack2
                newRow("packaging") = item.packaging
                newRow("contCode") = item.contCode
                newRow("container") = item.container
                newRow("sizeCode") = item.sizeCode
                newRow("size") = item.size
                newRow("gradeCode") = item.gradeCode
                newRow("grade") = item.grade
                newRow("floor") = item.floor
                newRow("received") = item.received
                newRow("unpack") = item.unpack
                newRow("repack") = item.repack
                newRow("shipped") = item.shipped
                newRow("shippedret") = item.shippedret
                newRow("unreceived") = item.unreceived
                newRow("endOfDay") = item.endOfDay

                DTResultado.Rows.Add(newRow)
            Next

            Console.WriteLine("### Se grabó la información de API en DTResultado ###")
            'Inserta DTResultado en BD
            vcResultado = fgGrabainventariosMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", sucursal), vdFecha)

            If vcResultado <> "" Then
                EscribeEnBitacora(vcResultado)
                Exit Function
            End If

            EscribeEnBitacora("Se inserto la informacion de Inventarios Marengo con Exito")

            Return True

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("rafael.gomez@aparedes.com.mx", Nothing, "Error en flObtenExistenciaMarengoWS", ex.Message)
            Return False
        End Try



    End Function


#End Region

#Region "Crystal Reports"
    Public Function LoginCR(ByVal cr As ReportDocument, ByVal sServidor As String, ByVal sBaseDatos As String, ByVal sUsuario As String, ByVal sPwd As String) As Boolean

        Dim oInfo As New ConnectionInfo

        Dim subObj As SubreportObject

        Dim obj As ReportObject

        oInfo.ServerName = sServidor

        oInfo.DatabaseName = sBaseDatos

        oInfo.UserID = sUsuario

        oInfo.Password = sPwd

        If Not AplicarCR(cr, oInfo) Then

            Return False

        End If

        For Each obj In cr.ReportDefinition.ReportObjects

            If obj.Kind = ReportObjectKind.SubreportObject Then

                subObj = obj

                If (Not AplicarCR(cr.OpenSubreport(subObj.SubreportName), oInfo)) Then

                    Return False

                End If

            End If

        Next obj

        Return True

    End Function

    Private Function AplicarCR(ByVal cr As ReportDocument, ByVal oInfo As ConnectionInfo) As Boolean

        Dim tInfo As TableLogOnInfo

        Dim tbl As Table

        'A cada tabla se le aplica logon info

        For Each tbl In cr.Database.Tables

            tInfo = tbl.LogOnInfo

            tInfo.ConnectionInfo = oInfo

            tbl.ApplyLogOnInfo(tInfo)

            'Verificar si el LOGIN fue correcto

            If tbl.TestConnectivity() Then

                'Cambiar Ubicación

                If tbl.Location.IndexOf(".") > 0 Then

                    tbl.Location = tbl.Location.Substring(tbl.Location.LastIndexOf(". ") + 1)

                Else

                    tbl.Location = tbl.Location

                End If
            Else

                Return False

            End If

        Next tbl

        Return True

    End Function

    Friend Sub AgregarParametro(ByVal sNomCampo As String, ByVal xValCampo As String, ByRef oRpt As ReportDocument)


        Dim xGrupoValor As CrystalDecisions.Shared.ParameterValues

        Dim xValor As CrystalDecisions.Shared.ParameterDiscreteValue

        xGrupoValor = New CrystalDecisions.Shared.ParameterValues

        xValor = New CrystalDecisions.Shared.ParameterDiscreteValue

        xValor.Value = xValCampo

        xGrupoValor.Add(xValor)

        oRpt.DataDefinition.ParameterFields(sNomCampo).ApplyCurrentValues(xGrupoValor)

    End Sub
#End Region

#Region "Eventos"

    Private Sub FrmM1505001_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If aplicacionExcel IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(aplicacionExcel)
            aplicacionExcel = Nothing
        End If

    End Sub

    Private Sub FrmM1505001_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DAO = Sistema.DataAccessCls.DevuelveInstancia
        'plObtenDatosVentasDiarias()
        'plObtenExistenciaMarengoAlterna("21/11/2016")
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Interval = 1000000

        Dim vnHora As DateTime
        Dim vbEnvioCorreo As Boolean = False
        Dim vDT As New DataTable
        Dim vDT2 As New DataTable
        Dim vdFecha As Date
        Dim vdFechaActual As Date

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")

        vnHora = DAO.RegresaFechaDelSistema
        vdFecha = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema)
        vdFechaActual = DAO.RegresaFechaDelSistema()

        'plObtenExistenciaMarengo("29/01/2016")
        'plObtenInventariosMarengo("20/01/2016")
        'plObtenDatosVentasDiariasMarengoAlterna()

        'If Not flEnviarMail("joser2203@gmail.com", Nothing, "test starttls", "nada") Then
        '    MsgBox("error", MsgBoxStyle.Information, "Atención...")
        'End If

        If vnHora.Hour >= 7 Then

            Dim vcSQL As String = ""


            'If fgObtenParametroEMB("ENVIAEXISTPISO",sucursal) = "S" Then
            '    EscribeEnBitacora("Envia informe de Existencia en Piso")

            '    vcSQL = "SELECT * FROM EYE_ENVIOEXISTENCIASPISO(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"


            '    DAO.RegresaConsultaSQL(vcSQL, vDT)
            '    vcSQL = ""

            '    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            '        plObtenDatosExistenciasPiso()
            '    End If
            'End If


            'If fgObtenParametroEMB("ENVIAPORCCALIDAD", sucursal) = "S" Then
            '    EscribeEnBitacora("Envia informe de Porcentaje de Calidad")

            '    vcSQL = "SELECT * FROM EYE_ENVIOPORCCALIDAD(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"


            '    DAO.RegresaConsultaSQL(vcSQL, vDT)
            '    vcSQL = ""

            '    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            '        plObtenPorcentajeCalidad()
            '    End If
            'End If

            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-MX")


            If fgObtenParametroEMB("ENVIAEXISTDIST", sucursal) = "S" Then
                EscribeEnBitacora("Trata de generar la informacion")

                vcSQL = "SELECT * FROM EYE_ENVIOEXISTENCIASDISTRIBUIDORAS(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112) AND CSUCURSAL = '" & sucursal & "'"

                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plObtenDatos() 'Distribuidora HM
                End If

            End If


                If fgObtenParametroEMB("ENVIAROTACIONDIST", sucursal) = "S" Then
                EscribeEnBitacora("Trata de generar la informacion")

                vcSQL = "SELECT * FROM EYE_ENVIOROTACIONPRODUCTOS(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"


                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plObtenDatosRotacionDistribuidoras()
                End If

            End If


            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-MX")


            'If fgObtenParametroEMB("ENVIAPREENFRIADOS", "001") = "S" Then
            '    vcSQL = "SELECT * FROM EYE_ENVIOPREENFRIADOS(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112) AND CSUCURSAL = '001'"
            '    DAO.RegresaConsultaSQL(vcSQL, vDT2)

            '    If vDT2 Is Nothing OrElse vDT2.Rows.Count = 0 Then
            '        plObtenEnfriados()
            '    End If
            'End If


            If fgObtenParametroEMB("ENVIAVENTASDIST", sucursal) = "S" Then
                vcSQL = "SELECT * FROM EYE_ENVIOVENTASDIARIASDISTRIBUIDORAS(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112) AND CSUCURSAL = '" & sucursal & "'"

                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plObtenDatosVentasDiarias() 'Comenzara flujo ventas Diarias marengo
                End If
            End If

                If fgObtenParametroEMB("ENVIAEXISTVENTADIST", sucursal) = "S" Then
                'EscribeEnBitacora("Envia informe de Porcentajes de Calidad")
                vcSQL = "SELECT * FROM INF_EXISTVENTADIST(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"
                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""
                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plEnviaExportacionVentaDisp()
                End If
            End If



            If fgObtenParametroEMB("ENVIAAVANCELAB", sucursal) = "S" Then
                vcSQL = "SELECT * FROM NOM_ENVIOAVANCELABORES(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"

                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plObtenDatosAvanceLabores()
                End If

            End If


            If fgObtenParametroEMB("ENVIARENTATRACT", sucursal) = "S" Then
                vcSQL = "SELECT * FROM NOM_ENVIORENTATRACTORES(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"

                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plObtenDatosRentaTractores()
                End If

            End If

            If fgObtenParametroEMB("ENVIAVENTASSEM", sucursal) = "S" Then
                Dim vnDiaSemana As Integer = DAO.RegresaDatoSQL("SELECT DATEPART(dw,GETDATE()) ")


                If vnDiaSemana = 1 Then
                    '' envio reportes de ventas semanales

                    vcSQL = "SELECT * FROM EYE_ENVIOVENTASEMANALESDISTRIBUIDORAS(NOLOCK)"
                    vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"
                    vcSQL += "AND CSUCURSAL = '" & sucursal & "'"

                    DAO.RegresaConsultaSQL(vcSQL, vDT)
                    vcSQL = ""

                    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                        plEnviaVentasSemanales()
                    End If

                End If

            End If


            'If fgObtenParametroEMB("ENVIAANTIGUEDADPISO",sucursal) = "S" Then
            '    vcSQL = "SELECT * FROM EYE_ENVIOANTIGUEDADPISO(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"
            '    DAO.RegresaConsultaSQL(vcSQL, vDT)


            '    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            '        plEnviaAntiguedadEnPiso()
            '    End If
            'End If

            '************
            'COMENTADO POR QUE NO EXISTE EN JALISCO 
            'vcSQL = "SELECT * FROM NOM_ELEGIBLESTARJETA(NOLOCK)"
            'vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"
            'DAO.RegresaConsultaSQL(vcSQL, vDT)

            'If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            'plObtenEmpleadosMayoresDeEdad()
            'End If

        End If


        If vnHora.Hour >= 8 Then

            Dim vnDiaSemana As Integer = DAO.RegresaDatoSQL("SELECT DATEPART(dw,GETDATE()) ")
            Dim vcSQL As String = ""

            'If vnDiaSemana = 1 Then
            '    EscribeEnBitacora("Envia informe de Cheques sin Facturas Cln")

            '    vcSQL = "SELECT * FROM EYE_ENVIOCHEQUESCLN(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"


            '    DAO.RegresaConsultaSQL(vcSQL, vDT)
            '    vcSQL = ""

            '    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            '        plObtenChequesSinFacturasCln()
            '    End If
            'End If

            'If vnDiaSemana = 1 Or vnDiaSemana = 5 Then
            '    EscribeEnBitacora("Envia informe de Cheques sin Facturas Jal")

            '    vcSQL = "SELECT * FROM EYE_ENVIOCHEQUESJAL(NOLOCK)"
            '    vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"


            '    DAO.RegresaConsultaSQL(vcSQL, vDT)
            '    vcSQL = ""

            '    If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
            '        plObtenChequesSinFacturasJal()
            '    End If
            'End If

        End If

        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-MX")

        'If fgObtenParametroEMB("ENVIAPORCEMP", sucursal) = "S" Then
        '    If (vnHora.Hour >= 8 And vnHora.Minute >= 30) Or (vnHora.Hour >= 9) Then

        '        Dim vcSQL As String = ""

        '        EscribeEnBitacora("Envia informe de Porcentajes de Calidad")

        '        vcSQL = "SELECT * FROM EYE_PORCENTAJESEMPAQUE(NOLOCK)"
        '        vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112)"

        '        DAO.RegresaConsultaSQL(vcSQL, vDT)
        '        vcSQL = ""

        '        If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
        '            plGeneraFormatoExcel()
        '        End If


        '    End If

        'End If

        If vnHora.Hour >= 9 Then
            'SE COMENTA POR QUE NO FUNCIONA PARA JALISCO
            Dim vcSQL As String = ""

            EscribeEnBitacora("Envia informe de Asistencia Diaria TSPV")

            'vcSQL = "SELECT * FROM TSPV_ENVIOASISTENCIADIARIA(NOLOCK)"
            ' vcSQL = vcSQL & vbCrLf & "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFechaActual, "yyyyMMdd") & "',112)"

            'DAO.RegresaConsultaSQL(vcSQL, vDT)
            'vcSQL = ""

            If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                'plEnviaAsistenciaTSPV() 'SE COMENTA POR QUE NO FUNCIONA PARA JALISCO
            End If

        End If
        If vnHora.Hour > 13 Then

            If fgObtenParametroEMB("ENVIACOMPNOMINA", sucursal) = "S" Then
                Dim vcSQL As String = ""

                EscribeEnBitacora("Envia informe de Complemento de Nomina")

                vcSQL = "SELECT * FROM NOM_ENVIOCOMPLEMENTONOMINA(NOLOCK)"
                vcSQL += "WHERE CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(vdFecha, "yyyyMMdd") & "',112) AND CSUCURSAL = '" & sucursal & "'"

                DAO.RegresaConsultaSQL(vcSQL, vDT)
                vcSQL = ""

                If vDT Is Nothing OrElse vDT.Rows.Count = 0 Then
                    plGeneraComplementoNomina()
                End If

            End If


        End If


    End Sub


    Private Sub plGeneraComplementoNomina()

        'Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", "001")
        'Dim oRpt = New ReportDocument
        'Dim Proc As New System.Diagnostics.Process
        'Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        'Dim vcArchivo As String = ""
        'Dim vcSQL As String = ""
        'Dim vDT As New DataTable
        'Dim vcSemana As String = ""
        'Dim vcDia As String = ""

        'Try

        '    vcSQL = "SELECT CSEMANA,DATEDIFF(DAY,dfec_ini,'" & Format(vdFecha, "yyyyMMdd") & "')+1 AS CDIA FROM CTL_Semanas WHERE ccve_temporada = '" & vcTemporada & "' AND ccve_nomina = '02' AND '" & Format(vdFecha, "yyyyMMdd") & "' BETWEEN CONVERT(VARCHAR(20),dfec_ini,112) AND CONVERT(VARCHAR(20),dfec_fin,112)"

        '    DAO.RegresaConsultaSQL(vcSQL, vDT)

        '    If Not vDT Is Nothing AndAlso vDT.Rows.Count > 0 Then
        '        vcSemana = vDT(0)("CSEMANA")
        '        vcDia = vDT(0)("CDIA")
        '    End If

        '    Dim oStrem As New System.IO.MemoryStream

        '    EscribeEnBitacora("Genera PDF de Complemento de Nomina")

        '    oRpt = New ReportDocument

        '    oRpt.Load("C:\CROP\RPT_COMPLEMENTONOMINAEMPAQUE.rpt")

        '    LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
        '    AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
        '    AgregarParametro("@PRMNOMINA", "02", oRpt)
        '    AgregarParametro("@PRMSEMANA", vcSemana, oRpt)
        '    AgregarParametro("@PRMDIA", vcDia, oRpt)

        '    oStrem = New System.IO.MemoryStream

        '    oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

        '    vcArchivo = "C:\ARCHIVOS\CLN COMPLEMENTO DE NOMINA DEL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

        '    If File.Exists(vcArchivo) Then
        '        File.Delete(vcArchivo)
        '    End If

        '    Dim vcAdjuntos As New ArrayList
        '    Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
        '    ArchivoPDF = New System.IO.FileStream(vcArchivo, IO.FileMode.Create)

        '    ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
        '    ArchivoPDF.Flush()
        '    ArchivoPDF.Close()
        '    ArchivoPDF.Dispose()
        '    ArchivoPDF = Nothing

        '    If File.Exists(vcArchivo) Then
        '        vcAdjuntos.Add(vcArchivo)
        '    End If

        '    If vcAdjuntos.Count > 0 Then

        '        EscribeEnBitacora("Se enviara correo de PDF de Complemento de Nomina")

        '        ' Enviamos correo
        '        flEnviarMail(vcCorreosComplementoNomina, vcAdjuntos, "CLN INFORME DE COMPLEMENTO DE NOMINA DEL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")

        '        DAO.EjecutaComandoSQL("INSERT NOM_ENVIOCOMPLEMENTONOMINA SELECT '" & Format(vdFecha, "yyyyMMdd") & "','001'")

        '        EscribeEnBitacora("Se inserta en tabla de COMPLEMENTO DE NOMINA")

        '    End If

        '    'If File.Exists(vcArchivo) Then
        '    '    File.Delete(vcArchivo)
        '    'End If

        'Catch ex As Exception
        '    EscribeEnBitacora(ex.Message)
        'End Try

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vcArchivo As String
        Dim vcSQL As String = ""
        Dim vDT As New DataTable
        Dim vcSemana As String = ""
        Dim vcDia As String = ""


        EscribeEnBitacora("Obteniendo informacion de Preenfriado")

        vcSQL = "SELECT CSEMANA,DATEDIFF(DAY,dfec_ini,'" & Format(vdFecha, "yyyyMMdd") & "')+1 AS CDIA FROM CTL_Semanas WHERE ccve_temporada = '" & vcTemporada & "' AND ccve_nomina = '02' AND '" & Format(vdFecha, "yyyyMMdd") & "' BETWEEN CONVERT(VARCHAR(20),dfec_ini,112) AND CONVERT(VARCHAR(20),dfec_fin,112)"

        DAO.RegresaConsultaSQL(vcSQL, vDT)

        If Not vDT Is Nothing AndAlso vDT.Rows.Count > 0 Then
            vcSemana = vDT(0)("CSEMANA")
            vcDia = vDT(0)("CDIA")
        End If
        vDT.Dispose()


        If vcSemana <> "" Then


            Try
                oRpt.Load("C:\CROP\RPT_COMPLEMENTONOMINAEMPAQUE.rpt")

                LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
                AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
                AgregarParametro("@PRMNOMINA", "02", oRpt)
                AgregarParametro("@PRMSEMANA", vcSemana, oRpt)
                AgregarParametro("@PRMDIA", vcDia, oRpt)


                Dim oStrem As New System.IO.MemoryStream

                oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

                vcArchivo = "C:\CROP\COMPLEMENTO DE NOMINA AL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

                'Si lo deseamos escribimos el pdf a disco.
                Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
                ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
                ArchivoPDF.Flush()
                ArchivoPDF.Close()

                EscribeEnBitacora("Se creo PDF de Complemento de Nomina")

                If File.Exists(vcArchivo) Then
                    Dim vcAdjuntos As New ArrayList()

                    vcAdjuntos.Add(vcArchivo)

                    EscribeEnBitacora("Se enviara coreo de PDF de Complemento de Nomina")

                    ' Enviamos correo
                    'plEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "PREENFRIADOS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                    flEnviarMail(vcCorreosComplementoNomina, vcAdjuntos, "INFORME DE COMPLEMENTO DE NOMINA AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                    DAO.EjecutaComandoSQL("INSERT NOM_ENVIOCOMPLEMENTONOMINA SELECT '" & Format(vdFecha, "yyyyMMdd") & "','001'")

                    EscribeEnBitacora("Se inserta en tabla de Complemento de Nomina")
                End If

                'If File.Exists(vcArchivo) Then
                '    File.Delete(vcArchivo)
                'End If

            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plGeneraComplementoNomina", ex.Message)
            End Try


        Else
            EscribeEnBitacora("No hubo datos para Complemento de Nomina")
            DAO.EjecutaComandoSQL("INSERT NOM_ENVIOCOMPLEMENTONOMINA SELECT '" & Format(vdFecha, "yyyyMMdd") & "','001'")
            EscribeEnBitacora("Se inserta en tabla de Complemento de Nomina")
            Exit Sub
        End If

        Proc.Dispose()
        'Proc.Kill()
        oRpt = Nothing


    End Sub

    Private Sub plEnviaVentasSemanales()
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")


        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vdFechaIni As Date = DateAdd(DateInterval.Day, -7, DAO.RegresaFechaDelSistema.Date)
        Dim vdFechaFin As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vcArchivo As String = ""


        Try

            EscribeEnBitacora("Genera PDF de ventas Semanales HM")

            oRpt.Load("C:\CROP\RPT_VENTASDIARIASHMSEMANAL.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFechaIni, oRpt)
            AgregarParametro("@PRMFECHAFIN", vdFechaFin, oRpt)

            Dim oStrem As New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\CLN VENTAS HM SEMANAL DEL " & UCase(Format(vdFechaIni, "dd-MMM-yy")) & " AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & " TEMP " & vcTemporada & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            'Si lo deseamos escribimos el pdf a disco.
            Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            Dim vcAdjuntos As New ArrayList()

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            EscribeEnBitacora("Se creo PDF de Ventas Semanales de HM")

            Proc.Dispose()
            oRpt = Nothing

            EscribeEnBitacora("Genera PDF de ventas Semanales MARENGO")

            oRpt = New ReportDocument

            oRpt.Load("C:\CROP\RPT_VENTASDIARIASMARENGOSEMANAL.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFechaIni, oRpt)
            AgregarParametro("@PRMFECHAFIN", vdFechaFin, oRpt)
            AgregarParametro("@PRMSUCURSAL", "CULIACAN", oRpt)

            oStrem = New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\CLN VENTAS MARENGO SEMANAL DEL " & UCase(Format(vdFechaIni, "dd-MMM-yy")) & " AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")) & " TEMP " & vcTemporada & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            ArchivoPDF = New System.IO.FileStream(vcArchivo, IO.FileMode.Create)

            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            If vcAdjuntos.Count > 0 Then

                EscribeEnBitacora("Se enviara correo de PDF de Ventas Semanales")

                ' Enviamos correo
                flEnviarMail(vcCorreosVentasSemanales, vcAdjuntos, "CLN INFORME DE VENTAS SEMANALES DEL " & UCase(Format(vdFechaIni, "dd-MMM-yy")) & " AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")
                'flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "INFORME DE VENTAS SEMANALES DEL " & UCase(Format(vdFechaIni, "dd-MMM-yy")) & " AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOVENTASEMANALESDISTRIBUIDORAS SELECT '" & Format(vdFechaFin, "yyyyMMdd") & "','001'")

                EscribeEnBitacora("Se inserta en tabla de VENTAS SEMANALES")

            End If

            'If File.Exists(vcArchivo) Then
            '    File.Delete(vcArchivo)
            'End If

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plEnviaVentasSemanales", ex.Message)
        End Try

    End Sub

    Private Sub plEnviaAntiguedadEnPiso()
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vcArchivo As String = ""


        Try

            'EscribeEnBitacora("Genera PDF de ventas Semanales HM")

            oRpt.Load("\\192.168.2.21\crop\EXES\INFORMES\RPT_BULTOS_PISO.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@ccve_temporada", vcTemporada, oRpt)
            AgregarParametro("@ccve_empaque", "001", oRpt)
            AgregarParametro("@ccve_agricultor", "0001", oRpt)
            AgregarParametro("@chep", "", oRpt)

            Dim oStrem As New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\REPORTE DE ANTIGUEDAD EN PISO " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            'Si lo deseamos escribimos el pdf a disco.
            Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            Dim vcAdjuntos As New ArrayList()

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            'EscribeEnBitacora("Se creo PDF de Ventas Semanales de HM")


            If vcAdjuntos.Count > 0 Then

                'EscribeEnBitacora("Se enviara correo de PDF de Ventas Semanales")

                ' Enviamos correo
                flEnviarMail(vcCorreosEnfriados, vcAdjuntos, "REPORTE DE ANTIGUEDAD EN PISO " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada, "SE ANEXA ARCHIVO")
                'flEnviarMail("<alan.zazueta@aparedes.com.mx>,<enriqueca@aparedes.com.mx>", vcAdjuntos, "REPORTE DE ANTIGUEDAD EN PISO " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada, "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_ENVIOANTIGUEDADPISO SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                'EscribeEnBitacora("Se inserta en tabla de VENTAS SEMANALES")

            End If

            'If File.Exists(vcArchivo) Then
            '    File.Delete(vcArchivo)
            'End If

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plEnviaAntiguedadEnPiso", ex.Message)
        End Try


    End Sub

    Private Sub plEnviaExportacionVentaDisp()

        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)
        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
        Dim vcArchivo As String = ""


        Try

            'EscribeEnBitacora("Genera PDF de ventas Semanales HM")

            oRpt.Load("C:\CROP\RPT_EXISTVENTDISTRIBUIDORASCLN.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMTEMPORADA", vcTemporada, oRpt)
            AgregarParametro("@PRMFECHA", vdFecha, oRpt)


            Dim oStrem As New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\CLN EXPORTACION DISPONIBLE PARA VENTA " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            'Si lo deseamos escribimos el pdf a disco.
            Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            Dim vcAdjuntos As New ArrayList()

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            'EscribeEnBitacora("Se creo PDF de Ventas Semanales de HM")


            If vcAdjuntos.Count > 0 Then

                'EscribeEnBitacora("Se enviara correo de PDF de Ventas Semanales")

                ' Enviamos correo

                flEnviarMail(vcCorreosDisponibleExportacion, vcAdjuntos, "CLN EXPORTACION DISPONIBLE PARA VENTA " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada, "SE ANEXA ARCHIVO")

                'flEnviarMail("<edwin@aparedes.com.mx>", vcAdjuntos, "CLN EXPORTACION DISPONIBLE PARA VENTA " & UCase(Format(vdFecha, "dd-MMM-yy")) & " TEMP " & vcTemporada, "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT INF_EXISTVENTADIST SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                'EscribeEnBitacora("Se inserta en tabla de VENTAS SEMANALES")

            End If

            'If File.Exists(vcArchivo) Then
            '    File.Delete(vcArchivo)
            'End If

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plEnviaExportacionVentaDisp", ex.Message)
        End Try

    End Sub

    Private Sub plEnviaAsistenciaTSPV()


        Dim oRpt = New ReportDocument
        Dim Proc As New System.Diagnostics.Process
        Dim vcAdjuntosPDF As New ArrayList
        Dim vdFecha As Date = DAO.RegresaFechaDelSistema.Date
        Dim vcArchivo As String = ""


        Try

            EscribeEnBitacora("Genera PDF de Asistencia Diaria TSPV")

            oRpt.Load("C:\CROP\RPT_ASISTENCIADIARIATSPV.rpt")

            LoginCR(oRpt, DAO.GetNombreServidor, DAO.GetNombreBaseDeDatos, DAO.GetLoginUsuario, DAO.GetPassUsuario)
            AgregarParametro("@PRMFECHA", vdFecha, oRpt)

            Dim oStrem As New System.IO.MemoryStream

            oStrem = CType(oRpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat), System.IO.MemoryStream)

            vcArchivo = "C:\ARCHIVOS\TSPV ASISTENCIA DIARIA DEL " & UCase(Format(vdFecha, "dd-MMM-yy")) & ".pdf"

            If File.Exists(vcArchivo) Then
                File.Delete(vcArchivo)
            End If

            'Si lo deseamos escribimos el pdf a disco.
            Dim ArchivoPDF As New System.IO.FileStream(vcArchivo, IO.FileMode.Create)
            ArchivoPDF.Write(oStrem.ToArray, 0, oStrem.ToArray.Length)
            ArchivoPDF.Flush()
            ArchivoPDF.Close()
            ArchivoPDF.Dispose()
            ArchivoPDF = Nothing

            Dim vcAdjuntos As New ArrayList()

            If File.Exists(vcArchivo) Then
                vcAdjuntos.Add(vcArchivo)
            End If

            EscribeEnBitacora("Se creo PDF de Asistencia Diaria TSPV")

            Proc.Dispose()
            oRpt = Nothing

            If vcAdjuntos.Count > 0 Then

                EscribeEnBitacora("Se enviara correo de Asistencia Diaria TSPV")

                ' Enviamos correo
                flEnviarMail(vcCorreosAsistenciaTSPV, vcAdjuntos, "TSPV INFORME DE ASISTENCIA DEL DIA " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")
                'flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "INFORME DE VENTAS SEMANALES DEL " & UCase(Format(vdFechaIni, "dd-MMM-yy")) & " AL " & UCase(Format(vdFechaFin, "dd-MMM-yy")), "SE ANEXAN ARCHIVOS")

                DAO.EjecutaComandoSQL("INSERT TSPV_ENVIOASISTENCIADIARIA SELECT '" & Format(vdFecha, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de ASISTENCIA DIARIA TSPV")

            End If

            'If File.Exists(vcArchivo) Then
            '    File.Delete(vcArchivo)
            'End If

        Catch ex As Exception
            EscribeEnBitacora(ex.Message)
            flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plEnviaAsistenciaTSPV", ex.Message)
        End Try

    End Sub

    'Private Sub plObtenInventariosMarengo(ByVal prmFecha As Date)

    '    Dim vcResultado As String = ""
    '    Dim DTResultado As New DataTable
    '    Dim wsservice As New wsMarengo.GrowerService
    '    Dim Parametro1 As String = ""
    '    Dim Parametro2 As String = ""

    '    DTResultado.Columns.Add("Grower", GetType(String))
    '    DTResultado.Columns.Add("Branch", GetType(String))
    '    DTResultado.Columns.Add("CommodityName", GetType(String))
    '    DTResultado.Columns.Add("PackStyle", GetType(String))
    '    DTResultado.Columns.Add("Label", GetType(String))
    '    DTResultado.Columns.Add("Size", GetType(String))
    '    DTResultado.Columns.Add("UoM", GetType(String))
    '    DTResultado.Columns.Add("Inventory", GetType(Double))
    '    DTResultado.Columns.Add("Ins", GetType(Double))
    '    DTResultado.Columns.Add("Outs", GetType(Double))

    '    Try
    '        wsservice.Url = "https://www.marengosite.com/marengowebservices/growerservice.asmx"

    '        '' Obteniendo Inventarios
    '        Parametro1 = "001"
    '        Parametro2 = "p@r3d3s"

    '        Dim vdFecha As Date = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)

    '        'vcFecha = "24/01/2015"

    '        Dim vObjMarengo() As wsMarengo.GrowerInventory

    '        vObjMarengo = wsservice.Inventory(Parametro1, Parametro2)

    '        For i As Integer = 0 To vObjMarengo.Length - 1
    '            Dim vRow As DataRow

    '            vRow = DTResultado.NewRow

    '            vRow("Grower") = vObjMarengo(i).Grower
    '            'vRow("Branch") = vObjMarengo(i).Branch
    '            'vRow("CommodityName") = vObjMarengo(i).CommodityName
    '            'vRow("PackStyle") = vObjMarengo(i).PackStyle
    '            'vRow("Label") = vObjMarengo(i).Label
    '            'vRow("Size") = vObjMarengo(i).Size
    '            'vRow("UoM") = vObjMarengo(i).UoM
    '            'vRow("Inventory") = vObjMarengo(i).Inventory
    '            'vRow("Ins") = vObjMarengo(i).Ins
    '            'vRow("Outs") = vObjMarengo(i).Outs

    '            DTResultado.Rows.Add(vRow)

    '        Next


    '        EscribeEnBitacora("Se obtiene la información de existencias del webservice")

    '        vcResultado = fgGrabainventariosMarengo(DTResultado, fgObtenParametroEMB("TEMPORADA", "001"), vdFecha)

    '        If vcResultado <> "" Then
    '            EscribeEnBitacora(vcResultado)
    '            Exit Sub
    '        End If

    '        EscribeEnBitacora("Se inserto la informacion de Inventarios Marengo con Exito")

    '    Catch ex As Exception
    '        EscribeEnBitacora(ex.Message)
    '    End Try

    '    'Dim ConnectionString As String = "Provider=SQLOLEDB.1;Password=marengo123;Persist Security Info=True;User ID=MarengoRO;Initial Catalog=MarengoFoods;Data Source=db.marengoserver.dyndns.org;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    '    'Dim vcFecha As String = Format(prmFecha, "MM/dd/yyyy")

    '    'Using connection As New OleDb.OleDbConnection(ConnectionString)

    '    '    EscribeEnBitacora("Se conecta al servidor de Marengo")

    '    '    connection.Open()

    '    '    Dim dAdpter As New OleDb.OleDbDataAdapter("EXEC sExcelGrowerInOuts 2,'" & vcFecha & "'", connection)
    '    '    Dim DT As New DataTable

    '    '    EscribeEnBitacora("Obtiene informacion de Existencia")
    '    '    dAdpter.Fill(DT)

    '    '    If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

    '    '        EscribeEnBitacora("Graba informacion de Inventarios de Marengo")
    '    '        vcResultado = fgGrabainventariosMarengo(DT, fgObtenParametroEMB("TEMPORADA"), prmFecha)

    '    '    End If

    '    'End Using


    'End Sub


    Private Sub plGeneraFormatoExcel()

        '***********************************************************
        ' ENRIQUE ALONSO CORRAL AGUILAR
        ' 11 / FEB / 2019
        ' GENERAR FORMATO DE PORCENTAJE DE EMPAQUE POR CALIDADES

        Dim xlApp As New Microsoft.Office.Interop.Excel.ApplicationClass
        Dim xlLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim xlHoja As Microsoft.Office.Interop.Excel.Worksheet


        Dim vcNombreArchivo As String
        Dim vnColumnaActual As Integer
        Dim vnColumnaFinal As Integer
        'se obtienen valores de tabla CONFIG_CONEXION_SUCURSAL        
        Dim sucursal As String = fgObtenerConexionBD("cve_sucursal")
        Dim vcTemporada As String = fgObtenParametroEMB("TEMPORADA", sucursal)


        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        Dim vdFechaInicio As Date
        Dim vdFechaFin As Date
        Dim vdFechaActual As Date


        xlLibro = xlApp.Workbooks.Open("\\192.168.2.21\crop\EXES\Porcentajes de empaque.xlsx")
        'xlLibro = xlApp.Workbooks.Open("C:\CROP\Porcentajes de empaque.xlsx")
        xlHoja = xlLibro.Worksheets.Application.Sheets("Hoja1")
        With xlHoja



            vdFechaInicio = DateAdd(DateInterval.Day, -8, DAO.RegresaFechaDelSistema.Date)
            vdFechaFin = DateAdd(DateInterval.Day, -1, DAO.RegresaFechaDelSistema.Date)
            vnColumnaActual = 2

            'xlApp.Visible = True

            Do While vdFechaInicio <= vdFechaFin
                vdFechaActual = vdFechaInicio


                Dim vcSQL As String
                Dim DS As New DataSet
                Dim DS2 As New DataSet
                Dim vcNombreDia As String

                vcNombreDia = DAO.RegresaDatoSQL("SELECT DBO.fgNombreDia('" & Format(vdFechaActual, "yyyyMMdd") & "') AS CDIA")

                Dim vParametros(2) As Object

                vParametros(0) = vcTemporada
                vParametros(1) = "001"
                vParametros(2) = vdFechaActual

                vcSQL = "SPANALISISPORCENTAJECORTEEMPAQUESEMANAL"
                If Not DAO.RegresaConsultaSQL(vcSQL, DS, vParametros) Then
                    Exit Sub
                End If

                vcSQL = "SPANALISISPORCENTAJECORTEEMPAQUESEMANALBER"
                If Not DAO.RegresaConsultaSQL(vcSQL, DS2, vParametros) Then
                    Exit Sub
                End If


                Dim objRango = xlHoja.Range(flLetraExcel(vnColumnaActual) & "5:" & flLetraExcel(vnColumnaActual) & "6")

                objRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone

                objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone

                With objRango.Cells
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlHAlign.xlHAlignJustify
                    .WrapText = True
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Fix(Excel.Constants.xlLTR)
                    .MergeCells = False
                    .ColumnWidth = 13.57
                    .Interior.Color = RGB(215, 215, 215)
                    .Font.Bold = True
                    .Font.Size = 12
                    .Merge()
                End With


                .Cells(5, vnColumnaActual).Value = vcNombreDia & " " & Format(vdFechaActual, "dd/MM/yyyy")

                Dim vRows() As DataRow

                '' BERENJENAS
                '***********************************************************
                ' ENRIQUE CORRAL
                ' 29 / OCT / 2022
                ' SE AGREGO LA SEPARACION DE BERENJENAS DE CAMPO ABIERTO Y MALLAS POR INSTRUCCION DEL ING. MONZON
                ' EN ESTE APARTADO ESTA TOMANDO EL CAMPO ABIERTO

                vRows = DS2.Tables(0).Select("CCVE_CULTIVO = '005' AND LOTE='16'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(7, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(8, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(9, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(10, vnColumnaActual, xlHoja, vRows(0)("Terceras"))
                Else

                    plLlenaValoresCeldaVacia(7, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(8, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(9, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(10, vnColumnaActual, xlHoja)

                End If

                '' BERENJENAS MALLA
                '***********************************************************
                ' EDWIN GANDARILLA
                ' 10 / NOV / 2021
                ' SE AGREGO LA SEPARACION DE BERENJENAS DE CAMPO ABIERTO Y MALLAS POR INSTRUCCION DEL ING. MONZON

                vRows = DS2.Tables(0).Select("CCVE_CULTIVO = '005' AND LOTE='09'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then
                    plLlenaValoresCelda(12, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(13, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(14, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(15, vnColumnaActual, xlHoja, vRows(0)("Terceras"))
                Else
                    plLlenaValoresCeldaVacia(12, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(13, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(14, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(15, vnColumnaActual, xlHoja)
                End If



                '' CHILE VERDE

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '003'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(17, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(18, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(19, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(20, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(17, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(18, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(19, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(20, vnColumnaActual, xlHoja)

                End If


                '' CHILE ROJO

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '008'")



                If Not vRows Is Nothing AndAlso vRows.Length > 0 AndAlso Not vRows(0)("Empaque") Is DBNull.Value Then

                    plLlenaValoresCelda(22, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(23, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(24, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(25, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(22, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(23, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(24, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(25, vnColumnaActual, xlHoja)

                End If


                '' CHILE AMARILLO

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '061'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(27, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(28, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(29, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(30, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(27, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(28, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(29, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(30, vnColumnaActual, xlHoja)

                End If


                '' CHILE NARANJA

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '010'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(32, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(33, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(34, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(35, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(32, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(33, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(34, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(35, vnColumnaActual, xlHoja)

                End If


                '' TOMATE BOLA

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '002'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(37, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(38, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(39, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(40, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(37, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(38, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(39, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(40, vnColumnaActual, xlHoja)

                End If


                '' TOMATE SALADETTE

                vRows = DS.Tables(0).Select("CCVE_CULTIVO = '004'")

                If Not vRows Is Nothing AndAlso vRows.Length > 0 Then

                    plLlenaValoresCelda(42, vnColumnaActual, xlHoja, vRows(0)("Empaque"))
                    plLlenaValoresCelda(43, vnColumnaActual, xlHoja, vRows(0)("Primeras"))
                    plLlenaValoresCelda(44, vnColumnaActual, xlHoja, vRows(0)("Segundas"))
                    plLlenaValoresCelda(45, vnColumnaActual, xlHoja, vRows(0)("Terceras"))

                Else

                    plLlenaValoresCeldaVacia(42, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(43, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(44, vnColumnaActual, xlHoja)
                    plLlenaValoresCeldaVacia(45, vnColumnaActual, xlHoja)

                End If

                vnColumnaFinal = vnColumnaActual

                vnColumnaActual = vnColumnaActual + 1
                vdFechaInicio = DateAdd(DateInterval.Day, 1, vdFechaInicio)
            Loop


        End With
        vcNombreArchivo = "C:\ARCHIVOS\CLN PORCENTAJE DE EMPAQUE AL DIA " & Format(vdFechaActual, "dd-MM-yyyy") & ".xlsx"

        If Len(Dir(vcNombreArchivo)) > 0 Then
            Kill(vcNombreArchivo)
        End If

        xlHoja.SaveAs(vcNombreArchivo)
        xlLibro.Close()
        xlApp.Quit()

        xlApp = Nothing
        xlLibro = Nothing
        xlHoja = Nothing


        If File.Exists(vcNombreArchivo) Then
            Try
                Dim vcAdjuntos As New ArrayList()

                vcAdjuntos.Add(vcNombreArchivo)

                EscribeEnBitacora("Se enviara correo de Excel de Porcentaje de Empaque")

                ' Enviamos correo
                flEnviarMail(vcCorreosPorcentajeEmpaque, vcAdjuntos, "CLN PORCENTAJE DE EMPAQUE AL " & UCase(Format(vdFechaActual, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'flEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "CLN PORCENTAJE DE EMPAQUE AL " & UCase(Format(vdFechaActual, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "NACIONAL DISPONIBLE AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")
                'plEnviarMail("enriqueca@aparedes.com.mx", vcAdjuntos, "EXISTENCIAS DISTRIBUIDORAS AL " & UCase(Format(vdFecha, "dd-MMM-yy")), "SE ANEXA ARCHIVO")

                DAO.EjecutaComandoSQL("INSERT EYE_PORCENTAJESEMPAQUE SELECT '" & Format(vdFechaActual, "yyyyMMdd") & "'")

                EscribeEnBitacora("Se inserta en tabla de Porcentaje de Empaque")


            Catch ex As Exception
                EscribeEnBitacora(ex.Message)
                flEnviarMail("alfredor@aparedes.com.mx", Nothing, "Error en plGeneraFormatoExcel", ex.Message)
            End Try

        End If

    End Sub

    Public Sub plChequesSinFacturas()

    End Sub


    Private Sub plLlenaValoresCelda(ByVal prmRenglon As Integer, ByVal prmColumna As Integer, ByVal xlHoja As Microsoft.Office.Interop.Excel.Worksheet, ByVal prmValor As Object, Optional ByVal prmPorcentaje As Boolean = True)

        Dim voValor As Object
        With xlHoja

            If prmPorcentaje Then
                .Cells(prmRenglon, prmColumna).Style = "Percent"
                If prmValor Is DBNull.Value Then
                    voValor = 0
                Else
                    voValor = prmValor / 100
                End If
                .Cells(prmRenglon, prmColumna).Value = voValor
                .Cells(prmRenglon, prmColumna).NumberFormat = "0.00%"
            Else
                .Cells(prmRenglon, prmColumna).Value = prmValor
            End If

            Dim objRango = xlHoja.Range(flLetraExcel(prmColumna) & prmRenglon & ":" & flLetraExcel(prmColumna) & prmRenglon)

            objRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            With objRango.Cells
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignJustify
            End With

        End With
    End Sub

    Private Sub plLlenaValoresCeldaCalidad(ByVal prmRenglon As Integer, ByVal prmColumna As Integer, ByVal xlHoja As Microsoft.Office.Interop.Excel.Worksheet, ByVal prmValor As Object, Optional ByVal prmPorcentaje As Boolean = True)


        With xlHoja

            If prmPorcentaje Then
                .Cells(prmRenglon, prmColumna).Style = "Percent"
                .Cells(prmRenglon, prmColumna).Value = prmValor / 100
                .Cells(prmRenglon, prmColumna).NumberFormat = "0.00%"
            Else
                .Cells(prmRenglon, prmColumna).Value = prmValor
            End If

            Dim objRango = xlHoja.Range(flLetraExcel(prmColumna) & prmRenglon & ":" & flLetraExcel(prmColumna) & prmRenglon)

            'objRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            'With objRango.Cells
            ' .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            '.VerticalAlignment = Excel.XlHAlign.xlHAlignJustify
            'End With

        End With
    End Sub

    Private Sub plBordesGruesos(ByVal prmRenglonInicio As Integer, ByVal prmRenglonFin As Integer, ByVal prmColumnaInicio As Integer, ByVal prmColumnaFin As Integer, ByVal xlHoja As Microsoft.Office.Interop.Excel.Worksheet)

        With xlHoja

            Dim objRango = xlHoja.Range(flLetraExcel(prmColumnaInicio) & prmRenglonInicio & ":" & flLetraExcel(prmColumnaFin) & prmRenglonFin)

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium

        End With



    End Sub


    Private Sub plLlenaValoresCeldaVacia(ByVal prmRenglon As Integer, ByVal prmColumna As Integer, ByVal xlHoja As Microsoft.Office.Interop.Excel.Worksheet)


        With xlHoja

            '.Range(Cells(prmRenglon, prmColumna), Cells(prmRenglon, prmColumna)).Select()

            .Cells(prmRenglon, prmColumna).Value = "-"

            Dim objRango = xlHoja.Range(flLetraExcel(prmColumna) & prmRenglon & ":" & flLetraExcel(prmColumna) & prmRenglon)

            objRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            objRango.Cells.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone

            With objRango.Cells
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlHAlign.xlHAlignJustify
            End With

        End With
    End Sub

    Private Function flLetraExcel(ByVal prmColumna As Integer) As String

        flLetraExcel = ""

        Select Case prmColumna

            Case 1
                flLetraExcel = "A"
            Case 2
                flLetraExcel = "B"
            Case 3
                flLetraExcel = "C"
            Case 4
                flLetraExcel = "D"
            Case 5
                flLetraExcel = "E"
            Case 6
                flLetraExcel = "F"
            Case 7
                flLetraExcel = "G"
            Case 8
                flLetraExcel = "H"
            Case 9
                flLetraExcel = "I"
            Case 10
                flLetraExcel = "J"
            Case 11
                flLetraExcel = "K"
            Case 12
                flLetraExcel = "L"
            Case 13
                flLetraExcel = "M"
            Case 14
                flLetraExcel = "N"
            Case 15
                flLetraExcel = "O"
            Case 16
                flLetraExcel = "P"
            Case 17
                flLetraExcel = "Q"
            Case 18
                flLetraExcel = "R"
            Case 19
                flLetraExcel = "S"
            Case 20
                flLetraExcel = "T"
            Case 21
                flLetraExcel = "U"
        End Select



    End Function





#End Region


End Class
