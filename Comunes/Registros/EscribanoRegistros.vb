Imports Sistema.Comunes.Comun
Imports Sistema.Comunes.Registros.FabricaRegistros
Imports Sistema.Comunes.Registros.EscribanoRegistros
Imports System.Net.Mail
Imports System.Net.Security
Imports Sistema.Comunes.Comun.ClsTools

Namespace Comunes.Registros

    Public Class EscribanoRegistros
        Private Shared Function fgObtenParametroEMB(ByRef prmClavePar As String, ByVal prmSucursal As String) As Object

            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            fgObtenParametroEMB = ""

            vcSQL = "SELECT CVALORPAR FROM EYE_PARAMETROS WHERE CCLAVEPAR = '" & prmClavePar & "' AND CSUCURSAL = '" & prmSucursal & "'"

            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing Then

                If DT.Rows.Count > 0 Then
                    fgObtenParametroEMB = DT.Rows(0)("CVALORPAR")
                End If
            End If

            DT = Nothing

        End Function


        Private Shared Function fgObtenParametroMail(ByRef prmClavePar As String) As Object

            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            fgObtenParametroMail = ""

            vcSQL = "SELECT CVALORPAR FROM MAIL_PARAMETROS WHERE CCLAVEPAR = '" & prmClavePar & "'"

            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing Then

                If DT.Rows.Count > 0 Then
                    fgObtenParametroMail = DT.Rows(0)("CVALORPAR")
                End If
            End If

            DT = Nothing

        End Function



        Public Shared Function flEnviarMail(ByVal prmCorreo As String, ByVal prmAdjuntos As ArrayList, ByVal prmTítulo As String, ByVal prmCuerpo As String) As Boolean

            Dim vcAdjunto As String
            Dim vcPass As String
            Dim gcServidorSMTP As String = ""
            Dim gnPuertoSMTP As Integer
            Dim gbUsaSSL As Boolean


            Try
                gcServidorSMTP = fgObtenParametroEMB("SERVIDORSMTP", "001")
                gnPuertoSMTP = fgObtenParametroEMB("PUERTOCORREOEMISOR", "001")
                gbUsaSSL = fgObtenParametroEMB("USASSL", "001")
                vcPass = fgObtenParametroMail("PASSNOREPLY")


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
                If prmAdjuntos IsNot Nothing Then
                    For vnPosicion As Integer = 0 To prmAdjuntos.Count - 1
                        vcAdjunto = prmAdjuntos.Item(vnPosicion)
                        oMsg.Attachments.Add(New System.Net.Mail.Attachment(vcAdjunto))
                    Next
                End If


                Servidor.Host = gcServidorSMTP
                Servidor.Port = gnPuertoSMTP
                Servidor.EnableSsl = gbUsaSSL
                Servidor.Credentials = New System.Net.NetworkCredential("noreply@aparedes.com.mx", vcPass)
                Servidor.Send(oMsg)

                oMsg = Nothing

                'EscribeEnBitacora("Se envio el correo")
                Return True

            Catch ex As Exception
                Return flEnviarMail
                'EscribeEnBitacora(ex.Message)
            End Try


        End Function

        Public Shared Function Guardar(ByVal prmFacturas As ClsRegistroFactura) As Boolean

            Dim vParam(49) As Object
            Dim vDs As New DataSet
            Dim vParamDetalle(14) As Object
            Dim vParamTraslados(4) As Object
            Dim vParamRetenciones(4) As Object
            Dim vlSQL As String
            Dim vnRenglon As Integer
            Dim vFolio As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vFolioGeneralFactura As Integer
            Dim vbActualizaFolio As Boolean = False

            If prmFacturas.FACTURA = 0 Then

                vFolioGeneralFactura = fgObtenFolioGeneralFacturas()
                vFolio = fgObtenFolioSerie(prmFacturas.NSERIE)
                vbActualizaFolio = True

            Else
                vFolioGeneralFactura = prmFacturas.nFactura
                vFolio = prmFacturas.FACTURA

            End If


            If vFolio = "" Then Return False


            prmFacturas.nFactura = vFolioGeneralFactura
            prmFacturas.FACTURA = vFolio

            vnRenglon = 0

            Try

                vParam(0) = prmFacturas.nFactura
                vParam(1) = prmFacturas.RFCEmisor
                vParam(2) = prmFacturas.SUCURSAL
                vParam(3) = prmFacturas.SERIE
                vParam(4) = prmFacturas.FACTURA
                vParam(5) = prmFacturas.NSERIE
                vParam(6) = prmFacturas.FECHA
                vParam(7) = prmFacturas.Cliente
                vParam(8) = prmFacturas.CCVE_COND
                vParam(9) = prmFacturas.SUBTOTAL
                vParam(10) = prmFacturas.DESCUENTO
                vParam(11) = prmFacturas.SUBTOTALDD
                vParam(12) = prmFacturas.IMPUESTOS
                vParam(13) = prmFacturas.IVARETENIDO
                vParam(14) = prmFacturas.ISRRETENIDO
                vParam(15) = prmFacturas.IMPORTE
                vParam(16) = prmFacturas.CSTATUS
                vParam(17) = prmFacturas.OBSERVACION
                vParam(18) = prmFacturas.CMONEDA
                vParam(19) = prmFacturas.PARIDAD
                vParam(20) = prmFacturas.CTIPODOCUMENTO
                vParam(21) = prmFacturas.CTIPOPAGO
                vParam(22) = prmFacturas.IMPORTECONLETRA
                vParam(23) = prmFacturas.REGIMEN
                vParam(24) = prmFacturas.LUGAREXPEDICION
                vParam(25) = prmFacturas.CCUENTA

                vParam(26) = prmFacturas.ClienteRFC
                vParam(27) = prmFacturas.ClienteDescripcion
                vParam(28) = prmFacturas.ClienteCalle
                vParam(29) = prmFacturas.ClienteNoInt
                vParam(30) = prmFacturas.ClienteNoExt
                vParam(31) = prmFacturas.ClienteColonia
                vParam(32) = prmFacturas.ClienteLocalidad
                vParam(33) = prmFacturas.Referencia1
                vParam(34) = prmFacturas.ClienteCiudad
                vParam(35) = prmFacturas.ClienteEstado
                vParam(36) = prmFacturas.ClientePais
                vParam(37) = prmFacturas.ClienteCodigoPostal
                vParam(38) = prmFacturas.NumeroAprobacion
                vParam(39) = prmFacturas.AnioAprobacion
                vParam(40) = prmFacturas.MetodoPago
                vParam(41) = prmFacturas.Referencia1
                vParam(42) = prmFacturas.Referencia2
                vParam(43) = prmFacturas.Referencia3


                vParam(44) = ""
                vParam(45) = ""
                vParam(46) = ""
                vParam(47) = ""
                vParam(48) = ""

                vParam(49) = prmFacturas.CERTIFICADO


                vlSQL = "Sp_MTTOFAC_ENC"

                If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParam) Then
                    Return False
                End If

                If Not DAO.EjecutaComandoSQL("DELETE FAC_DETFACTURAS WHERE nFactura = " & prmFacturas.nFactura & " AND nRFCEmisor = " & prmFacturas.RFCEmisor) Then
                    Return False
                End If

                For Each vRow As DataRow In prmFacturas.DTDetalle.Rows

                    vnRenglon += 1

                    vParam(0) = prmFacturas.nFactura
                    vParam(1) = prmFacturas.RFCEmisor
                    vParam(2) = vnRenglon
                    vParam(3) = vRow("nConcepto")
                    vParam(4) = vRow("cDescripcion")
                    vParam(5) = vRow("nCantidad")
                    vParam(6) = vRow("nPrecio")
                    vParam(7) = vRow("nSubtotal")
                    vParam(8) = IIf(vRow("nDescuento") Is DBNull.Value, 0, vRow("nDescuento"))
                    vParam(9) = IIf(vRow("nSubtotal") Is DBNull.Value, 0, vRow("nSubtotal")) - IIf(vRow("nDescuento") Is DBNull.Value, 0, vRow("nDescuento"))
                    vParam(10) = IIf(vRow("nIVA") Is DBNull.Value, 0, vRow("nIVA"))
                    vParam(11) = IIf(vRow("nImpuestos") Is DBNull.Value, 0, vRow("nImpuestos"))
                    vParam(12) = vRow("nImporte")
                    vParam(13) = vRow("cUnidad")
                    vParam(14) = vRow("cNumeroPredial")

                    vlSQL = "Sp_MTTOFAC_DET"
                    If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParam) Then
                        Return False
                    End If

                Next

                If Not DAO.EjecutaComandoSQL("DELETE FAC_IMPUESTOS WHERE nFactura = " & prmFacturas.nFactura & " AND nRFCEmisor = " & prmFacturas.RFCEmisor) Then
                    Return False
                End If


                If Not prmFacturas.DTTraslados Is Nothing AndAlso prmFacturas.DTTraslados.Rows.Count > 0 Then
                    For Each vRow As DataRow In prmFacturas.DTTraslados.Rows

                        vParamTraslados(0) = prmFacturas.nFactura
                        vParamTraslados(1) = prmFacturas.RFCEmisor
                        vParamTraslados(2) = vRow("nTasa")
                        vParamTraslados(3) = vRow("cImpuesto")
                        vParamTraslados(4) = vRow("nImporte")

                        vlSQL = "Sp_MTTOFAC_IMP"
                        If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamTraslados) Then
                            Return False
                        End If

                    Next
                End If

                If Not DAO.EjecutaComandoSQL("DELETE FAC_RETENCIONES WHERE nFactura = " & prmFacturas.nFactura & " AND nRFCEmisor = " & prmFacturas.RFCEmisor) Then
                    Return False
                End If


                If Not prmFacturas.DTRetenciones Is Nothing AndAlso prmFacturas.DTRetenciones.Rows.Count > 0 Then
                    For Each vRow As DataRow In prmFacturas.DTRetenciones.Rows

                        vParamRetenciones(0) = prmFacturas.nFactura
                        vParamRetenciones(1) = prmFacturas.RFCEmisor
                        vParamRetenciones(2) = vRow("nTasa")
                        vParamRetenciones(3) = vRow("cImpuesto")
                        vParamRetenciones(4) = vRow("nImporte")

                        vlSQL = "Sp_MTTOFAC_RET"
                        If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamRetenciones) Then
                            Return False
                        End If

                    Next

                End If


                If vbActualizaFolio Then
                    If Not fgActualizaFolioSerie(prmFacturas.NSERIE) Then
                        Return False
                    End If

                    fgActualizaFacturaGeneral()
                End If




            Catch ex As Exception
                Return False
            End Try

            Return True

        End Function

        Public Shared Function Cancelar(ByVal prmFacturas As ClsRegistroFactura) As Boolean

            Dim vParam(0) As Object
            Dim vlSQL As String
            Dim DS As New DataSet
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            Try

                vParam(0) = prmFacturas.nFactura

                vlSQL = "Sp_MTTOFAC_ENC_CANCELA"

                If Not DAO.RegresaConsultaSQL(vlSQL, DS, vParam) Then
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try

            Return True

        End Function

        Private Shared Function fgObtenFolioNuevoFlete() As Integer
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgObtenFolioNuevoFlete = DAO.RegresaDatoSQL("SELECT COALESCE(MAX(nMovimientoFlete),0)+1 AS NFOLIO FROM SPV_Fletes(NOLOCK)")
            Catch ex As Exception
                fgObtenFolioNuevoFlete = 0
            End Try
        End Function


        Public Shared Function Guardar(ByVal prmFlete As ClsRegistroFlete) As Boolean

            Dim vParam(52) As Object
            Dim vDs As New DataSet
            Dim vlSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vbActualizaFolio As Boolean = False

            If prmFlete.MovimientoFlete = 0 Then
                prmFlete.MovimientoFlete = fgObtenFolioNuevoFlete()
            End If

            Try

                vParam(0) = prmFlete.Factura
                vParam(1) = prmFlete.MovimientoFlete
                vParam(2) = prmFlete.Ruta
                vParam(3) = prmFlete.OrigenRemitente
                vParam(4) = prmFlete.DestinoDestinatario
                vParam(5) = prmFlete.Kilos
                vParam(6) = prmFlete.FechaCarga
                vParam(7) = prmFlete.FechaEntrega
                vParam(8) = prmFlete.DescripcionCarga
                vParam(9) = prmFlete.Operador
                vParam(10) = prmFlete.Unidad
                vParam(11) = prmFlete.PlacasUnidad
                vParam(12) = prmFlete.Remolque1
                vParam(13) = prmFlete.PlacasRemolque1
                vParam(14) = prmFlete.Dolly
                vParam(15) = prmFlete.Remolque2
                vParam(16) = prmFlete.PlacasRemolque2
                vParam(17) = prmFlete.Observaciones
                vParam(18) = prmFlete.IVAFlete
                vParam(19) = prmFlete.RETFlete
                vParam(20) = prmFlete.IVAManiobras
                vParam(21) = prmFlete.RETManiobras
                vParam(22) = prmFlete.IVARepartos
                vParam(23) = prmFlete.RETRepartos
                vParam(24) = prmFlete.IVARecolectas
                vParam(25) = prmFlete.RETRecolectas
                vParam(26) = prmFlete.IVAAutopistas
                vParam(27) = prmFlete.RETAutopistas
                vParam(28) = prmFlete.IVADemoras
                vParam(29) = prmFlete.RETDemoras
                vParam(30) = prmFlete.IVAOtros
                vParam(31) = prmFlete.RETOtros
                vParam(32) = prmFlete.IVASeguros
                vParam(33) = prmFlete.RETSeguros
                vParam(34) = prmFlete.IVARentas
                vParam(35) = prmFlete.RETRentas
                vParam(36) = prmFlete.PorcentajeIVA
                vParam(37) = prmFlete.PorcentajeRET
                vParam(38) = prmFlete.Flete
                vParam(39) = prmFlete.Maniobras
                vParam(40) = prmFlete.Repartos
                vParam(41) = prmFlete.Recolectas
                vParam(42) = prmFlete.Autopistas
                vParam(43) = prmFlete.Demoras
                vParam(44) = prmFlete.Otros
                vParam(45) = prmFlete.Seguros
                vParam(46) = prmFlete.Rentas
                vParam(47) = prmFlete.Subtotal
                vParam(48) = prmFlete.IVA
                vParam(49) = prmFlete.IVARet
                vParam(50) = prmFlete.Total
                vParam(51) = prmFlete.Bultos
                vParam(52) = prmFlete.Clasificacion

                vlSQL = "SPREGISTRASPV_Fletes"

                If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParam) Then
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try

            Return True

        End Function


        Public Shared Function ActualizaDatosTimbrado(ByVal prmFacturas As ClsRegistroFactura) As Boolean

            Dim vParam(3) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DS As New DataSet
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia


            vFolio = prmFacturas.FACTURA
            vFolio = Strings.Right(Strings.StrDup(10, "0") + vFolio, 10)

            Try

                vParam(0) = prmFacturas.nFactura
                vParam(1) = prmFacturas.RFCEmisor
                vParam(2) = prmFacturas.Cadena
                vParam(3) = prmFacturas.Sello

                vlSQL = "Sp_MTTOFAC_ENC_ACTUALIZADATOSTIMBRADO"

                If Not DAO.RegresaConsultaSQL(vlSQL, DS, vParam) Then
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try

            Return True

        End Function


        Public Shared Function ActualizaFechaCancelacion(ByVal prmFacturas As ClsRegistroFactura) As Boolean

            Dim vParam(2) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DS As New DataSet
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia


            vFolio = prmFacturas.FACTURA
            vFolio = Strings.Right(Strings.StrDup(10, "0") + vFolio, 10)

            Try

                vParam(0) = prmFacturas.SUCURSAL
                vParam(1) = prmFacturas.SERIE
                vParam(2) = vFolio

                vlSQL = "Sp_MTTOFAC_ENC_ACTUALIZACANCELACION"

                If Not DAO.RegresaConsultaSQL(vlSQL, DS, vParam) Then
                    Return False
                End If

            Catch ex As Exception
                Return False
            End Try

            Return True

        End Function


        Private Shared Function fgObtenFolioSerie(ByVal prmSerie As Integer) As Integer
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgObtenFolioSerie = DAO.RegresaDatoSQL("SELECT COALESCE(NFOLIOACTUAL,0) AS NFOLIO FROM FAC_SERIES WHERE NSERIE = " & prmSerie)
            Catch ex As Exception
                fgObtenFolioSerie = 0
            End Try
        End Function

        Private Shared Function fgObtenFolioGeneralFacturas() As Integer
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgObtenFolioGeneralFacturas = DAO.RegresaDatoSQL("SELECT COALESCE(NFOLIOCFD,0)+1 AS NFOLIO FROM FAC_FOLIOS(NOLOCK)")
            Catch ex As Exception
                fgObtenFolioGeneralFacturas = 0
            End Try
        End Function


        Public Shared Function fgObtenFolioActualSerie(ByVal prmSerie As Integer) As Integer
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgObtenFolioActualSerie = DAO.RegresaDatoSQL("SELECT NFOLIOACTUAL-1 FROM FAC_SERIES WHERE NSERIE = " & prmSerie)
            Catch ex As Exception
                fgObtenFolioActualSerie = 0
            End Try
        End Function
        Public Shared Function fgObtenerConexionBD(parametroGral As String) As Object
            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT " & parametroGral & " FROM CONFIG_CONEXIONES WHERE flag_activo = 1 " 'obtiene los datos de conexion activa

            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing Then
                If DT.Rows.Count > 0 Then
                    fgObtenerConexionBD = DT.Rows(0)(parametroGral)
                End If
            End If
            DT = Nothing
        End Function
        Public Shared Function fgObtenUltimoFolio(ByVal prmSerie As Integer) As Integer
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgObtenUltimoFolio = DAO.RegresaDatoSQL("SELECT NFOLIOFINAL FROM FAC_SERIES WHERE NSERIE = " & prmSerie)
            Catch ex As Exception
                fgObtenUltimoFolio = 0
            End Try
        End Function

        Private Shared Function fgActualizaFolioSerie(ByVal prmSerie As Integer) As Boolean
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgActualizaFolioSerie = DAO.EjecutaComandoSQL("UPDATE FAC_SERIES SET NFOLIOACTUAL = NFOLIOACTUAL + 1 WHERE NSERIE = " & prmSerie)
            Catch ex As Exception
                fgActualizaFolioSerie = False
            End Try
        End Function

        Private Shared Function fgActualizaFacturaGeneral() As Boolean
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Try
                fgActualizaFacturaGeneral = DAO.EjecutaComandoSQL("UPDATE FAC_FOLIOS SET NFOLIOCFD = NFOLIOCFD + 1")
            Catch ex As Exception
                fgActualizaFacturaGeneral = False
            End Try
        End Function

        Public Shared Function fgObtenParametrosGenerales(ByVal prmRFC As Integer, ByVal prmSERIE As Integer) As DataTable

            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT cClavePar,cValorPar FROM FAC_PARAMETROS"
            vcSQL = vcSQL & vbCrLf & "WHERE NRFCEMISOR = " & prmRFC & " AND NSERIE = " & prmSERIE

            DAO.RegresaConsultaSQL(vcSQL, DT)
            Return DT

        End Function

        Public Shared Function fgGrabaParametrosGenerales(ByVal prmDataTable As DataTable, ByVal prmRFC As Integer, ByVal prmSERIE As Integer) As Boolean

            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vcSQL As String
            Try


                If Not prmDataTable Is Nothing AndAlso prmDataTable.Rows.Count > 0 Then

                    For Each vRow As DataRow In prmDataTable.Rows

                        vcSQL = "UPDATE FAC_PARAMETROS SET CVALORPAR = '" & vRow("CVALORPAR") & "' WHERE CCLAVEPAR = '" & vRow("CCLAVEPAR") & "'"
                        vcSQL = vcSQL & vbCrLf & "AND NRFCEMISOR = " & prmRFC & " AND NSERIE = " & prmSERIE

                        DAO.EjecutaComandoSQL(vcSQL)
                    Next

                End If

                Return True
            Catch ex As Exception
                Return False
            End Try


        End Function

        Public Shared Function fgObtenReporteMensualSAT(ByVal prmRFC As Integer, ByVal prmEjercicio As Integer, ByVal prmPeriodo As Integer) As DataTable

            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT REPLACE(REPLACE(FE.CRFCCLIENTE,' ',''),'-','') AS CRFCCLIENTE,FE.CSERIE AS CSERIEEMISOR,FE.NFOLIO AS NFACTURA,FO.NANIOAUTORIZACION AS NANIO,"
            vcSQL = vcSQL & vbCrLf & "FO.NNUMEROAUTORIZACION,FE.NIMPORTE AS NTOTAL,FE.NIMPUESTOS AS NIVA,FE.CSTATUS,"
            vcSQL = vcSQL & vbCrLf & "SUBSTRING(CONVERT(VARCHAR(10),FE.DFECHA,103),1,6) + CONVERT(CHAR(4),YEAR(FE.DFECHA)) + ' ' + CONVERT(VARCHAR(8),FE.DFECHA,108) AS FechaFormato,"
            vcSQL = vcSQL & vbCrLf & "CASE WHEN FE.CTIPODOCUMENTO = 'F' THEN 'I' ELSE 'E' END AS CEFECTO"
            vcSQL = vcSQL & vbCrLf & "FROM FAC_ENCFACTURAS FE(NOLOCK) "
            vcSQL = vcSQL & vbCrLf & "INNER JOIN FAC_EMISORES EM(Nolock) ON (FE.NRFCEMISOR = EM.NRFCEMISOR) "
            vcSQL = vcSQL & vbCrLf & "INNER JOIN FAC_SERIES FO(nolock) ON (FE.NRFCEMISOR = FO.NRFCEMISOR AND FE.NSERIE = FO.NSERIE) "
            vcSQL = vcSQL & vbCrLf & "WHERE FE.NRFCEMISOR = " & prmRFC & " AND YEAR(FE.DFECHA) = " & prmEjercicio & " AND MONTH(FE.DFECHA) = " & prmPeriodo
            vcSQL = vcSQL & vbCrLf & ""
            vcSQL = vcSQL & vbCrLf & "UNION ALL"
            vcSQL = vcSQL & vbCrLf & ""
            vcSQL = vcSQL & vbCrLf & "SELECT REPLACE(REPLACE(FE.CRFCCLIENTE,' ',''),'-','') AS CRFCCLIENTE,FE.CSERIE AS CSERIEEMISOR,FE.NFOLIO AS NFACTURA,FO.NANIOAUTORIZACION AS NANIO,"
            vcSQL = vcSQL & vbCrLf & "FO.NNUMEROAUTORIZACION,FE.NIMPORTE AS NTOTAL,FE.NIMPUESTOS AS NIVA,FE.CSTATUS,"
            vcSQL = vcSQL & vbCrLf & "SUBSTRING(CONVERT(VARCHAR(10),FE.DFECHA,103),1,6) + CONVERT(CHAR(4),YEAR(FE.DFECHA)) + ' ' + CONVERT(VARCHAR(8),FE.DFECHA,108) AS FechaFormato,"
            vcSQL = vcSQL & vbCrLf & "CASE WHEN FE.CTIPODOCUMENTO = 'F' THEN 'I' ELSE 'E' END AS CEFECTO"
            vcSQL = vcSQL & vbCrLf & "FROM FAC_ENCFACTURAS FE(NOLOCK) "
            vcSQL = vcSQL & vbCrLf & "INNER JOIN FAC_EMISORES EM(Nolock) ON (FE.NRFCEMISOR = EM.NRFCEMISOR) "
            vcSQL = vcSQL & vbCrLf & "INNER JOIN FAC_SERIES FO(nolock) ON (FE.NRFCEMISOR = FO.NRFCEMISOR AND FE.NSERIE = FO.NSERIE) "
            vcSQL = vcSQL & vbCrLf & "WHERE FE.NRFCEMISOR = " & prmRFC & " AND FE.CSTATUS = 'C'"
            vcSQL = vcSQL & vbCrLf & "AND YEAR(FE.DFECHACAN) = " & prmEjercicio & " AND MONTH(FE.DFECHACAN) = " & prmPeriodo
            vcSQL = vcSQL & vbCrLf & ""
            vcSQL = vcSQL & vbCrLf & "ORDER BY FE.NFOLIO"

            DAO.RegresaConsultaSQL(vcSQL, DT)
            Return DT

        End Function

        Public Shared Function fgGrabainventarios(ByVal prmDT As DataTable, ByVal prmTemporada As String, ByVal prmFecha As Date) As String

            Dim vDs As New DataSet
            Dim vParamDetalle(34) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vFolioGeneralFactura As Integer
            Dim vnRenglon As Integer = 0

            Try

                For Each vRow As DataRow In prmDT.Rows

                    If vnRenglon = 0 Then
                        DAO.EjecutaComandoSQL("DELETE EYE_INVENTARIOSHM WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                    End If

                    vnRenglon += 1
                    vParamDetalle(0) = prmTemporada
                    vParamDetalle(1) = prmFecha

                    vParamDetalle(2) = vRow("PackerID")
                    vParamDetalle(3) = vRow("GrowerID")
                    vParamDetalle(4) = vRow("CommodityID")
                    vParamDetalle(5) = vRow("VarietyID")
                    vParamDetalle(6) = vRow("StyleID")
                    vParamDetalle(7) = vRow("SizeID")
                    vParamDetalle(8) = vRow("LabelID")
                    vParamDetalle(9) = vRow("LastNight")
                    vParamDetalle(10) = vRow("PkgsRec")
                    vParamDetalle(11) = vRow("PkgsShp")
                    vParamDetalle(12) = vRow("PkgsRepack")
                    vParamDetalle(13) = vRow("YTDPkgsRec")
                    vParamDetalle(14) = vRow("YTDPkgsShp")
                    vParamDetalle(15) = vRow("CurrentFloor")
                    vParamDetalle(16) = vRow("PkgsPendShp")
                    vParamDetalle(17) = vRow("ProductDesc")
                    vParamDetalle(18) = vRow("CommDesc")
                    vParamDetalle(19) = vRow("VarDesc")
                    vParamDetalle(20) = vRow("StyleDesc")
                    vParamDetalle(21) = vRow("LabelDesc")
                    vParamDetalle(22) = vRow("PackerName")
                    vParamDetalle(23) = vRow("GrowerName")
                    vParamDetalle(24) = vRow("GrowerPackerDesc")
                    vParamDetalle(25) = vRow("PkOrderKey")
                    vParamDetalle(26) = vRow("GrOrderKey")
                    vParamDetalle(27) = vRow("CoOrderKey")
                    vParamDetalle(28) = vRow("VaOrderKey")
                    vParamDetalle(29) = vRow("LbOrderKey")
                    vParamDetalle(30) = vRow("SyOrderKey")
                    vParamDetalle(31) = IIf(vRow("SzOrderKey") = "", 0, vRow("SzOrderKey"))
                    vParamDetalle(32) = vRow("GroCommCode")
                    vParamDetalle(33) = vRow("GroLabelCode")
                    vParamDetalle(34) = vRow("ExternalLabelDesc")

                    vlSQL = "SPINSERTAINVENTARIOSHM"
                    If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamDetalle) Then
                        Return "No inserto la informacion"
                    End If

                Next

            Catch ex As Exception
                Return ex.Message
            End Try

            Return ""

        End Function

        Public Shared Function fgGrabainventariosMarengo(ByVal prmDT As DataTable, ByVal prmTemporada As String, ByVal prmFecha As Date) As String
            Console.WriteLine(vbCrLf & "## fgGrabainventariosMarengo() ## " & vbCrLf)
            Dim vDs As New DataSet
            Dim vParamDetalle(29) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vFolioGeneralFactura As Integer
            Dim vnRenglon As Integer = 0

            Try
                ' Recorrer todas las filas de prmDT
                Console.WriteLine("### Inicia Insertar info a BD ### SP_Inserta_ApiTransactionsV3")
                For Each vRow As DataRow In prmDT.Rows

                    If vnRenglon = 0 Then
                        DAO.EjecutaComandoSQL("DELETE EYE_TRANSACTIONS_V3 WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                        'DAO.EjecutaComandoSQL("DELETE EYE_INVENTARIOSMARENGO WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                    End If

                    vnRenglon += 1
                    vParamDetalle(0) = prmTemporada
                    vParamDetalle(1) = prmFecha
                    vParamDetalle(2) = vRow("branch")
                    vParamDetalle(3) = vRow("reference")
                    vParamDetalle(4) = vRow("product")
                    vParamDetalle(5) = vRow("lotNo")
                    vParamDetalle(6) = vRow("palletIds")
                    vParamDetalle(7) = vRow("tSource")
                    vParamDetalle(8) = vRow("vendorProductCode")
                    vParamDetalle(9) = vRow("commodity")
                    vParamDetalle(10) = vRow("var")
                    vParamDetalle(11) = vRow("variety")
                    vParamDetalle(12) = vRow("pack")
                    vParamDetalle(13) = vRow("pack2")
                    vParamDetalle(14) = vRow("packaging")
                    vParamDetalle(15) = vRow("contCode")
                    vParamDetalle(16) = vRow("container")
                    vParamDetalle(17) = vRow("sizeCode")
                    vParamDetalle(18) = vRow("size")
                    vParamDetalle(19) = vRow("gradeCode")
                    vParamDetalle(20) = vRow("grade")
                    vParamDetalle(21) = vRow("floor")
                    vParamDetalle(22) = vRow("received")
                    vParamDetalle(23) = vRow("unpack")
                    vParamDetalle(24) = vRow("repack")
                    vParamDetalle(25) = vRow("shipped")
                    vParamDetalle(26) = vRow("shippedret")
                    vParamDetalle(27) = vRow("unreceived")
                    vParamDetalle(28) = vRow("endOfDay")

                    vlSQL = "SP_Inserta_ApiTransactionsV3"
                    If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamDetalle) Then
                        Return "No insertó la informacion ## SP_Inserta_ApiTransactionsV3"
                    End If
                Next
                Console.WriteLine("### FIN. Se insertó info en BD ###")
            Catch ex As Exception
                Return ex.Message
            End Try
            Return ""

        End Function

        Public Shared Function fgGrabaVentasDiariasMarengo(ByVal prmDT As DataTable, ByVal prmTemporada As String, ByVal prmFecha As Date) As String
            Console.WriteLine("### fgGrabaVentasDiariasMarengo() ### API GrowerNetSales")
            Dim vDs As New DataSet
            Dim vParamDetalle(20) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vFolioGeneralFactura As Integer
            Dim vnRenglon As Integer = 0
            Dim query As String

            Try
                Console.WriteLine("### Inicio Insertar info a BD ###")
                For Each vRow As DataRow In prmDT.Rows
                    If vnRenglon = 0 Then
                        DAO.EjecutaComandoSQL("DELETE EYE_GROWER_NETSALES WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) >= CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                        'DAO.EjecutaComandoSQL("DELETE EYE_VENTASDIARIASMARENGO WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) >= CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                    End If

                    vnRenglon += 1
                    vParamDetalle(0) = prmTemporada
                    vParamDetalle(1) = prmFecha

                    vParamDetalle(2) = vRow("vendor")
                    vParamDetalle(3) = vRow("growerDealNumber")
                    vParamDetalle(4) = vRow("postDate")
                    vParamDetalle(5) = vRow("reference")
                    vParamDetalle(6) = vRow("product")
                    vParamDetalle(7) = vRow("lotNo")
                    vParamDetalle(8) = vRow("palletID")
                    vParamDetalle(9) = vRow("commodity")
                    vParamDetalle(10) = vRow("variety")
                    vParamDetalle(11) = vRow("packaging")
                    vParamDetalle(12) = vRow("container")
                    vParamDetalle(13) = vRow("size")
                    vParamDetalle(14) = vRow("grade")
                    vParamDetalle(15) = vRow("qty")
                    vParamDetalle(16) = vRow("grossSales")
                    vParamDetalle(17) = vRow("netSales")
                    vParamDetalle(18) = vRow("totalAdjustments")

                    'vlSQL = "SPINSERTAVENTASDIARIASMARENGO"                    
                    vlSQL = "SP_Inserta_ApiGrowerNetSales" 'inserta datos de api

                    If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamDetalle) Then
                        Return "No insertó la información en BD"
                    End If
                Next
                Console.WriteLine("### FIN. Se insertó info a BD ### SP_Inserta_ApiGrowerNetSales")
            Catch ex As Exception
                Return ex.Message
            End Try


            Return ""

        End Function


        Private Shared Function fgObtenCultivoMarengo(ByVal prmCommodityName As String) As String
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT CCVE_CULTIVO FROM EYE_CULTIVOSMARENGO"
            vcSQL = vcSQL & vbCrLf & "WHERE CommodityName = '" & prmCommodityName.Trim & "'"

            Return DAO.RegresaDatoSQL(vcSQL)

        End Function

        Private Shared Function fgObtenEtiquetaMarengo(ByVal prmLabel As String) As String
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT CCVE_ETIQUETA FROM EYE_ETIQUETASMARENGO"
            vcSQL = vcSQL & vbCrLf & "WHERE Label = '" & prmLabel.Trim & "'"

            Return DAO.RegresaDatoSQL(vcSQL)
        End Function

        Private Shared Function fgObtenTamañoMarengo(ByVal prmCommodityName As String, ByVal prmSize As String) As String
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT CCVE_TAMAÑO FROM EYE_TAMAÑOSMARENGO"
            vcSQL = vcSQL & vbCrLf & "WHERE CommodityName = '" & prmCommodityName.Trim.ToUpper & "'"
            vcSQL = vcSQL & vbCrLf & "AND Size = '" & prmSize.Trim & "'"

            Return DAO.RegresaDatoSQL(vcSQL)
        End Function

        Private Shared Function fgObtenEnvaseMarengo(ByVal prmPackStyle As String) As String
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT CCVE_ENVASE FROM EYE_ENVASESMARENGO"
            vcSQL = vcSQL & vbCrLf & "WHERE PackStyle = '" & prmPackStyle.Trim & "'"

            Return DAO.RegresaDatoSQL(vcSQL)
        End Function

        Public Shared Function fgGrabaVentasDiarias(ByVal prmDT As DataTable, ByVal prmTemporada As String, ByVal prmFecha As Date) As Boolean
            Dim vDs As New DataSet
            Dim vParamDetalle(15) As Object
            Dim vlSQL As String
            Dim vFolio As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia
            Dim vFolioGeneralFactura As Integer
            Dim vnRenglon As Integer = 0
            Dim vcEtiqueta As String = ""

            Try
                For Each vRow As DataRow In prmDT.Rows

                    If vnRenglon = 0 Then
                        DAO.EjecutaComandoSQL("DELETE EYE_VENTASDIARIASHM WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) = CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112)")
                        'DAO.EjecutaComandoSQL("DELETE EYE_VENTASDIARIASHM WHERE CCVE_TEMPORADA = '" & prmTemporada & "' AND CONVERT(VARCHAR(20),DFECHA,112) = '" & Format(prmFecha, "YYYYMMDD") & "'")
                    End If

                    vcEtiqueta = IIf(vRow("GroLabelDesc") Is DBNull.Value, "", vRow("GroLabelDesc"))

                    If vcEtiqueta <> "PARIS-07" Then
                        vnRenglon += 1
                        vParamDetalle(0) = prmTemporada
                        vParamDetalle(1) = prmFecha

                        vParamDetalle(2) = vRow("Truck")
                        vParamDetalle(3) = vRow("ShpDate")
                        vParamDetalle(4) = vRow("ArvDate")
                        vParamDetalle(5) = vRow("GroReference")
                        vParamDetalle(6) = vRow("DistLotNum")
                        vParamDetalle(7) = vRow("GroProductCode")
                        vParamDetalle(8) = IIf(vRow("GroProductDesc") Is DBNull.Value, "", vRow("GroProductDesc"))
                        vParamDetalle(9) = vRow("GroLabelCode")
                        vParamDetalle(10) = IIf(vRow("GroLabelDesc") Is DBNull.Value, "", vRow("GroLabelDesc"))
                        vParamDetalle(11) = vRow("Prefix")
                        vParamDetalle(12) = vRow("PalletTNum")
                        vParamDetalle(13) = vRow("Pkgs")
                        vParamDetalle(14) = vRow("UPrice")
                        vParamDetalle(15) = vRow("Amount")

                        Console.WriteLine("### Insertando información a BD ### SPINSERTAVENTASDIARIASHM")
                        vlSQL = "SPINSERTAVENTASDIARIASHM"
                        If Not DAO.RegresaConsultaSQL(vlSQL, vDs, vParamDetalle) Then
                            Return False
                        End If

                    End If
                Next

            Catch ex As Exception
                Console.WriteLine(ex)
                Return False
            End Try

            Return True

        End Function


        Public Shared Function fgObtenParametrosSemanaEnBaseFecha(ByVal prmTemporada As String, ByVal prmFecha As Date) As String

            Dim DT As New DataTable
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

            vcSQL = "SELECT CSEMANA FROM CTL_SEMANAS"
            vcSQL = vcSQL & vbCrLf & "WHERE CCVE_TEMPORADA = '" & prmTemporada & "'"
            vcSQL = vcSQL & vbCrLf & "AND CCVE_NOMINA = '01'"
            vcSQL = vcSQL & vbCrLf & "AND CONVERT(VARCHAR(20),'" & Format(prmFecha, "yyyyMMdd") & "',112) BETWEEN "
            vcSQL = vcSQL & vbCrLf & "CONVERT(VARCHAR(20),DFEC_INI,112) AND CONVERT(VARCHAR(20),DFEC_FIN,112)"

            Return DAO.RegresaDatoSQL(vcSQL)

        End Function



    End Class


End Namespace

