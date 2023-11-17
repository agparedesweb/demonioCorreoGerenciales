Imports Sistema.Comunes.Comun

Namespace Comunes.Registros

    Public Class FabricaRegistros
        Inherits Comunes.Comun.Fabrica

        Public Shared Function ObtenRegistroFactura(ByVal prmFactura As Integer, ByVal prmRFCEmisor As Integer) As ClsRegistroFactura

            Dim DT As New DataTable
            Dim DTDetalle As New DataTable
            Dim DTRetenciones As New DataTable
            Dim DTTraslados As New DataTable
            Dim ret As New ClsRegistroFactura()
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia


            vcSQL = "SELECT * FROM FAC_ENCFACTURAS WHERE NFOLIO = " & prmFactura & " AND NRFCEMISOR = " & prmRFCEmisor

            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

                Dim prmrow As DataRow
                prmrow = DT.Rows(0)

                ret.nFactura = prmrow("nFactura")
                ret.Cliente = prmrow("NCLIENTE")
                ret.NSERIE = prmrow("NSERIE")
                ret.SUBTOTAL = prmrow("NSUBTOTAL")
                ret.DESCUENTO = prmrow("NDESCUENTO")
                ret.IMPUESTOS = prmrow("NIMPUESTOS")
                ret.IVARETENIDO = prmrow("NIVARETENIDO")
                ret.ISRRETENIDO = prmrow("NISRRETENIDO")
                ret.IMPORTE = prmrow("NIMPORTE")
                ret.OBSERVACION = prmrow("COBSERVACION")
                ret.PARIDAD = prmrow("NPARIDAD")
                ret.CSTATUS = prmrow("CSTATUS")
                ret.FECHA = prmrow("DFECHA")
                ret.CCVE_COND = prmrow("CCONDICIONPAGO")
                ret.CMONEDA = prmrow("CMONEDA")
                ret.SUBTOTALDD = prmrow("NSUBTOTALDD")
                ret.FECHACAN = IIf(prmrow("DFECHACAN") Is DBNull.Value, DAO.RegresaFechaDelSistema, prmrow("DFECHACAN"))
                ret.CERTIFICADO = IIf(prmrow("CCERTIFICADO") Is DBNull.Value, "", prmrow("CCERTIFICADO"))
                ret.Sello = IIf(prmrow("CSELLO") Is DBNull.Value, "", prmrow("CSELLO"))
                ret.Cadena = IIf(prmrow("CCADENA") Is DBNull.Value, "", prmrow("CCADENA"))
                ret.IMPORTECONLETRA = prmrow("CIMPORTECONLETRA")
                ret.UID = IIf(prmrow("UUID") Is DBNull.Value, "", prmrow("UUID"))
                ret.FECHATIMBRADO = IIf(prmrow("DFECHATIMBRADO") Is DBNull.Value, DAO.RegresaFechaDelSistema, prmrow("DFECHATIMBRADO"))
                ret.SELLOSAT = IIf(prmrow("CSELLOSAT") Is DBNull.Value, "", prmrow("CSELLOSAT"))
                ret.REGIMEN = prmrow("CREGIMEN")
                ret.LUGAREXPEDICION = prmrow("CLUGAREXPEDICION")
                ret.CTIPOPAGO = prmrow("CTIPOPAGO")
                ret.CCUENTA = IIf(prmrow("CCUENTA") Is DBNull.Value, "", prmrow("CCUENTA"))
                ret.SUCURSAL = IIf(prmrow("NEMISORSUCURSAL") Is DBNull.Value, 0, prmrow("NEMISORSUCURSAL"))
                ret.SERIE = IIf(prmrow("CSERIE") Is DBNull.Value, "", prmrow("CSERIE"))
                ret.FACTURA = IIf(prmrow("NFOLIO") Is DBNull.Value, 0, prmrow("NFOLIO"))
                ret.Referencia1 = IIf(prmrow("CREFERENCIA1") Is DBNull.Value, "", prmrow("CREFERENCIA1"))
                ret.Referencia2 = IIf(prmrow("CREFERENCIA2") Is DBNull.Value, "", prmrow("CREFERENCIA2"))
                ret.Referencia3 = IIf(prmrow("CREFERENCIA3") Is DBNull.Value, "", prmrow("CREFERENCIA3"))
                ret.MetodoPago = prmrow("CMETODOPAGO")

                vcSQL = "SELECT COALESCE(DET.nConcepto,0) AS nConcepto,COALESCE(DET.CCONCEPTO,'') AS cDescripcion,DET.NCANTIDAD AS nCantidad," & vbCrLf
                vcSQL = vcSQL & "DET.NPRECIO AS nPrecio,DET.NDESCUENTO AS nDescuento,DET.NIMPUESTOS AS nImpuestos,DET.NIMPORTE AS nImporte,DET.NSUBTOTAL AS nSubtotal," & vbCrLf
                vcSQL = vcSQL & "DET.NIVA AS nIVA,DET.CUNIDAD AS cUnidad,IMP.CDESCRIP AS cDescImpuesto,COALESCE(DET.cNumeroPredial,'') AS cNumeroPredial" & vbCrLf
                vcSQL = vcSQL & "FROM FAC_DETFACTURAS DET(NOLOCK)" & vbCrLf
                vcSQL = vcSQL & "JOIN IMPUESTOS IMP(NOLOCK) ON IMP.NIMPUESTO = DET.NIMPUESTOS" & vbCrLf
                vcSQL = vcSQL & "WHERE DET.nFactura = " & ret.nFactura & " AND DET.NRFCEMISOR = " & prmRFCEmisor & vbCrLf
                vcSQL = vcSQL & "ORDER BY DET.NRENGLON" & vbCrLf

                DAO.RegresaConsultaSQL(vcSQL, DTDetalle)

                If Not DTDetalle Is Nothing Then
                    ClsTools.copiaRows(DTDetalle.Select(""), ret.DTDetalle, DTDetalle.Columns)
                End If


                vcSQL = "SELECT nTasa,cImpuesto,nImporte" & vbCrLf
                vcSQL = vcSQL & "FROM FAC_IMPUESTOS(NOLOCK)" & vbCrLf
                vcSQL = vcSQL & "WHERE nFactura = " & ret.nFactura & " AND NRFCEMISOR = " & prmRFCEmisor & vbCrLf

                DAO.RegresaConsultaSQL(vcSQL, DTTraslados)

                If Not DTTraslados Is Nothing Then
                    ClsTools.copiaRows(DTTraslados.Select(""), ret.DTTraslados, DTTraslados.Columns)
                End If

                vcSQL = "SELECT nTasa,cImpuesto,nImporte" & vbCrLf
                vcSQL = vcSQL & "FROM FAC_RETENCIONES(NOLOCK)" & vbCrLf
                vcSQL = vcSQL & "WHERE nFactura = " & ret.nFactura & " AND NRFCEMISOR = " & prmRFCEmisor & vbCrLf

                DAO.RegresaConsultaSQL(vcSQL, DTRetenciones)

                If Not DTRetenciones Is Nothing Then
                    ClsTools.copiaRows(DTRetenciones.Select(""), ret.DTRetenciones, DTRetenciones.Columns)
                End If


            End If

            Return ret
        End Function

        Public Shared Function ObtenRegistroFlete(ByVal prmFactura As Long) As ClsRegistroFlete

            Dim DT As New DataTable
            Dim ret As New ClsRegistroFlete()
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia


            vcSQL = "SELECT * FROM SPV_Fletes WHERE nFactura = " & prmFactura

            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then

                Dim prmrow As DataRow
                prmrow = DT.Rows(0)

                ret.Factura = prmrow("nFactura")
                ret.MovimientoFlete = prmrow("nMovimientoFlete")
                ret.Ruta = prmrow("cRuta")
                ret.OrigenRemitente = prmrow("cOrigenRemitente")
                ret.DestinoDestinatario = prmrow("cDestinoDestinatario")
                ret.Kilos = prmrow("nKilos")
                ret.FechaCarga = prmrow("dFechaCarga")
                ret.FechaEntrega = prmrow("dFechaEntrega")
                ret.DescripcionCarga = prmrow("cDescripcionCarga")
                ret.Operador = prmrow("cOperador")
                ret.Unidad = prmrow("cUnidad")
                ret.PlacasUnidad = prmrow("cPlacasUnidad")
                ret.Remolque1 = prmrow("cRemolque1")
                ret.PlacasRemolque1 = prmrow("cPlacasRemolque1")
                ret.Dolly = prmrow("cDolly")
                ret.Remolque2 = prmrow("cRemolque2")
                ret.PlacasRemolque2 = prmrow("cPlacasRemolque2")
                ret.Observaciones = prmrow("cObservaciones")
                ret.IVAFlete = prmrow("bIVAFlete")
                ret.RETFlete = prmrow("bRETFlete")
                ret.IVAManiobras = prmrow("bIVAManiobras")
                ret.RETManiobras = prmrow("bRETManiobras")
                ret.IVARepartos = prmrow("bIVARepartos")
                ret.RETRepartos = prmrow("bRETRepartos")
                ret.IVARecolectas = prmrow("bIVARecolectas")
                ret.RETRecolectas = prmrow("bRETRecolectas")
                ret.IVAAutopistas = prmrow("bIVAAutopistas")
                ret.RETAutopistas = prmrow("bRETAutopistas")
                ret.IVADemoras = prmrow("bIVADemoras")
                ret.RETDemoras = prmrow("bRETDemoras")
                ret.IVAOtros = prmrow("bIVAOtros")
                ret.RETOtros = prmrow("bRETOtros")
                ret.IVASeguros = prmrow("bIVASeguros")
                ret.RETSeguros = prmrow("bRETSeguros")
                ret.IVARentas = prmrow("bIVARentas")
                ret.RETRentas = prmrow("bRETRentas")
                ret.PorcentajeIVA = prmrow("nPorcentajeIVA")
                ret.PorcentajeRET = prmrow("nPorcentajeRET")
                ret.Flete = prmrow("nFlete")
                ret.Maniobras = prmrow("nManiobras")
                ret.Repartos = prmrow("nRepartos")
                ret.Recolectas = prmrow("nRecolectas")
                ret.Autopistas = prmrow("nAutopistas")
                ret.Demoras = prmrow("nDemoras")
                ret.Otros = prmrow("nOtros")
                ret.Seguros = prmrow("nSeguros")
                ret.Rentas = prmrow("nRentas")
                ret.Subtotal = prmrow("nSubtotal")
                ret.IVA = prmrow("nIVA")
                ret.IVARet = prmrow("nIVARet")
                ret.Total = prmrow("nTotal")
                ret.Bultos = prmrow("nBultos")
                ret.Clasificacion = prmrow("cClasificacion")
            End If

            Return ret
        End Function

        Public Shared Function fgObtenerInventario(ByVal prmTemporada As String, ByVal prmCommoditie As String, ByVal prmCarton As String, ByVal prmEtiqueta As String, ByVal prmTamaño As String, ByVal prmTipoCarton As String) As Double
            Dim DT As New DataTable
            Dim DTDetalle As New DataTable
            Dim DTRetenciones As New DataTable
            Dim DTTraslados As New DataTable
            Dim ret As New ClsRegistroFactura()
            Dim vcSQL As String
            Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia


            vcSQL = "SELECT TOP 1 Floor FROM 
"
            vcSQL = vcSQL & vbCrLf & "WHERE CCVE_TEMPORADA = '" & prmTemporada & "'"
            vcSQL = vcSQL & vbCrLf & "AND CommodityName = '" & prmCommoditie & "'"
            vcSQL = vcSQL & vbCrLf & "AND PackStyle = '" & prmCarton & "'"
            vcSQL = vcSQL & vbCrLf & "AND Label = '" & prmEtiqueta & "'"
            vcSQL = vcSQL & vbCrLf & "AND Size = '" & prmTamaño & "'"
            vcSQL = vcSQL & vbCrLf & "AND UoM = '" & prmTipoCarton & "'"
            vcSQL = vcSQL & vbCrLf & "ORDER BY DFECHA DESC"


            DAO.RegresaConsultaSQL(vcSQL, DT)

            If Not DT Is Nothing AndAlso DT.Rows.Count > 0 Then
                Return DT(0)(0)
            End If

            Return 0
        End Function



    End Class


End Namespace
