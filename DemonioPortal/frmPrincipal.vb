Imports Sistema.Comunes.Comun.ClsTools
Imports Sistema.DataAccessCls
Imports Sistema
Imports Sistema.Comunes.Comun
Imports Sistema.Comunes.Registros.FabricaRegistros
Imports Sistema.Comunes.Registros.EscribanoRegistros
Imports System.Net.Mail
Imports System.Net.Security
Imports Herramientas.Archivos.Archivo
Imports INIFile
Imports System.IO

Public Class frmPrincipal

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
    End Sub

    Private Sub frmPrincipal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DAO As Sistema.DataAccessCls = DataAccessCls.DevuelveInstancia

        'Variables de Entorno / cambiar según sucursal
        Dim conexion As String = "CONEXION_BD_JALISCO"
        'Dim conexion As String = "CONEXION_BD_CULIACAN"
        Dim valorVariable As String = Environment.GetEnvironmentVariable(conexion)

        If valorVariable IsNot Nothing Then
            Console.WriteLine($"El valor de {conexion} es: {valorVariable}")
        Else
            Console.WriteLine($"La variable de entorno {conexion} no está definida.")
        End If

        Dim vcConexion As String = valorVariable
        'Dim SQLC As String
        'SQLC = fgObtenerConexionBD()
        'Dim vcSQL As String

        'vcConexion = "192.168.2.21\SQL2008:PAREDESEMP0708:sa:PaJeAr2012"
        'vcConexion = "192.168.0.21\SQL2008:PAREDESCG:sa:PaVeSe2012"
        'vcConexion = "192.168.2.21\SQL2008:PAREDESEMP0708DEV:sa:PaJeAr2012"
        'vcConexion = "EDWIN-PC\SQLSERVER2008:PAREDES0708:sa:123456"
        'vcConexion = "DELLENRIQUE\SQL2008:PAREDES0708:sa:paredes2012"

        If vcConexion = "" Then End

        DAO = DataAccessCls.DevuelveInstancia(vcConexion)

        If DAO Is Nothing Then
            End
        End If


        inicializa()
        Me.Location = Screen.PrimaryScreen.WorkingArea.Location
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size

        Application.CurrentCulture = New System.Globalization.CultureInfo("es-MX")
        HabilitaBotones(True, True, True, True, True, True)

        'Dim frm2 As New FrmEmpresas

        'If Not frm2.ShowDialog = Windows.Forms.DialogResult.OK Then
        '    End
        'End If
        'vcNombreEmpresa = frm2.Empresa
        'frm2 = Nothing

        Me.Text = "    Sistema de Existencias de Distribuidoras"

    End Sub

    Private Sub ProductosToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductosToolStripMenuItem1.Click
        Dim frm2 As New FrmM1505001
        CargarForma(frm2)
    End Sub

End Class
