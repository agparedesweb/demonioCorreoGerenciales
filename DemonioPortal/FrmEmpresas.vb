Imports Sistema
Imports Sistema.Comunes.Comun.ClsTools

Public Class FrmEmpresas

    Dim DAO As DataAccessCls
    Dim vcBaseDeDatos As String
    Private atrSerie As Integer
    Private atrEmpresa As String

    Private Sub BtnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAceptar.Click
        atrSerie = CboEmpresas.SelectedValue
        atrEmpresa = CboEmpresas.Text
        DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub BtnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancelar.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FrmEmpresas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DAO = DataAccessCls.DevuelveInstancia
        DAO.RegresaConsultaSQL("SELECT cNombreEmpresa,nSerie FROM FAC_SERIES WHERE BACTIVO = 1", CboEmpresas)
    End Sub

    Public ReadOnly Property Serie() As Integer
        Get
            Return atrSerie
        End Get
    End Property

    Public ReadOnly Property Empresa() As String
        Get
            Return atrEmpresa
        End Get
    End Property

End Class