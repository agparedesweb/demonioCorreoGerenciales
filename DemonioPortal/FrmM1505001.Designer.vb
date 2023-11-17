<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmM1505001
    Inherits Sistema.AccessForm

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmM1505001))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'lblToolTip
        '
        Me.lblToolTip.Size = New System.Drawing.Size(19, 25)
        '
        'Timer1
        '
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Items.AddRange(New Object() {"" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "HORA>=7 ", "1)  T EYE_ENVIOEXISTENCIASPISO              " & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN EXISTENCIANACIONAL_dd-MM-" & _
                        "yyyy  ", "2)  T EYE_ENVIOPORCCALIDAD                  " & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN PORCENTAJECALIDADES_dd-MM" & _
                        "-yyyy ", "3)  T EYE_ENVIOEXISTENCIASDISTRIBUIDORAS    " & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN EXISTENCIAS DISTRIBUIDORAS" & _
                        " AL dd-MMM-yy.pdf", "4)  T EYE_ENVIOROTACIONPRODUCTOS" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN ROTACIONDISTRIBUIDORAS_dd-MM-yyyy.xls" & _
                        "x", "5)  T EYE_ENVIOPREENFRIADOS" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN VENTAS DIARIAS HM AL dd-MMM-yy.pdf", "6)  T EYE_ENVIOVENTASDIARIASDISTRIBUIDORAS  " & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN VENTAS DIARIAS HM AL dd-MM" & _
                        "M-yy.pdf", "7)  T INF_EXISTVENTADIST" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN EXPORTACION DISPONIBLE PARA VENTA "" & UCase(" & _
                        "Format(vdFecha, ""dd-MMM-yy"")) & "" TEMP "" & vcTemporada & "".pdf", "8)  T NOM_ENVIOAVANCELABORES" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: AVANCE DE LABORES DEL "" & UCase(Format(vdFech" & _
                        "a, ""dd-MMM-yy"")) & "".pdf", "9)  T NOM_ENVIORENTATRACTORES" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: INFORME DE RISTAS DEL "" & UCase(Format" & _
                        "(vdFecha, ""dd-MMM-yy"")) & "".pdf", resources.GetString("ListBox1.Items"), "11) T EYE_ENVIOANTIGUEDADPISO" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: REPORTE DE ANTIGUEDAD EN PISO "" & UCase(Form" & _
                        "at(vdFecha, ""dd-MMM-yy"")) & "" TEMP "" & vcTemporada & "".pdf", "12) T NOM_ELEGIBLESTARJETA" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: ELIGIBLESTARJETA"" & Format(vdFecha, ""dd-MM-yyy" & _
                        "y"") & "".xlsx", "" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9), "" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "HORA>=8 ", "13) T EYE_PORCENTAJESEMPAQUE " & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: CLN PORCENTAJE DE EMPAQUE AL DIA "" & Format(" & _
                        "vdFechaActual, ""dd-MM-yyyy"") & "".xlsx", "" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "HORA>=9", "14) T TSPV_ENVIOASISTENCIADIARIA" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: TSPV ASISTENCIA DIARIA DEL "" & UCase(Form" & _
                        "at(vdFecha, ""dd-MMM-yy"")) & "".pdf", "" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "HORA>=13", "15) T NOM_ENVIOCOMPLEMENTONOMINA            " & Global.Microsoft.VisualBasic.ChrW(9) & "FILE: COMPLEMENTO DE NOMINA AL ""dd-M" & _
                        "MM-yy"".pdf"})
        Me.ListBox1.Location = New System.Drawing.Point(32, 17)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(1276, 354)
        Me.ListBox1.TabIndex = 2
        '
        'FrmM1505001
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1370, 460)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "FrmM1505001"
        Me.ReferenciaRapidaVisible = True
        Me.Text = "   Existencias Distribuidoras"
        Me.Controls.SetChildIndex(Me.ListBox1, 0)
        Me.Controls.SetChildIndex(Me.lblToolTip, 0)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox

End Class
