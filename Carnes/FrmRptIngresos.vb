Imports Microsoft.Reporting.WinForms

Public Class FrmRptIngresos

    Private Sub BtnSALIR_Click(sender As Object, e As EventArgs) Handles BtnSALIR.Click
        Me.Close()
    End Sub

    Private Sub BtnLimpiar_Click(sender As Object, e As EventArgs) Handles BtnLimpiar.Click
        DtpDesde.Value = CDate("01/" + Format(Month(Now), "00") + "/" + Format(Year(Now), "0000"))
        DtpDesde.Value = Now
        CMBEstado.Items.Clear()
        CMBEstado.Items.Add("")
        CMBEstado.Items.Add("ABIERTO")
        CMBEstado.Items.Add("CERRADO")
        CMBEstado.Items.Add("ANULADO")
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Dim FRM As New FrmReportes
        Dim DT As New DataTable
        Dim RD As New ReportDataSource()
        Dim ClsProc As New ClsOrdendeIngreso

        Try
            FRM.MdiParent = Me.MdiParent
            FRM.Text = "REPORTE DE INGRESOS"
            FRM.ReportViewer1.LocalReport.ReportEmbeddedResource = "Carnes.RptIngresos.rdlc"
            DT = ClsProc.RPTIngresos(CMBEstado.Text.Trim, DtpDesde.Value, DtpHasta.Value)
            RD.Value = DT
            RD.Name = "DataSet1"

            FRM.ReportViewer1.LocalReport.DataSources.Clear()
            FRM.ReportViewer1.LocalReport.DataSources.Add(RD)
            FRM.ReportViewer1.LocalReport.Refresh()

            Dim PRM1 As New List(Of Microsoft.Reporting.WinForms.ReportParameter)
            PRM1.Add(New Microsoft.Reporting.WinForms.ReportParameter("rangofechas", "Del " + DtpDesde.Value.ToShortDateString + " Al " + DtpHasta.Value.ToShortDateString, False))
            FRM.ReportViewer1.LocalReport.SetParameters(PRM1)

            Dim PRM2 As New List(Of Microsoft.Reporting.WinForms.ReportParameter)
            PRM2.Add(New Microsoft.Reporting.WinForms.ReportParameter("estado", CMBEstado.Text.ToUpper(), False))
            FRM.ReportViewer1.LocalReport.SetParameters(PRM2)


            FRM.ReportViewer1.RefreshReport()

            FRM.Show()
        Catch EX As Exception
            MessageBox.Show(EX.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub FrmRptIngresos_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub FrmRptIngresos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ENVIAR_ENTER(e.KeyChar)
    End Sub

    Private Sub FrmRptIngresos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BtnLimpiar_Click(sender, e)
    End Sub
End Class