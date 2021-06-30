Imports Microsoft.Reporting.WinForms

Public Class FrmPrincipal

    Private Sub FrmPrincipal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("¿Desea Salir del Sistema?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        Else
            End
        End If
    End Sub

    Private Sub FrmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Sistema de Ventas     Usuario: " + USUNOM + "     Base de datos: " + USUDB

        'AQUI SE BLOQUEAN MENUS DEL SISTEMA 
        If USUTIPO = "VENDEDOR" Then
            RegistroDeProveedoresToolStripMenuItem.Visible = False
            RegistroDeFamiliarToolStripMenuItem.Visible = False
            RegistroDeSubFamiliasToolStripMenuItem.Visible = False
            RegistroDeMarcasToolStripMenuItem.Visible = False
            RegistroDeArticulosToolStripMenuItem.Visible = False
            ImpuestoToolStripMenuItem.Visible = False
            SeriesToolStripMenuItem.Visible = False
            OrdenDeIngresoToolStripMenuItem.Visible = False
            ReporteDeVentasToolStripMenuItem.Visible = False
            ReporteDeClientesToolStripMenuItem.Visible = False
            ReporteDeProveedoresToolStripMenuItem.Visible = False
            ReporteDeArticulosToolStripMenuItem.Visible = False
            ReporteDeFamiliasToolStripMenuItem.Visible = False
            ReporteDeSubFamiliasToolStripMenuItem.Visible = False
            ReporteDeMarcasToolStripMenuItem.Visible = False
            RegistroDeUsuariosToolStripMenuItem.Visible = False
            ReporteDeIngresosToolStripMenuItem.Visible = False
            ReporteDeKardexToolStripMenuItem.Visible = False
        End If


    End Sub






    Private Sub CambiarPasswordToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        Dim FRM As New FrmCambiarPassword
        FRM.MdiParent = Me
        FRM.Show()
    End Sub

    Private Sub RegistroDeUsuariosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroDeUsuariosToolStripMenuItem.Click
        Dim FRM As New FrmUsuarios
        FRM.MdiParent = Me
        FRM.Show()
    End Sub

















    Private Sub ProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProveedoresToolStripMenuItem.Click

        Dim FRM As New FrmProductores
        FRM.MdiParent = Me
        FRM.Show()

    End Sub

    Private Sub LocalidadesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LocalidadesToolStripMenuItem.Click

        Dim FRM As New FrmLocalidades
        FRM.MdiParent = Me
        FRM.Show()

    End Sub

    Private Sub TablasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TablasToolStripMenuItem1.Click

        Dim FRM As New FrmTipo
        FRM.MdiParent = Me
        FRM.Show()

    End Sub

    Private Sub LogueoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogueoToolStripMenuItem.Click

        Dim FRM As New FrmLogueo
        FRM.MdiParent = Me
        FRM.Show()

    End Sub


End Class