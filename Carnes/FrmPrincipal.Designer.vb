<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPrincipal
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmPrincipal))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.TablasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeClientesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeFamiliarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeSubFamiliasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeMarcasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeArticulosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImpuestoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeriesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OrdenDeIngresoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VentasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeVentasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeIngresosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeKardexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeClientesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeArticulosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeFamiliasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeSubFamiliasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReporteDeMarcasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SeguridadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeUsuariosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TablasToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.LocalidadesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogueoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HaciendaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TablasToolStripMenuItem, Me.IngresoToolStripMenuItem, Me.VentasToolStripMenuItem, Me.ReportesToolStripMenuItem, Me.HaciendaToolStripMenuItem, Me.SeguridadToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(607, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'TablasToolStripMenuItem
        '
        Me.TablasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RegistroDeClientesToolStripMenuItem, Me.RegistroDeProveedoresToolStripMenuItem, Me.RegistroDeFamiliarToolStripMenuItem, Me.RegistroDeSubFamiliasToolStripMenuItem, Me.RegistroDeMarcasToolStripMenuItem, Me.RegistroDeArticulosToolStripMenuItem, Me.ImpuestoToolStripMenuItem, Me.SeriesToolStripMenuItem})
        Me.TablasToolStripMenuItem.Enabled = False
        Me.TablasToolStripMenuItem.Name = "TablasToolStripMenuItem"
        Me.TablasToolStripMenuItem.Size = New System.Drawing.Size(62, 20)
        Me.TablasToolStripMenuItem.Text = "Registro"
        '
        'RegistroDeClientesToolStripMenuItem
        '
        Me.RegistroDeClientesToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.clientes
        Me.RegistroDeClientesToolStripMenuItem.Name = "RegistroDeClientesToolStripMenuItem"
        Me.RegistroDeClientesToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeClientesToolStripMenuItem.Text = "Registro de Clientes"
        '
        'RegistroDeProveedoresToolStripMenuItem
        '
        Me.RegistroDeProveedoresToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.proveedor
        Me.RegistroDeProveedoresToolStripMenuItem.Name = "RegistroDeProveedoresToolStripMenuItem"
        Me.RegistroDeProveedoresToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeProveedoresToolStripMenuItem.Text = "Registro de Proveedores"
        '
        'RegistroDeFamiliarToolStripMenuItem
        '
        Me.RegistroDeFamiliarToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.familias
        Me.RegistroDeFamiliarToolStripMenuItem.Name = "RegistroDeFamiliarToolStripMenuItem"
        Me.RegistroDeFamiliarToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeFamiliarToolStripMenuItem.Text = "Registro de Familias"
        '
        'RegistroDeSubFamiliasToolStripMenuItem
        '
        Me.RegistroDeSubFamiliasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.subfamilias
        Me.RegistroDeSubFamiliasToolStripMenuItem.Name = "RegistroDeSubFamiliasToolStripMenuItem"
        Me.RegistroDeSubFamiliasToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeSubFamiliasToolStripMenuItem.Text = "Registro de SubFamilias"
        '
        'RegistroDeMarcasToolStripMenuItem
        '
        Me.RegistroDeMarcasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.marcas
        Me.RegistroDeMarcasToolStripMenuItem.Name = "RegistroDeMarcasToolStripMenuItem"
        Me.RegistroDeMarcasToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeMarcasToolStripMenuItem.Text = "Registro de Marcas"
        '
        'RegistroDeArticulosToolStripMenuItem
        '
        Me.RegistroDeArticulosToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.articulos
        Me.RegistroDeArticulosToolStripMenuItem.Name = "RegistroDeArticulosToolStripMenuItem"
        Me.RegistroDeArticulosToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RegistroDeArticulosToolStripMenuItem.Text = "Registro de Articulos"
        '
        'ImpuestoToolStripMenuItem
        '
        Me.ImpuestoToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.monedas
        Me.ImpuestoToolStripMenuItem.Name = "ImpuestoToolStripMenuItem"
        Me.ImpuestoToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ImpuestoToolStripMenuItem.Text = "Impuesto"
        '
        'SeriesToolStripMenuItem
        '
        Me.SeriesToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.series
        Me.SeriesToolStripMenuItem.Name = "SeriesToolStripMenuItem"
        Me.SeriesToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.SeriesToolStripMenuItem.Text = "Series"
        '
        'IngresoToolStripMenuItem
        '
        Me.IngresoToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OrdenDeIngresoToolStripMenuItem})
        Me.IngresoToolStripMenuItem.Enabled = False
        Me.IngresoToolStripMenuItem.Name = "IngresoToolStripMenuItem"
        Me.IngresoToolStripMenuItem.Size = New System.Drawing.Size(58, 20)
        Me.IngresoToolStripMenuItem.Text = "Ingreso"
        '
        'OrdenDeIngresoToolStripMenuItem
        '
        Me.OrdenDeIngresoToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.ingresos
        Me.OrdenDeIngresoToolStripMenuItem.Name = "OrdenDeIngresoToolStripMenuItem"
        Me.OrdenDeIngresoToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.OrdenDeIngresoToolStripMenuItem.Text = "Hacienda"
        '
        'VentasToolStripMenuItem
        '
        Me.VentasToolStripMenuItem.Enabled = False
        Me.VentasToolStripMenuItem.Name = "VentasToolStripMenuItem"
        Me.VentasToolStripMenuItem.Size = New System.Drawing.Size(53, 20)
        Me.VentasToolStripMenuItem.Text = "Ventas"
        '
        'ReportesToolStripMenuItem
        '
        Me.ReportesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReporteDeVentasToolStripMenuItem, Me.ReporteDeIngresosToolStripMenuItem, Me.ReporteDeKardexToolStripMenuItem, Me.ReporteDeClientesToolStripMenuItem, Me.ReporteDeProveedoresToolStripMenuItem, Me.ReporteDeArticulosToolStripMenuItem, Me.ReporteDeFamiliasToolStripMenuItem, Me.ReporteDeSubFamiliasToolStripMenuItem, Me.ReporteDeMarcasToolStripMenuItem})
        Me.ReportesToolStripMenuItem.Enabled = False
        Me.ReportesToolStripMenuItem.Name = "ReportesToolStripMenuItem"
        Me.ReportesToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.ReportesToolStripMenuItem.Text = "Reportes"
        '
        'ReporteDeVentasToolStripMenuItem
        '
        Me.ReporteDeVentasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeVentasToolStripMenuItem.Name = "ReporteDeVentasToolStripMenuItem"
        Me.ReporteDeVentasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeVentasToolStripMenuItem.Text = "Reportes Crystal"
        '
        'ReporteDeIngresosToolStripMenuItem
        '
        Me.ReporteDeIngresosToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeIngresosToolStripMenuItem.Name = "ReporteDeIngresosToolStripMenuItem"
        Me.ReporteDeIngresosToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeIngresosToolStripMenuItem.Text = "Reporte de Ingresos"
        '
        'ReporteDeKardexToolStripMenuItem
        '
        Me.ReporteDeKardexToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeKardexToolStripMenuItem.Name = "ReporteDeKardexToolStripMenuItem"
        Me.ReporteDeKardexToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeKardexToolStripMenuItem.Text = "Reporte de Kardex"
        '
        'ReporteDeClientesToolStripMenuItem
        '
        Me.ReporteDeClientesToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeClientesToolStripMenuItem.Name = "ReporteDeClientesToolStripMenuItem"
        Me.ReporteDeClientesToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeClientesToolStripMenuItem.Text = "Reporte de Clientes"
        '
        'ReporteDeProveedoresToolStripMenuItem
        '
        Me.ReporteDeProveedoresToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeProveedoresToolStripMenuItem.Name = "ReporteDeProveedoresToolStripMenuItem"
        Me.ReporteDeProveedoresToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeProveedoresToolStripMenuItem.Text = "Reporte de Proveedores"
        '
        'ReporteDeArticulosToolStripMenuItem
        '
        Me.ReporteDeArticulosToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeArticulosToolStripMenuItem.Name = "ReporteDeArticulosToolStripMenuItem"
        Me.ReporteDeArticulosToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeArticulosToolStripMenuItem.Text = "Reporte de Articulos"
        '
        'ReporteDeFamiliasToolStripMenuItem
        '
        Me.ReporteDeFamiliasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeFamiliasToolStripMenuItem.Name = "ReporteDeFamiliasToolStripMenuItem"
        Me.ReporteDeFamiliasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeFamiliasToolStripMenuItem.Text = "Reporte de Familias"
        '
        'ReporteDeSubFamiliasToolStripMenuItem
        '
        Me.ReporteDeSubFamiliasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeSubFamiliasToolStripMenuItem.Name = "ReporteDeSubFamiliasToolStripMenuItem"
        Me.ReporteDeSubFamiliasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeSubFamiliasToolStripMenuItem.Text = "Reporte de Sub Familias"
        '
        'ReporteDeMarcasToolStripMenuItem
        '
        Me.ReporteDeMarcasToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.reporte
        Me.ReporteDeMarcasToolStripMenuItem.Name = "ReporteDeMarcasToolStripMenuItem"
        Me.ReporteDeMarcasToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ReporteDeMarcasToolStripMenuItem.Text = "Reporte de Marcas"
        '
        'SeguridadToolStripMenuItem
        '
        Me.SeguridadToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RegistroDeUsuariosToolStripMenuItem, Me.ProveedoresToolStripMenuItem, Me.TablasToolStripMenuItem1, Me.LocalidadesToolStripMenuItem, Me.LogueoToolStripMenuItem})
        Me.SeguridadToolStripMenuItem.Name = "SeguridadToolStripMenuItem"
        Me.SeguridadToolStripMenuItem.Size = New System.Drawing.Size(100, 20)
        Me.SeguridadToolStripMenuItem.Text = "Administracion"
        '
        'RegistroDeUsuariosToolStripMenuItem
        '
        Me.RegistroDeUsuariosToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.usuario
        Me.RegistroDeUsuariosToolStripMenuItem.Name = "RegistroDeUsuariosToolStripMenuItem"
        Me.RegistroDeUsuariosToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.RegistroDeUsuariosToolStripMenuItem.Text = "Registro de Usuarios"
        '
        'ProveedoresToolStripMenuItem
        '
        Me.ProveedoresToolStripMenuItem.Image = Global.Carnes.My.Resources.Resources.proveedor
        Me.ProveedoresToolStripMenuItem.Name = "ProveedoresToolStripMenuItem"
        Me.ProveedoresToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ProveedoresToolStripMenuItem.Text = "Productores"
        '
        'TablasToolStripMenuItem1
        '
        Me.TablasToolStripMenuItem1.Name = "TablasToolStripMenuItem1"
        Me.TablasToolStripMenuItem1.Size = New System.Drawing.Size(181, 22)
        Me.TablasToolStripMenuItem1.Text = "Tipo Datos"
        '
        'LocalidadesToolStripMenuItem
        '
        Me.LocalidadesToolStripMenuItem.Name = "LocalidadesToolStripMenuItem"
        Me.LocalidadesToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.LocalidadesToolStripMenuItem.Text = "Localidades"
        '
        'LogueoToolStripMenuItem
        '
        Me.LogueoToolStripMenuItem.Name = "LogueoToolStripMenuItem"
        Me.LogueoToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.LogueoToolStripMenuItem.Text = "Logueo"
        '
        'HaciendaToolStripMenuItem
        '
        Me.HaciendaToolStripMenuItem.Name = "HaciendaToolStripMenuItem"
        Me.HaciendaToolStripMenuItem.Size = New System.Drawing.Size(69, 20)
        Me.HaciendaToolStripMenuItem.Text = "Hacienda"
        '
        'FrmPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Carnes.My.Resources.Resources.FONDO
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(607, 401)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FrmPrincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SISTEMA DE VENTAS"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents TablasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeClientesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeFamiliarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeSubFamiliasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeMarcasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeArticulosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImpuestoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SeriesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OrdenDeIngresoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VentasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SeguridadToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeVentasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeClientesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeArticulosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeFamiliasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeSubFamiliasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeMarcasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeUsuariosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeIngresosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReporteDeKardexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProveedoresToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TablasToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents LocalidadesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LogueoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HaciendaToolStripMenuItem As ToolStripMenuItem
End Class
