<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUsuarios
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmUsuarios))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtCodigo = New System.Windows.Forms.TextBox()
        Me.txtusuario = New System.Windows.Forms.TextBox()
        Me.TxtPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtnombres = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtapellidos = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.DgUsuarios = New System.Windows.Forms.DataGridView()
        Me.CmbIdSuperior = New System.Windows.Forms.ComboBox()
        Me.CmbIdFrigorifico = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckActivo = New System.Windows.Forms.CheckBox()
        Me.DgModuloAsignado = New System.Windows.Forms.DataGridView()
        Me.DgModulo = New System.Windows.Forms.DataGridView()
        Me.CmbJerarquia = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.BtnModuloAgregar = New System.Windows.Forms.Button()
        Me.BtnModuloEliminar = New System.Windows.Forms.Button()
        Me.ChkDgModuloAsignado = New System.Windows.Forms.CheckBox()
        Me.chkDgModulo = New System.Windows.Forms.CheckBox()
        Me.LblModulo = New System.Windows.Forms.Label()
        Me.PanelModulo = New System.Windows.Forms.Panel()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.BtnGrabar = New System.Windows.Forms.Button()
        Me.BtnNuevo = New System.Windows.Forms.Button()
        Me.BtnModificar = New System.Windows.Forms.Button()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        CType(Me.DgUsuarios, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgModuloAsignado, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgModulo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelModulo.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(781, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Codigo"
        '
        'TxtCodigo
        '
        Me.TxtCodigo.BackColor = System.Drawing.Color.Bisque
        Me.TxtCodigo.Location = New System.Drawing.Point(871, 33)
        Me.TxtCodigo.Name = "TxtCodigo"
        Me.TxtCodigo.ReadOnly = True
        Me.TxtCodigo.Size = New System.Drawing.Size(100, 20)
        Me.TxtCodigo.TabIndex = 1
        '
        'txtusuario
        '
        Me.txtusuario.Location = New System.Drawing.Point(529, 59)
        Me.txtusuario.MaxLength = 30
        Me.txtusuario.Name = "txtusuario"
        Me.txtusuario.Size = New System.Drawing.Size(166, 20)
        Me.txtusuario.TabIndex = 2
        '
        'TxtPassword
        '
        Me.TxtPassword.Location = New System.Drawing.Point(784, 59)
        Me.TxtPassword.MaxLength = 30
        Me.TxtPassword.Name = "TxtPassword"
        Me.TxtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtPassword.Size = New System.Drawing.Size(187, 20)
        Me.TxtPassword.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(720, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(465, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Usuario"
        '
        'txtnombres
        '
        Me.txtnombres.Location = New System.Drawing.Point(529, 85)
        Me.txtnombres.MaxLength = 100
        Me.txtnombres.Name = "txtnombres"
        Me.txtnombres.Size = New System.Drawing.Size(166, 20)
        Me.txtnombres.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(465, 85)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Nombres"
        '
        'txtapellidos
        '
        Me.txtapellidos.Location = New System.Drawing.Point(784, 85)
        Me.txtapellidos.MaxLength = 100
        Me.txtapellidos.Name = "txtapellidos"
        Me.txtapellidos.Size = New System.Drawing.Size(187, 20)
        Me.txtapellidos.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(720, 85)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Apellidos"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(465, 111)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(46, 13)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Superior"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(721, 119)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(52, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Frigorifico"
        '
        'DgUsuarios
        '
        Me.DgUsuarios.AllowUserToAddRows = False
        Me.DgUsuarios.AllowUserToDeleteRows = False
        Me.DgUsuarios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgUsuarios.Location = New System.Drawing.Point(12, 38)
        Me.DgUsuarios.Name = "DgUsuarios"
        Me.DgUsuarios.Size = New System.Drawing.Size(447, 150)
        Me.DgUsuarios.TabIndex = 27
        '
        'CmbIdSuperior
        '
        Me.CmbIdSuperior.FormattingEnabled = True
        Me.CmbIdSuperior.Location = New System.Drawing.Point(529, 111)
        Me.CmbIdSuperior.Name = "CmbIdSuperior"
        Me.CmbIdSuperior.Size = New System.Drawing.Size(166, 21)
        Me.CmbIdSuperior.TabIndex = 6
        '
        'CmbIdFrigorifico
        '
        Me.CmbIdFrigorifico.FormattingEnabled = True
        Me.CmbIdFrigorifico.Location = New System.Drawing.Point(784, 111)
        Me.CmbIdFrigorifico.Name = "CmbIdFrigorifico"
        Me.CmbIdFrigorifico.Size = New System.Drawing.Size(187, 21)
        Me.CmbIdFrigorifico.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 13)
        Me.Label4.TabIndex = 252
        Me.Label4.Text = "Listado de Usuarios"
        '
        'CheckActivo
        '
        Me.CheckActivo.AutoSize = True
        Me.CheckActivo.Location = New System.Drawing.Point(468, 169)
        Me.CheckActivo.Name = "CheckActivo"
        Me.CheckActivo.Size = New System.Drawing.Size(56, 17)
        Me.CheckActivo.TabIndex = 8
        Me.CheckActivo.Text = "Activo"
        Me.CheckActivo.UseVisualStyleBackColor = True
        '
        'DgModuloAsignado
        '
        Me.DgModuloAsignado.AllowUserToAddRows = False
        Me.DgModuloAsignado.AllowUserToDeleteRows = False
        Me.DgModuloAsignado.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgModuloAsignado.Location = New System.Drawing.Point(9, 22)
        Me.DgModuloAsignado.Name = "DgModuloAsignado"
        Me.DgModuloAsignado.Size = New System.Drawing.Size(367, 160)
        Me.DgModuloAsignado.TabIndex = 254
        '
        'DgModulo
        '
        Me.DgModulo.AllowUserToAddRows = False
        Me.DgModulo.AllowUserToDeleteRows = False
        Me.DgModulo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgModulo.Location = New System.Drawing.Point(517, 22)
        Me.DgModulo.Name = "DgModulo"
        Me.DgModulo.Size = New System.Drawing.Size(280, 160)
        Me.DgModulo.TabIndex = 255
        '
        'CmbJerarquia
        '
        Me.CmbJerarquia.FormattingEnabled = True
        Me.CmbJerarquia.Location = New System.Drawing.Point(803, 57)
        Me.CmbJerarquia.Name = "CmbJerarquia"
        Me.CmbJerarquia.Size = New System.Drawing.Size(142, 21)
        Me.CmbJerarquia.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(803, 41)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 13)
        Me.Label8.TabIndex = 257
        Me.Label8.Text = "Jerarquia"
        '
        'BtnModuloAgregar
        '
        Me.BtnModuloAgregar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnModuloAgregar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModuloAgregar.Location = New System.Drawing.Point(395, 27)
        Me.BtnModuloAgregar.Name = "BtnModuloAgregar"
        Me.BtnModuloAgregar.Size = New System.Drawing.Size(107, 37)
        Me.BtnModuloAgregar.TabIndex = 15
        Me.BtnModuloAgregar.Text = " << Agregar"
        Me.BtnModuloAgregar.UseVisualStyleBackColor = True
        '
        'BtnModuloEliminar
        '
        Me.BtnModuloEliminar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnModuloEliminar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModuloEliminar.Location = New System.Drawing.Point(395, 153)
        Me.BtnModuloEliminar.Name = "BtnModuloEliminar"
        Me.BtnModuloEliminar.Size = New System.Drawing.Size(107, 29)
        Me.BtnModuloEliminar.TabIndex = 16
        Me.BtnModuloEliminar.Text = "Quitar >>"
        Me.BtnModuloEliminar.UseVisualStyleBackColor = True
        '
        'ChkDgModuloAsignado
        '
        Me.ChkDgModuloAsignado.AutoSize = True
        Me.ChkDgModuloAsignado.Location = New System.Drawing.Point(9, 194)
        Me.ChkDgModuloAsignado.Name = "ChkDgModuloAsignado"
        Me.ChkDgModuloAsignado.Size = New System.Drawing.Size(89, 17)
        Me.ChkDgModuloAsignado.TabIndex = 12
        Me.ChkDgModuloAsignado.Text = "Marca Todos"
        Me.ChkDgModuloAsignado.UseVisualStyleBackColor = True
        '
        'chkDgModulo
        '
        Me.chkDgModulo.AutoSize = True
        Me.chkDgModulo.Location = New System.Drawing.Point(517, 197)
        Me.chkDgModulo.Name = "chkDgModulo"
        Me.chkDgModulo.Size = New System.Drawing.Size(92, 17)
        Me.chkDgModulo.TabIndex = 15
        Me.chkDgModulo.Text = "Marcar Todos"
        Me.chkDgModulo.UseVisualStyleBackColor = True
        '
        'LblModulo
        '
        Me.LblModulo.AutoSize = True
        Me.LblModulo.BackColor = System.Drawing.SystemColors.ControlDark
        Me.LblModulo.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblModulo.Location = New System.Drawing.Point(5, 0)
        Me.LblModulo.Name = "LblModulo"
        Me.LblModulo.Size = New System.Drawing.Size(68, 19)
        Me.LblModulo.TabIndex = 262
        Me.LblModulo.Text = "Modulos"
        '
        'PanelModulo
        '
        Me.PanelModulo.Controls.Add(Me.DgModuloAsignado)
        Me.PanelModulo.Controls.Add(Me.LblModulo)
        Me.PanelModulo.Controls.Add(Me.ChkDgModuloAsignado)
        Me.PanelModulo.Controls.Add(Me.chkDgModulo)
        Me.PanelModulo.Controls.Add(Me.DgModulo)
        Me.PanelModulo.Controls.Add(Me.BtnModuloEliminar)
        Me.PanelModulo.Controls.Add(Me.CmbJerarquia)
        Me.PanelModulo.Controls.Add(Me.BtnModuloAgregar)
        Me.PanelModulo.Controls.Add(Me.Label8)
        Me.PanelModulo.Location = New System.Drawing.Point(16, 214)
        Me.PanelModulo.Name = "PanelModulo"
        Me.PanelModulo.Size = New System.Drawing.Size(955, 218)
        Me.PanelModulo.TabIndex = 263
        '
        'BtnSalir
        '
        Me.BtnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSalir.Image = Global.Carnes.My.Resources.Resources.salir
        Me.BtnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSalir.Location = New System.Drawing.Point(884, 446)
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.Size = New System.Drawing.Size(87, 32)
        Me.BtnSalir.TabIndex = 20
        Me.BtnSalir.Text = "&Salir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'BtnGrabar
        '
        Me.BtnGrabar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGrabar.Image = Global.Carnes.My.Resources.Resources.grabar
        Me.BtnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnGrabar.Location = New System.Drawing.Point(884, 138)
        Me.BtnGrabar.Name = "BtnGrabar"
        Me.BtnGrabar.Size = New System.Drawing.Size(92, 32)
        Me.BtnGrabar.TabIndex = 13
        Me.BtnGrabar.Text = "&Aceptar"
        Me.BtnGrabar.UseVisualStyleBackColor = True
        '
        'BtnNuevo
        '
        Me.BtnNuevo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnNuevo.Location = New System.Drawing.Point(533, 154)
        Me.BtnNuevo.Name = "BtnNuevo"
        Me.BtnNuevo.Size = New System.Drawing.Size(92, 32)
        Me.BtnNuevo.TabIndex = 10
        Me.BtnNuevo.Text = "&Nuevo"
        Me.BtnNuevo.UseVisualStyleBackColor = True
        '
        'BtnModificar
        '
        Me.BtnModificar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnModificar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnModificar.Location = New System.Drawing.Point(643, 154)
        Me.BtnModificar.Name = "BtnModificar"
        Me.BtnModificar.Size = New System.Drawing.Size(92, 32)
        Me.BtnModificar.TabIndex = 11
        Me.BtnModificar.Text = "&Modificar"
        Me.BtnModificar.UseVisualStyleBackColor = True
        '
        'BtnCancelar
        '
        Me.BtnCancelar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCancelar.Image = Global.Carnes.My.Resources.Resources.eliminar
        Me.BtnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnCancelar.Location = New System.Drawing.Point(884, 176)
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.Size = New System.Drawing.Size(92, 32)
        Me.BtnCancelar.TabIndex = 14
        Me.BtnCancelar.Text = "&Cancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = True
        '
        'BtnEliminar
        '
        Me.BtnEliminar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEliminar.Location = New System.Drawing.Point(752, 154)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(92, 32)
        Me.BtnEliminar.TabIndex = 12
        Me.BtnEliminar.Text = "&Eliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'FrmUsuarios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(995, 490)
        Me.Controls.Add(Me.BtnCancelar)
        Me.Controls.Add(Me.BtnModificar)
        Me.Controls.Add(Me.PanelModulo)
        Me.Controls.Add(Me.CheckActivo)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DgUsuarios)
        Me.Controls.Add(Me.CmbIdFrigorifico)
        Me.Controls.Add(Me.CmbIdSuperior)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtapellidos)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtnombres)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtPassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtusuario)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtCodigo)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.BtnEliminar)
        Me.Controls.Add(Me.BtnGrabar)
        Me.Controls.Add(Me.BtnNuevo)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "FrmUsuarios"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Usuarios"
        CType(Me.DgUsuarios, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgModuloAsignado, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgModulo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelModulo.ResumeLayout(False)
        Me.PanelModulo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnNuevo As System.Windows.Forms.Button
    Friend WithEvents BtnGrabar As System.Windows.Forms.Button
    Friend WithEvents BtnSalir As System.Windows.Forms.Button
    Friend WithEvents TxtCodigo As System.Windows.Forms.TextBox
    Friend WithEvents txtusuario As System.Windows.Forms.TextBox
    Friend WithEvents TxtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtnombres As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtapellidos As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DgUsuarios As System.Windows.Forms.DataGridView
    Friend WithEvents CmbIdSuperior As ComboBox
    Friend WithEvents CmbIdFrigorifico As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents CheckActivo As CheckBox
    Friend WithEvents DgModuloAsignado As DataGridView
    Friend WithEvents DgModulo As DataGridView
    Friend WithEvents CmbJerarquia As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents BtnModuloAgregar As Button
    Friend WithEvents BtnModuloEliminar As Button
    Friend WithEvents ChkDgModuloAsignado As CheckBox
    Friend WithEvents chkDgModulo As CheckBox
    Friend WithEvents LblModulo As Label
    Friend WithEvents PanelModulo As Panel
    Friend WithEvents BtnCancelar As Button
    Friend WithEvents BtnModificar As Button
    Friend WithEvents BtnEliminar As Button
End Class
