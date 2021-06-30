<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmTipo
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DGTipo = New System.Windows.Forms.DataGridView()
        Me.DgTipoDato = New System.Windows.Forms.DataGridView()
        Me.PnlTipoDatos = New System.Windows.Forms.Panel()
        Me.BtnEliminarDato = New System.Windows.Forms.Button()
        Me.TxtDescripcion = New System.Windows.Forms.TextBox()
        Me.BtnGrabarDato = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.BtnNuevoDato = New System.Windows.Forms.Button()
        Me.TxtTipoDato = New System.Windows.Forms.TextBox()
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.TxtIdTipoDato = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me.lblCodigo = New System.Windows.Forms.Label()
        Me.lblLocalidad = New System.Windows.Forms.Label()
        Me.PnlTipo = New System.Windows.Forms.Panel()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.BtnGrabar = New System.Windows.Forms.Button()
        Me.BtnNuevo = New System.Windows.Forms.Button()
        Me.TxtDescrAfip = New System.Windows.Forms.TextBox()
        Me.LblDescrAfip = New System.Windows.Forms.Label()
        Me.TxtDefault = New System.Windows.Forms.TextBox()
        Me.LblDefault = New System.Windows.Forms.Label()
        Me.TxtCodigoAfip = New System.Windows.Forms.TextBox()
        Me.lblCodigoAfip = New System.Windows.Forms.Label()
        Me.TxtAtributo = New System.Windows.Forms.TextBox()
        Me.LblTipo = New System.Windows.Forms.Label()
        Me.TxtCodigoTipo = New System.Windows.Forms.TextBox()
        Me.LblAtributo = New System.Windows.Forms.Label()
        Me.TxtIdTipo = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TxtNombre = New System.Windows.Forms.TextBox()
        Me.LblNombre = New System.Windows.Forms.Label()
        Me.LblCodigoTipo = New System.Windows.Forms.Label()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        CType(Me.DGTipo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgTipoDato, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlTipoDatos.SuspendLayout()
        Me.PnlTipo.SuspendLayout()
        Me.SuspendLayout()
        '
        'DGTipo
        '
        Me.DGTipo.AllowUserToAddRows = False
        Me.DGTipo.AllowUserToDeleteRows = False
        Me.DGTipo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGTipo.Location = New System.Drawing.Point(4, 30)
        Me.DGTipo.Name = "DGTipo"
        Me.DGTipo.Size = New System.Drawing.Size(475, 172)
        Me.DGTipo.TabIndex = 254
        '
        'DgTipoDato
        '
        Me.DgTipoDato.AllowUserToAddRows = False
        Me.DgTipoDato.AllowUserToDeleteRows = False
        Me.DgTipoDato.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgTipoDato.Location = New System.Drawing.Point(3, 46)
        Me.DgTipoDato.Name = "DgTipoDato"
        Me.DgTipoDato.Size = New System.Drawing.Size(559, 152)
        Me.DgTipoDato.TabIndex = 255
        '
        'PnlTipoDatos
        '
        Me.PnlTipoDatos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PnlTipoDatos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PnlTipoDatos.Controls.Add(Me.BtnEliminarDato)
        Me.PnlTipoDatos.Controls.Add(Me.TxtDescripcion)
        Me.PnlTipoDatos.Controls.Add(Me.BtnGrabarDato)
        Me.PnlTipoDatos.Controls.Add(Me.Label7)
        Me.PnlTipoDatos.Controls.Add(Me.BtnNuevoDato)
        Me.PnlTipoDatos.Controls.Add(Me.TxtTipoDato)
        Me.PnlTipoDatos.Controls.Add(Me.DgTipoDato)
        Me.PnlTipoDatos.Controls.Add(Me.lblDescripcion)
        Me.PnlTipoDatos.Controls.Add(Me.TxtIdTipoDato)
        Me.PnlTipoDatos.Controls.Add(Me.Label1)
        Me.PnlTipoDatos.Controls.Add(Me.txtCodigo)
        Me.PnlTipoDatos.Controls.Add(Me.lblCodigo)
        Me.PnlTipoDatos.Controls.Add(Me.lblLocalidad)
        Me.PnlTipoDatos.Location = New System.Drawing.Point(12, 25)
        Me.PnlTipoDatos.Name = "PnlTipoDatos"
        Me.PnlTipoDatos.Size = New System.Drawing.Size(848, 230)
        Me.PnlTipoDatos.TabIndex = 271
        '
        'BtnEliminarDato
        '
        Me.BtnEliminarDato.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEliminarDato.Image = Global.Carnes.My.Resources.Resources.eliminar
        Me.BtnEliminarDato.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEliminarDato.Location = New System.Drawing.Point(730, 140)
        Me.BtnEliminarDato.Name = "BtnEliminarDato"
        Me.BtnEliminarDato.Size = New System.Drawing.Size(94, 32)
        Me.BtnEliminarDato.TabIndex = 292
        Me.BtnEliminarDato.Text = "&Eliminar"
        Me.BtnEliminarDato.UseVisualStyleBackColor = True
        '
        'TxtDescripcion
        '
        Me.TxtDescripcion.Location = New System.Drawing.Point(658, 114)
        Me.TxtDescripcion.MaxLength = 100
        Me.TxtDescripcion.Name = "TxtDescripcion"
        Me.TxtDescripcion.Size = New System.Drawing.Size(166, 20)
        Me.TxtDescripcion.TabIndex = 4
        '
        'BtnGrabarDato
        '
        Me.BtnGrabarDato.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGrabarDato.Image = Global.Carnes.My.Resources.Resources.grabar
        Me.BtnGrabarDato.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnGrabarDato.Location = New System.Drawing.Point(732, 178)
        Me.BtnGrabarDato.Name = "BtnGrabarDato"
        Me.BtnGrabarDato.Size = New System.Drawing.Size(94, 32)
        Me.BtnGrabarDato.TabIndex = 291
        Me.BtnGrabarDato.Text = "&Grabar"
        Me.BtnGrabarDato.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Gray
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(4, 23)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(116, 20)
        Me.Label7.TabIndex = 268
        Me.Label7.Text = "Tipos de Datos"
        '
        'BtnNuevoDato
        '
        Me.BtnNuevoDato.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNuevoDato.Image = Global.Carnes.My.Resources.Resources.nuevo
        Me.BtnNuevoDato.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnNuevoDato.Location = New System.Drawing.Point(618, 140)
        Me.BtnNuevoDato.Name = "BtnNuevoDato"
        Me.BtnNuevoDato.Size = New System.Drawing.Size(92, 32)
        Me.BtnNuevoDato.TabIndex = 290
        Me.BtnNuevoDato.Text = "&Nuevo"
        Me.BtnNuevoDato.UseVisualStyleBackColor = True
        '
        'TxtTipoDato
        '
        Me.TxtTipoDato.Location = New System.Drawing.Point(658, 55)
        Me.TxtTipoDato.MaxLength = 100
        Me.TxtTipoDato.Name = "TxtTipoDato"
        Me.TxtTipoDato.Size = New System.Drawing.Size(166, 20)
        Me.TxtTipoDato.TabIndex = 2
        '
        'lblDescripcion
        '
        Me.lblDescripcion.AutoSize = True
        Me.lblDescripcion.Location = New System.Drawing.Point(568, 117)
        Me.lblDescripcion.Name = "lblDescripcion"
        Me.lblDescripcion.Size = New System.Drawing.Size(63, 13)
        Me.lblDescripcion.TabIndex = 267
        Me.lblDescripcion.Text = "Descripcion"
        '
        'TxtIdTipoDato
        '
        Me.TxtIdTipoDato.BackColor = System.Drawing.Color.Bisque
        Me.TxtIdTipoDato.Location = New System.Drawing.Point(724, 10)
        Me.TxtIdTipoDato.Name = "TxtIdTipoDato"
        Me.TxtIdTipoDato.ReadOnly = True
        Me.TxtIdTipoDato.Size = New System.Drawing.Size(100, 20)
        Me.TxtIdTipoDato.TabIndex = 1
        Me.TxtIdTipoDato.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Gray
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(644, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 16)
        Me.Label1.TabIndex = 255
        Me.Label1.Text = "Detalle"
        '
        'txtCodigo
        '
        Me.txtCodigo.Location = New System.Drawing.Point(658, 83)
        Me.txtCodigo.MaxLength = 100
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(166, 20)
        Me.txtCodigo.TabIndex = 3
        '
        'lblCodigo
        '
        Me.lblCodigo.AutoSize = True
        Me.lblCodigo.Location = New System.Drawing.Point(568, 90)
        Me.lblCodigo.Name = "lblCodigo"
        Me.lblCodigo.Size = New System.Drawing.Size(40, 13)
        Me.lblCodigo.TabIndex = 262
        Me.lblCodigo.Text = "Codigo"
        '
        'lblLocalidad
        '
        Me.lblLocalidad.AutoSize = True
        Me.lblLocalidad.Location = New System.Drawing.Point(568, 55)
        Me.lblLocalidad.Name = "lblLocalidad"
        Me.lblLocalidad.Size = New System.Drawing.Size(69, 13)
        Me.lblLocalidad.TabIndex = 261
        Me.lblLocalidad.Text = "Tipo de Dato"
        '
        'PnlTipo
        '
        Me.PnlTipo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PnlTipo.Controls.Add(Me.BtnSalir)
        Me.PnlTipo.Controls.Add(Me.BtnEliminar)
        Me.PnlTipo.Controls.Add(Me.BtnGrabar)
        Me.PnlTipo.Controls.Add(Me.BtnNuevo)
        Me.PnlTipo.Controls.Add(Me.TxtDescrAfip)
        Me.PnlTipo.Controls.Add(Me.LblDescrAfip)
        Me.PnlTipo.Controls.Add(Me.TxtDefault)
        Me.PnlTipo.Controls.Add(Me.LblDefault)
        Me.PnlTipo.Controls.Add(Me.TxtCodigoAfip)
        Me.PnlTipo.Controls.Add(Me.lblCodigoAfip)
        Me.PnlTipo.Controls.Add(Me.TxtAtributo)
        Me.PnlTipo.Controls.Add(Me.LblTipo)
        Me.PnlTipo.Controls.Add(Me.TxtCodigoTipo)
        Me.PnlTipo.Controls.Add(Me.LblAtributo)
        Me.PnlTipo.Controls.Add(Me.TxtIdTipo)
        Me.PnlTipo.Controls.Add(Me.Label4)
        Me.PnlTipo.Controls.Add(Me.TxtNombre)
        Me.PnlTipo.Controls.Add(Me.LblNombre)
        Me.PnlTipo.Controls.Add(Me.LblCodigoTipo)
        Me.PnlTipo.Controls.Add(Me.DGTipo)
        Me.PnlTipo.Location = New System.Drawing.Point(12, 273)
        Me.PnlTipo.Name = "PnlTipo"
        Me.PnlTipo.Size = New System.Drawing.Size(848, 248)
        Me.PnlTipo.TabIndex = 272
        '
        'BtnSalir
        '
        Me.BtnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSalir.Image = Global.Carnes.My.Resources.Resources.salir
        Me.BtnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSalir.Location = New System.Drawing.Point(732, 170)
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.Size = New System.Drawing.Size(92, 32)
        Me.BtnSalir.TabIndex = 289
        Me.BtnSalir.Text = "&Salir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'BtnEliminar
        '
        Me.BtnEliminar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEliminar.Image = Global.Carnes.My.Resources.Resources.eliminar
        Me.BtnEliminar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEliminar.Location = New System.Drawing.Point(732, 122)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(94, 32)
        Me.BtnEliminar.TabIndex = 288
        Me.BtnEliminar.Text = "&Eliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'BtnGrabar
        '
        Me.BtnGrabar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGrabar.Image = Global.Carnes.My.Resources.Resources.grabar
        Me.BtnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnGrabar.Location = New System.Drawing.Point(730, 80)
        Me.BtnGrabar.Name = "BtnGrabar"
        Me.BtnGrabar.Size = New System.Drawing.Size(94, 32)
        Me.BtnGrabar.TabIndex = 287
        Me.BtnGrabar.Text = "&Grabar"
        Me.BtnGrabar.UseVisualStyleBackColor = True
        '
        'BtnNuevo
        '
        Me.BtnNuevo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNuevo.Image = Global.Carnes.My.Resources.Resources.nuevo
        Me.BtnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnNuevo.Location = New System.Drawing.Point(732, 36)
        Me.BtnNuevo.Name = "BtnNuevo"
        Me.BtnNuevo.Size = New System.Drawing.Size(92, 32)
        Me.BtnNuevo.TabIndex = 286
        Me.BtnNuevo.Text = "&Nuevo"
        Me.BtnNuevo.UseVisualStyleBackColor = True
        '
        'TxtDescrAfip
        '
        Me.TxtDescrAfip.Location = New System.Drawing.Point(560, 181)
        Me.TxtDescrAfip.MaxLength = 100
        Me.TxtDescrAfip.Name = "TxtDescrAfip"
        Me.TxtDescrAfip.Size = New System.Drawing.Size(166, 20)
        Me.TxtDescrAfip.TabIndex = 285
        '
        'LblDescrAfip
        '
        Me.LblDescrAfip.AutoSize = True
        Me.LblDescrAfip.Location = New System.Drawing.Point(494, 184)
        Me.LblDescrAfip.Name = "LblDescrAfip"
        Me.LblDescrAfip.Size = New System.Drawing.Size(56, 13)
        Me.LblDescrAfip.TabIndex = 284
        Me.LblDescrAfip.Text = "Descr Afip"
        '
        'TxtDefault
        '
        Me.TxtDefault.Location = New System.Drawing.Point(560, 129)
        Me.TxtDefault.MaxLength = 100
        Me.TxtDefault.Name = "TxtDefault"
        Me.TxtDefault.Size = New System.Drawing.Size(166, 20)
        Me.TxtDefault.TabIndex = 280
        '
        'LblDefault
        '
        Me.LblDefault.AutoSize = True
        Me.LblDefault.Location = New System.Drawing.Point(494, 132)
        Me.LblDefault.Name = "LblDefault"
        Me.LblDefault.Size = New System.Drawing.Size(41, 13)
        Me.LblDefault.TabIndex = 282
        Me.LblDefault.Text = "Default"
        '
        'TxtCodigoAfip
        '
        Me.TxtCodigoAfip.Location = New System.Drawing.Point(560, 155)
        Me.TxtCodigoAfip.MaxLength = 100
        Me.TxtCodigoAfip.Name = "TxtCodigoAfip"
        Me.TxtCodigoAfip.Size = New System.Drawing.Size(166, 20)
        Me.TxtCodigoAfip.TabIndex = 281
        '
        'lblCodigoAfip
        '
        Me.lblCodigoAfip.AutoSize = True
        Me.lblCodigoAfip.Location = New System.Drawing.Point(494, 158)
        Me.lblCodigoAfip.Name = "lblCodigoAfip"
        Me.lblCodigoAfip.Size = New System.Drawing.Size(61, 13)
        Me.lblCodigoAfip.TabIndex = 280
        Me.lblCodigoAfip.Text = "Codigo Afip"
        '
        'TxtAtributo
        '
        Me.TxtAtributo.Location = New System.Drawing.Point(560, 103)
        Me.TxtAtributo.MaxLength = 100
        Me.TxtAtributo.Name = "TxtAtributo"
        Me.TxtAtributo.Size = New System.Drawing.Size(166, 20)
        Me.TxtAtributo.TabIndex = 279
        '
        'LblTipo
        '
        Me.LblTipo.AutoSize = True
        Me.LblTipo.BackColor = System.Drawing.Color.Gray
        Me.LblTipo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTipo.Location = New System.Drawing.Point(4, 9)
        Me.LblTipo.Name = "LblTipo"
        Me.LblTipo.Size = New System.Drawing.Size(39, 20)
        Me.LblTipo.TabIndex = 278
        Me.LblTipo.Text = "Tipo"
        '
        'TxtCodigoTipo
        '
        Me.TxtCodigoTipo.Location = New System.Drawing.Point(560, 44)
        Me.TxtCodigoTipo.MaxLength = 100
        Me.TxtCodigoTipo.Name = "TxtCodigoTipo"
        Me.TxtCodigoTipo.Size = New System.Drawing.Size(166, 20)
        Me.TxtCodigoTipo.TabIndex = 269
        '
        'LblAtributo
        '
        Me.LblAtributo.AutoSize = True
        Me.LblAtributo.Location = New System.Drawing.Point(494, 106)
        Me.LblAtributo.Name = "LblAtributo"
        Me.LblAtributo.Size = New System.Drawing.Size(43, 13)
        Me.LblAtributo.TabIndex = 277
        Me.LblAtributo.Text = "Atributo"
        '
        'TxtIdTipo
        '
        Me.TxtIdTipo.BackColor = System.Drawing.Color.Bisque
        Me.TxtIdTipo.Location = New System.Drawing.Point(724, 9)
        Me.TxtIdTipo.Name = "TxtIdTipo"
        Me.TxtIdTipo.ReadOnly = True
        Me.TxtIdTipo.Size = New System.Drawing.Size(100, 20)
        Me.TxtIdTipo.TabIndex = 268
        Me.TxtIdTipo.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Gray
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(644, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 16)
        Me.Label4.TabIndex = 273
        Me.Label4.Text = "Detalle"
        '
        'TxtNombre
        '
        Me.TxtNombre.Location = New System.Drawing.Point(560, 72)
        Me.TxtNombre.MaxLength = 100
        Me.TxtNombre.Name = "TxtNombre"
        Me.TxtNombre.Size = New System.Drawing.Size(166, 20)
        Me.TxtNombre.TabIndex = 270
        '
        'LblNombre
        '
        Me.LblNombre.AutoSize = True
        Me.LblNombre.Location = New System.Drawing.Point(494, 79)
        Me.LblNombre.Name = "LblNombre"
        Me.LblNombre.Size = New System.Drawing.Size(44, 13)
        Me.LblNombre.TabIndex = 275
        Me.LblNombre.Text = "Nombre"
        '
        'LblCodigoTipo
        '
        Me.LblCodigoTipo.AutoSize = True
        Me.LblCodigoTipo.Location = New System.Drawing.Point(494, 44)
        Me.LblCodigoTipo.Name = "LblCodigoTipo"
        Me.LblCodigoTipo.Size = New System.Drawing.Size(40, 13)
        Me.LblCodigoTipo.TabIndex = 274
        Me.LblCodigoTipo.Text = "Codigo"
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 521)
        Me.Splitter1.TabIndex = 273
        Me.Splitter1.TabStop = False
        '
        'FrmTipo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(885, 521)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.PnlTipo)
        Me.Controls.Add(Me.PnlTipoDatos)
        Me.Name = "FrmTipo"
        Me.Text = "TipoDatos"
        CType(Me.DGTipo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgTipoDato, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlTipoDatos.ResumeLayout(False)
        Me.PnlTipoDatos.PerformLayout()
        Me.PnlTipo.ResumeLayout(False)
        Me.PnlTipo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DGTipo As DataGridView
    Friend WithEvents DgTipoDato As DataGridView
    Friend WithEvents PnlTipoDatos As Panel
    Friend WithEvents TxtTipoDato As TextBox
    Friend WithEvents lblDescripcion As Label
    Friend WithEvents TxtIdTipoDato As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtCodigo As TextBox
    Friend WithEvents lblCodigo As Label
    Friend WithEvents lblLocalidad As Label
    Friend WithEvents PnlTipo As Panel
    Friend WithEvents TxtCodigoTipo As TextBox
    Friend WithEvents LblAtributo As Label
    Friend WithEvents TxtIdTipo As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TxtNombre As TextBox
    Friend WithEvents LblNombre As Label
    Friend WithEvents LblCodigoTipo As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents LblTipo As Label
    Friend WithEvents Splitter1 As Splitter
    Friend WithEvents TxtDescripcion As TextBox
    Friend WithEvents TxtAtributo As TextBox
    Friend WithEvents TxtDescrAfip As TextBox
    Friend WithEvents LblDescrAfip As Label
    Friend WithEvents TxtDefault As TextBox
    Friend WithEvents LblDefault As Label
    Friend WithEvents TxtCodigoAfip As TextBox
    Friend WithEvents lblCodigoAfip As Label
    Friend WithEvents BtnSalir As Button
    Friend WithEvents BtnEliminar As Button
    Friend WithEvents BtnGrabar As Button
    Friend WithEvents BtnNuevo As Button
    Friend WithEvents BtnEliminarDato As Button
    Friend WithEvents BtnGrabarDato As Button
    Friend WithEvents BtnNuevoDato As Button
End Class
