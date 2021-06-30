<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLocalidades
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
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DgLocalidades = New System.Windows.Forms.DataGridView()
        Me.CmbIdProvincia = New System.Windows.Forms.ComboBox()
        Me.CmbIdPais = New System.Windows.Forms.ComboBox()
        Me.txtCodLoc = New System.Windows.Forms.TextBox()
        Me.lblCodigo = New System.Windows.Forms.Label()
        Me.lblLocalidad = New System.Windows.Forms.Label()
        Me.TxtIdLocalidad = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.BtnGrabar = New System.Windows.Forms.Button()
        Me.BtnNuevo = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtnombre = New System.Windows.Forms.TextBox()
        Me.lblPais = New System.Windows.Forms.Label()
        Me.idProvincia = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        CType(Me.DgLocalidades, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(27, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 13)
        Me.Label4.TabIndex = 254
        Me.Label4.Text = "Listado de Localidades"
        '
        'DgLocalidades
        '
        Me.DgLocalidades.AllowUserToAddRows = False
        Me.DgLocalidades.AllowUserToDeleteRows = False
        Me.DgLocalidades.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgLocalidades.Location = New System.Drawing.Point(30, 34)
        Me.DgLocalidades.Name = "DgLocalidades"
        Me.DgLocalidades.Size = New System.Drawing.Size(454, 189)
        Me.DgLocalidades.TabIndex = 253
        '
        'CmbIdProvincia
        '
        Me.CmbIdProvincia.FormattingEnabled = True
        Me.CmbIdProvincia.Location = New System.Drawing.Point(101, 149)
        Me.CmbIdProvincia.Name = "CmbIdProvincia"
        Me.CmbIdProvincia.Size = New System.Drawing.Size(166, 21)
        Me.CmbIdProvincia.TabIndex = 4
        '
        'CmbIdPais
        '
        Me.CmbIdPais.FormattingEnabled = True
        Me.CmbIdPais.Location = New System.Drawing.Point(101, 117)
        Me.CmbIdPais.Name = "CmbIdPais"
        Me.CmbIdPais.Size = New System.Drawing.Size(166, 21)
        Me.CmbIdPais.TabIndex = 5
        '
        'txtCodLoc
        '
        Me.txtCodLoc.Location = New System.Drawing.Point(101, 86)
        Me.txtCodLoc.MaxLength = 100
        Me.txtCodLoc.Name = "txtCodLoc"
        Me.txtCodLoc.Size = New System.Drawing.Size(166, 20)
        Me.txtCodLoc.TabIndex = 3
        '
        'lblCodigo
        '
        Me.lblCodigo.AutoSize = True
        Me.lblCodigo.Location = New System.Drawing.Point(11, 93)
        Me.lblCodigo.Name = "lblCodigo"
        Me.lblCodigo.Size = New System.Drawing.Size(40, 13)
        Me.lblCodigo.TabIndex = 262
        Me.lblCodigo.Text = "Codigo"
        '
        'lblLocalidad
        '
        Me.lblLocalidad.AutoSize = True
        Me.lblLocalidad.Location = New System.Drawing.Point(11, 58)
        Me.lblLocalidad.Name = "lblLocalidad"
        Me.lblLocalidad.Size = New System.Drawing.Size(53, 13)
        Me.lblLocalidad.TabIndex = 261
        Me.lblLocalidad.Text = "Localidad"
        '
        'TxtIdLocalidad
        '
        Me.TxtIdLocalidad.BackColor = System.Drawing.Color.Bisque
        Me.TxtIdLocalidad.Location = New System.Drawing.Point(167, 13)
        Me.TxtIdLocalidad.Name = "TxtIdLocalidad"
        Me.TxtIdLocalidad.ReadOnly = True
        Me.TxtIdLocalidad.Size = New System.Drawing.Size(100, 20)
        Me.TxtIdLocalidad.TabIndex = 1
        Me.TxtIdLocalidad.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(87, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 255
        Me.Label1.Text = "Codigo"
        '
        'BtnSalir
        '
        Me.BtnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSalir.Image = Global.Carnes.My.Resources.Resources.salir
        Me.BtnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSalir.Location = New System.Drawing.Point(167, 73)
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.Size = New System.Drawing.Size(84, 32)
        Me.BtnSalir.TabIndex = 9
        Me.BtnSalir.Text = "&Salir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'BtnEliminar
        '
        Me.BtnEliminar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEliminar.Image = Global.Carnes.My.Resources.Resources.eliminar
        Me.BtnEliminar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEliminar.Location = New System.Drawing.Point(39, 73)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(94, 32)
        Me.BtnEliminar.TabIndex = 8
        Me.BtnEliminar.Text = "&Eliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'BtnGrabar
        '
        Me.BtnGrabar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnGrabar.Image = Global.Carnes.My.Resources.Resources.grabar
        Me.BtnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnGrabar.Location = New System.Drawing.Point(167, 20)
        Me.BtnGrabar.Name = "BtnGrabar"
        Me.BtnGrabar.Size = New System.Drawing.Size(84, 32)
        Me.BtnGrabar.TabIndex = 7
        Me.BtnGrabar.Text = "&Grabar"
        Me.BtnGrabar.UseVisualStyleBackColor = True
        '
        'BtnNuevo
        '
        Me.BtnNuevo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNuevo.Image = Global.Carnes.My.Resources.Resources.nuevo
        Me.BtnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnNuevo.Location = New System.Drawing.Point(41, 20)
        Me.BtnNuevo.Name = "BtnNuevo"
        Me.BtnNuevo.Size = New System.Drawing.Size(92, 32)
        Me.BtnNuevo.TabIndex = 6
        Me.BtnNuevo.Text = "&Nuevo"
        Me.BtnNuevo.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtnombre)
        Me.Panel1.Controls.Add(Me.lblPais)
        Me.Panel1.Controls.Add(Me.idProvincia)
        Me.Panel1.Controls.Add(Me.TxtIdLocalidad)
        Me.Panel1.Controls.Add(Me.CmbIdPais)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtCodLoc)
        Me.Panel1.Controls.Add(Me.CmbIdProvincia)
        Me.Panel1.Controls.Add(Me.lblCodigo)
        Me.Panel1.Controls.Add(Me.lblLocalidad)
        Me.Panel1.Location = New System.Drawing.Point(497, 34)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(278, 189)
        Me.Panel1.TabIndex = 270
        '
        'txtnombre
        '
        Me.txtnombre.Location = New System.Drawing.Point(101, 58)
        Me.txtnombre.MaxLength = 100
        Me.txtnombre.Name = "txtnombre"
        Me.txtnombre.Size = New System.Drawing.Size(166, 20)
        Me.txtnombre.TabIndex = 2
        '
        'lblPais
        '
        Me.lblPais.AutoSize = True
        Me.lblPais.Location = New System.Drawing.Point(11, 120)
        Me.lblPais.Name = "lblPais"
        Me.lblPais.Size = New System.Drawing.Size(27, 13)
        Me.lblPais.TabIndex = 267
        Me.lblPais.Text = "Pais"
        '
        'idProvincia
        '
        Me.idProvincia.AutoSize = True
        Me.idProvincia.Location = New System.Drawing.Point(11, 149)
        Me.idProvincia.Name = "idProvincia"
        Me.idProvincia.Size = New System.Drawing.Size(51, 13)
        Me.idProvincia.TabIndex = 266
        Me.idProvincia.Text = "Provincia"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.BtnGrabar)
        Me.Panel2.Controls.Add(Me.BtnNuevo)
        Me.Panel2.Controls.Add(Me.BtnSalir)
        Me.Panel2.Controls.Add(Me.BtnEliminar)
        Me.Panel2.Location = New System.Drawing.Point(497, 240)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(278, 137)
        Me.Panel2.TabIndex = 271
        '
        'FrmLocalidades
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DgLocalidades)
        Me.Name = "FrmLocalidades"
        Me.Text = "FrmLocalidades"
        CType(Me.DgLocalidades, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label4 As Label
    Friend WithEvents DgLocalidades As DataGridView
    Friend WithEvents CmbIdProvincia As ComboBox
    Friend WithEvents CmbIdPais As ComboBox
    Friend WithEvents txtCodLoc As TextBox
    Friend WithEvents lblCodigo As Label
    Friend WithEvents lblLocalidad As Label
    Friend WithEvents TxtIdLocalidad As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BtnSalir As Button
    Friend WithEvents BtnEliminar As Button
    Friend WithEvents BtnGrabar As Button
    Friend WithEvents BtnNuevo As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lblPais As Label
    Friend WithEvents idProvincia As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents txtnombre As TextBox
End Class
