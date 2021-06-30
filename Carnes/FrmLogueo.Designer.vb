<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLogueo
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
        Me.ListArchivosDatos = New System.Windows.Forms.ListBox()
        Me.GroupBoxResultado = New System.Windows.Forms.GroupBox()
        Me.ListArchivos = New System.Windows.Forms.ListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblArchivo = New System.Windows.Forms.Label()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.GroupBoxResultado.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListArchivosDatos
        '
        Me.ListArchivosDatos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListArchivosDatos.FormattingEnabled = True
        Me.ListArchivosDatos.Location = New System.Drawing.Point(6, 41)
        Me.ListArchivosDatos.Name = "ListArchivosDatos"
        Me.ListArchivosDatos.Size = New System.Drawing.Size(686, 212)
        Me.ListArchivosDatos.TabIndex = 0
        '
        'GroupBoxResultado
        '
        Me.GroupBoxResultado.Controls.Add(Me.ListArchivosDatos)
        Me.GroupBoxResultado.Controls.Add(Me.lblArchivo)
        Me.GroupBoxResultado.Controls.Add(Me.BtnEliminar)
        Me.GroupBoxResultado.Location = New System.Drawing.Point(44, 192)
        Me.GroupBoxResultado.Name = "GroupBoxResultado"
        Me.GroupBoxResultado.Size = New System.Drawing.Size(698, 272)
        Me.GroupBoxResultado.TabIndex = 17
        Me.GroupBoxResultado.TabStop = False
        Me.GroupBoxResultado.Text = "Resultado"
        '
        'ListArchivos
        '
        Me.ListArchivos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListArchivos.FormattingEnabled = True
        Me.ListArchivos.Location = New System.Drawing.Point(6, 19)
        Me.ListArchivos.Name = "ListArchivos"
        Me.ListArchivos.Size = New System.Drawing.Size(686, 134)
        Me.ListArchivos.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ListArchivos)
        Me.GroupBox1.Location = New System.Drawing.Point(44, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(698, 165)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Resultado"
        '
        'lblArchivo
        '
        Me.lblArchivo.AutoSize = True
        Me.lblArchivo.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblArchivo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArchivo.Location = New System.Drawing.Point(6, 22)
        Me.lblArchivo.Name = "lblArchivo"
        Me.lblArchivo.Size = New System.Drawing.Size(55, 16)
        Me.lblArchivo.TabIndex = 19
        Me.lblArchivo.Text = "Label1"
        '
        'BtnSalir
        '
        Me.BtnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSalir.Image = Global.Carnes.My.Resources.Resources.salir
        Me.BtnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnSalir.Location = New System.Drawing.Point(650, 470)
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.Size = New System.Drawing.Size(92, 32)
        Me.BtnSalir.TabIndex = 322
        Me.BtnSalir.Text = "&Salir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'BtnEliminar
        '
        Me.BtnEliminar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnEliminar.Image = Global.Carnes.My.Resources.Resources.eliminar
        Me.BtnEliminar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnEliminar.Location = New System.Drawing.Point(598, 6)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(94, 32)
        Me.BtnEliminar.TabIndex = 321
        Me.BtnEliminar.Text = "&Eliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'FrmLogueo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 529)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxResultado)
        Me.Name = "FrmLogueo"
        Me.Text = "FrmLogueo"
        Me.GroupBoxResultado.ResumeLayout(False)
        Me.GroupBoxResultado.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListArchivosDatos As ListBox
    Friend WithEvents GroupBoxResultado As GroupBox
    Friend WithEvents ListArchivos As ListBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblArchivo As Label
    Friend WithEvents BtnSalir As Button
    Friend WithEvents BtnEliminar As Button
End Class
