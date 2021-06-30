Public Class FrmSesion

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        'If Now().Year = 2017 And Now().Month = 2 And (Now().Day = 8 Or Now().Day = 9) Then
        Try
            Dim CUsuarios As New ClsGenerica
            Dim DT As New DataTable
            DT = CUsuarios.TraerDatos("UL_spC_ClaveUsuario_Valida", txtusuario.Text.Trim, txtclave.Text.Trim).Tables(0)
            If DT.Rows.Count = 0 Then

                MessageBox.Show("Incorrecto Vuelva a Ingresar los datos correctamente", "Aviso", MessageBoxButtons.OK)
            Else

                DT = CUsuarios.TraerDatos("UL_spC_Usuario", txtusuario.Text.Trim).Tables(0)



                USUCOD = DT.Rows(0).Item("IDUsuario").ToString
                USULOGIN = DT.Rows(0).Item("login").ToString

                USUNOM = DT.Rows(0).Item("Nombre").ToString
                USUPAS = DT.Rows(0).Item("Password").ToString
                USUTIPO = DT.Rows(0).Item("descrip_jerarquia").ToString
                USUDB = "urien"
                Me.Visible = False
                Dim FRM As New FrmPrincipal
                FRM.Show()


            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        ' Else
        ' MsgBox("SISTEMA VENCIDO")
        '    End If
    End Sub

    Private Sub FrmSesion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub FrmSesion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MessageBox.Show("¿Desea Salir del Sistema?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        Else
            End
        End If
    End Sub

    Private Sub FrmSesion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ENVIAR_ENTER(e.KeyChar)
    End Sub

    Private Sub FrmSesion_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class
