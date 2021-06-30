Public Class FrmCambiarPassword

    Private Sub Limpiar()
        TxtPassword.Clear()
        txtNuevoPassword.Clear()
        TxtConfirmarPassword.Clear()
    End Sub


    Private Function Acceso() As Integer
        Dim Retorno As Integer = 0
        Try
            Dim CUsuarios As New ClsGenerica
            Dim DT As New DataTable
            DT = CUsuarios.TraerDatos("UL_spC_ClaveUsuario_Valida", txtusuario.Text.Trim, TxtPassword.Text.Trim).Tables(0)
            If DT.Rows.Count = 0 Then
                MessageBox.Show("Password Incorrecto Vuelva a Ingresar los datos correctamente", "Aviso", MessageBoxButtons.OK)
                TxtPassword.Clear()
                TxtPassword.Focus()
                Retorno = 0
            Else
                Retorno = 1
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Retorno = 0
        End Try
        Return Retorno
    End Function

    Private Sub FrmCambiarPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ENVIAR_ENTER(e.KeyChar)
    End Sub

    Private Sub FrmCambiarPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub FrmCambiarPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtusuario.Text = USULOGIN
        TxtPassword.Focus()
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub

    Private Sub BtnGrabar_Click(sender As Object, e As EventArgs) Handles BtnGrabar.Click
        If TxtPassword.Text.Trim = "" Then
            MessageBox.Show("Ingrese Password", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtPassword.Focus()
            Exit Sub
        End If
        If txtNuevoPassword.Text.Trim = "" Then
            MessageBox.Show("Ingrese Nuevo Password", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtNuevoPassword.Focus()
            Exit Sub
        End If
        If TxtConfirmarPassword.Text.Trim = "" Then
            MessageBox.Show("Confirmar Password", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtConfirmarPassword.Focus()
            Exit Sub
        End If

        If txtNuevoPassword.Text.Trim <> TxtConfirmarPassword.Text.Trim Then
            MessageBox.Show("No Coinciden: Nuevo Password y Confirmar Password", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtNuevoPassword.Clear()
            TxtConfirmarPassword.Clear()
            txtNuevoPassword.Focus()
            Exit Sub
        End If
        If Acceso() = 1 Then
            Dim Retorno As String = ""
            Dim CUsuarios As New ClsGenerica
            Try
                Retorno = CUsuarios.EjecutarSP("GRL_USUARIO_CAMBIOPASS", txtusuario.Text.Trim, txtNuevoPassword.Text.Trim)
                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call Limpiar()
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If
    End Sub
End Class