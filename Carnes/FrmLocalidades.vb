Public Class FrmLocalidades

    Dim Modo As String = ""
    Dim OLocalidad(6) As Object
    Private Sub FrmLocalidades_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RellenarLista()

        ComboPais()
        ComboProvincia()


    End Sub

#Region "Grilla Localidad"


    Private Sub DgLocalidades_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgLocalidades.CellClick

        On Error GoTo ErrLinea
        Dim Fila As Integer
        If DgLocalidades.Rows(DgLocalidades.CurrentRow.Index).Cells(0).Value.ToString <> "" Then
            Fila = DgLocalidades.CurrentRow.Index
            BuscaLocalidad(DgLocalidades.Rows(Fila).Cells(0).Value.ToString())

        End If
        Return
ErrLinea:
        MessageBox.Show("No hay registros en la lista de busqueda", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)


    End Sub

#End Region


#Region "Rellenar Grilla Localidad"

    Private Sub RellenarLista()



        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("Ven_localidad_Consul", 0).Tables(0)
            DgLocalidades.DataSource = DT
            DgLocalidades.ReadOnly = True
            DgLocalidades.Columns("idprovincia").Visible = False
            DgLocalidades.Columns("idpais").Visible = False


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub BuscaLocalidad(Codigo As Integer)
        Try

            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("Ven_localidad_Consul", Codigo).Tables(0)

            Modo = "M"

            If DT.Rows.Count > 0 Then

                TxtIdLocalidad.Text = DT.Rows(0).Item("IdLocalidad").ToString()
                txtnombre.Text = DT.Rows(0).Item("nombre").ToString()
                txtCodLoc.Text = DT.Rows(0).Item("CodLoc").ToString()
                CmbIdProvincia.SelectedValue = DT.Rows(0).Item("IdProvincia").ToString()
                CmbIdPais.SelectedValue = DT.Rows(0).Item("IdPais").ToString()





            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub



#End Region


#Region "Rutinas Botones"


    Private Sub BtnNuevo_Click(sender As Object, e As EventArgs) Handles BtnNuevo.Click
        Limpiar()
        RellenarLista()
    End Sub

    Private Sub BtnGrabar_Click(sender As Object, e As EventArgs) Handles BtnGrabar.Click

        If TxtNombre.Text.Trim = "" Then
            MessageBox.Show("Ingrese Nombre Localidad", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtNombre.Focus()
            Exit Sub
        End If
        '         @idlocalidad 	    int          = null,
        '   @nombre             varchar(100) = NULL,
        '   @codloc             varchar(10)  = null,
        '@idProvincia		int          = null,
        '@idPais        		int          = null,
        '   @usuario	        varchar(20),
        '   @Modo 		        char(1)


        OLocalidad(0) = TxtIdLocalidad.Text
        OLocalidad(1) = txtnombre.Text
        OLocalidad(2) = txtCodLoc.Text
        OLocalidad(3) = CmbIdProvincia.SelectedValue
        OLocalidad(4) = CmbIdPais.SelectedValue
        OLocalidad(5) = USULOGIN




        Dim Retorno As Integer



        Dim CLocalidad As New ClsGenerica
        Dim DT As New DataTable

        Try
            If TxtIdLocalidad.Text.Trim = "" Then
                OLocalidad(0) = 0

                OLocalidad(6) = "A"

                Retorno = CLocalidad.EjecutarSP("ven_localidad_Abm", OLocalidad)


                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Agregado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)
                End If
            Else

                OLocalidad(6) = "M"
                Retorno = CLocalidad.EjecutarSP("ven_localidad_Abm", OLocalidad)

                If Val(Retorno) > 0 Then

                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click

        OLocalidad(0) = TxtIdLocalidad.Text
        OLocalidad(1) = txtnombre.Text
        OLocalidad(2) = txtCodLoc.Text
        OLocalidad(3) = CmbIdProvincia.SelectedValue
        OLocalidad(4) = CmbIdPais.SelectedValue
        OLocalidad(5) = USULOGIN
        OLocalidad(6) = "B"


        If TxtIdLocalidad.Text.Trim = "" Then
            MessageBox.Show("Buscar Registro de Localidad", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtIdLocalidad.Focus()
            Exit Sub
        End If

        If MessageBox.Show("¿Desea Eliminar el Registro Seleccionado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim Retorno As Integer
        Dim CLocalidad As New ClsGenerica

        Try
            Retorno = CLocalidad.EjecutarSP("ven_localidad_Abm", OLocalidad)
            If Val(Retorno) > 0 Then
                MessageBox.Show("Registro Eliminado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call BtnNuevo_Click(sender, e)
            Else
                MessageBox.Show("No se pudo Realizar la Operación", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub
#End Region

#Region "CARGAR COMBOS"
    ' Cargar Combo Localidad


    ' Cargar Combo Provincia
    Private Sub ComboProvincia()
        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.EjecutarSql("select idprovincia, nombre From grl_provincia").Tables(0)
            With CmbIdProvincia
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idprovincia"
            End With




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso- Provincia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    ' Cargar Combo Localidad
    Private Sub ComboPais()
        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.EjecutarSql("select idPais, nombre From grl_pais").Tables(0)
            With CmbIdPais
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idPais"
            End With
        Catch ex As Exception


            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso- Pais", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub



#End Region
#Region "Rutinas"

    Private Sub Limpiar()

        txtCodLoc.Clear()
        TxtIdLocalidad.Clear()
        TxtNombre.Clear()



        CmbIdProvincia.SelectedIndex = 0
        CmbIdPais.SelectedIndex = 0




    End Sub



#End Region

End Class