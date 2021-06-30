Public Class FrmTipo

    Dim Modo As String = ""

    Dim Otipo(9) As Object
    Dim OtipoDatos(5) As Object




    Dim VarIdTipoDato As Integer



    Private Sub FrmTipoDato_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CargarGrillaTipoDatos()


        PnlTipo.Visible = False

    End Sub


#Region "Rutinas de Grillas"

    ' Grilla GdvTipoDato
    Private Sub GdvTipoDato_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgTipoDato.CellClick

        On Error GoTo ErrLinea
        Dim Fila As Integer
        If DgTipoDato.Rows(DgTipoDato.CurrentRow.Index).Cells(0).Value.ToString <> "" Then
            Fila = DgTipoDato.CurrentRow.Index

            BuscarTipoDato(DgTipoDato.Rows(Fila).Cells(0).Value.ToString())

            CargarGrillaTipo(DgTipoDato.Rows(Fila).Cells(0).Value.ToString())



        End If
        Return
ErrLinea:
        MessageBox.Show("No hay registros en la lista de busqueda", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    End Sub





    Private Sub GdvTipo_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGTipo.CellClick

        On Error GoTo ErrLinea

        Dim Fila As Integer
        If DGTipo.Rows(DGTipo.CurrentRow.Index).Cells(0).Value.ToString <> "" Then
            Fila = DGTipo.CurrentRow.Index
            BuscarTipo(DGTipo.Rows(Fila).Cells(0).Value.ToString())



        End If
        Return
ErrLinea:
        MessageBox.Show("No hay registros en la lista de busqueda", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    End Sub

#End Region

#Region "Rutinas"

    Private Sub BuscarTipoDato(Codigo As Integer)
        Try

            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("Gen_Tipodatos_Consul", Codigo).Tables(0)

            Modo = "M"

            If DT.Rows.Count > 0 Then


                TxtIdTipoDato.Text = DT.Rows(0).Item("IdTipodato").ToString()
                TxtIdTipoDato.Enabled = False
                TxtTipoDato.Text = DT.Rows(0).Item("tipoDato").ToString()
                txtCodigo.Text = DT.Rows(0).Item("codigo").ToString()

                TxtDescripcion.Text = DT.Rows(0).Item("descripcion").ToString()






            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    Private Sub BuscarTipo(Codigo As Integer)
        Try

            ' MsgBox(Codigo)

            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("Gen_Tipos_ou", Codigo).Tables(0)

            Modo = "M"

            If DT.Rows.Count > 0 Then

                TxtIdTipo.Text = DT.Rows(0).Item("IdTipo").ToString()
                TxtCodigoTipo.Text = DT.Rows(0).Item("Codigo").ToString()
                TxtNombre.Text = DT.Rows(0).Item("Nombre").ToString()
                TxtAtributo.Text = DT.Rows(0).Item("Atributo").ToString()
                TxtDefault.Text = DT.Rows(0).Item("DefaultTipo").ToString()
                TxtCodigoAfip.Text = DT.Rows(0).Item("CodigoAfip").ToString()
                TxtDescrAfip.Text = DT.Rows(0).Item("DescrAfip").ToString()






            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


#End Region

#Region "Rellenar Grilla"

    ' Grilla Tipo De Datos
    Private Sub CargarGrillaTipoDatos()



        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("Gen_Tipodatos_Consul", 0).Tables(0)
            DgTipoDato.DataSource = DT
            DgTipoDato.ReadOnly = True
            ' DgLocalidades.Columns("idprovincia").Visible = False
            ' DgLocalidades.Columns("idpais").Visible = False


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    ' Grilla Tipo 

    Private Sub CargarGrillaTipo(Codigo As Integer)
        Try
            VarIdTipoDato = Codigo

            PnlTipo.Visible = True

            Dim CUsuario As New ClsGenerica
            Dim DT As New DataTable
            DT = CUsuario.TraerDatos("Gen_Tipos_Consul", Codigo).Tables(0)

            DGTipo.DataSource = DT

            DGTipo.Columns("idTipodato").Visible = False
            DGTipo.Columns("fechaModificacion").Visible = False
            DGTipo.Columns("fechaCreacion").Visible = False
            DGTipo.Columns("UsuarioCreacion").Visible = False
            DGTipo.Columns("UsuarioModificacion").Visible = False

            If DT.Rows.Count > 0 Then

                BuscarTipo(DT.Rows(0)("idtipo"))

            End If




            'DgModuloAsignado.Columns("idUsuario").Visible = False
            ' DgModuloAsignado.Columns("jerarquia").Visible = False






        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

#Region "Rutinas Botones"


    Private Sub BtnNuevo_Click(sender As Object, e As EventArgs) Handles BtnNuevo.Click

        LimpiarFormularioTipo()


    End Sub

    Private Sub BtnGrabar_Click(sender As Object, e As EventArgs) Handles BtnGrabar.Click

        If TxtNombre.Text.Trim = "" Then
            MessageBox.Show("Ingrese Nombre Localidad", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtNombre.Focus()
            Exit Sub
        End If



        Otipo(0) = TxtIdTipo.Text
        Otipo(1) = TxtIdTipoDato.Text
        Otipo(2) = TxtCodigoTipo.Text
        Otipo(3) = TxtNombre.Text
        Otipo(4) = TxtAtributo.Text
        Otipo(5) = TxtDefault.Text
        Otipo(6) = USULOGIN
        Otipo(7) = Modo
        Otipo(8) = TxtCodigoAfip.Text
        Otipo(9) = TxtDescrAfip.Text





        Dim Retorno As Integer



        Dim Ctipo As New ClsGenerica
        Dim DT As New DataTable

        Try
            If TxtIdTipo.Text.Trim = "" Then

                Otipo(0) = 0
                Otipo(7) = "A"

                Retorno = Ctipo.EjecutarSP("Gen_Tipos_Abm", Otipo)


                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Agregado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)
                    CargarGrillaTipo(Otipo(1))
                    BuscarTipo(Retorno)
                End If
            Else

                Otipo(7) = "M"
                Retorno = Ctipo.EjecutarSP("Gen_Tipos_Abm", Otipo)

                If Val(Retorno) > 0 Then

                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)


                    CargarGrillaTipo(Otipo(1))
                    BuscarTipo(Retorno)
                End If
            End If
            ClsFormularios.GrillaPosicionarme(DGTipo, Retorno)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click

        Otipo(0) = TxtIdTipo.Text
        Otipo(1) = TxtIdTipoDato.Text
        Otipo(2) = TxtCodigoTipo.Text
        Otipo(3) = TxtNombre.Text
        Otipo(4) = TxtAtributo.Text
        Otipo(5) = TxtDefault.Text
        Otipo(6) = USULOGIN
        Otipo(7) = "B"
        Otipo(8) = TxtCodigoAfip.Text
        Otipo(9) = TxtDescrAfip.Text



        If TxtIdTipo.Text.Trim = "" Then
            MessageBox.Show("Buscar Registro de Tipo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtIdTipo.Focus()
            Exit Sub
        End If

        If MessageBox.Show("¿Desea Eliminar el Registro Seleccionado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim Retorno As Integer
        Dim CLocalidad As New ClsGenerica

        Try
            Retorno = CLocalidad.EjecutarSP("Gen_Tipos_Abm", Otipo)
            If Val(Retorno) > 0 Then
                MessageBox.Show("Registro Eliminado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)


                LimpiarFormulario()
                CargarGrillaTipo(Otipo(1))


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



    Private Sub LimpiarFormulario()

        TxtIdTipoDato.Clear()
        TxtTipoDato.Clear()
        txtCodigo.Clear()
        TxtDescripcion.Clear()


    End Sub


    Private Sub LimpiarFormularioTipo()


        TxtIdTipo.Clear()
        TxtCodigoTipo.Clear()
        TxtNombre.Clear()
        TxtAtributo.Clear()
        TxtDefault.Clear()
        TxtCodigoAfip.Clear()
        TxtDescrAfip.Clear()





    End Sub

    Private Sub BtnNuevoDato_Click(sender As Object, e As EventArgs) Handles BtnNuevoDato.Click

        LimpiarFormulario()
        LimpiarFormularioTipo()
        CargarGrillaTipoDatos()



    End Sub

    Private Sub BtnEliminarDato_Click(sender As Object, e As EventArgs) Handles BtnEliminarDato.Click

        CargarTipoDatos()

        OtipoDatos(5) = "B"

        If TxtTipoDato.Text.Trim = "" Then
            MessageBox.Show("Buscar Registro de Tipo de Dato", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtTipoDato.Focus()
            Exit Sub
        End If

        If MessageBox.Show("¿Desea Eliminar el Registro Seleccionado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim Retorno As Integer
        Dim CLocalidad As New ClsGenerica

        Try
            Retorno = CLocalidad.EjecutarSP("Gen_Tipodatos_Abm", OtipoDatos)
            If Val(Retorno) > 0 Then
                MessageBox.Show("Registro Eliminado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)


                LimpiarFormulario()


                CargarGrillaTipoDatos()


            Else
                MessageBox.Show("No se pudo Realizar la Operación", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub





    Private Sub BtnGrabarDato_Click(sender As Object, e As EventArgs) Handles BtnGrabarDato.Click

        If TxtTipoDato.Text.Trim = "" Then
            MessageBox.Show("Ingrese el Tipo de Dato", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtTipoDato.Focus()
            Exit Sub
        End If


        CargarTipoDatos()




        Dim Retorno As Integer



        Dim Ctipo As New ClsGenerica
        Dim DT As New DataTable

        Try
            If TxtIdTipoDato.Text.Trim = "" Then

                OtipoDatos(5) = "A"


                Retorno = Ctipo.EjecutarSP("Gen_Tipodatos_Abm", OtipoDatos)


                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Agregado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevoDato_Click(sender, e)
                    BuscarTipoDato(Retorno)


                End If
            Else
                OtipoDatos(5) = "M"

                Retorno = Ctipo.EjecutarSP("Gen_Tipodatos_Abm", OtipoDatos)

                If Val(Retorno) > 0 Then

                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevoDato_Click(sender, e)

                    BuscarTipoDato(Retorno)
                End If
            End If

            ClsFormularios.GrillaPosicionarme(DgTipoDato, Retorno)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub


    Private Sub CargarTipoDatos()




        OtipoDatos(0) = Val(TxtIdTipoDato.Text)
        OtipoDatos(1) = TxtTipoDato.Text
        OtipoDatos(2) = txtCodigo.Text
        OtipoDatos(3) = TxtDescripcion.Text
        OtipoDatos(4) = USULOGIN

    End Sub











#End Region

End Class