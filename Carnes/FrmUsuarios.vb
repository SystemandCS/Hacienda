Public Class FrmUsuarios



    Public Modo As String


    Private Sub FrmUsuarios_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RellenarComboSuperior()
        RellenarComboFrigorifico()

        RellenarComboJerarquica()

        LblModulo.Text = "Modulos"

        PanelModulo.Visible = False

        Call BtnNuevo_Click(sender, e)

    End Sub


#Region "Rutinas de la Grillas"


    Private Sub DgUsuarios_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgUsuarios.CellClick

        On Error GoTo ErrLinea
        Dim Fila As Integer
        If DgUsuarios.Rows(DgUsuarios.CurrentRow.Index).Cells(0).Value.ToString <> "" Then
            Fila = DgUsuarios.CurrentRow.Index
            BuscaRegistro(DgUsuarios.Rows(Fila).Cells(0).Value.ToString())


            RellenarModuloAsigados(DgUsuarios.Rows(Fila).Cells(0).Value.ToString())

            RellenarModulo(DgUsuarios.Rows(Fila).Cells(0).Value.ToString())

            ClsFormularios.Bloquear(Me)


        End If
        Return
ErrLinea:
        MessageBox.Show("No hay registros en la lista de busqueda", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    End Sub





#End Region

#Region "Rutinas del Formulario"

    Private Sub Limpiar()
        txtusuario.Clear()
        TxtPassword.Clear()
        TxtCodigo.Clear()
        txtnombres.Clear()
        txtapellidos.Clear()
        txtusuario.Focus()
    End Sub


    Private Sub FrmUsuarios_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ENVIAR_ENTER(e.KeyChar)
    End Sub

    Private Sub FrmUsuarios_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

#End Region


#Region "Rutinas de Carga Datos "

    Private Sub BuscaRegistro(Codigo As Integer)
        Try
            Dim CUsuario As New ClsUsuario
            Dim DT As New DataTable
            DT = CUsuario.ObtenerUno(Codigo).Tables(0)
            If DT.Rows.Count > 0 Then
                TxtCodigo.Text = DT.Rows(0).Item("idUsuario").ToString()
                txtusuario.Text = DT.Rows(0).Item("Usuario").ToString()
                txtapellidos.Text = DT.Rows(0).Item("Apellido").ToString()
                txtnombres.Text = DT.Rows(0).Item("Nombre").ToString()
                TxtPassword.Text = DT.Rows(0).Item("Password").ToString()
                CmbIdSuperior.SelectedValue = DT.Rows(0).Item("idsuperior").ToString()
                CmbIdFrigorifico.SelectedValue = DT.Rows(0).Item("idfrigorifico").ToString()

                LblModulo.Text = "Modulos para " & txtnombres.Text & " " & txtapellidos.Text

                PanelModulo.Visible = True

                BtnEliminar.Enabled = True



            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    ' Rellena Grilla usuarios
    Private Sub RellenarLista(ByVal Valor As String)
        Try
            Dim CUsuario As New ClsGenerica
            Dim DT As New DataTable
            'DT = CUsuario.EjecutarSP("of_lista", Valor).Tables(0)

            Dim ObjUsuario(1) As Object
            ObjUsuario(0) = ""
            ObjUsuario(1) = 1

            DT = CUsuario.TraerDatos("UL_spC_Usuario_Modulo", ObjUsuario).Tables(0)


            DgUsuarios.DataSource = DT
            DgUsuarios.ReadOnly = True
            ' DgUsuarios.Columns("idUsuario").Visible = False
            'DgUsuarios.Columns("jerarquia").Visible = False
            ' DgUsuarios.Columns("Acceso").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    Private Sub RellenarModuloAsigados(Codigo As Integer)
        Try
            Dim CUsuario As New ClsGenerica
            Dim DT As New DataTable
            DT = CUsuario.TraerDatos("UL_spC_ModuloXUsuario", Codigo, "").Tables(0)


            DgModuloAsignado.Columns.Clear()

            DgModuloAsignado.DataSource = DT

            DgModuloAsignado.Columns("idUsuario").Visible = False
            DgModuloAsignado.Columns("jerarquia").Visible = False

            'Add a CheckBox Column to the DataGridView at the first position.
            Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            checkBoxColumn.HeaderText = ""
            checkBoxColumn.Width = 30
            checkBoxColumn.Name = "ChkCodigoModuloAsignado"
            DgModuloAsignado.Columns.Insert(0, checkBoxColumn)




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    Private Sub RellenarModulo(Codigo As Integer)
        Try

            Dim CUsuario As New ClsGenerica
            Dim DT As New DataTable
            DT = CUsuario.TraerDatos("UL_spC_ModuloXUsuario", Codigo, "B").Tables(0)



            DgModulo.Columns.Clear()

            DgModulo.DataSource = DT

            'Add a CheckBox Column to the DataGridView at the first position.
            Dim checkBoxColumn As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            checkBoxColumn.HeaderText = ""
            checkBoxColumn.Width = 30
            checkBoxColumn.Name = "ChkCodigoModulo"
            DgModulo.Columns.Insert(0, checkBoxColumn)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub






#End Region

#Region "Rutinas de Carga de Combos"


    Private Sub RellenarComboSuperior()
        Try
            Dim CFacCab As New ClsUsuario
            Dim DT As New DataTable
            DT = CFacCab.ObtenerTodos.Tables(0)
            With CmbIdSuperior
                .DataSource = DT
                .DisplayMember = "Usuario"
                .ValueMember = "idUsuario"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    ' traer todos los modulos asignados


    Private Sub RellenarComboFrigorifico()
        Try
            Dim CFacCab As New ClsGenerica
            Dim DT As New DataTable
            DT = CFacCab.TraerDatos("GRL_TIPO_OF_TIPO", 1).Tables(0)
            With CmbIdFrigorifico
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idtipo"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    Private Sub RellenarComboJerarquica()
        Try
            Dim CFacCab As New ClsGenerica
            Dim DT As New DataTable
            DT = CFacCab.TraerDatos("GRL_TIPODATO_OF_CODIGO", "jera").Tables(0)
            With CmbJerarquia
                .DataSource = DT
                .DisplayMember = "descripcion"
                .ValueMember = "idtipoDato"
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

#End Region

#Region "Botones"


    Private Sub BtnNuevo_Click(sender As Object, e As EventArgs) Handles BtnNuevo.Click

        Limpiar()
        RellenarLista(USULOGIN)
        ClsFormularios.Habilitar(Me)

    End Sub

    Private Sub BtnGrabar_Click(sender As Object, e As EventArgs) Handles BtnGrabar.Click
        If txtusuario.Text.Trim = "" Then
            MessageBox.Show("Ingrese Usuario", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtusuario.Focus()
            Exit Sub
        End If

        If TxtPassword.Text.Trim = "" Then
            MessageBox.Show("Ingrese Password", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtusuario.Focus()
            Exit Sub
        End If


        If txtnombres.Text.Trim = "" Then
            MessageBox.Show("Ingrese Nombre", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtnombres.Focus()
            Exit Sub
        End If



        If txtapellidos.Text.Trim = "" Then
            MessageBox.Show("Ingrese Apellido", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtnombres.Focus()
            Exit Sub
        End If




        Dim Retorno As Integer
        Dim CUsuarios As New ClsUsuario
        Dim CVendedor As New ClsVendedor
        Try
            If TxtCodigo.Text.Trim = "" Then
                Retorno = CUsuarios.Insertar(0, txtnombres.Text, txtapellidos.Text, txtusuario.Text, TxtPassword.Text.Trim, 0,
                                     CmbIdSuperior.SelectedValue, 0, USULOGIN.Trim, CmbIdFrigorifico.SelectedValue)




                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Agregado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)
                    BuscaRegistro(Retorno)
                    DgUsuarios.Enabled = True
                End If

            Else


                Retorno = CUsuarios.Actualizar(TxtCodigo.Text, txtnombres.Text, txtapellidos.Text, txtusuario.Text, TxtPassword.Text.Trim, 0,
                                     CmbIdSuperior.SelectedValue, USULOGIN.Trim, 0, CmbIdFrigorifico.SelectedValue)
                If Val(Retorno) > 0 Then

                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)
                    BuscaRegistro(Retorno)



                End If
            End If

            ClsFormularios.GrillaPosicionarme(DgUsuarios, Retorno)
            RellenarModuloAsigados(Retorno)
            RellenarModulo(Retorno)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        If TxtCodigo.Text.Trim = "" Then
            MessageBox.Show("Buscar Registro", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtCodigo.Focus()
            Exit Sub
        End If

        If MessageBox.Show("¿Desea Eliminar el Registro Seleccionado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim Retorno As Integer
        Dim CUsuarios As New ClsUsuario
        Try
            Retorno = CUsuarios.Eliminar(TxtCodigo.Text, USULOGIN)
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

    Private Sub BtnModificar_Click(sender As Object, e As EventArgs) Handles BtnModificar.Click
        ClsFormularios.Habilitar(Me)

        BtnEliminar.Enabled = False

    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click

        Call FrmUsuarios_Load(sender, e)
        BtnEliminar.Enabled = True


    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub





#End Region


#Region "Botones Del Modulo De Roles"

    Private Sub BtnModuloAgregar_Click(sender As Object, e As EventArgs) Handles BtnModuloAgregar.Click

        Try

            Dim Omodulo As New ClsGenerica
            Dim ObjModulo(5) As Object


            ObjModulo(1) = ""
            ObjModulo(2) = 0
            ObjModulo(3) = USULOGIN.Trim
            ObjModulo(4) = CmbJerarquia.SelectedValue
            ObjModulo(5) = "A"


            For i As Integer = 0 To Me.DgModulo.Rows.Count - 1
                If DgModulo.Rows(i).Cells("ChkCodigoModulo").Value = True Then
                    ObjModulo(0) = Convert.ToInt32(TxtCodigo.Text)
                    'Agregar Modulos al usuario
                    ObjModulo(1) = Convert.ToInt32(DgModulo.Rows(i).Cells("codigo").Value)

                    Omodulo.EjecutarSP("UL_spIUD_Usuariomodulo", ObjModulo)

                End If
            Next

            RellenarModuloAsigados(TxtCodigo.Text)
            RellenarModulo(TxtCodigo.Text)

            MessageBox.Show("Se Agregó el Modulo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub

    Private Sub BtnModuloEliminar_Click(sender As Object, e As EventArgs) Handles BtnModuloEliminar.Click


        Try
            Dim Omodulo As New ClsGenerica
            Dim ObjModulo(5) As Object


            ObjModulo(1) = ""
            ObjModulo(2) = 0
            ObjModulo(3) = USULOGIN.Trim
            ObjModulo(4) = CmbJerarquia.SelectedValue
            ObjModulo(5) = "D"


            For i As Integer = 0 To Me.DgModuloAsignado.Rows.Count - 1

                If DgModuloAsignado.Rows(i).Cells("ChkCodigoModuloAsignado").Value = True Then
                    ObjModulo(0) = Convert.ToInt32(TxtCodigo.Text)
                    'Agregar Modulos al usuario
                    ObjModulo(1) = Convert.ToInt32(DgModuloAsignado.Rows(i).Cells("codigo").Value)
                    Omodulo.EjecutarSP("UL_spIUD_Usuariomodulo", ObjModulo)


                End If
            Next

            RellenarModuloAsigados(TxtCodigo.Text)
            RellenarModulo(TxtCodigo.Text)


            MessageBox.Show("Se Quitó el Modulo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub


#End Region



    Private Sub chkDgModulo_CheckedChanged(sender As Object, e As EventArgs) Handles chkDgModulo.CheckedChanged

        If chkDgModulo.Checked Then

            For i As Integer = 0 To Me.DgModulo.Rows.Count - 1
                DgModulo.Rows(i).Cells("ChkCodigoModulo").Value = True
            Next

        Else

            For i As Integer = 0 To Me.DgModulo.Rows.Count - 1
                DgModulo.Rows(i).Cells("ChkCodigoModulo").Value = False
            Next


        End If


    End Sub

    Private Sub ChkDgModuloAsignado_CheckedChanged(sender As Object, e As EventArgs) Handles ChkDgModuloAsignado.CheckedChanged


        If ChkDgModuloAsignado.Checked Then


            For i As Integer = 0 To Me.DgModuloAsignado.Rows.Count - 1
                DgModuloAsignado.Rows(i).Cells("ChkCodigoModuloAsignado").Value = True
            Next
        Else

            For i As Integer = 0 To Me.DgModuloAsignado.Rows.Count - 1
                DgModuloAsignado.Rows(i).Cells("ChkCodigoModuloAsignado").Value = False
            Next

        End If


    End Sub


End Class



'Option Explicit On
'Dim IDUsuario As String

'Private Sub chkAdmin_Click()
'    On Error GoTo ErrorGrabar

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub chkMtodo_Click()
'    Dim I As Integer
'    If chkMtodo.Value = 1 Then
'        For I = 1 To lswVend1.ListItems.Count
'            lswVend1.ListItems(I).Checked = True
'        Next
'    Else
'        For I = 1 To lswVend1.ListItems.Count
'            lswVend1.ListItems(I).Checked = False
'        Next
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub chkMUtodo_Click()
'    Dim I As Integer
'    If chkMUtodo.Value = 1 Then
'        For I = 1 To lswVend.ListItems.Count
'            lswVend.ListItems(I).Checked = True
'        Next
'    Else
'        For I = 1 To lswVend.ListItems.Count
'            lswVend.ListItems(I).Checked = False
'        Next
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CmdAgregar_Click()
'    Dim strConnStr As String
'    Dim strReturn As String
'    Dim lngErrNum As Long
'    Dim strErrDesc As String
'    Dim rs As ADODB.Recordset
'    Dim sSql As String
'    On Error GoTo ErrorGrabar

'    If camposRequeridos(Me) Then
'        If valida_clave() Then
'            If TxtModo.Text = "N" Or TxtModo.Text = "O" Then

'                ' sSql = "select orden  From grl_tipodato where idtipodato = " & _
'                CmbJerar.ItemData(CmbJerar.ListIndex)
'                '    If ADO.ExecuteSQL(sSql, , , , , , , lngErrNum, strErrDesc, rs, , 50, True) = True Then
'                '        If Not rs.EOF Then
'                '           If Val(sorden) < Val(rs!orden) Then
'                '            MsgBox "No puede establecer una jerarquía superior a la suya ", vbExclamation, App.Title
'                '            Exit Sub
'                '           End If
'                '        End If
'                '    End If
'                ABM_Usuario()
'            End If
'            If Jerarquia = "B" Then
'                cmdNuevo.Enabled = False
'                CmdBorrar.Enabled = False
'            Else
'                cmdNuevo.Enabled = True
'                CmdBorrar.Enabled = True
'            End If
'            CmdCancelar.Enabled = False
'            CmdAgregar.Enabled = False
'            CmdModificar.Enabled = True
'            CmdCambiar.Enabled = True
'            lswUsu.Enabled = True
'            lswVend.Enabled = True
'            CargaUsuario()
'            lswUsu_Click()
'            seteaCampos(False)
'        End If
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Function valida_clave() As Boolean
'    On Error GoTo ErrorGrabar
'    ReDim strParams(2)
'    ReDim varValues(2)

'    strParams(0) = "Clave"
'    strParams(1) = "IDUsuario"
'    varValues(0) = TxtClave
'    varValues(1) = IIf(TxtIdUsuario = vbNullString, 0, TxtIdUsuario)

'    If ADO.ExecuteStoredProcEx("UL_spC_ClaveUsuario_Valida", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then
'        If Not Return_RecordSet Is Nothing Then
'            If Return_RecordSet.EOF Then
'                MsgBox "No se pudo Validar", vbCritical
'            TxtNombre.SetFocus
'                Exit Function
'            Else
'                If Return_RecordSet!cantidad <> 0 Then
'                    MsgBox "Ya existe un Usuario con este login " & TxtClave, vbInformation
'               valida_clave = False
'                Else
'                    valida_clave = True
'                End If
'            End If
'        End If
'        If Return_RecordSet.State = adStateOpen Then Return_RecordSet.Close
'    End If
'    Exit Function

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Function


'Private Sub CmdBorrarRela_Click()
'    Dim xItem As ListItem
'    Dim chequeado As Integer
'    Dim I As Integer
'    On Error GoTo ErrorGrabar
'    TxtModoRela = "D"
'    chequeado = 0
'    For Each xItem In lswVend.ListItems
'        If xItem.Checked Then
'            ABM_UsuarioVendedor(xItem.Text)
'            chequeado = chequeado + 1
'        End If
'    Next
'    If chequeado = 0 Then
'        MsgBox "Debe marcar el/los modulos a quitar ", vbExclamation, App.Title
'End If

'    '  CmdBorrarRela.Visible = False
'    ' CmdNuevoRela.Visible = False
'    '  chkMtodo.Visible = False

'    CargaModUsuario()
'    lswVend1.Visible = False
'    chkMUtodo.Visible = False
'    Lblmodu.Caption = "Modulos"
'    CmdCambiar.Caption = "Cambiar"
'    CmdNuevoRela.Visible = False
'    CmdBorrarRela.Visible = False
'    lswUsu.Enabled = True
'    lbljerar.Visible = False
'    CmbJerar.Visible = False
'    lswVend1.ListItems.Clear
'    cmdNuevo.Enabled = True
'    CmdModificar.Enabled = True
'    CmdBorrar.Enabled = True

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description

'End Sub

'Private Sub CmdCambiar_Click()
'    On Error GoTo ErrorGrabar
'    If Jerarquia = "A" Then
'        If CmdCambiar.Caption = "Cambiar" Then
'            Lblmodu.Caption = "Modulos para " & TxtNombre.Text & " " & TxtApellido.Text
'            lswUsu.Enabled = False
'            CmdNuevoRela.Visible = True
'            CmdBorrarRela.Visible = True
'            TxtModoRela = "A"
'            lswVend1.Visible = True
'            chkMtodo.Visible = True
'            chkMUtodo.Visible = True
'            CmbJerar.Visible = True
'            lbljerar.Visible = True
'            cmdNuevo.Enabled = False
'            CmdModificar.Enabled = False
'            CmdBorrar.Enabled = False
'            CargaModulosAgregar()
'        End If

'        If CmdCambiar.Caption = "Terminar" Then
'            Lblmodu.Caption = "Modulos"
'            lswVend1.Visible = False
'            chkMtodo.Visible = False
'            CmdNuevoRela.Visible = False
'            CmdBorrarRela.Visible = False
'            lbljerar.Visible = False
'            CmbJerar.Visible = False
'            lswUsu.Enabled = True
'            cmdNuevo.Enabled = True
'            CmdModificar.Enabled = True
'            CmdBorrar.Enabled = True

'        End If
'        If CmdCambiar.Caption = "Cambiar" Then
'            CmdCambiar.Caption = "Terminar"
'        Else
'            CmdCambiar.Caption = "Cambiar"
'        End If
'    Else
'        MsgBox "No tiene privilegios suficientes para esta acción", vbExclamation, App.Title
'      CmdCambiar.Enabled = False
'    End If
'    Exit Sub
'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CmdNuevoRela_Click()
'    Dim xItem As ListItem
'    Dim chequeado As Integer
'    Dim I As Integer
'    Dim s As Boolean

'    On Error GoTo ErrorGrabar

'    TxtModoRela = "A"
'    chequeado = 0
'    For Each xItem In lswVend1.ListItems
'        s = True
'        If xItem.Checked Then
'            ABM_UsuarioVendedor(xItem.Text)
'            chequeado = chequeado + 1
'        End If
'    Next
'    If chequeado = 0 And s Then
'        MsgBox "Debe marcar el/los modulos a insertar ", vbExclamation, App.Title
'End If

'    '  CmdBorrarRela.Visible = False
'    ' CmdNuevoRela.Visible = False
'    '  chkMtodo.Visible = False

'    CargaModUsuario()
'    lswVend1.Visible = False
'    chkMtodo.Visible = False
'    Lblmodu.Caption = "Modulos"
'    CmdCambiar.Caption = "Cambiar"
'    lbljerar.Visible = False
'    CmbJerar.Visible = False
'    CmdNuevoRela.Visible = False
'    CmdBorrarRela.Visible = False
'    lswUsu.Enabled = True
'    lswVend1.ListItems.Clear
'    cmdNuevo.Enabled = True
'    CmdModificar.Enabled = True
'    CmdBorrar.Enabled = True
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub


'Private Sub Command1_Click()
'    If Confirma_Salir Then
'        Unload Me
'End If
'End Sub

''Private Sub Form_Resize()
''On Error GoTo ErrorGrabar
''    Ubicar Me
''Exit Sub
''
''ErrorGrabar:
''    RecuperarError Err.Number, Err.Description
''End Sub
''
''
''Private Sub Form_Unload(Cancel As Integer)
''On Error GoTo ErrorGrabar
''   Unload Me
''    MDIMainCRM.cierraMenu
''Exit Sub
''
''ErrorGrabar:
''    RecuperarError Err.Number, Err.Description
''End Sub
''

'Private Sub Form_Activate()
'    On Error GoTo ErrorGrabar
'    '   MSFlexGridSectores_GotFocus
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Unload Me
'End Sub
'Private Sub Form_Load()
'    'Width = 11325
'    'Height = 6600
'    CentrarForm Me
''Ubicar Me
'    FlagInicio = True
'    FlagInicioAgrega = True
'    Call CargarComboSTKCorral(tipoDatoFrigorifico, cboFrigorificoHC, True)

'    CargaUsuario()

'    lswUsu_Click()

'End Sub







'Private Sub CmdBorrar_Click()
'    On Error GoTo ErrorGrabar
'    Dim I As Integer

'    TxtModo.Text = "D"
'    'CmdCambiar.Enabled = False
'    If (MsgBox("¿Realmente desea eliminar el usuario?", vbYesNo + vbQuestion, App.Title) = vbYes) Then
'        ABM_Usuario()
'    End If

'    CmdCancelar.Enabled = False
'    CmdAgregar.Enabled = False
'    cmdNuevo.Enabled = True
'    CmdBorrar.Enabled = True
'    lswUsu.Enabled = True
'    CargaUsuario()

'    If lswUsu.ListItems.Count > 1 Then
'        lswUsu.ListItems(1).Selected = True
'        lswUsu_Click()
'    End If
'    seteaCampos(False)

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub cmdCancelar_Click()
'    Dim I As Integer

'    On Error GoTo ErrorGrabar
'    If Jerarquia = "B" Then
'        cmdNuevo.Enabled = False
'        CmdBorrar.Enabled = False
'    Else
'        cmdNuevo.Enabled = True
'        CmdBorrar.Enabled = True
'    End If
'    CmdCancelar.Enabled = False
'    CmdAgregar.Enabled = False
'    CmdModificar.Enabled = True
'    CmdCambiar.Enabled = True
'    lswVend.Enabled = True
'    TxtModo.Text = "C"
'    TxtPwd.Enabled = False
'    If Jerarquia <> "B" Then
'        seteaCampos(False)
'    End If
'    TxtApellido.Tag = 0
'    TxtClave.Tag = 0
'    TxtIdUsuario.Tag = 0
'    TxtNombre.Tag = 0
'    lswUsu.Enabled = True
'    IDUsuario = TxtIdUsuario
'    SeleccionaFila()
'    lswUsu_Click()
'    '  For i = 1 To lswUsu.ListItems.Count
'    '    If lswUsu.ListItems(i).Selected Then
'    '       lswUsu.ListItems(i).Selected = False
'    '    End If
'    '  Next
'    '  If lswUsu.ListItems.Count > 1 Then
'    '     lswUsu.ListItems(1).Selected = True
'    '     lswUsu_Click
'    '  End If
'    If Return_RecordSet.State = adStateOpen Then Return_RecordSet.Close
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CmdModificar_Click()
'    On Error GoTo ErrorGrabar
'    cmdNuevo.Enabled = False
'    CmdBorrar.Enabled = False
'    CmdModificar.Enabled = False
'    CmdCambiar.Enabled = False
'    lswVend.Enabled = False
'    lswUsu.Enabled = False
'    TxtModo.Text = "O"
'    seteaCampos(True)
'    Exit Sub
'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CmdNuevo_Click()
'    Dim I As Integer
'    On Error GoTo ErrorGrabar
'    cmdNuevo.Enabled = False
'    CmdBorrar.Enabled = False
'    CmdModificar.Enabled = False
'    CmdCambiar.Enabled = False
'    lswVend.ListItems.Clear
'    lswVend.Enabled = False
'    For I = 1 To lswUsu.ListItems.Count
'        If lswUsu.ListItems(I).Selected Then
'            lswUsu.ListItems(I).Selected = False
'        End If
'    Next
'    lswUsu.Enabled = False
'    TxtModo.Text = "N"
'    seteaCampos(True)
'    BlanqueaCampos()
'    TxtModo.Text = "N"

'    Exit Sub
'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub
'Sub seteaCampos(valor As Boolean)
'    On Error GoTo ErrorGrabar

'    Select Case Jerarquia
'        Case "A", "M"
'            CmdAgregar.Enabled = valor
'            CmdCancelar.Enabled = valor
'            TxtApellido.Enabled = valor
'            TxtClave.Enabled = valor
'            TxtNombre.Enabled = valor
'            TxtPwd.Enabled = valor
'            CmbJerar.Enabled = valor
'            CmbSuperior.Enabled = valor

'            cboFrigorificoHC.Enabled = valor
'        Case "B"
'            CmdModificar.Enabled = True
'            If idLogin = TxtIdUsuario.Text And TxtModo.Text = "O" Then
'                CmdAgregar.Enabled = valor
'                CmdCancelar.Enabled = valor
'                TxtPwd.Enabled = valor
'            Else
'                MsgBox "No tiene permitido efectuar modificaciones ", vbExclamation, App.Title
'                lswUsu.Enabled = True
'            End If
'    End Select

'    '        If Admin = 0 And Sistemas = 0 And (TxtModo.Text = "N" Or TxtModo.Text = "C" Or TxtModo.Text = vbNullString) Then
'    '            TxtApellido.Enabled = valor
'    '            TxtClave.Enabled = valor
'    '            TxtNombre.Enabled = valor
'    '            TxtPwd.Enabled = valor
'    '        Else
'    '           If Admin = 0 And Sistemas = 0 And TxtModo.Text = "O" And chkSistemas.Value <> 1 And chkAdmin.Value <> 1 Then
'    '              TxtApellido.Enabled = valor
'    '              TxtNombre.Enabled = valor
'    '              TxtClave.Enabled = valor
'    '              TxtPwd.Enabled = valor
'    '           Else
'    '             If Admin = 0 And Sistemas = 0 And TxtModo.Text = "O" Then
'    '               TxtApellido.Enabled = valor
'    '               TxtNombre.Enabled = valor
'    '             End If
'    '           End If
'    '        End If
'    '        If Admin = 1 And (TxtModo.Text = "O" Or TxtModo.Text = "C") Then
'    '            If chkSistemas.Value = 1 Then
'    '               chkAdmin.Enabled = valor
'    '               chkObs.Enabled = valor
'    '               TxtApellido.Enabled = valor
'    '               TxtNombre.Enabled = valor
'    '            Else
'    '               chkAdmin.Enabled = valor
'    '               chkObs.Enabled = valor
'    '               TxtApellido.Enabled = valor
'    '               TxtClave.Enabled = valor
'    '               TxtNombre.Enabled = valor
'    '               TxtPwd.Enabled = valor
'    '            End If
'    '        End If
'    '        If Admin = 1 And (TxtModo.Text = "N" Or TxtModo.Text = vbNullString) Then
'    '           chkAdmin.Enabled = valor
'    '           chkObs.Enabled = valor
'    '           TxtApellido.Enabled = valor
'    '           TxtClave.Enabled = valor
'    '           TxtNombre.Enabled = valor
'    '           TxtPwd.Enabled = valor
'    '        End If
'    '        If Sistemas = 1 Then
'    '            chkSistemas.Enabled = valor
'    '            chkAdmin.Enabled = valor
'    '            chkObs.Enabled = valor
'    '            TxtApellido.Enabled = valor
'    '            TxtClave.Enabled = valor
'    '            TxtNombre.Enabled = valor
'    '            TxtPwd.Enabled = valor
'    '        End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub
'Sub BlanqueaCampos()
'    On Error GoTo ErrorGrabar
'    TxtApellido.Text = vbNullString
'    TxtClave.Text = vbNullString
'    TxtIdUsuario.Text = vbNullString
'    TxtNombre.Text = vbNullString
'    TxtPwd.Text = vbNullString
'    TxtModo.Text = vbNullString
'    Chkactivo.Value = 1
'    CmbSuperior.ListIndex = -1
'    CmbJerar.ListIndex = -1
'    'CargaModUsuario
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub
'Private Sub ABM_Usuario()
'    On Error GoTo ErrorGrabar

'    ReDim strParams(10)
'    ReDim varValues(10)


'    strParams(0) = "IDUsuario"
'    strParams(1) = "Nombre"
'    strParams(2) = "Apellido"
'    strParams(3) = "login"
'    strParams(4) = "Pwd"
'    strParams(5) = "jera"
'    strParams(6) = "superior"
'    strParams(7) = "usuario"
'    strParams(8) = "modo"
'    strParams(9) = "idFrigorifico"


'    varValues(0) = IIf(TxtIdUsuario = vbNullString, 0, TxtIdUsuario)
'    varValues(1) = TxtNombre
'    varValues(2) = TxtApellido
'    varValues(3) = TxtClave
'    varValues(4) = TxtPwd
'    If CmbJerar.ListIndex = -1 Then
'        varValues(5) = 0
'    Else
'        varValues(5) = CmbJerar.ItemData(CmbJerar.ListIndex)
'    End If
'    'varValues(5) = IIf(CmbJerar.ListIndex = -1, 0, CmbJerar.ItemData(CmbJerar.ListIndex))
'    If CmbSuperior.ListIndex = -1 Then
'        varValues(6) = 0
'    Else
'        varValues(6) = CmbSuperior.ItemData(CmbSuperior.ListIndex)
'    End If
'    ' varValues(6) = IIf(CmbSuperior.ListIndex = -1, 0, CmbSuperior.ItemData(CmbSuperior.ListIndex))
'    varValues(7) = Usuario
'    varValues(8) = TxtModo
'    varValues(9) = cboFrigorificoHC.ItemData(cboFrigorificoHC.ListIndex)

'    If ADO.ExecuteStoredProcEx("UL_spIUD_Usuarios", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then
'        If Not Return_RecordSet Is Nothing Then
'            If Return_RecordSet.EOF Then
'                If TxtModo.Text = "N" Then MsgBox "No se pudo Insertar el Usuario", vbCritical
'            If TxtModo.Text = "O" Then MsgBox "No se pudo Modificar el Usuario", vbCritical
'            If TxtModo.Text = "D" Then MsgBox "No se pudo Eliminar el Usuario", vbCritical
'            TxtNombre.SetFocus
'                Exit Sub
'            Else
'                TxtIdUsuario = Return_RecordSet!IDUsuario
'            End If
'        End If
'        If Return_RecordSet.State = adStateOpen Then Return_RecordSet.Close
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CargaUsuario()
'    On Error GoTo ErrorGrabar

'    Dim SP As String
'    Dim I As Integer
'    Dim mcolum As MSComctlLib.ListItem
'    Dim mItem As MSComctlLib.ListItem
'    Dim sSql As String
'    Dim rs As ADODB.Recordset
'    Dim strConnStr As String
'    Dim strReturn As String
'    Dim lngErrNum As Long
'    Dim strErrDesc As String
'    IDUsuario = TxtIdUsuario
'    If Jerarquia <> "B" Then
'        cmdNuevo.Enabled = True
'        CmdBorrar.Enabled = True
'    End If

'    ReDim strParams(0) As String
'ReDim varValues(0) As Variant

'strParams(0) = ""
'    varValues(0) = ""
'    If Jerarquia = "A" Then
'        TxtPwd.Locked = False
'    End If

'    Screen.MousePointer = vbHourglass

'    If ADO.ExecuteStoredProcEx("UL_spC_Usuario", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then

'        If Return_RecordSet.EOF Or Return_RecordSet.BOF Then
'            MsgBox "No hay Usuarios Cargados ", 48
'         Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        '/Columnas del listview de Supervisor
'        If Jerarquia <> "A" Then
'            TxtPwd.PasswordChar = "*"
'        Else
'            TxtPwd.PasswordChar = "*"
'        End If
'        With lswUsu
'            .OLEDragMode = ccOLEDragAutomatic
'            .Sorted = True
'            .ColumnHeaders.Clear
'            .ColumnHeaders.Add 1, "ID", " ", 0
'    .ColumnHeaders.Add 2, "Nombre", "Nombre", 1500
'    .ColumnHeaders.Add 3, "APELLIDO", "Apellido", 1500
'    .ColumnHeaders.Add 4, "LOGIN", "Login", 1500
'    .ColumnHeaders.Add 5, "", "", 0
'    .ColumnHeaders.Add 6, "", "", 0
'    .ColumnHeaders.Add 7, "SUPERIOR", "Superior", 1500, lvwColumnRight
'    .ColumnHeaders.Add 8, "FECHABAJA", "Fecha Baja", 1500
'    .ColumnHeaders.Add 9, "", "", 0
'    .ColumnHeaders.Add 10, "", "", 0
'    .ColumnHeaders.Add 11, "", "", 0

''   If FlagInicio And jerarquia = "A" Then
'            '       .ColumnHeaders.Add 7, "PASSWORD", "Password", 800
'            '   End If

'            '    Set mItem = .Add(, , IIf(IsNull(rsExp!FechaDeCarga) = False, rsExp!FechaDeCarga, "0"), , iAtraso)
'            '    With mItem.ListSubItems

'            CmbSuperior.Clear
'            CmbSuperior.AddItem("No asignado ")
'            CmbSuperior.ItemData(CmbSuperior.NewIndex) = 0

'            With .ListItems
'                .Clear
'                '/Cargo los items del listview
'                Do Until Return_RecordSet.EOF
'        Set mItem = .Add()
'        mItem.Text = Return_RecordSet!IDUsuario
'                    mItem.SubItems(1) = RTrim(Return_RecordSet!nombre)
'                    mItem.SubItems(2) = "" & RTrim(Return_RecordSet!apellido)
'                    mItem.SubItems(3) = "" & RTrim(Return_RecordSet!login) 'Format("" & Return_RecordSet!FechaIngreso, cFormatoFecha)
'                    mItem.SubItems(4) = ""
'                    mItem.SubItems(5) = ""
'                    mItem.SubItems(6) = "" & IIf(IsNull(Return_RecordSet!superior) = True, " ", Return_RecordSet!superior)
'                    mItem.SubItems(7) = "" & IIf(IsNull(Return_RecordSet!bajalogica) = True, " ", Return_RecordSet!bajalogica)
'                    mItem.SubItems(8) = "" & RTrim(Return_RecordSet!Password)
'                    mItem.SubItems(9) = IIf(IsNull(Return_RecordSet!idsuperior) = True, 0, Return_RecordSet!idsuperior)
'                    'If idLogin <> Return_RecordSet!IDUsuario Then
'                    CmbSuperior.AddItem(RTrim(Return_RecordSet!apellido) & " " & RTrim(Return_RecordSet!nombre))
'                    CmbSuperior.ItemData(CmbSuperior.NewIndex) = Return_RecordSet!IDUsuario
'                    'End If
'                    mItem.SubItems(10) = IIf(IsNull(Return_RecordSet!idfrigorifico) = True, 0, Return_RecordSet!idfrigorifico)

'                    Return_RecordSet.MoveNext
'                Loop
'                CmbSuperior.ListIndex = 0
'            End With
'        End With

'        sSql = "select idtipodato,rtrim(tipodato)+' '+rtrim(descripcion)as " &
'        "descripcion, orden From grl_tipodato where codigo ='jera' "

'        If ADO.ExecuteSQL(sSql, , , , , , , lngErrNum, strErrDesc, rs, , 50, True) = True Then

'            CmbJerar.Clear
'            Do Until rs.EOF
'                CmbJerar.AddItem(rs!descripcion)
'                CmbJerar.ItemData(CmbJerar.NewIndex) = rs!idTipoDato
'                rs.MoveNext
'            Loop

'            With StsVendedor
'                .Panels(1).Text = "Total Usuarios : " + CStr(Return_RecordSet.RecordCount)
'            End With

'            If Not FlagInicio Then
'                SeleccionaFila()
'            End If
'        End If
'        Screen.MousePointer = vbNormal

'    End If

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub SeleccionaFila()
'    On Error GoTo ErrorGrabar
'    Dim I As Integer

'    If IDUsuario <> vbNullString Then
'        For I = 1 To lswUsu.ListItems.Count
'            lswUsu.ListItems(I).Selected = False
'            If lswUsu.ListItems.Item(I).Text = IDUsuario Then
'                lswUsu.ListItems(I).Selected = True
'            End If
'        Next
'    Else
'        If lswUsu.ListItems.Count > 1 Then
'            lswUsu.ListItems(1).Selected = True
'        End If
'    End If
'    IDUsuario = vbNullString
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub lswUsu_Click()
'    Dim s As Integer
'    On Error GoTo ErrorGrabar

'    Dim I As Integer

'    For I = 1 To lswUsu.ListItems.Count
'        If lswUsu.ListItems(I).Selected Then
'            TxtIdUsuario.Text = lswUsu.ListItems.Item(I).Text
'            TxtNombre.Text = lswUsu.ListItems.Item(I).ListSubItems(1).Text
'            TxtApellido.Text = lswUsu.ListItems.Item(I).ListSubItems(2).Text
'            TxtClave.Text = lswUsu.ListItems.Item(I).ListSubItems(3).Text
'            TxtPwd.Text = lswUsu.ListItems.Item(I).ListSubItems(8).Text
'            cboFrigorificoHC.ListIndex = vBuscarenItemData(cboFrigorificoHC, lswUsu.ListItems.Item(I).ListSubItems(10).Text)
'            If Jerarquia <> "A" Then
'                If idLogin = TxtIdUsuario.Text Then
'                    TxtPwd.PasswordChar = ""
'                    TxtPwd.Locked = False
'                Else
'                    TxtPwd.PasswordChar = "*"
'                    TxtPwd.Locked = True
'                End If
'            End If
'            '            For s = 0 To CmbJerar.ListCount - 1
'            '                If CmbJerar.ItemData(s) = lswUsu.ListItems.Item(i).ListSubItems(4).Text Then
'            CmbJerar.ListIndex = s
'            '                End If
'            '            Next
'            For s = 0 To CmbSuperior.ListCount - 1
'                If CmbSuperior.ItemData(s) = lswUsu.ListItems.Item(I).ListSubItems(9).Text Then
'                    CmbSuperior.ListIndex = s
'                End If
'            Next
'            Chkactivo.Value = IIf(lswUsu.ListItems.Item(I).ListSubItems(7).Text = " ", 1, 0)
'            CargaModUsuario()
'        End If
'    Next
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub lswUsu_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    OrdenaLv lswUsu, ColumnHeader
'  '  Ordenar
'End Sub
'Private Sub lswVend_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    OrdenaLv lswVend, ColumnHeader
'End Sub

'Private Sub lswVend_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim xItem As ListItem
'    Dim chequeado As Integer

'    chequeado = 0
'    For Each xItem In lswVend.ListItems
'        If xItem.Checked Then
'            chequeado = chequeado + 1
'        End If
'    Next
'    If chequeado > 0 Then
'        CmdNuevoRela.Enabled = False
'    Else
'        CmdNuevoRela.Enabled = True
'    End If
'End Sub

'Private Sub lswVend1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    OrdenaLv lswVend1, ColumnHeader
'End Sub
'Private Sub CargaModUsuario()
'    On Error GoTo ErrorGrabar

'    Dim SP As String
'    Dim I As Integer
'    Dim s As Integer
'    Dim mcolum As MSComctlLib.ListItem
'    Dim mItem As MSComctlLib.ListItem


'    ReDim strParams(1) As String
'ReDim varValues(1) As Variant

'valida_modifVend()
'    strParams(1) = "IDUsuario"
'    varValues(1) = IIf(TxtIdUsuario.Text = vbNullString, 0, TxtIdUsuario.Text)

'    Screen.MousePointer = vbHourglass

'    If ADO.ExecuteStoredProcEx("UL_spC_ModuloXUsuario", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then

'        If Return_RecordSet.EOF Or Return_RecordSet.BOF Then
'            'MsgBox "No hay Vendedores Cargados para este Usuario", 48
'            lswVend.ListItems.Clear
'            With StsVendedor
'                .Panels(2).Text = "Total modulos por Usuario : " + "0"
'            End With
'            Screen.MousePointer = vbDefault
'            With lswVend
'                .MultiSelect = True
'                .View = lvwReport                    'Ansicht: Details
'                .Gridlines = True
'                If Jerarquia <> "B" Then
'                    .CheckBoxes = True
'                Else
'                    .CheckBoxes = False
'                End If
'                .OLEDropMode = ccOLEDropManual
'                .Sorted = True
'                .Refresh
'                If FlagInicio Then
'                    .ColumnHeaders.Add 1, "CODIGO", "Codigo", 1000
'              .ColumnHeaders.Add 2, "NOMBRE", "Nombre", 1500
'              .ColumnHeaders.Add 3, "ACCESO", "Acceso", 2000
'              .ColumnHeaders.Add 4, "USUARIO", "Usuario", 0
'              FlagInicio = False
'                End If
'            End With
'            Exit Sub
'        End If
'        '/Columnas del listview de Modulo

'        With lswVend
'            .MultiSelect = True
'            .View = lvwReport                    'Ansicht: Details
'            .Gridlines = True
'            '    If jerarquia <> "B" Then
'            .CheckBoxes = True
'            '   Else
'            '     .Checkboxes = False
'            '  End If
'            .OLEDropMode = ccOLEDropManual
'            .Sorted = True
'            .Refresh

'            If FlagInicio Then
'                .ColumnHeaders.Add 1, "CODIGO", "Codigo", 1000
'      .ColumnHeaders.Add 2, "NOMBRE", "Nombre", 1500
'      .ColumnHeaders.Add 3, "ACCESO", "Acceso", 2000
'      .ColumnHeaders.Add 4, "USUARIO", "Usuario", 0
'      FlagInicio = False
'            End If
'            With .ListItems
'                .Clear
'                '/Cargo los items del listview
'                Do Until Return_RecordSet.EOF
'        Set mItem = .Add()
'        mItem.Text = " " & Return_RecordSet!codigo
'                    mItem.SubItems(1) = RTrim(Return_RecordSet!nombre)

'                    For s = 0 To CmbJerar.ListCount - 1
'                        If CmbJerar.ItemData(s) = Return_RecordSet!Jerarquia Then
'                            CmbJerar.ListIndex = s
'                        End If
'                    Next
'                    mItem.SubItems(2) = CmbJerar.Text 'Return_RecordSet!IDUsuario
'                    mItem.SubItems(3) = Return_RecordSet!IDUsuario
'                    Return_RecordSet.MoveNext
'                Loop
'            End With
'        End With


'        With StsVendedor
'            .Panels(2).Text = "Total modulos que accede : " + CStr(Return_RecordSet.RecordCount)
'        End With


'        Screen.MousePointer = vbNormal
'    End If

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub lswVend1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    'Dim xItem As ListItem
'    Dim xItem As MSComctlLib.ListItem
'    Dim chequeado As Integer
'    chequeado = Item.Index
'    For Each xItem In lswVend1.ListItems

'        If xItem.Checked Then
'            If xItem.Index <> chequeado Then
'                xItem.Checked = False
'                xItem.ForeColor = vbBlack
'            Else
'                CmbJerar.Enabled = True
'            End If
'        End If
'    Next

'End Sub

'Private Sub TxtClave_KeyPress(KeyAscii As Integer)

'    Formu_KeyPress Me, KeyAscii

'End Sub
'Private Sub lswVend_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
'    imgDyd1.Visible = True
'End Sub
'Private Sub ABM_UsuarioVendedor(idModulo As String)
'    On Error GoTo ErrorGrabar

'    ReDim strParams(4)
'    ReDim varValues(4)

'    strParams(0) = "IDUsuario"
'    strParams(1) = "Idmodulo"
'    strParams(2) = "Usuario"
'    strParams(3) = "idJerarquia"
'    strParams(4) = "Modo"


'    varValues(0) = IIf(TxtIdUsuario = vbNullString, 0, TxtIdUsuario)
'    varValues(1) = idModulo
'    varValues(2) = Usuario
'    varValues(3) = CmbJerar.ItemData(CmbJerar.ListIndex)
'    varValues(4) = TxtModoRela

'    If ADO.ExecuteStoredProcEx("UL_spIUD_Usuariomodulo", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then
'        If Not Return_RecordSet Is Nothing Then
'            If Return_RecordSet.EOF Then
'                If TxtModo.Text = "N" Then MsgBox "No se pudo Insertar el Modulo", vbCritical
'            If TxtModo.Text = "O" Then MsgBox "No se pudo Modificar el Modulo", vbCritical
'            If TxtModo.Text = "D" Then MsgBox "No se pudo Eliminar el Modulo", vbCritical
'            'TxtNombre.SetFocus
'                Exit Sub
'            Else
'                'If TxtModoRela.Text = "N" Then MsgBox "Se Insertó un Supervisor", vbInformation
'                'If TxtModoRela.Text = "O" Then MsgBox "Se Modificó el Supervisor", vbInformation
'                'If TxtModoRela.Text = "D" Then MsgBox "Se Eliminó el Supervisor", vbInformation
'                'TxtIDUsuario = Return_RecordSet!IDUsuario
'            End If
'        End If
'        If Return_RecordSet.State = adStateOpen Then Return_RecordSet.Close
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub CargaModulosAgregar()
'    On Error GoTo ErrorGrabar

'    Dim SP As String
'    Dim I As Integer
'    Dim mcolum As MSComctlLib.ListItem
'    Dim mItem As MSComctlLib.ListItem


'    ReDim strParams(2) As String
'ReDim varValues(2) As Variant


'strParams(1) = "IDUsuario"
'    strParams(2) = "Modo"

'    varValues(1) = IIf(TxtIdUsuario.Text = vbNullString, 0, TxtIdUsuario.Text)
'    varValues(2) = TxtModoRela

'    Screen.MousePointer = vbHourglass

'    If ADO.ExecuteStoredProcEx("UL_spC_ModuloXUsuario", strParams(), varValues(), True, , , , , , , , Return_RecordSet, False) = True Then

'        If Return_RecordSet.EOF Or Return_RecordSet.BOF Then
'            lswVend1.Visible = False
'            CmdBorrarRela.Visible = True
'            CmdNuevoRela.Visible = True
'            chkMtodo.Visible = False
'            If sbande Then
'                MsgBox "No hay Modulos para adicionarle a este Usuario", 48
'        End If
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        '/Columnas del listview de Supervisor

'        With lswVend1
'            .MultiSelect = True
'            .View = lvwReport
'            .Gridlines = True
'            .CheckBoxes = True
'            .OLEDragMode = ccOLEDragAutomatic
'            .Sorted = True
'            .Refresh


'            If FlagInicioAgrega Then
'                .ColumnHeaders.Add 1, "COD", "Codigo", 1000
'      .ColumnHeaders.Add 2, "NOMBRE", "Nombre", 1500
'      .ColumnHeaders.Add 3, "IDUSUARIO", "idUsuario", 0
'      FlagInicioAgrega = False
'            End If
'            With .ListItems
'                .Clear
'                '/Cargo los items del listview
'                Do Until Return_RecordSet.EOF
'        Set mItem = .Add()
'        mItem.Text = Return_RecordSet!codigo
'                    mItem.SubItems(1) = RTrim(Return_RecordSet!nombre)
'                    '        mItem.SubItems(2) = RTrim(Return_RecordSet!Nombre)
'                    '        mItem.SubItems(3) = "" & Return_RecordSet!idVendedor 'Format("" & Return_RecordSet!FechaFinPaso, cFormatoFecha)
'                    Return_RecordSet.MoveNext
'                Loop
'            End With
'        End With


'        Screen.MousePointer = vbNormal
'    End If
'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub
'Private Sub lswVend_OLEDragDrop(Data As MSComctlLib.DataObject,
'Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'    Dim myItems() As String
'    Dim myItemsData As String
'    Dim mySubItems() As String
'    Dim objItem As ListItem
'    Dim I As Integer
'    Dim j As Integer

'    If Data.GetFormat(vbCFText) Then
'        Text1.Text = Data.GetData(vbCFText)
'    End If

'    If Data.GetFormat(vbCFText) Then
'        myItemsData = Data.GetData(vbCFText)
'        myItems = Split(myItemsData, vbCrLf, -1, vbBinaryCompare)
'        For I = LBound(myItems) To UBound(myItems)
'            mySubItems = Split(myItems(I), vbTab, -1, vbBinaryCompare)
'            Set objItem = lswVend.ListItems.Add(, , mySubItems(0))
'            For j = LBound(mySubItems) + 1 To UBound(mySubItems)
'                objItem.ListSubItems.Add j, , mySubItems(j)
'            Next j
'        Next
'        lswVend.Refresh
'        ABMVendUsuario()
'    End If
'End Sub



'Private Sub lswVend1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
'    ' Beim klick auf ein Listview1 eintrag diesen "nehmen"
'    ' oder bei Mehreren diese durch Enter ("vbCrLf") trennen

'    Dim I As Integer
'    Dim j As Integer
'    Dim zw As String

'    For I = 1 To lswVend1.ListItems.Count
'        If lswVend1.ListItems(I).Selected Then
'            If Len(zw) <> 0 Then
'                zw = zw & vbCrLf
'            End If
'            zw = zw & lswVend1.ListItems(I).Text
'            For j = 1 To lswVend1.ListItems(I).ListSubItems.Count
'                zw = zw & Chr(9) & lswVend1.ListItems(I).ListSubItems(j).Text
'            Next j
'        End If
'    Next

'    Data.SetData zw
'    DataFormat = vbCFText

'End Sub
'Private Sub ABMVendUsuario()
'    On Error GoTo ErrorGrabar
'    Dim I As Integer
'    TxtModoRela.Text = "N"
'    For I = 1 To lswVend.ListItems.Count
'        If lswVend.ListItems.Item(I).Text = 0 Then
'            TxtIdVendedor.Text = lswVend.ListItems.Item(I).ListSubItems(3).Text
'            ABM_UsuarioVendedor(TxtIdVendedor.Text)
'        End If
'    Next
'    CargaModUsuario()
'    '    If lswVend.ListItems.Count > 1 Then
'    '         lswVend.ListItems(1).Selected = True
'    '    End If
'    lswVend1.Visible = False
'    CmdBorrarRela.Visible = True
'    CmdNuevoRela.Visible = True
'    chkMtodo.Visible = True

'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

'Private Sub valida_modifVend()
'    On Error GoTo ErrorGrabar

'    CmdNuevoRela.Enabled = True
'    CmdBorrarRela.Enabled = True
'    chkMtodo.Enabled = True


'    Exit Sub

'ErrorGrabar:
'    RecuperarError Err.Number, Err.Description
'End Sub

