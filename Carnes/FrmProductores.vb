
Imports System.Runtime.InteropServices
Public Class FrmProductores


    Private currentButton As Button
    Private random As Random
    Private tempIndex As Integer


    Dim Modo As String
    Dim OProductor(47) As Object

    Public Sub New()
        InitializeComponent()
        random = New Random()
        currentButton = New Button()

        Me.Text = String.Empty
        Me.ControlBox = False
        Me.MaximizedBounds = Screen.FromHandle(Me.Handle).WorkingArea
        'Logger.Nivel = Logger.TipoMensaje.WARNING

    End Sub

    Private Sub FrmProveedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RellenarLista()
        ComboPais()
        ComboProvincia()
        ComboLocalidad()
        CargarCondIva()
        CargarTipoProveedor()




    End Sub



#Region "Funciones "

    Private Sub OProductorCargar()


        Try

            '***************************************

            OProductor(0) = Val(TxtidProveedor.Text) '@idproveedor int = null ,
            OProductor(1) = Val(TxtProvNro.Text)   '@provNro int = NULL ,
            OProductor(2) = TxtRazonSocial.Text '@razonSocial varchar (100) = null ,
            OProductor(3) = TxtCuit.Text  '@CUIT varchar (20) = NULL ,
            OProductor(4) = Val(CmbcondIva.SelectedValue)  '@condIva int = NULL ,
            OProductor(5) = TxtCalle.Text '@calle varchar (100) = NULL ,
            OProductor(6) = TxtNumero.Text '@numero varchar (5) = NULL ,
            OProductor(7) = Txtpiso.Text '@piso varchar (5) = NULL ,
            OProductor(8) = Val(CmbLocalidad.SelectedValue) '@idLocalidad int = NULL ,
            OProductor(9) = Val(CmbProvincia.SelectedValue) '@idprovincia int = NULL ,



            '@idproveedor int = null ,
            '@provNro int = NULL ,
            '@razonSocial varchar (100) = null ,
            '@CUIT varchar (20) = NULL ,
            '@condIva int = NULL ,
            '@calle varchar (100) = NULL ,
            '@numero varchar (5) = NULL ,
            '@piso varchar (5) = NULL ,
            '@idLocalidad int = NULL ,
            '@idprovincia int = NULL ,

            '***************************************

            OProductor(10) = Val(CmbPais.SelectedValue) '@idPais int = NULL ,
            OProductor(11) = txtTelPart.Text '@telPart varchar (50) = null ,
            OProductor(12) = TxtTelFax.Text '@telfax varchar (50) = null ,
            OProductor(13) = TxtTelCel.Text '@telCel varchar (50) = null ,
            OProductor(14) = TxtCPostal.Text '@cpostal varchar(12)=null,
            OProductor(15) = txtRsComercial.Text '@RSComercial varchar(100) = null,
            OProductor(16) = txtDomicom.Text '@DomiciCom varchar (100) = NULL ,
            OProductor(17) = Val(cmbIdLocaCom.Text) '@idLocaCom int = NULL ,
            OProductor(18) = txtTelCom.Text '@TelCom varchar (50) = null ,
            OProductor(19) = TxtTelFaxCom.Text '@TelFaxCom varchar (50) = null ,



            '@idPais int = NULL ,
            '@telPart varchar (50) = null ,
            '@telfax varchar (50) = null ,
            '@telCel varchar (50) = null ,
            '@cpostal varchar(12)=null,
            '@RSComercial varchar(100) = null,
            '@DomiciCom varchar (100) = NULL ,
            '@idLocaCom int = NULL ,
            '@TelCom varchar (50) = null ,
            '@TelFaxCom varchar (50) = null ,

            '***************************************

            OProductor(20) = txtCPostalCom.Text '@cPostalCom varchar(12)=null,
            OProductor(21) = Convert.ToDecimal(IIf(TxtComision.Text = vbNullString, 0, TxtComision.Text)) '@Comision decimal(15,2)=null,
            OProductor(22) = Convert.ToDecimal(IIf(TxtImpSellos.Text = vbNullString, 0, TxtImpSellos.Text)) '@ImpSellos decimal(15,2)=null,
            OProductor(23) = Convert.ToDecimal(IIf(txtDerRegistro.Text = vbNullString, 0, txtDerRegistro.Text)) '@DerRegistro decimal(15,2)=null,
            OProductor(24) = Convert.ToDecimal(IIf(TxtMbs.Text = vbNullString, 0, TxtMbs.Text)) '@mbs decimal(15,2)=null,
            OProductor(25) = Convert.ToDecimal(IIf(TxtRec.Text = vbNullString, 0, TxtRec.Text)) '@Rec decimal(15,2)=null,
            OProductor(26) = Convert.ToDecimal(IIf(TxtComision2.Text = vbNullString, 0, TxtComision2.Text)) '@Comision2 decimal(15,2)=null,
            OProductor(27) = Convert.ToDecimal(IIf(TxtFGarantia.Text = vbNullString, 0, TxtFGarantia.Text)) '@FGarantia decimal(15,2)=null,
            OProductor(28) = Convert.ToDecimal(IIf(TxtGastoFrigor.Text = vbNullString, 0, TxtGastoFrigor.Text)) '@GastoFrigor decimal(15,2)=null,
            OProductor(29) = Convert.ToDecimal(IIf(txtGuia.Text = vbNullString, 0, txtGuia.Text)) '@Guia decimal(15,2)=null, 




            '@cPostalCom varchar(12)=null,
            '@Comision decimal(15,2)=null,
            '@ImpSellos decimal(15,2)=null,
            '@DerRegistro decimal(15,2)=null,
            '@mbs decimal(15,2)=null,
            '@Rec decimal(15,2)=null,
            '@Comision2 decimal(15,2)=null,
            '@FGarantia decimal(15,2)=null,
            '@GastoFrigor decimal(15,2)=null,
            '@Guia decimal(15,2)=null,

            '***************************************





            OProductor(30) = Convert.ToDecimal(IIf(txtotros.Text = vbNullString, 0, txtotros.Text)) '@Otros decimal(15,2)=null,
            OProductor(31) = Convert.ToDecimal(IIf(txtFlete.Text = vbNullString, 0, txtFlete.Text)) '@Flete decimal(15,2)=null,
            OProductor(32) = If(ChkIngBrutos.Checked, "S", "N")   '@IngBrutos char(1)=null,
            OProductor(33) = TxtProcedencia.Text '@Procedencia varchar(100)=null,
            OProductor(34) = If(ChkGanancias.Checked, "S", "N") '@Ganancias char(1)=null,
            OProductor(35) = Convert.ToDecimal(IIf(TxtIBPorcen.Text = vbNullString, 0, TxtIBPorcen.Text))   '@IBPorcen decimal(9,3)=null,'@IngBrutos char(1)=null,
            OProductor(36) = TxtObservacion1.Text '@observacion1 nvarchar (50) = NULL ,
            OProductor(37) = TxtObservacion2.Text '@observacion2 nvarchar (50) = NULL ,
            OProductor(38) = IIf(TxtEmail1.Text = vbNullString, " ", TxtEmail1.Text) '@email1 varchar (50) = NULL ,
            OProductor(39) = IIf(TxtEmail2.Text = vbNullString, " ", TxtEmail2.Text) '@email2 varchar (50) = NULL ,


            '@Otros decimal(15,2)=null,
            '@Flete decimal(15,2)=null,
            '@IngBrutos char(1)=null,
            '@Procedencia varchar(100)=null,
            '@Ganancias char(1)=null,
            '@IBPorcen decimal(9,3)=null,
            '@observacion1 nvarchar (50) = NULL ,
            '@observacion2 nvarchar (50) = NULL ,
            '@email1 varchar (50) = NULL ,        
            '@email2 varchar (50) = NULL ,





            OProductor(40) = "123" '@usuario varchar(20) ,
            OProductor(41) = "" '@Modo char(1),
            OProductor(42) = Val(If(chktodosReg.Checked, "1", "0")) '@todosReg int=null, 
            OProductor(43) = Val(If(ChktodosSell.Checked, "1", "0")) '@todosSell int=null,
            OProductor(44) = TxtRuca.Text '@ruca varchar(50)=null,
            OProductor(45) = txtRenspa.Text '@renspa varchar(50)=null,
            OProductor(46) = Val(CmbTipoProveedor.SelectedValue) '@tipoCaracter int=3,
            OProductor(47) = Val(TxtPuntoVenta.Text)  '@puntoVenta int=16

            ''@usuario varchar(20) ,
            '@Modo char(1),
            '@todosReg int=null,
            '@todosSell int=null,
            '@ruca varchar(50)=null,
            '@renspa varchar(50)=null,
            '@tipoCaracter int=3,
            '@puntoVenta int=16

            '***************************************

        Catch ex As Exception
            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub



    Private Sub RellenarLista()
        Try

            Dim OProductor(0) As Object
            OProductor(0) = ""
            ' OProductor(1) = ""
            ' OProductor(2) = ""


            Dim CPProveedores As New ClsGenerica
            Dim DT As New DataTable
            DT = CPProveedores.TraerDatos("Gen_Proveedor_Consul", 0, "", "").Tables(0)
            DgProveedores.DataSource = DT
            DgProveedores.ReadOnly = True
            'DgProveedores.Columns("idUsuario").Visible = False
            'DgProveedores.Columns("idfrigorifico").Visible = False
            'DgProveedores.Columns("idSuperior").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

#End Region
#Region "Grilla Proveedores"


    Private Sub DgProveedores_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgProveedores.CellClick

        Try


            Dim Fila As Integer
            If DgProveedores.Rows(DgProveedores.CurrentRow.Index).Cells(0).Value.ToString <> "" Then

                Fila = DgProveedores.CurrentRow.Index

                BuscarProveedor(DgProveedores.Rows(Fila).Cells(0).Value.ToString())


                ClsFormularios.Bloquear(Me)
                ClsFormularios.Bloquear(GroupBoxContacto)
                ClsFormularios.Bloquear(GroupBoxDomLegal)
                ClsFormularios.Bloquear(GroupBoxDatos)


            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Busqueda", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub


#End Region



#Region "Rutinas de los Botones "



    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click

        OProductorCargar()

        If TxtidProveedor.Text.Trim = "" Then
            MessageBox.Show("Buscar Proveedor", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtidProveedor.Focus()
            Exit Sub
        End If

        If MessageBox.Show("¿Desea Eliminar el Registro Seleccionado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If





        Dim Retorno As Integer
        Dim Cproveedor As New ClsGenerica





        Try

            OProductor(41) = "B"
            Retorno = Cproveedor.EjecutarSP("Gen_proveedor_Abm", OProductor)

            If Val(Retorno) > 0 Then
                MessageBox.Show("Registro Eliminado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call BtnNuevo_Click(sender, e)
            Else
                MessageBox.Show("No se pudo Realizar la Operación", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            OProductor(0) = ""
            LimpiarFormulario()
            RellenarLista()


        Catch ex As Exception
            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try





    End Sub

    Private Sub BtnGrabar_Click(sender As Object, e As EventArgs) Handles BtnGrabar.Click

        If TxtProvNro.Text.Trim = "" Then
            MessageBox.Show("Ingrese Proveedor", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtProvNro.Focus()
            Exit Sub
        End If

        If TxtRazonSocial.Text.Trim = "" Then
            MessageBox.Show("Ingrese Razon Social", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtRazonSocial.Focus()
            Exit Sub
        End If


        OProductorCargar()


        Dim Retorno As Integer
        Dim Cproveedor As New ClsGenerica



        If VerificarProveedor(OProductor(1), OProductor(3)) Then

            Modo = "M"

        Else

            Modo = "A"

        End If
        Try
            If TxtidProveedor.Text.Trim = "" And Modo = "A" Then



                OProductor(41) = Modo

                Retorno = Cproveedor.EjecutarSP("Gen_proveedor_in", OProductor)


                If Val(Retorno) > 0 Then
                    MessageBox.Show("Registro Agregado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Call BtnNuevo_Click(sender, e)


                End If
            Else


                OProductor(41) = Modo
                Retorno = Cproveedor.EjecutarSP("Gen_proveedor_Abm", OProductor)

                If Val(Retorno) > 0 Then

                    MessageBox.Show("Registro Actualizado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)


                    Call BtnNuevo_Click(sender, e)

                End If
            End If

            ClsFormularios.GrillaPosicionarme(DgProveedores, Retorno)
            BuscarProveedor(Retorno)

        Catch ex As Exception

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

    End Sub

    Private Function VerificarProveedor(ByVal ProvNro As Integer, ByVal Cuit As String) As Boolean

        Dim retorno As Boolean = False

        Try

            Dim Odatatable As New DataTable

            Dim Cproveedor As New ClsGenerica
            Dim strsql As String = "SELECT TOP 1 idProveedor FROM GRL_PROVEEDOR WHERE CUIT ='" & Cuit & "'"
            strsql = strsql & " Or provNro =" & ProvNro

            Odatatable = Cproveedor.EjecutarSql(strsql).Tables(0)

            If Odatatable.Rows.Count > 0 Then

                OProductor(0) = Odatatable.Rows(0)("idProveedor")

                retorno = True

            Else
                retorno = False

            End If

        Catch ex As Exception
            retorno = False

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Verificar Proveedor", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try



        Return retorno


    End Function


    Private Sub BtnNuevo_Click(sender As Object, e As EventArgs) Handles BtnNuevo.Click

        'OPCION NUEVO SE LIMPIA EL FORMULARIO Y SE 
        'COLOCA OPCION MODO ="A"  QUE ES EL ALTA

        Modo = "A"
        '  LimpiarFormulario()
        RellenarLista()
        GroupBoxGrilla.Enabled = True
        DgProveedores.Enabled = True

        ClsFormularios.Limpiar(Me)
        ClsFormularios.Limpiar(GroupBoxContacto)
        ClsFormularios.Limpiar(GroupBoxDomLegal)
        ClsFormularios.Limpiar(GroupBoxDatos)






    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()

    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click

        Dim OProductor(2) As Object

        OProductor(0) = 0

        If RdbRazonSocial.Checked = True Then
            OProductor(1) = txtBuscarRazonSocial.Text
        Else
            OProductor(1) = ""
        End If


        If RdbCuit.Checked = True Then

            OProductor(2) = TxtBuscarCuit.Text
        Else
            OProductor(2) = ""

        End If

        Try



            Dim CPProveedores As New ClsGenerica
            Dim DT As New DataTable
            DT = CPProveedores.TraerDatos("Gen_Proveedor_Consul", OProductor).Tables(0)
            DgProveedores.DataSource = DT
            DgProveedores.ReadOnly = True

            LimpiarFormulario()


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Busqueda", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub



    Private Sub BtnModificar_Click(sender As Object, e As EventArgs) Handles BtnModificar.Click

        'OPCION NUEVO SE LIMPIA EL FORMULARIO Y SE 
        'COLOCA OPCION MODO ="A"  QUE ES EL ALTA

        Modo = "M"

        ClsFormularios.Habilitar(Me)
        ClsFormularios.Habilitar(GroupBoxContacto)
        ClsFormularios.Habilitar(GroupBoxDomLegal)
        ClsFormularios.Habilitar(GroupBoxDatos)




    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click

        GroupBoxGrilla.Enabled = True
        DgProveedores.Enabled = True


    End Sub



#End Region


#Region "Funciones y Rutinas para el Formulario"

    Private Sub BuscarProveedor(Codigo As Integer)
        Try
            Dim CProveedor As New ClsGenerica
            Dim DT As New DataTable
            DT = CProveedor.TraerDatos("Gen_Proveedor_Consul", Codigo, "", "").Tables(0)


            If DT.Rows.Count > 0 Then

                TxtidProveedor.Text = DT.Rows(0).Item("idProveedor").ToString()
                TxtProvNro.Text = DT.Rows(0).Item("ProvNro").ToString()
                TxtRazonSocial.Text = DT.Rows(0).Item("RazonSocial").ToString()
                TxtProvNro.Text = DT.Rows(0).Item("ProvNro").ToString()
                TxtCuit.Text = DT.Rows(0).Item("CUIT").ToString()
                TxtProcedencia.Text = DT.Rows(0).Item("procedencia").ToString()
                CmbcondIva.SelectedValue = DT.Rows(0).Item("condIva").ToString()
                CmbTipoProveedor.SelectedValue = DT.Rows(0).Item("tipocaracter").ToString()
                TxtProcedencia.Text = DT.Rows(0).Item("procedencia").ToString()

                ' falta el chequed de ganancias 
                ' falta el chequed de Ingresos Brutos 




                TxtCalle.Text = DT.Rows(0).Item("calle").ToString()
                TxtNumero.Text = DT.Rows(0).Item("numero").ToString()
                Txtpiso.Text = DT.Rows(0).Item("piso").ToString()
                TxtCPostal.Text = DT.Rows(0).Item("CPostal").ToString()
                CmbPais.SelectedValue = DT.Rows(0).Item("idpais").ToString()
                CmbProvincia.SelectedValue = DT.Rows(0).Item("IdProvincia").ToString()
                CmbLocalidad.SelectedValue = DT.Rows(0).Item("idlocalidad").ToString()
                txtTelPart.Text = DT.Rows(0).Item("TelPart").ToString()
                TxtTelFax.Text = DT.Rows(0).Item("Telfax").ToString()
                TxtTelCel.Text = DT.Rows(0).Item("Telcel").ToString()

                ' PARTE COMERCIAL
                txtRsComercial.Text = DT.Rows(0).Item("RsComercial").ToString()
                txtDomicom.Text = DT.Rows(0).Item("Domicicom").ToString()
                cmbIdLocaCom.SelectedValue = DT.Rows(0).Item("IdLocaCom").ToString()

                txtTelCom.Text = DT.Rows(0).Item("TelCom").ToString()
                TxtTelFaxCom.Text = DT.Rows(0).Item("TelFaxCom").ToString()
                txtCPostalCom.Text = DT.Rows(0).Item("CPostalCom").ToString()

                TxtObservacion1.Text = DT.Rows(0).Item("Observacion1").ToString()
                TxtObservacion2.Text = DT.Rows(0).Item("Observacion2").ToString()
                TxtEmail1.Text = DT.Rows(0).Item("Email1").ToString()
                TxtEmail2.Text = DT.Rows(0).Item("Email2").ToString()
                TxtRuca.Text = DT.Rows(0).Item("Ruca").ToString()


                TxtComision.Text = DT.Rows(0).Item("comision").ToString()
                TxtComision2.Text = DT.Rows(0).Item("comision2").ToString()
                TxtImpSellos.Text = DT.Rows(0).Item("ImpSellos").ToString()
                txtDerRegistro.Text = DT.Rows(0).Item("DerRegistro").ToString()
                TxtMbs.Text = DT.Rows(0).Item("Mbs").ToString()
                TxtRec.Text = DT.Rows(0).Item("rec").ToString()
                TxtFGarantia.Text = DT.Rows(0).Item("FGarantia").ToString()
                TxtGastoFrigor.Text = DT.Rows(0).Item("GastoFrigor").ToString()
                txtGuia.Text = DT.Rows(0).Item("guia").ToString()
                txtotros.Text = DT.Rows(0).Item("otros").ToString()
                txtFlete.Text = DT.Rows(0).Item("flete").ToString()
                TxtIBPorcen.Text = DT.Rows(0).Item("IBPorcen").ToString()




            Else
                MessageBox.Show("No se encontró ningun Registros", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)


            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub




#Region "CARGAR COMBOS"
    ' Cargar Combo Localidad
    Private Sub ComboLocalidad()
        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.EjecutarSql("select idlocalidad, nombre From grl_localidad").Tables(0)
            With CmbLocalidad
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idlocalidad"
            End With

            With cmbIdLocaCom
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idlocalidad"
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Localidad", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' Cargar Combo Provincia
    Private Sub ComboProvincia()
        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.EjecutarSql("select idprovincia, nombre From grl_provincia").Tables(0)
            With CmbProvincia
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
            With CmbPais
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idPais"
            End With
        Catch ex As Exception


            ' Logger.e("Error con excepción", ex)
            'Logger.d("Debug con traza", New StackFrame(True))
            ' Logger.i("Info sin traza", New StackFrame(True))
            ' Logger.e("Error con excepción y traza", ex, New StackFrame(True))
            'Logger.e(ex.Message, ex)
            'Logger.w(ex.Message,)
            'Logger.e(ex.Message, ex)
            'Logger.e(ex.Message, ex)

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso- Pais", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    Private Sub CargarCondIva()


        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("condiva_ot").Tables(0)
            With CmbcondIva
                .DataSource = DT
                .DisplayMember = "Descripcion"
                .ValueMember = "IdcondIva"
            End With
        Catch ex As Exception


            ' Logger.e("Error con excepción", ex)
            'Logger.d("Debug con traza", New StackFrame(True))
            ' Logger.i("Info sin traza", New StackFrame(True))
            ' Logger.e("Error con excepción y traza", ex, New StackFrame(True))
            'Logger.e(ex.Message, ex)
            'Logger.w(ex.Message,)
            'Logger.e(ex.Message, ex)
            'Logger.e(ex.Message, ex)

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso- Pais", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub CargarTipoProveedor()


        Try
            Dim CLocalidad As New ClsGenerica
            Dim DT As New DataTable
            DT = CLocalidad.TraerDatos("ProveedorTipo_Ot").Tables(0)
            With CmbTipoProveedor
                .DataSource = DT
                .DisplayMember = "Descripcion"
                .ValueMember = "IdProveedorTipo"
            End With
        Catch ex As Exception


            ' Logger.e("Error con excepción", ex)
            'Logger.d("Debug con traza", New StackFrame(True))
            ' Logger.i("Info sin traza", New StackFrame(True))
            ' Logger.e("Error con excepción y traza", ex, New StackFrame(True))
            'Logger.e(ex.Message, ex)
            'Logger.w(ex.Message,)
            'Logger.e(ex.Message, ex)
            'Logger.e(ex.Message, ex)

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso- Pais", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

#End Region


    Private Sub LimpiarFormulario()


        TxtidProveedor.Clear()
        TxtProvNro.Clear()
        TxtRazonSocial.Clear()
        TxtCuit.Clear()
        CmbcondIva.SelectedIndex = 0
        TxtCalle.Clear()
        TxtNumero.Clear()
        Txtpiso.Clear()
        CmbLocalidad.SelectedIndex = 0
        CmbPais.SelectedIndex = 0
        CmbProvincia.SelectedIndex = 0
        txtTelPart.Clear()
        TxtTelFax.Clear()
        TxtTelCel.Clear()
        TxtCPostal.Clear()
        TxtEmail1.Clear()
        TxtEmail2.Clear()







        'txtusuario.Clear()
        'TxtPassword.Clear()
        'TxtCodigo.Clear()
        'txtnombres.Clear()
        'txtapellidos.Clear()


        'txtusuario.Focus()
    End Sub

    Private Sub TxtProvNro_TextChanged(sender As Object, e As EventArgs) Handles TxtProvNro.TextChanged

    End Sub



#End Region


End Class