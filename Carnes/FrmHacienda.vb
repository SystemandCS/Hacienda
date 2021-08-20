

Public Class FrmHacienda



    '' definicio de varialbles


    Dim intStock As Integer
    Dim idMRDET As Integer
    Dim esAlta As Boolean
    Dim altaSTKCorralFueManual
    Dim idMR As Long
    Dim idsc As Long
    Dim intIdCorte As Integer
    Dim intIdFrigoCuart As Integer
    Dim idFrigoDespos As Integer
    Dim blnEsCabeceraCuarteo As Boolean
    Const cTexto As String = "Texto"
    Const cNumero As String = "Número"
    Const cMoneda As String = "Moneda"
    Const cFecha As String = "Fecha"
    ' Los formatos a aplicar según el tipo de datos que contiene
    Const cFormatoFecha As String = "dd/mm/yyyy"
    Const cFormatoNumero As String = "###,##0"      ' "###"
    Const cFormatoMoneda As String = "#####0.00"   ' "###.00"
    ' La cantidad de cifras a tener en cuenta en los números
    Const cCuantasCifras As Long = 20&







#Region "Funciones y Rutinas para el Formulario"

    Private Sub _CargaHCorral()

        Try
            Dim OHacienda(4) As Object
            OHacienda(0) = ""

            Dim CCorral As New ClsHacienda
            Dim DT As New DataTable
            DT = CCorral.Filtrado("HCorral", OHacienda, "").Tables(0)





        Catch ex As Exception
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


#End Region

#Region "Funciones y Rutinas para el Formulario"


    Private Sub verDatosTropa(ByVal intIdSC As Integer)

        Dim OHacienda(4) As Object
        OHacienda(0) = ""

        Try



            Dim CStock As New ClsStock
            Dim DT As New DataTable
        DT = CStock.Filtrado("GET_CORRAL", intIdSC).Tables(0)

            If DT.Rows.Count > 0 Then


            Else

                MessageBox.Show("Error inesperado en verdatosTropa. No se encontró el registro a mostrar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            End If

        Catch ex As Exception
        MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try




    End Sub


#End Region





#Region "Load del Formulario"



    Private Sub FrmHacienda_Load(sender As Object, e As EventArgs) Handles Me.Load


        Try


            intCantCabezas = 0

            '  DTFechaDesdeCorral.Value = DateAdd(Day, -30, Date.Now).ToString

            ' DTfechaHastaCorral = Dim CurrentDateTime As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            '  DTFechaDesdeFaena = Date - 30
            '
            'DTFechaHAstaFaena = Date

            ' DTFechaDesdeCuarteo = Date - 30

            ' DTFechaHastaCuarteo = Date



            'SSTab1.TabVisible(3) = True
            'SSTab1.Tab = 0
            'intFichaactiva = 1
            'If MDIHacienda.mnuStock.Tag <> "A" Then
            '    FrameCabeceraC.Enabled = False
            '    'FrameDespoCab.Enabled = False
            '    'FrameDesposDet.Enabled = False
            '    Framedetalle.Enabled = False
            '    FrmHCDetalle(0).Enabled = False
            '    FrmHCDetalle(1).Enabled = False
            'End If



            If intidUsrFrigorifico <> 0 Then
                'SSTab1.TabVisible(3) = False

            End If
            ' ClsCombos.CargarComboSTKCorral(tipoDatoCorte, cmbCortes, False)

            'Call CargarComboSTKCorral(tipoDatoCorte, cmbCorteDespo, False)

            'Call CargarComboSTKCorral("CORTEDEPOS", cmbCorteDesposProd, False)
            Call cargaDatosparaABMSTKCorral()


            ClsCombos.CargarComboProveedor(cboProveedor, True)
            ClsCombos.CargarComboProveedor(cboConsignatario, True)
            ClsCombos.CargarComboProveedor(cboTitular, True)

            cboFiltroStock.SelectedIndex = -1

            '  MSFlexGrid1.Visible = True
            ' MSFlexGrid1.Enabled = True
            'lblseparadesp.Visible = True
            'DTfechasepDesp.Visible = True
            ' cboTitular.SelectedIndex = ClsCombos.vBuscarCombo(cboTitular, "MALEFU AGROPECUARIA S.R.L. 33-69757684-9", False)
            ' cboConsignatario.SelectedIndex = ClsCombos.vBuscarCombo(cboConsignatario, "URIEN LOZA SA 30-69215795-4", False)

            ' ClsCombos.CargarComboProvincia(idgrlprovinciasdest, False)

            idgrlprovinciasdest.Visible = True
            idgrllocalidaddest.Visible = True
            idgrlprovinciasdest.SelectedIndex = ClsCombos.vBuscarCombo(idgrlprovinciasdest, UCase("Buenos Aires"), False)
            idgrllocalidaddest.SelectedIndex = ClsCombos.vBuscarCombo(idgrllocalidaddest, UCase("San Fernando"), False)
            txtDestinoF.Text = "San Fernando"








        Catch ex As Exception
            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


    End Sub










#End Region






    Private Sub cmdFiltrarHC_Click(sender As Object, e As EventArgs) Handles cmdFiltrarHC.Click

        'If FrmHCDetalle(0).Enabled = True Then

        Dim idStockFiltro As Integer = cboFiltroStock.SelectedValue
        Dim idFrigFiltro As Integer = cboFrigorificoFil.SelectedValue


        'With StsFaenaTot

        '    .Panels(1).Text = "Total Romaneos: 0"
        '    .Panels(2).Text = "Total Cabezas faenadas: 0"
        '    .Panels(3).Text = "Total Kg faenados: 0"

        'End With

        Call inicializarControlesHC()

        CargarGrilla(2)

        'If intidUsrFrigorifico = 0 Then
        '    '= idFrigFiltro
        '    cboFrigorificoFil.ListIndex = vBuscarenItemData(cboFrigorificoFil, idFrigFiltro)
        'Else
        '    cboFrigorificoFil.ListIndex = 0
        'End If



    End Sub


    Sub inicializarControlesHC()
        esAlta = True

        TxtFechaMovHC.Value = Date.Now()
        cboFrigorificoHC.SelectedIndex = -1
        cboCategoriaHC.SelectedIndex = -1
        cboRazaHC.SelectedIndex = -1

        txtProveedorHC.Text = ""

        ' TxtHiltonHC = ""

        TxtDTAHC.Text = ""
        TxtGuiaHC.Text = ""

        TxtTropaHC.Text = ""
        TxtCantHC.Text = ""
        TxtCantCaidasHC.Text = ""
        TxtCantMuertasHC.Text = ""

        txtIdSCHC.Text = ""
        TxtCantRemanenteHC.Text = ""
        TxtKgTotalHC.Text = ""
        TxtKgPromedioHC.Text = ""
        TxtObservacionHC.Text = ""
        chkCorralCompleto.Checked = False
        altaSTKCorralFueManual = True
        ChkComisionista.Checked = 0

    End Sub





    Private Sub CargarGrilla(tipo As Integer)


        'Dim servidor As SvrDatosADO.ServidorADO
        Dim arrID() As VariantType

        Dim CantRojo As Integer
        Dim CantAmarillo As Integer
        Dim CantVerde As Integer
        Dim CantCorral As Integer
        Dim CantFaenacompleta As Integer
        Dim CantKgFaenado As Double
        Dim CantRecibida As Integer

        'Eze
        altaSTKCorralFueManual = False

        'ReDim arrID(1, 2)
        ReDim arrID(1)
        CantRecibida = 0
        CantFaenacompleta = 0
        CantKgFaenado = 0
        CantCorral = 0

        ' lswfaena.ListItems.Clear

        CantRojo = 0
        CantAmarillo = 0
        CantVerde = 0




        'ListView1.Columns.Add("Emp Name", 100, HorizontalAlignment.Left)
        'ListView1.Columns.Add("Emp Address", 150, HorizontalAlignment.Left)
        'ListView1.Columns.Add("Title", 60, HorizontalAlignment.Left)
        'ListView1.Columns.Add("Salary", 50, HorizontalAlignment.Left)
        'ListView1.Columns.Add("Department", 60, HorizontalAlignment.Left)



        '    ListView1.Columns.Add("", 250, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Tropa", 800, HorizontalAlignment.Right)
        '    ListView1.Columns.Add("Fecha de Ingreso", 1300, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Frigorifico", 1300, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Proveedor", 1300, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Cant. Vivas", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Cant. Muertas", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Stock", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Raza", 1000, HorizontalAlignment.Left)

        '    ListView1.Columns.Add("Categoria", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Kilos Vivos", 1000, HorizontalAlignment.Left)

        '    ListView1.Columns.Add("Raza", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Cant. Muertas", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Kilos Faenados", 1000, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Rendimiento", 1000, HorizontalAlignment.Left)

        '    ListView1.Columns.Add("Provincia", 1800, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("Departamento", 1800, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("idproveedor", 0, HorizontalAlignment.Left)
        '    ListView1.Columns.Add("comisionista", 0, HorizontalAlignment.Left)


        '    ListViewMakeColumnHeaders(ListView1,
        '"Title", 200, HorizontalAlignment.Left,
        '"URL", 100, HorizontalAlignment.Left,
        '"ISBN", 100, HorizontalAlignment.Left,
        '"Picture", 200, HorizontalAlignment.Left,
        '"Pages", 50, HorizontalAlignment.Right,
        '"Year", 50, HorizontalAlignment.Right)





        Dim Filtro As String = ""


        Select Case cboFiltroStock.SelectedValue
            Case 2
                Filtro = "STK_CORRALES_OF_SIN_STOCK"
            Case 1
                Filtro = "STK_CORRALES_OF_CON_STOCK"
            Case 0
                Filtro = "STK_CORRALES_OF_TODOS"
        End Select


        Dim CDatos As New ClsGenerica
        Dim DT As New DataTable
        DT = CDatos.TraerDatos(Filtro).Tables(0)


        If DT.Rows.Count > 0 Then

            DgStk.DataSource = DT





            For Each fila As DataGridViewRow In DgStk.Rows

                If fila.Cells("cantidad_stock").Value = 0 Then
                    fila.DefaultCellStyle.BackColor = Color.Red

                Else
                    fila.DefaultCellStyle.BackColor = Color.Green
                End If

            Next






            'For Each row As DataRow In DT.Rows


            '    ' DgStk.Rows.Add(False, Color.Green, row.Item("tropanro"), row.Item("ingreso"))

            '    If row("cantidad_stock") = 0 Then


            '        DgStk.Columns(1).DefaultCellStyle.BackColor = Color.Red

            '    Else

            '        DgStk.Columns(1).DefaultCellStyle.BackColor = Color.Green


            '    End If


            'Next

        Else

            DgStk.DataSource = Nothing



        End If







        ' RellenarListview(DT)

        'For i = 0 To DT.Rows.Count - 1

        '    If CInt(DT.Rows(i).Item("cantidad_stock") = 0) Then

        '        CantRojo = CantRojo + 1
        '        ' .SmallIcon = 1
        '    Else

        '        CantVerde = CantVerde + 1



        '    End If
        '    '                        End If
        '    '.SmallIcon = 3

        '    CantVerde = CantVerde + 1


        '    Dim Item As ListViewItem
        '    Item = New ListViewItem(DT.Rows(i).Item("ingreso").ToString)
        '    Item.SubItems.Add(DT.Rows(i).Item("cuitproveedor").ToString)

        '    ListView1.Items.Add(Item)




        'Next




        'End If



        '        oDatos.Args = arrID
        'Set oRecord = oDatos.Leer_SP_Sin_Arg
        'a = 0
        '        Do While Not oRecord.EOF


        '            If (idFrigFiltro = 0 Or idFrigFiltro = oRecord("idFrigorifico")) And CDate(oRecord(1)) >= DTFechaDesdeCorral And CDate(oRecord(1)) <= DTfechaHastaCorral Then

        '                With ListView1(tipo).ListItems.Add(, , IIf(oRecord("cantidad_stock") <= 0, 0, 1))

        '                    If CInt(oRecord("cantidad_stock")) = 0 Then

        '                        CantRojo = CantRojo + 1
        '                        .SmallIcon = 1
        '                    Else
        '                        'If CInt(oRecord("cantidadstock")) = CInt(oRecord("cantidad")) - CInt(IIf(IsNull(oRecord("cantidadmuertas")), 0, oRecord("cantidadmuertas"))) Then

        '                        'CantRojo = CantRojo + 1
        '                        '.SmallIcon = 1
        '                        'Else
        '                        CantVerde = CantVerde + 1
        '                        .SmallIcon = 3
        '                    End If
        '                    '                        End If
        '                    '.SmallIcon = 3

        '                    CantVerde = CantVerde + 1
        '                    ' Cada subitem debe corresponder con cada una de las cabeceras
        '                    ' la segunda cabecera es el Subitems(1) y así sucesivamente

        '                    .SubItems(1) = IIf(IsNull(oRecord(2)) Or oRecord(2) = 0, "0", Format(oRecord(2), cFormatoNumero))
        '                    .SubItems(2) = IIf(IsNull(oRecord(1)), "", oRecord(1))
        '                    .SubItems(3) = IIf(IsNull(oRecord("Frigorifico")), "", oRecord("Frigorifico"))
        '                    .SubItems(4) = IIf(IsNull(oRecord(3)), "", oRecord(3))

        '                    .SubItems(5) = IIf(IsNull(oRecord(4)), "", oRecord(4))
        '                    .SubItems(6) = IIf(IsNull(oRecord(5)), "", oRecord(5))
        '                    .SubItems(7) = IIf(IsNull(oRecord(11)), "", oRecord(11))
        '                    .SubItems(8) = IIf(IsNull(oRecord(6)), "", oRecord(6))
        '                    .SubItems(9) = IIf(IsNull(oRecord(7)), "", oRecord(7))

        '                    .SubItems(10) = IIf(IsNull(oRecord("kgsTotales")), "0", Format(oRecord("kgsTotales"), cFormatoNumero))
        '                    .SubItems(11) = IIf(IsNull(oRecord("kgsFaenados")), "0", Format(oRecord("kgsFaenados"), cFormatoNumero))
        '                    .SubItems(12) = IIf(IsNull(oRecord("rendimiento")), "0", Format(oRecord("rendimiento"), cFormatoMoneda))

        '                    .SubItems(13) = IIf(IsNull(oRecord(8)), "", oRecord(8))
        '                    .SubItems(14) = IIf(IsNull(oRecord(9)), "", oRecord(9))
        '                    .SubItems(15) = IIf(IsNull(oRecord(10)), "", oRecord(10))
        '                    .SubItems(16) = IIf(IsNull(oRecord("idproveedor")), "", oRecord("idproveedor"))
        '                    .SubItems(17) = IIf(IsNull(oRecord("idcorral")), "0", oRecord("idcorral"))

        '                    ' Si quieres probar con números con decimales
        '                    '.SubItems(2)
        '                    .Tag = oRecord(0)
        '                End With
        '                If oRecord("cantidad_stock") = 0 Then
        '                    CantFaenacompleta = CantFaenacompleta + 1
        '                End If
        '                CantRecibida = CantRecibida + 1
        '                CantKgFaenado = CantKgFaenado + IIf(IsNull(oRecord("kgsFaenados")), 0, oRecord("kgsFaenados"))
        '            End If
        '            oRecord.MoveNext

        '        Loop
        '        With stBarCorr

        '            .Panels(1).Text = "Total Entradas en Corral: " + CStr(CantRecibida)
        '            .Panels(2).Text = "Compras con faena completa: " + CStr(CantFaenacompleta)
        '            .Panels(3).Text = "Total KG faenados " + CStr(CantKgFaenado)

        '        End With

        '        If ListView1(tipo).ListItems.Count > 0 Then
        '            ListView1(tipo).HoverSelection = False
        '            ListView1(tipo).ListItems(1).Selected = True
        '        End If

        '        Call ListView1_Click(tipo)

        'For a = 0 To ListView1(tipo).ListItems(0).SubItems(1)

        'Set oRecord = oDatos.Leer_SP_Sin_Arg
        Exit Sub
    End Sub




    Private Sub DgStk_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgStk.CellFormatting


        Try
            'If DgStk.Columns(e.ColumnIndex).Name = "cantidad_stock" Then
            '    Dim _filaDGV As DataGridViewRow = DgStk.Rows(e.RowIndex)
            '    If CStr(_filaDGV.Cells("cantidad_stock").Value) = 0 Then

            '        _filaDGV.DefaultCellStyle.ForeColor = Color.Red

            '        _filaDGV.Cells(0).

            '    Else
            '        _filaDGV.DefaultCellStyle.ForeColor = Color.Green


            '    End If
            'End If



            If DgStk.Columns(e.ColumnIndex).Name <> "cantidad_stock" Then

                Dim cell As DataGridViewCell = DgStk.Rows(e.RowIndex).Cells("cantidad_stock")

                If CStr(cell.Value) = 0 Then

                    e.CellStyle.ForeColor = Color.Red
                End If

                If CStr(cell.Value) > 0 Then

                    e.CellStyle.ForeColor = Color.Green

                End If


            End If



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub


    Sub cargaDatosparaABMSTKCorral()


        ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cboFrigorificoHC, False)
        ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cboFrgoF, False)
        ClsCombos.CargarComboSTKCorral(tipoDatoCategoria, cboCategoriaHC, True)
        ClsCombos.CargarComboSTKCorral(tipoDatoRaza, cboRazaHC, True)
        ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cboFrigorificoFil, True)
        ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cbofrigorificoFil2, True)
        ' ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cboFrigorificoFilCuarteos, True)
        'ClsCombos.CargarComboSTKCorral(tipoDatoFrigorifico, cboFrigorificoFilCuarteos1, True)
        ' Call CargarComboSTKCorral(tipoDatoFrigorifico, cmbFilDesp, True)

        'cboFrigorificoFilCuarteos.ListIndex = 0
        ' cbofrigorificoFil2.ListIndex = 0
        ' cbostockfil2.ListIndex = 0


    End Sub

    Private Sub cboProveedor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboProveedor.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles catAfip.SelectedIndexChanged

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub




    'Function RellenarListview(ByVal dt As DataTable) As Boolean
    '    Dim bolResultado As Boolean = True
    '    Dim lstElemento As ListViewItem
    '    Try
    '        ' Me.ListView1.Items.Clear()
    '        ' Me.ListView1.Columns.Clear()
    '        For Each col As DataColumn In dt.Columns
    '            ListView1.Columns.Add(col.ColumnName, col.ColumnName)
    '        Next
    '        For Each row As DataRow In dt.Rows
    '            lstElemento = New ListViewItem
    '            lstElemento.Text = row("INGRESO").ToString()

    '            For intcontador As Integer = 1 To dt.Columns.Count - 1
    '                lstElemento.SubItems.Add(row(intcontador).ToString())
    '            Next

    '            ListView1.Items.Add(lstElemento)
    '        Next
    '        Me.ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
    '    Catch ex As Exception
    '        bolResultado = False
    '    End Try
    '    Return bolResultado
    'End Function

End Class