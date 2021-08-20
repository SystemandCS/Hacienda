Public Class ClsCombos


    Public Shared Sub CargarComboSTKCorral(ByVal TipoDato As String, ByVal combo As ComboBox, sinEspecificar As Boolean)
        Try


            '
            If (combo Is Nothing) Then
                Throw New ArgumentNullException("combo")
            End If

            If (String.IsNullOrEmpty(TipoDato)) Then
                Throw New ArgumentException("No se ha especificado el tipo de Dato")
            End If



            Dim CFrigorifico As New ClsGenerica
            Dim DT As New DataTable
            DT = CFrigorifico.TraerDatos("GRL_GET_TIPOS_FOR_COMBOS", TipoDato).Tables(0)

            With combo
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idtipo"
            End With


        Catch ex As Exception

            Throw New ApplicationException("No se pudo cargar el combo" + TipoDato)

        End Try


    End Sub



    Public Shared Sub CargarComboProvincia(ByVal combo As ComboBox, sinEspecificar As Boolean)
        Try


            '
            If (combo Is Nothing) Then
                Throw New ArgumentNullException("combo")
            End If





            Dim CFrigorifico As New ClsGenerica
            Dim DT As New DataTable
            DT = CFrigorifico.TraerDatos("GRL_PROVINCIA_OL").Tables(0)

            With combo
                .DataSource = DT
                .DisplayMember = "Nombre"
                .ValueMember = "idprovincia"
            End With


        Catch ex As Exception

            Throw New ApplicationException("No se pudo cargar el combo Provincia")

        End Try


    End Sub


    Public Shared Sub CargarComboProveedor(combo As ComboBox, sinEspecificar As Boolean)

        Try


            '
            If (combo Is Nothing) Then
                Throw New ArgumentNullException("combo")
            End If





            Dim CFrigorifico As New ClsGenerica
            Dim DT As New DataTable
            DT = CFrigorifico.TraerDatos("GRL_PROVEEDOR_OL").Tables(0)

            With combo
                .DataSource = DT
                .DisplayMember = "razonsocial"
                .ValueMember = "idproveedor"
            End With


        Catch ex As Exception

            Throw New ApplicationException("No se pudo cargar el combo " & combo.Name)

        End Try



    End Sub



    Public Shared Function vBuscarCombo(ByVal Combo As Object, ByVal texto As String, ByVal UsarItemData As Boolean) As Integer
        Dim i As Integer
        On Error GoTo ErrTextoCombo
        If Combo.ListCount > 0 Then
            For i = 0 To Combo.ListCount - 1
                If UsarItemData Then
                    If Combo.ItemData(i) = Val(texto) Then Exit For
                Else
                    If Combo.List(i) = texto Then Exit For
                End If
            Next i
        Else
            i = -1
        End If
        If i >= Combo.ListCount Then i = -1
        vBuscarCombo = i
        Exit Function
ErrTextoCombo:
        vBuscarCombo = -1
    End Function


End Class
