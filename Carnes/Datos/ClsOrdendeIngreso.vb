Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class ClsOrdendeIngreso

    Dim cnx As New SqlConnection(ConfigurationManager.ConnectionStrings("cn").ConnectionString)

    Public Function InsertaOrdendeIngreso(ORDSERIE As String, PROCOD As String,
    IMPUESTO As Double, OBSERVACION As String, NROCOTIZACION As String, USUREG As String) As String
        Dim retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_INSERTA_ORDENDEINGRESO", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDSERIE", ORDSERIE)
                .Parameters.AddWithValue("@PROCOD", PROCOD)
                .Parameters.AddWithValue("@IMPUESTO", IMPUESTO)
                .Parameters.AddWithValue("@OBSERVACION", OBSERVACION)
                .Parameters.AddWithValue("@NROCOTIZACION", NROCOTIZACION)
                .Parameters.AddWithValue("@USUREG", USUREG)
                retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return retorno
    End Function



    Public Function ActualizaOrdendeIngreso(ORDINGCOD As String, PROCOD As String, OBSERVACION As String,
    NROCOTIZACION As String, USUMOD As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ACTUALIZA_ORDENDEINGRESO", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Parameters.AddWithValue("@PROCOD", PROCOD)
                .Parameters.AddWithValue("@OBSERVACION", OBSERVACION)
                .Parameters.AddWithValue("@NROCOTIZACION", NROCOTIZACION)
                .Parameters.AddWithValue("@USUMOD", USUMOD)
                Retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Retorno
    End Function


    Public Function EliminaOrdendeIngreso(ORDINGCOD As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ELIMINA_ORDENDEINGRESO", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                Retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Retorno
    End Function

    Public Function ListaOrdendeIngreso(ORDINGCOD As String, ORDSERIE As String, ORDCORRELATIVO As String,
    PROCOD As String, DESDE As Date, HASTA As Date) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_LISTA_ORDENDEINGRESO", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .SelectCommand.Parameters.AddWithValue("@ORDSERIE", ORDSERIE)
                .SelectCommand.Parameters.AddWithValue("@ORDCORRELATIVO", ORDCORRELATIVO)
                .SelectCommand.Parameters.AddWithValue("@PROCOD", PROCOD)
                .SelectCommand.Parameters.AddWithValue("@DESDE", DESDE)
                .SelectCommand.Parameters.AddWithValue("@HASTA", HASTA)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function


    Public Function BuscaOrdendeIngreso(ORDINGCOD As String) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_BUSCA_ORDENDEINGRESO", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function

    Public Function SerieOrdendeIngreso() As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_SERIE_ORDENDEINGRESO", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function

    Public Function CerrarOrdendeIngreso(ORDINGCOD As String, USUCIERRA As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_CERRAR_ORDENDEINGRESO", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Parameters.AddWithValue("@USUCIERRA", USUCIERRA)
                Retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Retorno
    End Function

    Public Function AnularOrdendeIngreso(ORDINGCOD As String, USUANULA As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ANULAR_ORDENDEINGRESO", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Parameters.AddWithValue("@USUANULA", USUANULA)
                Retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Retorno
    End Function

    Public Function ActualizarTotales(ORDINGCOD As String, SUBTOTAL As Double,
    SUBTOTALIMPUESTO As Double, TOTALGENERAL As Double, TOTALLETRAS As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ACTUALIZA_ORDENDEINGRESO_TOT", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Parameters.AddWithValue("@SUBTOTAL", SUBTOTAL)
                .Parameters.AddWithValue("@SUBTOTALIMPUESTO", SUBTOTALIMPUESTO)
                .Parameters.AddWithValue("@TOTALGENERAL", TOTALGENERAL)
                .Parameters.AddWithValue("@TOTALLETRAS", TOTALLETRAS)
                Retorno = .ExecuteScalar
                If cnx.State = ConnectionState.Open Then
                    cnx.Close()
                End If
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Retorno
    End Function

    Public Function RPTOrdIng(ORDINGCOD As String) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_RPT_ORDING", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@ORDINGCOD", ORDINGCOD)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function


    Public Function RPTIngresos(ESTADO As String, DESDE As Date, HASTA As Date) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_RPT_INGRESOS", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@ESTADO", ESTADO)
                .SelectCommand.Parameters.AddWithValue("@DESDE", DESDE)
                .SelectCommand.Parameters.AddWithValue("@HASTA", HASTA)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function

End Class
