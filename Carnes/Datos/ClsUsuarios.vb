Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration

Public Class ClsUsuarios
    Dim cnx As New SqlConnection(ConfigurationManager.ConnectionStrings("cn").ConnectionString)

    Public Function DB() As String
        Return cnx.Database
    End Function

    Public Function Acceso(USUNOM As String, USUPAS As String) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_USUARIOS", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@USUNOM", USUNOM)
                .SelectCommand.Parameters.AddWithValue("@USUPAS", USUPAS)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function

    Public Function InsertaUsuarios(USUNOM As String, USUPAS As String,
    USUTIPO As String, USUNOMBRES As String, USUAPELLIDOS As String, USUDIRECCION As String,
    USUTELEFONO As String, USUCELULAR As String, USUEMAIL As String, USUDOCIDEN As String) As String
        Dim retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_INSERTA_USUARIOS", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@USUNOM", USUNOM)
                .Parameters.AddWithValue("@USUPAS", USUPAS)
                .Parameters.AddWithValue("@USUTIPO", USUTIPO)
                .Parameters.AddWithValue("@USUNOMBRES", USUNOMBRES)
                .Parameters.AddWithValue("@USUAPELLIDOS", USUAPELLIDOS)
                .Parameters.AddWithValue("@USUDIRECCION", USUDIRECCION)
                .Parameters.AddWithValue("@USUTELEFONO", USUTELEFONO)
                .Parameters.AddWithValue("@USUCELULAR", USUCELULAR)
                .Parameters.AddWithValue("@USUEMAIL", USUEMAIL)
                .Parameters.AddWithValue("@USUDOCIDEN", USUDOCIDEN)
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

    Public Function ActualizaUsuarios(USUCOD As String, USUNOM As String, USUPAS As String,
USUTIPO As String, USUNOMBRES As String, USUAPELLIDOS As String, USUDIRECCION As String,
USUTELEFONO As String, USUCELULAR As String, USUEMAIL As String, USUDOCIDEN As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ACTUALIZA_USUARIOS", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@USUCOD", USUCOD)
                .Parameters.AddWithValue("@USUNOM", USUNOM)
                .Parameters.AddWithValue("@USUPAS", USUPAS)
                .Parameters.AddWithValue("@USUTIPO", USUTIPO)
                .Parameters.AddWithValue("@USUNOMBRES", USUNOMBRES)
                .Parameters.AddWithValue("@USUAPELLIDOS", USUAPELLIDOS)
                .Parameters.AddWithValue("@USUDIRECCION", USUDIRECCION)
                .Parameters.AddWithValue("@USUTELEFONO", USUTELEFONO)
                .Parameters.AddWithValue("@USUCELULAR", USUCELULAR)
                .Parameters.AddWithValue("@USUEMAIL", USUEMAIL)
                .Parameters.AddWithValue("@USUDOCIDEN", USUDOCIDEN)
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


    Public Function EliminaUsuarios(USUCOD As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_ELIMINA_USUARIOS", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@USUCOD", USUCOD)
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

    Public Function ListaUsuarios(Valor As String) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_LISTA_USUARIOS", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@VALOR", Valor)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function

    Public Function BuscaUsuarios(USUCOD As String) As DataTable
        Dim Dt As New DataTable
        Try
            Dim Sda As New SqlDataAdapter("SP_BUSCA_USUARIOS", cnx)
            With Sda
                .SelectCommand.CommandType = CommandType.StoredProcedure
                .SelectCommand.Parameters.AddWithValue("@USUCOD", USUCOD)
                .Fill(Dt)
            End With
        Catch ex As SqlException
            Throw ex
        End Try
        Return Dt
    End Function



    Public Function CambioPasswordUsuarios(USUNOM As String, USUPAS As String) As String
        Dim Retorno As String = ""
        Try
            If cnx.State = ConnectionState.Open Then
                cnx.Close()
            Else
                cnx.Open()
            End If
            Dim Cmd As New SqlCommand("SP_USUARIOS_CAMBIOPASS", cnx)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@USUNOM", USUNOM)
                .Parameters.AddWithValue("@USUPAS", USUPAS)
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

End Class
