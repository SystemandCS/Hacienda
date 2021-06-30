Imports SvrDatosNet
Public Class ClsGenerica

    Public App As String

    Protected m_Servidor As Servidor


    Sub New(Optional App As String = Nothing)

        If App Is Nothing Then
            Me.App = "CARNES"
        Else
            Me.App = App

        End If

    End Sub


    Public Overridable Function EjecutarSP(ByVal Nombre As String, ByVal ParamArray Parametros() As Object) As Integer
        Dim cant As Integer
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then

                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If

            cant = srv.EjecutarSP(Nombre, Parametros)

            ''If cant <> 1 Then
            ''Throw New ApplicationException("No se encontro el Registro buscado")
            ''End If

            Return Parametros(0)

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Overridable Function TraerDatos(ByVal Nombre As String,
       ByVal ParamArray Parametros() As Object) As System.Data.DataSet
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.TraerDatos(Nombre, Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function

    Public Overridable Function EjecutarSql(ByVal StrSql As String) As System.Data.DataSet
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.EjecutarSql(StrSql)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function




End Class