Public MustInherit Class ServidorNET

    Public Tabla As String
    Public FuenteDatos As String
    Public App As String

    Protected m_Servidor As Servidor


    Sub New(ByVal Tabla As String, Optional App As String = Nothing)
        Me.Tabla = Tabla
        If App Is Nothing Then
            Me.App = "CARNES"

        Else
            Me.App = App

        End If

    End Sub

    Sub New(ByVal oServidor As Servidor, ByVal Tabla As String)
        Me.Tabla = Tabla
        m_Servidor = oServidor
    End Sub

#Region "Metodos de Actualizacion"
    Public Overridable Function Insertar(ByVal ParamArray Parametros() As Object) As Integer
        Dim cant As Integer
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then

                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If

            cant = srv.EjecutarSP(Tabla & "_IN", Parametros)

            ' If cant <> 1 Then
            'Throw New ApplicationException("No se encontro el Registro buscado")
            'End If

            Return cant

        Catch ex As Exception
            Throw New ApplicationException("Excepción en el metodo Insertar: " & ex.Message, ex)

        End Try

    End Function



    Public Overridable Function Actualizar(ByVal ParamArray Parametros() As Object) As Integer
        Dim cant As Integer
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If

            cant = srv.EjecutarSP(Tabla & "_AC", Parametros)


            Return Parametros(0)

            'If cant = 0 Then

            '  Throw New ApplicationException("No se encontro el Registro buscado")

            ' End If
        Catch ex As Exception
            Throw New ApplicationException("Excepción en el metodo Actualizar: " & ex.Message, ex)
        End Try

    End Function

    Public Overridable Function Eliminar(ByVal ParamArray Parametros() As Object) As Integer
        Dim cant As Integer
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            cant = srv.EjecutarSP(Tabla & "_EL", Parametros)

            Return Parametros(0)
        Catch ex As Exception
            Throw New ApplicationException("Excepción en el metodo Eliminar: " & ex.Message, ex)

        End Try

    End Function

#End Region


#Region "Metodos de Lectura"

    Public Overridable Function ObtenerUno(ByVal ParamArray Parametros() As Object) As System.Data.DataSet

        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If

            ds = srv.TraerDatos(Tabla & "_OU", Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function

    Public Overridable Function ObtenerTodos(ByVal ParamArray Parametros() As Object) As System.Data.DataSet
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.TraerDatos(Tabla & "_OT", Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function



    Public Overridable Function Filtrado(ByVal ParamArray Parametros() As Object) As System.Data.DataSet
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.TraerDatos(Tabla & "_filtrado", Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function

    'Ejecuto un stores procedure
    Public Overridable Function EjecutarSp(ByVal sp As String, ByVal ParamArray Parametros() As Object) As System.Data.DataSet
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.TraerDatos(Tabla & "_" & sp, Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function



    'Ejecuto un store Procedure y lo transformo en resultado XML
    Public Overridable Function EjecutarSpXml(ByVal sp As String, ByVal ParamArray Parametros() As Object) As Xml.XmlDocument
        Dim Xml As Xml.XmlDocument
        Dim ds As Data.DataSet
        Try
            Dim srv As Servidor
            If m_Servidor Is Nothing Then
                srv = Servidor.AbrirConexion(App)
            Else
                srv = m_Servidor
            End If
            ds = srv.TraerDatos(Tabla & "_" & sp, Parametros)
        Catch ex As Exception
            Throw ex
        End Try
        Return Xml
    End Function

#End Region
End Class
