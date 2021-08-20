Imports SvrDatosNet
Public Class ClsStock

    Inherits ServidorNET

    Sub New()
        MyBase.New("STK")

    End Sub



    Public Overrides Function Insertar(ByVal ParamArray Parametros() As Object) As Integer

        Dim clsVendedor As New ClsVendedor
        ' Dim ParamVendedores() As Object

        ' En este puede ir la regla de negocio de insertar en la tabla vendedor

        Try




            MyBase.Insertar(Parametros)

            ' ParamVendedores(0) = Parametros(0)
            ' ParamVendedores(1) = Parametros(8)

            ' clsVendedor.Insertar(ParamVendedores)

            Return Parametros(0)






        Catch ex As Exception
            Throw ex
        End Try
    End Function



End Class