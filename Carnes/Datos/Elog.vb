Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Web
Imports System.IO
Imports System.Diagnostics





Public Class ELog

    Public FilePath As String


    Public Shared Sub Grabar(ByVal obj As Object, ByVal ex As Exception)

        Dim REF As String


        Dim FilePath As String
        FilePath = System.AppDomain.CurrentDomain.BaseDirectory


        Dim numAleatorio As New Random()
        REF = System.Convert.ToString(numAleatorio.Next)


        Dim fecha As String = System.DateTime.Now.ToString("yyyyMMdd")
        Dim hora As String = System.DateTime.Now.ToString("HH:mm:ss")
        Dim path As String = FilePath & ("log\" & fecha & ".txt")

        Dim sw As StreamWriter = New StreamWriter(path, True)
        Dim stacktrace As StackTrace = New StackTrace()



        sw.WriteLine("")
        sw.WriteLine("----------MENSAJE  REF:" & REF & "----------")
        sw.WriteLine("Usuario:" & USULOGIN)

        sw.WriteLine("Equipo:" & USUDB)
        sw.WriteLine("Fecha Hora:" & fecha + " " + hora)
        sw.WriteLine("Formulario:" & obj.[GetType]().FullName)
        sw.WriteLine("Metodo:" & stacktrace.GetFrame(1).GetMethod().Name)
        sw.WriteLine("Mensaje:" & ex.Message)
        sw.WriteLine("HelpLink:" & ex.HelpLink)
        sw.WriteLine("Source:" & ex.Source)
        sw.WriteLine("StackTrace:" & ex.StackTrace)
        sw.WriteLine("TargetSite:" & ex.TargetSite.ToString)
        sw.WriteLine("Informacion Extra")


        If Not IsNothing(ex.InnerException) Then

            sw.WriteLine("Inner Exception: ")

            sw.WriteLine("Message: " & ex.InnerException.ToString)
            sw.WriteLine("Message: " & ex.InnerException.GetType.GUID.ToString)


        End If

        sw.WriteLine("----------FIN MENSAJE-------------")
        sw.WriteLine("")


        sw.Flush()
        sw.Close()



    End Sub

    Public Shared Function ObtenerUbicacion() As String
        Dim FilePath As String
        'FilePath = server.MapPath("/Precursores2010")
        FilePath = System.AppDomain.CurrentDomain.BaseDirectory

        Return (FilePath & ("log\"))

    End Function


    Public Shared Sub EliminarArchivo(ByVal ArchivoBorrar As String)

        'comprobamos que el archivo existe

        Try
            If System.IO.File.Exists(ArchivoBorrar) = True Then
                System.IO.File.Delete(ArchivoBorrar)
            End If


        Catch ex As Exception


        End Try


    End Sub






    Private Function IsDbNullOrEmpty(ByVal value As Object) As Boolean
        Return IsDBNull(value) OrElse String.IsNullOrWhiteSpace(value.ToString())
    End Function





End Class
