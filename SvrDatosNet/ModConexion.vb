
Imports Microsoft.VisualBasic
Imports System.Xml

Module modConexion




    Public Function ObtenerStringConexion(ByVal strConexion As String) As String
        Dim FilePath As String
        'FilePath = server.MapPath("/Precursores2010")
        FilePath = System.AppDomain.CurrentDomain.BaseDirectory
        Return ObtenerStringDeArchivoXML(FilePath & "conectar/cnxml.xml", strConexion, "connectionstring")

    End Function

    Public Function ObtenerStringDeArchivoXML(ByVal strPath As String, _
    ByVal strAplicacion As String, ByVal strParametro As String) As String
        Dim oDom As New XmlDocument()
        oDom.Load(strPath)
        Return oDom.SelectSingleNode("/aplicaciones/" & strAplicacion & "/" & strParametro).InnerText
        oDom = Nothing
    End Function

   



End Module
