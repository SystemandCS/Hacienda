':::Libreria necesaria para manejo de directorios y archivos
Imports System.IO

Public Class FrmLogueo
    Private Sub FrmLogueo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CargarArchivos()
        'LeerArchivo()



    End Sub

    Private Sub CargarArchivos()


        Try

            ListArchivos.Items.Clear()

            ':::Contamos cuanto archivos de texto hay en la carpeta
            Dim Total = My.Computer.FileSystem.GetFiles(ELog.ObtenerUbicacion(), FileIO.SearchOption.SearchAllSubDirectories, "*.txt")
            ' LblTotal.Text = "Total Archivos de Texto: " + CStr(Total.Count)

            ':::Realizamos la búsqueda de la ruta de cada archivo de texto y los agregamos al ListBox
            For Each archivos As String In My.Computer.FileSystem.GetFiles(ELog.ObtenerUbicacion(), FileIO.SearchOption.SearchAllSubDirectories, "*.txt")
                ListArchivos.Items.Add(archivos)
            Next
        Catch ex As Exception
            MsgBox("No se realizó la operación por: " & ex.Message)
        End Try

        GroupBoxResultado.Visible = False


    End Sub


    Sub LeerArchivo(ByVal Archivo As String)
        ':::Creamos nuestro objeto de tipo StreamReader que nos permite leer archivos
        Dim leer As New StreamReader(Archivo) 'ELog.ObtenerUbicacion() & "/20210414.txt")

        ':::Limpiamos nuestro ListBox
        ListArchivosDatos.Items.Clear()

        Try
            ':::Indicamos mediante un While que mientras no sea el ultimo caracter repita el proceso
            While leer.Peek <> -1
                ':::Leemos cada linea del archivo TXT
                Dim linea As String = leer.ReadLine()
                ':::Validamos que la linea no este vacia
                If String.IsNullOrEmpty(linea) Then
                    Continue While
                End If
                ':::Agregramos los registros encontrados
                ListArchivosDatos.Items.Add(linea)
            End While

            leer.Close()
            ':::Total de registros cargados al ListBox
            '  LblTotal.Text = ListBoxClientes.Items.Count
        Catch ex As Exception
            MsgBox("Se presento un problema al leer el archivo: " & ex.Message, MsgBoxStyle.Critical, ":::Aprendamos de Programación:::")
        End Try
    End Sub



    Private Sub ListArchivos_Click(sender As Object, e As EventArgs) Handles ListArchivos.Click

        If ListArchivos.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a file.")
            Exit Sub
        End If

        ' Obtain the file path from the list box selection.
        Dim filePath = ListArchivos.SelectedItem.ToString

        ' Verify that the file was not removed since the
        ' Browse button was clicked.
        If My.Computer.FileSystem.FileExists(filePath) = False Then
            MessageBox.Show("File Not Found: " & filePath)
            Exit Sub
        End If

        ' Obtain file information in a string.
        Dim fileInfoText As String = GetTextForOutput(filePath)

        ' Show the file information.
        'MessageBox.Show(fileInfoText)
        lblArchivo.Text = filePath


        GroupBoxResultado.Visible = True

        LeerArchivo(filePath)


    End Sub


    Private Function GetTextForOutput(ByVal filePath As String) As String
        ' Verify that the file exists.
        If My.Computer.FileSystem.FileExists(filePath) = False Then
            Throw New Exception("File Not Found: " & filePath)
        End If

        ' Create a new StringBuilder, which is used
        ' to efficiently build strings.
        Dim sb As New System.Text.StringBuilder()

        ' Obtain file information.
        Dim thisFile As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(filePath)

        ' Add file attributes.
        sb.Append("File: " & thisFile.FullName)
        sb.Append(vbCrLf)
        sb.Append("Modified: " & thisFile.LastWriteTime.ToString)
        sb.Append(vbCrLf)
        sb.Append("Size: " & thisFile.Length.ToString & " bytes")
        sb.Append(vbCrLf)

        ' Open the text file.
        Dim sr As System.IO.StreamReader =
        My.Computer.FileSystem.OpenTextFileReader(filePath)

        ' Add the first line from the file.
        If sr.Peek() >= 0 Then
            sb.Append("First Line: " & sr.ReadLine())
        End If
        sr.Close()

        Return sb.ToString
    End Function

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click


        Try


            If lblArchivo.Text.Trim = "" Then
                MessageBox.Show("Debe seleccionar una Archivo", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            ELog.EliminarArchivo(lblArchivo.Text)




            MessageBox.Show("Registro Eliminado", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)

            CargarArchivos()

        Catch ex As Exception

            ELog.Grabar(Me, ex)
            MessageBox.Show(ex.Message, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try




    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click

        Me.Close()

    End Sub
End Class