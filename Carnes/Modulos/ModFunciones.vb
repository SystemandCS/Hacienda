Module ModFunciones

    Public USUCOD As String
    Public USULOGIN As String
    Public USUNOM As String
    Public USUPAS As String
    Public USUTIPO As String
    Public USUDB As String

    Public Sub ENVIAR_ENTER(ByVal xKeyChar As Char)
        If Asc(xKeyChar) = 13 Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    'FUNCION QUE PERMITE EL INGRESO DE SOLO NUMEROS: UTILIZAR PARA DNI, RUC, Y OTROS QUE SEAN SOLO NUMEROS
    Public Function SoloNumeros(ByRef KeyAscii As Integer)
        Dim retorno As Integer = 1
        If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
            retorno = 0
        End If
        Return retorno
    End Function

    'FUNCION QUE PERMITE EL INGRESO DE SOLO MONEDAS CON PUNTO DECIMAL
    Public Function SoloMonedas(ByRef TXT As TextBox, ByRef KeyAscii As Integer)
        Dim retorno As Integer = 1
        Dim cadena As String = Chr(KeyAscii)
        If InStr(TXT.Text, Chr(KeyAscii)) > 0 And Chr(KeyAscii) = "." Then
            If TXT.SelectionLength > 0 Then
                If TXT.SelectionLength <> TXT.Text.Length Then
                    If InStr(Mid(TXT.Text, TXT.SelectionStart + 1, TXT.SelectionLength), ".") = 0 Then
                        If InStr(TXT.Text, ".") > 0 And Chr(KeyAscii) = "." Then
                            retorno = 0
                        End If
                    End If
                End If
            Else
                retorno = 0
            End If
        ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
            retorno = 0
        End If
        Return retorno
    End Function




End Module
