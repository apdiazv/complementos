Option Explicit On

Imports System.Management

Public Class frmValidador

    Private Sub frmValidador_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.WriteLine(GenerarClave(MostrarInformacionDeDisco))
        Debug.WriteLine(LeeRegistro())

        If GenerarClave(MostrarInformacionDeDisco) = LeeRegistro() Then
            MsgBox("OK")
        Else
            MsgBox("BAD")
        End If
        Debug.WriteLine(GenerarClave(MostrarInformacionDeDisco))
        Debug.WriteLine(LeeRegistro())
    End Sub

    ''' <summary>
    ''' Genera una clave del tipo xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx a partir de un string.
    ''' </summary>
    ''' <param name="sClave">Clave a Encriptar</param>
    ''' <returns>String del tipo xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx en hexadecimal</returns>
    ''' <remarks></remarks>
    Private Function GenerarClave(ByVal sClave As String) As String
        ' Obtenemos la longitud de la cadena de usuario
        Dim longitud As Byte = sClave.Length
        ' Declaramos valorEntrada para obtener el valor general
        ' correspondiente a la entrada de usuario
        Dim valorEntrada As Long = 0
        ' Recorremos la cadena entera para sumar el valor
        ' total de sus cASCII
        For I As Byte = 0 To longitud - 1
            valorEntrada += Asc(sClave.Substring(I, 1))
        Next
        ' Dividimos el valor final resultante de la suma de
        ' sus valores ASCII entre la longitud de la cadena
        valorEntrada \= longitud
        ' Obtenemos un valor base que corresponde con el
        ' cdel producto entre el valor de entrada 
        ' anteriormente calcula por su longitud
        Dim valorBase As Integer = valorEntrada * longitud
        Dim key As String = ""
        ' Empezamos obteniento valores
        ' Obtenemos el valor hexadecimal
        Dim valor As String = Hex(valorBase + (123 * 10000))
        key &= valor.Substring(valor.Length - 6, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (98 * 12500))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(0, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (77 * 15000))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (121 * 17500))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(0, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (55 * 20000))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (134 * 22500))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(0, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (63 * 25000))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(valor.Length - 6, 6)
        ' Obtenemos el valor hexadecimal
        valor = Hex(valorBase + (117 * 27500))
        ' Obtenemos el valor de clave
        key &= "-" & valor.Substring(0, 6)
        ' Devolvemos el valor final (xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx-xxxxxx)
        Return key
    End Function




    Public Function MostrarInformacionDeDisco() As String
        Dim sDriveLetter As String = "C"
        Dim sSerial As String
        Dim oGetVol As New Volume.GetVol


        sSerial = oGetVol.GetVolumeSerial(sDriveLetter)

        Return sSerial

    End Function

    Private Function LeeRegistro() As String

        Return GetSetting("Derivados", "Licencia", "Value")

    End Function


    Private Sub GrabaRegistro()
        SaveSetting("Derivados", "Licencia", "Value", GenerarClave(MostrarInformacionDeDisco))
    End Sub


    Private Sub BorraRegistro()
        DeleteSetting("Derivados", "Licencia", "Value")
    End Sub


    Private Sub btnCrear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCrear.Click
        GrabaRegistro()
    End Sub

    Private Sub btnBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrar.Click
        BorraRegistro()
    End Sub


    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub


    Private Sub btnMostrar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMostrar.Click
        MsgBox(LeeRegistro())
    End Sub

    Private Declare Function GetVolumeInformation Lib _
   "kernel32.dll" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Integer, _
   ByVal lpVolumeSerialNumber As Long, _
   ByVal lpMaximumComponentLength As Long, _
   ByVal lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long

    Public Function DriveSerialNumber(ByVal Drive As String) As Long

        'usage: SN = DriveSerialNumber("C:")

        Dim lAns As Long
        Dim lRet As Long
        Dim sVolumeName As String, sDriveType As String
        Dim sDrive As String

        'Deal with one and two character input values
        sDrive = Drive

        sVolumeName = "" 'String$(255, Chr$(0))
        sDriveType = "" 'String$(255, Chr$(0))

        lRet = GetVolumeInformation(sDrive, sVolumeName, _
        255, lAns, 0, 0, sDriveType, 255)

        DriveSerialNumber = lAns
    End Function




End Class
