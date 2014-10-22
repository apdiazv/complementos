Option Strict Off
Imports Extensibility
imports System.Runtime.InteropServices

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the $SAFEOBJNAME$Setup project, 
' right click the project in the Solution Explorer, then choose install.
#End Region

<GuidAttribute("19F21CB0-D25B-4b77-8E41-99956B2F8B9A"), ProgIdAttribute("CAddin.Connect")> _
Public Class Connect
    Implements Extensibility.IDTExtensibility2
    Private applicationObject As Object
    Private addInInstance As Object
    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
    End Sub
    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    End Sub
    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete
    End Sub
    Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnDisconnection
    End Sub
    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection
        applicationObject = application
        addInInstance = addInInst
        Dim path As String
        Debug.WriteLine(Me.GetType.Assembly.FullName)
        path = System.IO.Path.GetDirectoryName(Me.GetType.Assembly.Location)
        MessageBox.Show(path)
        path = "C:\Archivos de programa\Creasys\Funciones"
        'MessageBox.Show("12345678")
        'MessageBox.Show(application.registerxll(path & "\proxy.xll"))
        application.registerxll(path & "\proxy.xll")
    End Sub
End Class
