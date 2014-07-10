Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Class1
    '------------------------------------------------------
    ' Declaration found in Microsoft.Win32.Win32Native
    '------------------------------------------------------
    Friend Declare Auto Function GetVolumeInformation Lib "kernel32.dll" (ByVal drive As String, <Out()> ByVal volumeName As StringBuilder, ByVal volumeNameBufLen As Integer, <Out()> ByRef volSerialNumber As Integer, <Out()> ByRef maxFileNameLen As Integer, <Out()> ByRef fileSystemFlags As Integer, <Out()> ByVal fileSystemName As StringBuilder, ByVal fileSystemNameBufLen As Integer) As Boolean

    '------------------------------------------------------
    ' Test in my Form class
    '------------------------------------------------------
    Public Sub Check()
        Try
            Dim volumeName As StringBuilder = New StringBuilder(50)
            Dim stringBuilder As StringBuilder = New StringBuilder(50)
            Dim volSerialNumber As Integer
            Dim maxFileNameLen As Integer
            Dim fileSystemFlags As Integer
            If Not GetVolumeInformation("C:\", volumeName, 50, volSerialNumber, maxFileNameLen, fileSystemFlags, stringBuilder, 50) Then
                Dim lastWin32Error As Integer = Marshal.GetLastWin32Error()
                MsgBox("Error number:" & lastWin32Error)
            Else
                MsgBox(volSerialNumber.ToString("X"))
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
End Class
