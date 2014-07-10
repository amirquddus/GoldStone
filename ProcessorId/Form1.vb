Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Management

Public Class Form1

    Dim objMOS As ManagementObjectSearcher

    Dim objMOC As Management.ManagementObjectCollection

    Dim objMO As Management.ManagementObject

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        objMOS = New ManagementObjectSearcher("Select * From Win32_Processor")

        objMOC = objMOS.Get

        For Each objMO In objMOC

            MessageBox.Show("CPU ID = " & objMO("ProcessorID"))

        Next

        objMOS.Dispose()

        objMOS = Nothing

        objMO.Dispose()

        objMO = Nothing
    End Sub

    Private Shared Sub EncryptFile(ByVal sInputFilename As String, ByVal sOutputFilename As String, ByVal sKey As String)
        Dim fsInput As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)

        Dim fsEncrypted As New FileStream(sOutputFilename, FileMode.Create, FileAccess.Write)
        Dim DES As New DESCryptoServiceProvider()
        DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
        Dim desencrypt As ICryptoTransform = DES.CreateEncryptor()
        Dim cryptostream As New CryptoStream(fsEncrypted, desencrypt, CryptoStreamMode.Write)

        Dim bytearrayinput As Byte() = New Byte(fsInput.Length - 1) {}
        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Close()
        fsInput.Close()
        fsEncrypted.Close()
    End Sub

    Private Shared Sub DecryptFile(ByVal sInputFilename As String, ByVal sOutputFilename As String, ByVal sKey As String)
        Dim DES As New DESCryptoServiceProvider()
        'A 64 bit key and IV is required for this provider.
        'Set secret key For DES algorithm.
        DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        'Set initialization vector.
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

        'Create a file stream to read the encrypted file back.
        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
        'Create a DES decryptor from the DES instance.
        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
        'Create crypto stream set to read and do a 
        'DES decryption transform on incoming bytes.
        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
        'Print the contents of the decrypted file.
        Dim fsDecrypted As New StreamWriter(sOutputFilename)
        fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd())
        fsDecrypted.Flush()
        fsDecrypted.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim hdCollection As Collection = New Collection()
        Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")

        For Each wmi_HD As ManagementObject In searcher.Get()
            Dim hd As HardDrive = New HardDrive()
            hd.Model = wmi_HD("Model").ToString()
            hd.Type = wmi_HD("InterfaceType").ToString()
            hd.SerialNumber = wmi_HD("SerialNumber").ToString()
            hd.Caption = wmi_HD("Caption").ToString()
            hdCollection.Add(hd)
        Next


        Dim modelNo As String = identifier("Win32_DiskDrive", "Model")
        Dim manufatureID As String = identifier("Win32_DiskDrive", "Manufacturer")
        Dim signature As String = identifier("Win32_DiskDrive", "Signature")
        Dim totalHeads As String = identifier("Win32_DiskDrive", "TotalHeads")
        Dim serialNumber As String = identifier("Win32_DiskDrive", "SerialNumber")
        Dim ProcessorID As String = identifier("Win32_Processor", "ProcessorID")
    End Sub

    Private Function identifier(ByVal wmiClass As String, ByVal wmiProperty As String) As String
        'Return a hardware identifier
        Dim result As String = ""
        Dim mc As System.Management.ManagementClass = New System.Management.ManagementClass(wmiClass)
        Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances()
        For Each mo As System.Management.ManagementObject In moc
            'Only get the first one
            If result = "" Then
                Try
                    result = mo(wmiProperty).ToString()
                    Exit For
                Catch ex As Exception

                End Try
            End If
        Next

        Return result
    End Function

    Public Class HardDrive
        Public Model As String
        Public Type As String
        Public SerialNumber As String
        Public Caption As String
    End Class

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim x As Class1 = New Class1()
        x.Check()
    End Sub
End Class
