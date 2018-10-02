Imports System.Data
Imports System.IO
Imports System.ComponentModel
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Management
Imports System.Net.Http
Imports System.DirectoryServices
Imports System.Windows.Forms
Imports System.Net

Class MainWindow


    Private _stringBuilder As StringBuilder = New StringBuilder

    Private strServername As String = "ladata2"
    Private strUsername As String = "ebell"
    Private strPassword As String = "6B3LL2017*"
    Private strClassSelection As String
    Private rcOptions As ConnectionOptions
    Private moCollection As ManagementObjectCollection
    Private oQuery As ObjectQuery
    Private moSearcher As ManagementObjectSearcher
    Private mScope As ManagementScope


    Private Sub Button1_Click(sender As Object, e As RoutedEventArgs)
        ListShares()
        Try
            Dim s As StreamWriter = File.AppendText("\\GSODATA1\USER\ebell\Config\Windows\Desktop\Output\output.txt")
            s.WriteLine(_stringBuilder.ToString)
            s.Flush()
            s.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    'SECTION FOR GATHERING NETWORK INFORMATION 
    <StructLayoutAttribute(LayoutKind.Sequential)>
    Public Structure SHARE_INFO
        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public shi_netname As String
        Public shi_type As UInteger
        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public shi_remark As String
        Public shi_permissions As Integer
        Public shi_max_uses As Integer
        Public shi_current_uses As Integer
        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public shi_path As String
        <MarshalAsAttribute(UnmanagedType.LPWStr)> Public shi_passwd As String
        Public shi502_reserved As Integer
    End Structure

    Private Structure SERVER_INFO
        Dim sv100_platform_id As Integer
        Dim sv100_nam
    End Structure



    Public Sub ListShares()
        Dim Path As ManagementPath = New ManagementPath
        Dim Shares As ManagementClass = Nothing
        Dim CO As ConnectionOptions = New ConnectionOptions
        CO.Impersonation = ImpersonationLevel.Impersonate
        CO.Authentication = AuthenticationLevel.Default
        CO.EnablePrivileges = True
        CO.Username = "cmsvc"
        CO.Password = "cm2016lw!"
        Path.Server = "LVDDATA2" ' use . for local server, else server name

        Path.NamespacePath = "\root\cimv2"
        Path.RelativePath = "Win32_Share"

        Dim password As String = "password"
        Dim username As String = "username"
        Dim domain As String = "domain"


        Dim theNetworkCredential As NetworkCredential = New NetworkCredential(username, password, domain)
        Dim theNetcache As CredentialCache = New CredentialCache()
        CredentialCache.DefaultNetworkCredentials.Domain = domain
        CredentialCache.DefaultNetworkCredentials.UserName = username
        CredentialCache.DefaultNetworkCredentials.Password = password


        Dim lNetWorkCredential As List(Of System.Net.NetworkCredential) = New System.Collections.Generic.List(Of System.Net.NetworkCredential)
        lNetWorkCredential.Add(theNetworkCredential)



        Dim Scope As ManagementScope = New ManagementScope(Path, CO)
        Scope.Connect()

        If (Scope.IsConnected) Then
            MsgBox("Connected")
        Else
            MsgBox("Error Connecting")
        End If
        Dim os As ManagementClass = New ManagementClass("Win32_OperatingSystem")
        Dim inParams As ManagementBaseObject = os.GetMethodParameters("Win32Shutdown")
        inParams("Flags") = "2"
        inParams("Reserved") = "0"




        Dim Options As ObjectGetOptions = New ObjectGetOptions(Nothing, New TimeSpan(0, 0, 0, 5), True)
        Try
            Shares = New ManagementClass(Scope, Path, Options)
            Dim MOC As ManagementObjectCollection = Shares.GetInstances()

            For Each mo As ManagementObject In MOC

                Console.WriteLine("{0} - {1} - {2}", mo("Name"), mo("Description"), mo("Path"))

            Next

        Catch ex As Exception

            Console.WriteLine(ex.Message)

        Finally

            If Shares IsNot Nothing Then

                Shares.Dispose()
            End If
        End Try

    End Sub
End Class
