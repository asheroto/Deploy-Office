﻿Imports System.IO
Imports System.Net

Module Functions

    Function CreateShortcut(ByVal TargetName As String, ByVal ShortcutPath As String, ByVal ShortcutName As String) As Boolean
        Dim oShell As Object
        Dim oLink As Object

        Try
            Dim ShortcutFullPath As String = Path.Combine(ShortcutPath, ShortcutName & ".lnk")

            oShell = CreateObject("WScript.Shell")
            oLink = oShell.CreateShortcut(ShortcutFullPath)

            oLink.TargetPath = TargetName
            oLink.WindowStyle = 1
            oLink.Save()

            Application.DoEvents()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Sub RunSetup()
        'Run setup
        Dim SetupPath As String = Path.Combine(Main.TempPath, Main.SetupFileName)
        Dim x As New Process
        x.StartInfo.FileName = SetupPath
        x.StartInfo.Arguments = "/configure configuration.xml"
        x.StartInfo.WorkingDirectory = Main.TempPath
        x.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        x.StartInfo.CreateNoWindow = True
        x.Start()

        'Wait until installation is finished
        Do Until x.HasExited
            Application.DoEvents()
            Threading.Thread.Sleep(500)
        Loop
    End Sub

    Sub CreateDesktopShortcuts()
        Dim OfficeApps As New Dictionary(Of String, String)
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", "Word")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", "Excel")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE", "PowerPoint")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE", "Outlook")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE", "OneNote")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE", "Publisher")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE", "Access")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE", "Visio")

        For Each exe As KeyValuePair(Of String, String) In OfficeApps
            If File.Exists(exe.Key) Then
                CreateShortcut(exe.Key, Environment.GetFolderPath(Environment.SpecialFolder.Desktop), exe.Value)
            End If
        Next
    End Sub

    Sub Cleanup()
        Dim SetupPath As String = Path.Combine(Main.TempPath, Main.SetupFileName)
        Dim ConfigPath As String = Path.Combine(Main.TempPath, Main.ConfigFileName)
        Dim AssetsPath As String = Path.Combine(Main.TempPath, Main.AssetsFilename)

        DeleteFileIfExist(SetupPath)
        DeleteFileIfExist(ConfigPath)
        DeleteFileIfExist(AssetsPath)
    End Sub

    Sub DeleteFileIfExist(path As String)
        Try
            If File.Exists(path) Then
                File.Delete(path)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub WaitToClose()
        Dim c As New Stopwatch
        c.Start()
        Do Until c.Elapsed.Seconds >= 30
            Application.DoEvents()
        Loop
        c.Stop()
    End Sub

    Sub LogAppend(text As String)
        Main.LogTextbox.AppendText(text & "..." & vbNewLine)
    End Sub

    Sub TLSchannelFix()
        'Enforce TLS secure channel
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"" /v ""DisabledByDefault"" /t REG_DWORD /d ""0"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"" /v ""Enabled"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"" /v ""DisabledByDefault"" /t REG_DWORD /d ""0"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"" /v ""Enabled"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319"" /v ""SchUseStrongCrypto"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"" /v ""SchUseStrongCrypto"" /t REG_DWORD /d ""1"" /f", vbHidden)
    End Sub

End Module