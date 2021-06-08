Imports System.IO
Imports System.Net

Module Functions

    Function IsProcessRunning(ProcessName As String)
        Dim p() As Process
        p = Process.GetProcessesByName(ProcessName)
        If p.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function CreateShortcut(ByVal TargetName As String, ByVal ShortcutPath As String, ByVal ShortcutName As String) As Boolean
        Dim oShell As Object
        Dim oLink As Object

        Try
            oShell = CreateObject("WScript.Shell")
            oLink = oShell.CreateShortcut(Path.Combine(ShortcutPath, ShortcutName & ".lnk"))

            oLink.TargetPath = TargetName
            oLink.WindowStyle = 1
            oLink.Save()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Sub Download(URL As String, FileName As String)
        Dim DownloadPath As String = Path.Combine(main.TempPath, FileName)

        If File.Exists(DownloadPath) Then
            Try
                File.Delete(DownloadPath)
            Catch ex As Exception
                MsgBox(String.Format("Failed to delete existing {0}. Please end all setup.exe processes or restart your computer and try again." & vbNewLine & vbNewLine & ex.Message, FileName), vbExclamation)
                End
            End Try

        End If

        Try
            Dim wc As WebClient = New WebClient()
            wc.DownloadFile(URL, DownloadPath)
        Catch ex As Exception
            MsgBox(String.Format("Failed to download {0}. Please ensure you have Internet connectivity or restart your computer and try again." & vbNewLine & vbNewLine & ex.Message, FileName), vbExclamation)
            End
        End Try

        If File.Exists(DownloadPath) = False Then
            MsgBox(String.Format("Problem downloading {0}. Please ensure you have Internet connectivity or restart your computer and try again.", FileName), vbExclamation)
        End If
    End Sub

    Sub RunSetup()
        Dim x As New Process
        x.StartInfo.FileName = Path.Combine(Application.StartupPath, Main.SetupFileName)
        x.StartInfo.Arguments = "/configure configuration.xml"
        x.StartInfo.WorkingDirectory = Application.StartupPath
        x.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        x.StartInfo.CreateNoWindow = True
        x.Start()

        Threading.Thread.Sleep(10000)

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
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE", "Project")
        OfficeApps.Add("C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE", "Visio")

        For Each exe As KeyValuePair(Of String, String) In OfficeApps
            CreateShortcut(exe.Key, Environment.GetFolderPath(Environment.SpecialFolder.Desktop), exe.Value)
        Next
    End Sub

    Sub Cleanup()
        Dim SetupPath As String = Path.Combine(Main.TempPath, Main.SetupFileName)
        Dim ConfigPath As String = Path.Combine(Main.TempPath, Main.ConfigFileName)

        Try
            If File.Exists(SetupPath) Then
                File.Delete(SetupPath)
            End If
        Catch ex As Exception

        End Try

        Try
            If File.Exists(ConfigPath) Then
                File.Delete(ConfigPath)
            End If
        Catch ex As Exception

        End Try
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