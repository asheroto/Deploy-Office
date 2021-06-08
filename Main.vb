Imports System.IO
Imports System.Net

Public Class Main
    Const SetupFileName As String = "setup.exe"
    Const ConfigFileName As String = "configuration.xml"
    Const SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/raw/master/setup.exe"
    Const ConfigURL As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/configuration.xml"
    Dim TempPath As String = Path.GetTempPath

    Sub Download(URL As String, FileName As String)
        Dim DownloadPath As String = Path.Combine(TempPath, FileName)

        If File.Exists(DownloadPath) Then
            Try
                File.Delete(DownloadPath)
            Catch ex As Exception
                MsgBox(String.Format("Failed to delete existing {0}. Please end all setup.exe processes or restart your computer and try again." & vbNewLine & ex.Message, FileName), vbExclamation)
                End
            End Try

        End If

        Try
            Dim wc As WebClient = New WebClient()
            wc.DownloadFile(URL, DownloadPath)
        Catch ex As Exception
            MsgBox(String.Format("Failed to download {0}. Please ensure you have Internet connectivity or restart your computer and try again." & vbNewLine & ex.Message, FileName), vbExclamation)
            End
        End Try

        If File.Exists(DownloadPath) = False Then
            MsgBox(String.Format("Problem downloading {0}. Please ensure you have Internet connectivity or restart your computer and try again.", FileName), vbExclamation)
        End If
    End Sub

    Sub RunSetup()
        Dim x As New Process
        x.StartInfo.FileName = Path.Combine(Application.StartupPath, SetupFileName)
        x.StartInfo.Arguments = "/configure configuration.xml"
        x.StartInfo.WorkingDirectory = Application.StartupPath
        x.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        x.StartInfo.CreateNoWindow = True
        x.Start()
        x.WaitForExit()
    End Sub

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
        Dim SetupPath As String = Path.Combine(TempPath, SetupFileName)
        Dim ConfigPath As String = Path.Combine(TempPath, ConfigFileName)

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
        LogTextbox.AppendText(text & "..." & vbNewLine)
    End Sub

    Sub GoRun()
        CountdownLabel.Text = "Running..."
        Application.DoEvents()

        'Download setup.exe
        LogAppend("Downloading setup.exe")
        Download(SetupURL, SetupFileName)

        'Download configuration.xml
        LogAppend("Downloading configuration.xml")
        Download(ConfigURL, ConfigFileName)

        'Run setup
        LogAppend("Running setup")
        RunSetup()

        'Create desktop shortcuts
        LogAppend("Creating desktop shortcuts")
        CreateDesktopShortcuts()

        'Cleanup
        LogAppend("Cleaning up")
        Cleanup()

        'Finished
        LogAppend("Finished")
        CountdownLabel.Text = "Finished"
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Enforce TLS secure channel
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"" /v ""DisabledByDefault"" /t REG_DWORD /d ""0"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"" /v ""Enabled"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"" /v ""DisabledByDefault"" /t REG_DWORD /d ""0"" /f", vbHidden)
        Shell("reg add ""HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"" /v ""Enabled"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SOFTWARE\Microsoft\.NETFramework\v4.0.30319"" /v ""SchUseStrongCrypto"" /t REG_DWORD /d ""1"" /f", vbHidden)
        Shell("reg add ""HKLM\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"" /v ""SchUseStrongCrypto"" /t REG_DWORD /d ""1"" /f", vbHidden)
    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
        CountdownLabel.Text = Integer.Parse(CountdownLabel.Text) - 1

        If CountdownLabel.Text = 0 Then
            CountdownTimer.Enabled = False

            GoRun()
        End If
    End Sub

End Class