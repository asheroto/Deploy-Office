Imports System.IO
Imports System.Net

Public Class Main
    Const SetupFileName As String = "setup.exe"
    Const ConfigFileName As String = "configuration.xml"
    Const SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/raw/master/setup.exe"
    Const ConfigURL As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/configuration.xml"

    Sub Download(URL As String, FileName As String)
        Dim DownloadPath As String = Path.Combine(Application.StartupPath, FileName)

        If File.Exists(DownloadPath) Then
            Try
                File.Delete(DownloadPath)
            Catch ex As Exception
                MsgBox(String.Format("Failed to delete existing {0}. Please end all setup.exe processes or restart your computer and try again.", FileName), vbExclamation)
                End
            End Try

        End If

        Try
            Dim wc As WebClient = New WebClient()
            wc.DownloadFile(URL, DownloadPath)
        Catch ex As Exception
            MsgBox(String.Format("Failed to download {0}. Please ensure you have Internet connectivity or restart your computer and try again.", FileName), vbExclamation)
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
        x.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        x.Start()
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

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
        CountdownLabel.Text = Integer.Parse(CountdownLabel.Text) - 1

        If CountdownLabel.Text = 0 Then
            CountdownTimer.Enabled = False

            CountdownLabel.Text = "Running..."

            'Download setup.exe
            Download(SetupURL, SetupFileName)

            'Download configuration.xml
            Download(ConfigURL, ConfigFileName)

            'Run setup
            RunSetup()

            'Create desktop shortcuts
            Dim OfficeApps() As String =
                {
                    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
                    "C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE",
                }
            End
        End If
    End Sub

End Class