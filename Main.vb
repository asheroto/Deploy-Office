Imports System.IO

Public Class Main
    Public ApplicationName = "Deploy Office 2019"

    Public Const SetupFileName As String = "setup.exe"
    Public Const ConfigFileName As String = "configuration.xml"
    Public Const SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/raw/master/setup.exe"
    Public Const ConfigURL As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/configuration.xml"
    Public TempPath As String = Path.GetTempPath

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Adjust height
        Me.Height = 145

        'Fix TLS 1.2 not enabled
        TLSchannelFix()
    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
        CountdownLabel.Text = Integer.Parse(CountdownLabel.Text) - 1

        If CountdownLabel.Text = 0 Then
            CountdownTimer.Enabled = False
            GoRun()
        End If
    End Sub

    Sub GoRun()
        'Adjust form
        CountdownLabel.Text = "Running..."
        Me.Height = 500
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

        'Closing window
        LogAppend("Closing this window in 1 minute")
        Dim c As New Stopwatch
        c.Start()
        Do Until c.Elapsed.Minutes >= 1
            Application.DoEvents()
        Loop
        c.Stop()
        End

    End Sub

End Class