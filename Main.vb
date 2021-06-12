Imports System.IO

Public Class Main
    Public ApplicationName = "Deploy Office 2019"

    Public Const SetupFileName As String = "setup.exe"
    Public Const ConfigFileName As String = "configuration.xml"
    Public Const SetupURL As String = "https://github.com/asheroto/Deploy-Office-2019/raw/master/setup.exe"
    Public Const ConfigURL_AllPlusVisio As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/Configurations/ConfigurationAllPlusVisio.xml"
    Public Const ConfigURL_All As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/Configurations/ConfigurationAll.xml"
    Public Const ConfigURL_PowerPointOnly As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/Configurations/ConfigurationPowerPointOnly.xml"
    Public Const ConfigURL_WordExcelPowerPointOutlookOnly As String = "https://raw.githubusercontent.com/asheroto/Deploy-Office-2019/master/Configurations/ConfigurationWordExcelPowerPointOutlookOnly.xml"
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
        Dim ConfigURL As String = Nothing
        Select Case EditionSelector.SelectedItem.ToString
            Case "All & Visio"
                ConfigURL = ConfigURL_AllPlusVisio
            Case "All"
                ConfigURL = ConfigURL_All
            Case "PowerPoint Only"
                ConfigURL = ConfigURL_PowerPointOnly
            Case "Word, Excel, PowerPoint, Outlook Only"
                ConfigURL = ConfigURL_WordExcelPowerPointOutlookOnly
        End Select
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
        CloseWindowAfterOneMinute()
        End

    End Sub

End Class