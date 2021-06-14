Imports System.IO
Imports System.IO.Compression

Public Class Main
    Public ApplicationName = "Deploy Office 2019"

    Public Const AssetsFilename As String = "assets.zip"
    Public Const SetupFilename As String = "setup.exe"
    Public Const ConfigFilename As String = "configuration.xml"
    Public ProductID As New Dictionary(Of String, String)
    Public TempPath As String = Path.GetTempPath

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Adjust height
        Me.Height = 145

        'Default selection
        EditionSelector.SelectedItem = "Professional Plus - Volume"

        'Configure product ID dictionary
        With ProductID
            .Add("Home & Business", "HomeBusiness2019Retail")
            .Add("Home & Student", "HomeStudent2019Retail")
            .Add("Personal", "Personal2019Retail")
            .Add("Professional", "Professional2019Retail")
            .Add("Professional Plus", "ProPlus2019Retail")
            .Add("Professional Plus - Volume", "ProPlus2019Volume")
            .Add("Standard", "Standard2019Retail")
            .Add("Standard - Volume", "Standard2019Volume")
            .Add("Visio Standard", "VisioStd2019Retail")
            .Add("Visio Standard - Volume", "VisioStd2019Volume")
            .Add("Visio Professional", "VisioPro2019Retail")
            .Add("Visio Professional - Volume", "VisioPro2019Volume")
        End With

        'Fix TLS 1.2 not enabled
        TLSchannelFix()
    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
        CountdownLabel.Text = Integer.Parse(CountdownLabel.Text) - 1

        If CountdownLabel.Text = 0 Then
            CountdownTimer.Enabled = False
            EditionSelector.Enabled = False
            GoRun()
        End If
    End Sub

    Sub GoRun()
        'Adjust form
        CountdownLabel.Text = "Running..."
        Me.Height = 500
        Application.DoEvents()

        'Prepare folder
        LogAppend("Preparing")
        Cleanup()

        'Extract assets
        LogAppend("Extracting assets")
        Dim AssetsPath As String = Path.Combine(TempPath, AssetsFilename)
        File.WriteAllBytes(AssetsPath, My.Resources.Assets)
        ZipFile.ExtractToDirectory(AssetsPath, TempPath)

        'Write configuration.xml
        LogAppend("Writing configuration.xml")
        Dim ConfigPath As String = Path.Combine(TempPath, ConfigFilename)
        Dim Configuration As String = File.ReadAllText(ConfigPath)
        Dim ProductIDValue As String = Nothing
        ProductID.TryGetValue(EditionSelector.Text, ProductIDValue)
        Configuration = Configuration.Replace("{PRODUCTID}", ProductIDValue)
        File.WriteAllText(ConfigPath, Configuration)

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

        'End
        End

    End Sub

    Private Sub EditionSelector_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EditionSelector.SelectedIndexChanged
        CountdownTimer.Enabled = False
        CountdownLabel.Text = 10
        CountdownTimer.Enabled = True
    End Sub

    Private Sub EditionSelector_DropDown(sender As Object, e As EventArgs) Handles EditionSelector.DropDown
        CountdownTimer.Enabled = False
        CountdownLabel.Text = "Paused"
    End Sub
End Class