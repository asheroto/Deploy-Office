Imports System.IO
Imports System.IO.Compression

Public Class Main
    Public ApplicationName = "Deploy Office"

    Public Const AssetsFilename As String = "Assets.zip"
    Public Const SetupFilename As String = "setup.exe"
    Public Const ConfigFilename As String = "configuration.xml"
    Public ProductID As New Dictionary(Of String, String)
    Public TempPath As String = Path.GetTempPath

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Adjust height
        Me.Height = 145

        'Configure product ID dictionary
        With ProductID
            .Add("Microsoft 365 Family/Personal", "O365HomePremRetail")
            .Add("Microsoft 365 Small Business", "O365SmallBusPremRetail")
            .Add("Microsoft 365 Education", "O365EduCloudRetail")
            .Add("Microsoft 365 Enterprise", "O365ProPlusRetail")
            .Add("Home & Business", "HomeBusiness{YYYY}Retail")
            .Add("Home & Student", "HomeStudent{YYYY}Retail")
            .Add("Personal", "Personal{YYYY}Retail")
            .Add("Professional", "Professional{YYYY}Retail")
            .Add("Professional Plus", "ProPlus{YYYY}Retail")
            .Add("Professional Plus - Volume", "ProPlus{YYYY}Volume")
            .Add("Standard", "Standard{YYYY}Retail")
            .Add("Standard - Volume", "Standard{YYYY}Volume")
            .Add("Visio Standard", "VisioStd{YYYY}Retail")
            .Add("Visio Standard - Volume", "VisioStd{YYYY}Volume")
            .Add("Visio Professional", "VisioPro{YYYY}Retail")
            .Add("Visio Professional - Volume", "VisioPro{YYYY}Volume")
            .Add("Project Standard", "ProjectStd{YYYY}Retail")
            .Add("Project Standard - Volume", "ProjectStd{YYYY}Volume")
            .Add("Project Professional", "ProjectPro{YYYY}Retail")
            .Add("Project Professional - Volume", "ProjectPro{YYYY}Volume")
            .Add("Access", "Access{YYYY}Retail")
            .Add("Access - Volume", "Access{YYYY}Volume")
            .Add("Excel", "Excel{YYYY}Retail")
            .Add("Excel - Volume", "Excel{YYYY}Volume")
            .Add("Outlook", "Outlook{YYYY}Retail")
            .Add("Outlook - Volume", "Outlook{YYYY}Volume")
            .Add("PowerPoint", "PowerPoint{YYYY}Retail")
            .Add("PowerPoint - Volume", "PowerPoint{YYYY}Volume")
            .Add("Publisher", "Publisher{YYYY}Retail")
            .Add("Publisher - Volume", "Publisher{YYYY}Volume")
            .Add("Word", "Word{YYYY}Retail")
            .Add("Word - Volume", "Word{YYYY}Volume")
        End With

        'Configure drop-down
        For Each item In ProductID
            EditionSelector.Items.Add(item.Key)
        Next

        'Default edition selection
        Try
            EditionSelector.SelectedItem = "Professional Plus - Volume"
        Catch ex As Exception

        End Try

        'Deploy config edition selection
        Dim DeployConfigPath As String = Path.Combine(Application.StartupPath, "Deploy-Office.txt")
        Dim DeployConfig() As String = Nothing
        If File.Exists(DeployConfigPath) Then
            Try
                DeployConfig = File.ReadAllLines(DeployConfigPath)
                Dim DeployConfigProduct As String = DeployConfig(0)
                If DeployConfig(0).StartsWith("2019-") Then
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False
                    DeployConfigProduct = DeployConfig(0).Replace("2019-", String.Empty)
                ElseIf DeployConfig(0).StartsWith("2021-") Then
                    RadioButton1.Checked = False
                    RadioButton2.Checked = True
                    DeployConfigProduct = DeployConfig(0).Replace("2021-", String.Empty)
                Else
                    Throw New Exception
                End If
                If DeployConfigProduct.Length > 0 Then
                    If DeployConfigProduct <= ProductID.Count - 1 Then
                        EditionSelector.SelectedIndex = DeployConfigProduct
                    End If
                End If
            Catch ex As Exception

            End Try
        End If

        'Fix TLS 1.2 not enabled
        TLSchannelFix()
    End Sub

    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles CountdownTimer.Tick
        CountdownLabel.Text = Integer.Parse(CountdownLabel.Text) - 1

        If CountdownLabel.Text = 0 Then
            Start()
        End If
    End Sub
    Sub Start()
        Button_Start.Visible = False
        CountdownTimer.Enabled = False
        EditionSelector.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        GoRun()
    End Sub

    Sub GoRun()
        Try
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
            Dim YYYY As String = Nothing
            If RadioButton1.Checked = True Then
                YYYY = "2019"
            ElseIf RadioButton2.Checked = True Then
                YYYY = "2021"
            End If
            ProductIDValue = ProductIDValue.Replace("{YYYY}", YYYY)
            If EditionSelector.Text.Contains("Volume") Then
                Configuration = Configuration.Replace("{CHANNEL}", "PerpetualVL" & YYYY)
            Else
                Configuration = Configuration.Replace("{CHANNEL}", "Current")
            End If
            Configuration = Configuration.Replace("{PRODUCTID}", ProductIDValue)
            File.WriteAllText(ConfigPath, Configuration)

            Shell("notepad " & ConfigPath)
            End
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
            WaitOneMinute()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'End
        End
    End Sub

    Private Sub EditionSelector_SelectedIndexChanged(sender As Object, e As EventArgs) Handles EditionSelector.SelectedIndexChanged
        CountdownTimer.Enabled = False
        CountdownLabel.Text = 30
        CountdownTimer.Enabled = True
    End Sub

    Private Sub EditionSelector_DropDown(sender As Object, e As EventArgs) Handles EditionSelector.DropDown
        CountdownTimer.Enabled = False
        CountdownLabel.Text = "Paused"
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        CountdownTimer.Enabled = False
        CountdownLabel.Text = "Paused"
    End Sub

    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        Start()
    End Sub
End Class