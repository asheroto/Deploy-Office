Imports System.IO
Imports System.Xml

Public Class Main
    Public Const ApplicationName As String = "Deploy Office"
    Public Const ErrorLog As String = "Deploy-Office.log"
    Public Const AutomatedDeployment As String = "Deploy-Office.ini"
    Public Const OldAutomatedDeployment As String = "Deploy-Office.txt"
    Public Const SetupFilename As String = "setup.exe"
    Public Const ConfigFilename As String = "configuration.xml"
    Public Const SetupURL As String = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Public Const DefaultYear As String = "2021"
    Public Const DefaultEdition As String = "Professional Plus - Volume"
    Public CountdownValue As Integer = 30
    Public ProductID As New Dictionary(Of String, String)
    Public TempPath As String = Path.GetTempPath
    Public SetupPath As String
    Public ConfigPath As String

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Config vars
        SetupPath = Path.Combine(TempPath, SetupFilename)
        ConfigPath = Path.Combine(TempPath, ConfigFilename)

        'Prep comboboxes
        ComboBox_Edition.DrawMode = DrawMode.OwnerDrawFixed
        AddHandler ComboBox_Edition.DrawItem, AddressOf DrawItem_Edition

        'Set year default to 2021
        ComboBox_Year.SelectedItem = DefaultYear

        'Configure year drop-down
        ComboBox_Year.Items.Add("")
        ComboBox_Year.Items.Add("2019")
        ComboBox_Year.Items.Add("2021")
        ComboBox_Year.Items.Add("2024")
        ComboBox_Year.Items.Add("Microsoft 365")

        ' Initialize the product ID dictionary with key-value pairs
        ProductID = New Dictionary(Of String, String) From {
            {"Microsoft 365 Family/Personal", "O365HomePremRetail"},
            {"Microsoft 365 Small Business", "O365SmallBusPremRetail"},
            {"Microsoft 365 Education", "O365EduCloudRetail"},
            {"Microsoft 365 Enterprise", "O365ProPlusRetail"},
            {"Home & Business", "HomeBusiness{YYYY}Retail"},
            {"Home & Student", "HomeStudent{YYYY}Retail"},
            {"Personal", "Personal{YYYY}Retail"},
            {"Professional", "Professional{YYYY}Retail"},
            {"Professional Plus", "ProPlus{YYYY}Retail"},
            {"Professional Plus - Volume", "ProPlus{YYYY}Volume"},
            {"Standard", "Standard{YYYY}Retail"},
            {"Standard - Volume", "Standard{YYYY}Volume"},
            {"Visio Standard", "VisioStd{YYYY}Retail"},
            {"Visio Standard - Volume", "VisioStd{YYYY}Volume"},
            {"Visio Professional", "VisioPro{YYYY}Retail"},
            {"Visio Professional - Volume", "VisioPro{YYYY}Volume"},
            {"Project Standard", "ProjectStd{YYYY}Retail"},
            {"Project Standard - Volume", "ProjectStd{YYYY}Volume"},
            {"Project Professional", "ProjectPro{YYYY}Retail"},
            {"Project Professional - Volume", "ProjectPro{YYYY}Volume"},
            {"Access", "Access{YYYY}Retail"},
            {"Access - Volume", "Access{YYYY}Volume"},
            {"Excel", "Excel{YYYY}Retail"},
            {"Excel - Volume", "Excel{YYYY}Volume"},
            {"Outlook", "Outlook{YYYY}Retail"},
            {"Outlook - Volume", "Outlook{YYYY}Volume"},
            {"PowerPoint", "PowerPoint{YYYY}Retail"},
            {"PowerPoint - Volume", "PowerPoint{YYYY}Volume"},
            {"Publisher", "Publisher{YYYY}Retail"},
            {"Publisher - Volume", "Publisher{YYYY}Volume"},
            {"Word", "Word{YYYY}Retail"},
            {"Word - Volume", "Word{YYYY}Volume"}
        }

        ' Configure drop-down
        ComboBox_Edition.Items.Clear()
        ComboBox_Edition.Items.Add("")
        For Each kvp As KeyValuePair(Of String, String) In ProductID
            ComboBox_Edition.Items.Add(kvp.Key)
        Next

        'Default year selection
        ComboBox_Year.SelectedItem = DefaultYear

        'Default edition selection
        ComboBox_Edition.SelectedItem = DefaultEdition

        'Handle Deploy-Office.ini
        HandleDeployOfficeIni()

        'Fix TLS 1.2 not enabled
        TLSchannelFix()
    End Sub

    Private Sub DrawItem_Edition(sender As Object, e As DrawItemEventArgs)
        If e.Index < 0 Then Return

        Dim comboBox As ComboBox = CType(sender, ComboBox)
        Dim itemText As String = comboBox.Items(e.Index).ToString()
        Dim yearSelected As String = ComboBox_Year.SelectedItem?.ToString()
        Dim isDisabled As Boolean = False

        If yearSelected = "Microsoft 365" Then
            ' Disable items that do not contain "Microsoft 365"
            isDisabled = Not itemText.Contains("Microsoft 365")
        ElseIf Not String.IsNullOrEmpty(yearSelected) Then
            ' Disable items that contain "Microsoft 365" when a specific year is selected
            isDisabled = itemText.Contains("Microsoft 365")
        End If

        ' Drawing item background
        e.DrawBackground()

        ' Set text color based on disabled state
        Dim itemColor As Brush = If(isDisabled, Brushes.Gray, Brushes.Black)
        e.Graphics.DrawString(itemText, e.Font, itemColor, e.Bounds)

        ' Draw focus rectangle if item has focus
        e.DrawFocusRectangle()
    End Sub

    Private Sub HandleDeployOfficeIni()
        ' Check for deprecated configuration file
        Dim oldDeployConfigPath As String = Path.Combine(Application.StartupPath, OldAutomatedDeployment)
        If File.Exists(oldDeployConfigPath) Then
            LogError("Deploy-Office.txt file is detected and is no longer supported. Please use Deploy-Office.ini instead. Note that the product indexing has changed. Refer to the README on the repository for details and instructions on updating your configuration.")
        End If

        ' Set the path for the current deployment configuration
        Dim deployConfigPath As String = Path.Combine(Application.StartupPath, AutomatedDeployment)
        If File.Exists(deployConfigPath) Then
            Try
                ' Read configuration from the INI file
                Dim deployConfig As Dictionary(Of String, String) = ReadIniFile(deployConfigPath)

                ' Validate and set the year
                Dim yearValue As String = GetValueFromConfig(deployConfig, "year")
                SetComboBoxValue(ComboBox_Year, yearValue)

                ' Validate and set the edition
                Dim editionValue As String = GetValueFromConfig(deployConfig, "edition")
                SetComboBoxValue(ComboBox_Edition, editionValue)

                ' Handle shortcuts setting
                Dim shortcutsValue As String = GetValueFromConfig(deployConfig, "shortcuts", False)
                If shortcutsValue IsNot Nothing Then
                    CheckBox_CreateDesktopShortcuts.Checked = IsTrueValue(shortcutsValue)
                End If

                'Exclude Teams
                Dim excludeTeamsValue As String = GetValueFromConfig(deployConfig, "exclude_teams", False)
                If excludeTeamsValue IsNot Nothing Then
                    CheckBox_ExcludeTeams.Checked = IsTrueValue(excludeTeamsValue)
                End If

                'Exclude OneDrive
                Dim excludeOneDriveValue As String = GetValueFromConfig(deployConfig, "exclude_onedrive", False)
                If excludeOneDriveValue IsNot Nothing Then
                    CheckBox_ExcludeOneDrive.Checked = IsTrueValue(excludeOneDriveValue)
                End If

            Catch ex As Exception
                ' Log any errors encountered during configuration handling
                LogError("[Deploy-Office.ini] " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub SetComboBoxValue(comboBox As ComboBox, value As String)
        ' First, check if the value is numeric and treat it differently based on length or value
        Dim index As Integer
        If Integer.TryParse(value, index) Then
            ' Check if it looks like a year (four-digit number)
            If value.Length = 4 Then
                ' Attempt to match as a year
                Dim yearMatchIndex As Integer = comboBox.Items.Cast(Of String)().ToList().FindIndex(Function(item) item.Equals(value))
                If yearMatchIndex <> -1 Then
                    comboBox.SelectedIndex = yearMatchIndex
                    Return
                End If
            Else
                ' Adjust for 1-based indexing if necessary (remove if using 0-based)
                If index >= 0 AndAlso index < comboBox.Items.Count Then
                    comboBox.SelectedIndex = index
                    Return
                Else
                    ' Log error if index is out of range
                    LogError("[Deploy-Office.ini] Index value is out of range: " & value)
                    Return
                End If
            End If
        End If

        ' If not a year or index, attempt to match it case-insensitively with the items
        Dim items As List(Of String) = comboBox.Items.Cast(Of String)().ToList()
        Dim matchIndex As Integer = items.FindIndex(Function(item) String.Equals(item, value, StringComparison.OrdinalIgnoreCase))
        If matchIndex <> -1 Then
            comboBox.SelectedIndex = matchIndex
        Else
            ' Log error if no match is found
            LogError("[Deploy-Office.ini] Value not found in ComboBox: " & value)
        End If
    End Sub

    Private Function IsTrueValue(value As String) As Boolean
        ' Check if the value represents a true state
        Return value.ToLower() = "true" OrElse value = "1"
    End Function

    Private Function GetValueFromConfig(config As Dictionary(Of String, String), key As String, Optional ErrorOnMissing As Boolean = True) As String
        Dim value As String = Nothing
        If config.TryGetValue(key, value) Then
            Return value
        ElseIf ErrorOnMissing Then
            Throw New KeyNotFoundException("Key not found in the INI file: " & key)
        Else
            Return Nothing
        End If
    End Function

    Sub DisableControls()
        ComboBox_Edition.Enabled = False
        ComboBox_Year.Enabled = False
        CheckBox_CreateDesktopShortcuts.Enabled = False
        CheckBox_ExcludeTeams.Enabled = False
        CheckBox_ExcludeOneDrive.Enabled = False
        Button_Start.Enabled = False
    End Sub

    Sub Start()
        DisableControls()

        'Disable countdown
        Timer_Countdown.Enabled = False
        CountdownLabel.Text = "Running"

        Application.DoEvents()

        GoRun()
    End Sub

    Sub GoRun()
        Try
            ' Prepare folder
            OutputAppend("Preparing")
            Cleanup()

            ' Download setup.exe
            OutputAppend("Downloading setup.exe from Microsoft")
            Using wc As New Net.WebClient
                wc.DownloadFile(SetupURL, SetupPath)
            End Using

            ' Extract configuration.xml from resources
            OutputAppend("Extracting configuration.xml")
            ' Read the configuration XML into a variable directly from the resource
            Dim Configuration As String = My.Resources.configuration_xml
            Dim ProductIDValue As String = Nothing

            ' Retrieve the selected edition and year
            ProductID.TryGetValue(ComboBox_Edition.Text, ProductIDValue)
            Dim SelectedYear As String = ComboBox_Year.SelectedItem.ToString()

            ' Replace placeholders in the configuration with actual values
            If SelectedYear <> "Microsoft 365" Then
                ProductIDValue = ProductIDValue.Replace("{YYYY}", SelectedYear)
                If ComboBox_Edition.Text.Contains("Volume") Then
                    Configuration = Configuration.Replace("{CHANNEL}", "PerpetualVL" & SelectedYear)
                Else
                    Configuration = Configuration.Replace("{CHANNEL}", "Current")
                End If
            End If
            Configuration = Configuration.Replace("{PRODUCTID}", ProductIDValue)

            ' Load the XML configuration into an XmlDocument
            Dim xmlDoc As New XmlDocument()
            xmlDoc.LoadXml(Configuration)

            ' Create XmlWriterSettings to omit the XML declaration
            Dim settings As New XmlWriterSettings()
            settings.OmitXmlDeclaration = True ' Omit the XML declaration line

            ' Check if CheckBoxes are checked and create the exclusion nodes
            Dim productNode As XmlNode = xmlDoc.SelectSingleNode("/Configuration/Add/Product")

            If CheckBox_ExcludeOneDrive.Checked Then
                Dim onedriveNode As XmlElement = xmlDoc.CreateElement("ExcludeApp")
                onedriveNode.SetAttribute("ID", "OneDrive")
                productNode.AppendChild(onedriveNode)
            End If

            If CheckBox_ExcludeTeams.Checked Then
                Dim teamsNode As XmlElement = xmlDoc.CreateElement("ExcludeApp")
                teamsNode.SetAttribute("ID", "Teams")
                productNode.AppendChild(teamsNode)
            End If

            ' Convert the XmlDocument back to a string without the XML declaration
            Dim configurationWithoutDeclaration As String
            Using stringWriter As New StringWriter()
                Using writer As XmlWriter = XmlWriter.Create(stringWriter, settings)
                    xmlDoc.Save(writer)
                End Using
                configurationWithoutDeclaration = stringWriter.ToString()
            End Using

            ' Write the modified configuration to the file
            Dim ConfigResourcePath As String = Path.Combine(TempPath, ConfigFilename)
            File.WriteAllText(ConfigResourcePath, configurationWithoutDeclaration)

            ' Uncomment if you want to view the configuration file before proceeding
            'Shell("notepad " & ConfigPath, AppWinStyle.NormalFocus)

            ' Run setup
            OutputAppend("Running setup")
            OutputAppend("Setup may run in background (system tray)")
            RunSetup()

            ' Create desktop shortcuts if needed
            If CheckBox_CreateDesktopShortcuts.Checked Then
                OutputAppend("Creating desktop shortcuts")
                ShortcutHelper.CreateDesktopShortcuts()
            End If

            ' Cleanup
            OutputAppend("Cleaning up")
            Cleanup()

            ' Finished
            OutputAppend("Finished")
            CountdownLabel.Text = "Finished"

            ' Closing window
            OutputAppend("Closing this window in 30 seconds")
            WaitToClose()
            Application.Exit()

        Catch ex As Exception
            LogError(ex.Message)
        End Try
    End Sub

    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        Start()
    End Sub

    Private Sub ComboBox_Year_DropDown(sender As Object, e As EventArgs) Handles ComboBox_Year.DropDown
        PauseCountdown()
    End Sub

    Private Sub ComboBox_Year_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Year.SelectedIndexChanged
        ResetCountdown()

        ' Check if the selected year is empty and enable/disable the edition ComboBox accordingly
        If ComboBox_Year.SelectedItem Is Nothing OrElse String.IsNullOrWhiteSpace(ComboBox_Year.SelectedItem.ToString()) Then
            ComboBox_Edition.Enabled = False
            ComboBox_Edition.SelectedIndex = -1 ' Optionally clear the selection
        Else
            ComboBox_Edition.Enabled = True
        End If

        ' Refresh to update the disabled state of items in the edition ComboBox
        ComboBox_Edition.Refresh()

        CanStartBeEnabled()
    End Sub

    Private Sub ComboBox_Edition_DropDown(sender As Object, e As EventArgs) Handles ComboBox_Edition.DropDown
        PauseCountdown()

        Button_Start.Focus()
    End Sub

    Private Sub ComboBox_Edition_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Edition.SelectedIndexChanged
        ResetCountdown()

        If ComboBox_Edition.SelectedIndex < 0 Then Return

        Dim itemText As String = ComboBox_Edition.SelectedItem.ToString()
        Dim yearSelected As String = ComboBox_Year.SelectedItem.ToString()
        Dim isDisabled As Boolean = False

        If yearSelected = "Microsoft 365" Then
            isDisabled = Not itemText.Contains("Microsoft 365")
        ElseIf Not String.IsNullOrEmpty(yearSelected) Then
            isDisabled = itemText.Contains("Microsoft 365")
        End If

        If isDisabled Then
            ComboBox_Edition.SelectedIndex = -1 ' Reset or set to a valid index

            LogError("The selected edition is not compatible with the selected year.")

            PauseCountdown()
        End If

        CanStartBeEnabled()

        Button_Start.Focus()
    End Sub

    Public Sub PauseCountdown()
        Timer_Countdown.Enabled = False
        CountdownLabel.Text = "Paused"
    End Sub

    Public Sub ResetCountdown()
        Timer_Countdown.Enabled = False
        CountdownLabel.Text = CountdownValue.ToString()
        Timer_Countdown.Enabled = True
    End Sub

    Sub CanStartBeEnabled()
        If ComboBox_Edition.SelectedIndex > 0 And ComboBox_Year.SelectedIndex > 0 Then
            Button_Start.Enabled = True
        Else
            Button_Start.Enabled = False
        End If
    End Sub

    Private Sub Timer_Countdown_Tick(sender As Object, e As EventArgs) Handles Timer_Countdown.Tick
        CountdownValue -= 1
        CountdownLabel.Text = CountdownValue.ToString()

        If CountdownValue = 0 Then
            Timer_Countdown.Stop() ' Stop the timer when countdown reaches zero
            Start() ' Call your Start method or any other action for when the countdown finishes
        End If
    End Sub

    Private Async Function CheckForUpdates() As Task(Of String)
        Dim latestVersion As String = Nothing

        Try
            ' Check for updates
            Dim updater As New GitHubRepoUpdater("asheroto", "Deploy-Office")
            latestVersion = Await updater.CheckForLatestReleaseAsync()
            Debug.WriteLine($"Latest release: {latestVersion}")
        Catch ex As Exception
            Debug.WriteLine("[CheckForUpdates] " & ex.Message)
        End Try

        Return latestVersion
    End Function

    Private Async Sub LinkLabel_CheckForUpdates_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_CheckForUpdates.LinkClicked
        LinkLabel_CheckForUpdates.Text = "Checking..."

        ' Call CheckForUpdates asynchronously
        Dim latestVersion As String = Await CheckForUpdates()

        LinkLabel_CheckForUpdates.Text = "Check for Updates"

        If latestVersion IsNot Nothing AndAlso New Version(latestVersion) > New Version(Application.ProductVersion) Then
            Dim result As DialogResult = MessageBox.Show($"A new version of {ApplicationName} is available. Would you like to download it now?", ApplicationName, MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If result = DialogResult.Yes Then
                ' Open the releases page in the default browser
                TryStartProcess("https://github.com/asheroto/Deploy-Office/releases")
            End If
        Else
            MessageBox.Show($"No updates found. You are running the latest version of {ApplicationName}.", ApplicationName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub LinkLabel_Repo_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_Repo.LinkClicked
        ' Open repo homepage in the default browser
        TryStartProcess("https://github.com/asheroto/Deploy-Office")
    End Sub

    Private Sub Main_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        End
    End Sub
End Class