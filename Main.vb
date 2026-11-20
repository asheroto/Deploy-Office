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
    Public Const DefaultYear As String = "Microsoft 365"
    Public Const DefaultEdition As String = "Microsoft 365 Enterprise"
    Public ProductID As New Dictionary(Of String, String)
    ' Display name -> ExcludeApp ID used by the Office Deployment Tool, filled in Main_Load
    Public AppIDs As New Dictionary(Of String, String)
    Public TempPath As String = Path.GetTempPath
    Public SetupPath As String
    Public ConfigPath As String
    Public autostart As Boolean
    Private autostartSecondsLeft As Integer = 10
    Private settingsHadErrors As Boolean
    Public autoclose As Boolean
    Public silent As Boolean

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Config vars
        SetupPath = Path.Combine(TempPath, SetupFilename)
        ConfigPath = Path.Combine(TempPath, ConfigFilename)

        'Prep comboboxes
        ComboBox_Edition.DrawMode = DrawMode.OwnerDrawFixed
        ComboBox_Edition.ItemHeight = ComboBox_Year.ItemHeight
        AddHandler ComboBox_Edition.DrawItem, AddressOf DrawItem_Edition

        'Set year default to 365
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

        'Apps to install, all checked by default
        AppIDs = New Dictionary(Of String, String) From {
            {"Word", "Word"},
            {"Excel", "Excel"},
            {"PowerPoint", "PowerPoint"},
            {"Outlook", "Outlook"},
            {"New Outlook", "OutlookForWindows"},
            {"OneNote", "OneNote"},
            {"Access", "Access"},
            {"Publisher", "Publisher"},
            {"Teams", "Teams"},
            {"OneDrive", "OneDrive"},
            {"Skype", "Lync"}
        }
        For Each app As String In AppIDs.Keys
            CheckedListBox_Apps.Items.Add(app, True)
        Next

        'Two links in one label: "Select all" checks every app, "Select none" clears them
        LinkLabel_AppsSelect.Links.Clear()
        LinkLabel_AppsSelect.Links.Add(0, 10, True)
        LinkLabel_AppsSelect.Links.Add(13, 11, False)

        'Shortcuts go to the Public Desktop by default
        ComboBox_ShortcutScope.SelectedIndex = 0

        'Fix TLS 1.2 not enabled
        TLSchannelFix()

        'Layout bits that Windows Forms does not scale with the display: four even columns of apps,
        'exactly three rows tall, and divider lines the same width as the Install button
        CheckedListBox_Apps.ColumnWidth = CheckedListBox_Apps.Width \ 4
        CheckedListBox_Apps.Height = (CheckedListBox_Apps.Font.Height + 2) * 3 + 6
        Dim slack As Integer = Label_Sep2.Top - CheckedListBox_Apps.Bottom - CheckedListBox_Apps.Font.Height \ 2
        For Each c As Control In Controls
            If c.Top > CheckedListBox_Apps.Top Then c.Top -= slack
        Next
        Height -= slack
        ActiveControl = Button_Start
        For Each divider As Label In {Label_Sep1, Label_Sep2, Label_Sep3}
            divider.SetBounds(Button_Start.Left, divider.Top, Button_Start.Width, 1)
        Next

        'Handle Deploy-Office.ini
        HandleDeployOfficeIni()

        'Silent + autostart means no windows at all: hide ours too and exit with the result when done
        If autostart AndAlso silent Then
            autoclose = True
            Opacity = 0
            ShowInTaskbar = False
        End If

        'Only greet when there is nothing more important (a settings error) to show
        If LogTextbox.TextLength = 0 Then LogTextbox.Text = "Ready. Choose your options and click Install."
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
        Dim selected As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Dim itemColor As Color = If(isDisabled, Color.Gray, If(selected, SystemColors.HighlightText, SystemColors.WindowText))
        TextRenderer.DrawText(e.Graphics, itemText, e.Font, e.Bounds, itemColor, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

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
        Dim errorsBefore As Integer = ErrorCount
        Dim wantLog As Boolean = False
        Dim deployConfig As New Dictionary(Of String, String)
        If File.Exists(deployConfigPath) Then deployConfig = ReadIniFile(deployConfigPath)

        ' Command-line parameters use the same names as the INI and win over it
        For Each setting As KeyValuePair(Of String, String) In ParseArguments(Environment.GetCommandLineArgs().Skip(1))
            deployConfig(setting.Key) = setting.Value
        Next

        If deployConfig.Count > 0 Then
            Try
                ' Year and edition keep the window's defaults when not given
                Dim yearValue As String = GetValueFromConfig(deployConfig, "year", False)
                If yearValue IsNot Nothing Then SetComboBoxValue(ComboBox_Year, yearValue)

                Dim editionValue As String = GetValueFromConfig(deployConfig, "edition", False)
                If editionValue IsNot Nothing Then SetComboBoxValue(ComboBox_Edition, editionValue)

                ' Handle shortcuts setting
                Dim shortcutsValue As String = GetValueFromConfig(deployConfig, "shortcuts", False)
                If shortcutsValue IsNot Nothing Then
                    CheckBox_CreateDesktopShortcuts.Checked = IsTrueValue(shortcutsValue)
                End If

                'Shortcut scope: all (Public Desktop) or user (current user's Desktop)
                Dim scopeValue As String = GetValueFromConfig(deployConfig, "shortcuts_scope", False)
                If scopeValue IsNot Nothing Then
                    ComboBox_ShortcutScope.SelectedIndex = If(scopeValue.Trim().ToLower() = "user", 1, 0)
                End If

                'Exclude apps, comma-separated list of app names or ODT IDs
                Dim excludeAppsValue As String = GetValueFromConfig(deployConfig, "exclude_apps", False)
                If excludeAppsValue IsNot Nothing Then
                    For Each app As String In excludeAppsValue.Split(","c)
                        SetAppChecked(app.Trim(), False)
                    Next
                End If

                'Legacy keys, still honored
                Dim excludeTeamsValue As String = GetValueFromConfig(deployConfig, "exclude_teams", False)
                If excludeTeamsValue IsNot Nothing Then
                    SetAppChecked("Teams", Not IsTrueValue(excludeTeamsValue))
                End If
                Dim excludeOneDriveValue As String = GetValueFromConfig(deployConfig, "exclude_onedrive", False)
                If excludeOneDriveValue IsNot Nothing Then
                    SetAppChecked("OneDrive", Not IsTrueValue(excludeOneDriveValue))
                End If

                'Add-ons
                Dim addVisioValue As String = GetValueFromConfig(deployConfig, "add_visio", False)
                If addVisioValue IsNot Nothing Then
                    CheckBox_AddVisio.Checked = IsTrueValue(addVisioValue)
                End If
                Dim addProjectValue As String = GetValueFromConfig(deployConfig, "add_project", False)
                If addProjectValue IsNot Nothing Then
                    CheckBox_AddProject.Checked = IsTrueValue(addProjectValue)
                End If

                'Remove MSI-based Office first
                Dim removeMsiValue As String = GetValueFromConfig(deployConfig, "remove_msi", False)
                If removeMsiValue IsNot Nothing Then
                    CheckBox_RemoveMSI.Checked = IsTrueValue(removeMsiValue)
                End If

                'Unattended options
                Dim silentValue As String = GetValueFromConfig(deployConfig, "silent", False)
                If silentValue IsNot Nothing Then
                    silent = IsTrueValue(silentValue)
                End If
                Dim autoCloseValue As String = GetValueFromConfig(deployConfig, "autoclose", False)
                If autoCloseValue IsNot Nothing Then
                    autoclose = IsTrueValue(autoCloseValue)
                End If

                'Log file
                Dim logValue As String = GetValueFromConfig(deployConfig, "log", False)
                If logValue IsNot Nothing Then
                    wantLog = IsTrueValue(logValue)
                End If

                'Auto Start
                Dim iniAutoStart As String = GetValueFromConfig(deployConfig, "autostart", False)
                If iniAutoStart IsNot Nothing Then
                    autostart = IsTrueValue(iniAutoStart)
                End If

            Catch ex As Exception
                ' Log any errors encountered during configuration handling
                LogError("[Settings] " & ex.Message)
            End Try
        End If

        settingsHadErrors = ErrorCount > errorsBefore

        ' Write Deploy-Office.log only when asked, or when the window will not be there to read afterwards
        If wantLog OrElse silent OrElse autoclose Then EnableLog()
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
                    LogError("[Settings] Index value is out of range: " & value)
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
            LogError("[Settings] Not a valid year or edition: " & value)
        End If
    End Sub

    Private Sub SetAppChecked(app As String, isChecked As Boolean)
        ' Accept either the display name or the ODT ID (e.g. Skype or Lync)
        For i As Integer = 0 To CheckedListBox_Apps.Items.Count - 1
            Dim name As String = CheckedListBox_Apps.Items(i).ToString()
            If String.Equals(name, app, StringComparison.OrdinalIgnoreCase) OrElse String.Equals(AppIDs(name), app, StringComparison.OrdinalIgnoreCase) Then
                CheckedListBox_Apps.SetItemChecked(i, isChecked)
                Return
            End If
        Next
        LogError("[Settings] Unknown app: " & app)
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
        ComboBox_ShortcutScope.Enabled = False
        CheckedListBox_Apps.Enabled = False
        LinkLabel_AppsSelect.Enabled = False
        CheckBox_AddVisio.Enabled = False
        CheckBox_AddProject.Enabled = False
        CheckBox_RemoveMSI.Enabled = False
        Button_Start.Enabled = False
    End Sub

    Sub EnableControls()
        ComboBox_Year.Enabled = True
        ComboBox_Edition.Enabled = True
        CheckBox_CreateDesktopShortcuts.Enabled = True
        ComboBox_ShortcutScope.Enabled = CheckBox_CreateDesktopShortcuts.Checked
        CheckedListBox_Apps.Enabled = True
        LinkLabel_AppsSelect.Enabled = True
        CheckBox_AddVisio.Enabled = True
        CheckBox_AddProject.Enabled = True
        CheckBox_RemoveMSI.Enabled = True
        CanStartBeEnabled()
    End Sub

    Sub Start()
        DisableControls()
        LogTextbox.Clear()
        Button_Start.Text = "Installing..."
        ProgressBar_Install.Visible = True

        Application.DoEvents()

        Dim exitCode As Integer = GoRun()

        ' Unattended runs hand setup's result back to whatever launched us (RMM, script)
        If autoclose Then Environment.Exit(exitCode)

        ProgressBar_Install.Visible = False
        Button_Start.Text = "Install"
        EnableControls()
    End Sub

    Private Sub Button_Start_EnabledChanged(sender As Object, e As EventArgs) Handles Button_Start.EnabledChanged
        ' Accent blue when it can be clicked, gray when it cannot
        Button_Start.BackColor = If(Button_Start.Enabled, Color.FromArgb(0, 103, 192), Color.FromArgb(204, 204, 204))
    End Sub

    ''' <summary>Runs the install. Returns 0 on success, setup.exe's exit code if setup failed, or 1 for any other failure.</summary>
    Function GoRun() As Integer
        Try
            ' Prepare folder
            OutputAppend("Preparing")
            Cleanup()

            ' Download setup.exe
            OutputAppend("Downloading setup.exe from Microsoft")
            Using wc As New Net.WebClient
                ' Download in the background so the window keeps responding
                Dim download As Task = wc.DownloadFileTaskAsync(SetupURL, SetupPath)
                Do Until download.IsCompleted
                    Application.DoEvents()
                    Threading.Thread.Sleep(30)
                Loop
                download.GetAwaiter().GetResult()
            End Using

            ' setup.exe runs elevated, so make sure it really is Microsoft's before running it
            OutputAppend("Verifying setup.exe is signed by Microsoft")
            If Not IsSignedByMicrosoft(SetupPath) Then
                Throw New Exception("setup.exe does not have a valid Microsoft signature, so it was not run.")
            End If

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
            Else
                Configuration = Configuration.Replace("{CHANNEL}", "Current")
            End If
            Configuration = Configuration.Replace("{PRODUCTID}", ProductIDValue)

            ' Load the XML configuration into an XmlDocument
            Dim xmlDoc As New XmlDocument()
            xmlDoc.LoadXml(Configuration)

            ' Create XmlWriterSettings to omit the XML declaration
            Dim settings As New XmlWriterSettings()
            settings.OmitXmlDeclaration = True ' Omit the XML declaration line

            ' Add Visio/Project as extra products, matching the main product's year and license type
            Dim addNode As XmlNode = xmlDoc.SelectSingleNode("/Configuration/Add")
            Dim mainProduct As XmlNode = addNode.SelectSingleNode("Product")
            Dim addOns As New List(Of String)
            If CheckBox_AddVisio.Checked Then addOns.Add("VisioPro")
            If CheckBox_AddProject.Checked Then addOns.Add("ProjectPro")
            For Each addOn As String In addOns
                ' Skip if the main product already is Visio/Project
                If ProductIDValue.StartsWith(If(addOn = "VisioPro", "Visio", "Project")) Then Continue For
                Dim addOnID As String
                If SelectedYear = "Microsoft 365" Then
                    addOnID = addOn & "Retail"
                Else
                    addOnID = addOn & SelectedYear & If(ComboBox_Edition.Text.Contains("Volume"), "Volume", "Retail")
                End If
                Dim newProduct As XmlNode = mainProduct.CloneNode(True)
                CType(newProduct, XmlElement).SetAttribute("ID", addOnID)
                addNode.AppendChild(newProduct)
            Next

            ' Unchecked apps become ExcludeApp entries on every product so the ODT applies them consistently
            For Each productNode As XmlNode In addNode.SelectNodes("Product")
                For Each app As String In AppIDs.Keys
                    If Not CheckedListBox_Apps.CheckedItems.Contains(app) Then
                        Dim excludeNode As XmlElement = xmlDoc.CreateElement("ExcludeApp")
                        excludeNode.SetAttribute("ID", AppIDs(app))
                        productNode.AppendChild(excludeNode)
                    End If
                Next
            Next

            ' Remove MSI-based Office (2016 and earlier) before installing
            If CheckBox_RemoveMSI.Checked Then
                xmlDoc.DocumentElement.AppendChild(xmlDoc.CreateElement("RemoveMSI"))
            End If

            ' Hide the Office setup window
            If silent Then
                CType(xmlDoc.SelectSingleNode("/Configuration/Display"), XmlElement).SetAttribute("Level", "None")
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
            Dim exitCode As Integer = RunSetup()

            If exitCode = 0 Then
                ' Create desktop shortcuts if needed
                If CheckBox_CreateDesktopShortcuts.Checked Then
                    OutputAppend("Creating desktop shortcuts")
                    ShortcutHelper.CreateDesktopShortcuts(ComboBox_ShortcutScope.SelectedIndex = 0)
                End If
            End If

            ' Cleanup
            OutputAppend("Cleaning up")
            Cleanup()

            If exitCode = 0 Then
                OutputAppend("Done. Office installed successfully.")
            Else
                LogError("Office setup failed with exit code " & exitCode & ". Nothing else was changed. Microsoft's setup logs are in " & TempPath & " and C:\Windows\Temp.")
            End If
            Return exitCode

        Catch ex As Exception
            LogError(ex.Message)
            Cleanup()
            Return 1
        End Try
    End Function

    Private Sub Button_Start_Click(sender As Object, e As EventArgs) Handles Button_Start.Click
        ' A click during the autostart countdown cancels it
        If Timer_autostart.Enabled Then
            CancelAutostart()
            Return
        End If

        Start()
    End Sub

    Private Sub CancelAutostart()
        If Not Timer_autostart.Enabled Then Return
        Timer_autostart.Enabled = False
        Button_Start.Text = "Install"
        LogTextbox.Text = "Automatic start canceled. Click Install when you are ready."
    End Sub

    ' Touching any option during the countdown means a person is here and choosing, so stop it
    Private Sub Options_Touched(sender As Object, e As EventArgs) Handles _
            ComboBox_Year.DropDown, ComboBox_Year.SelectionChangeCommitted,
            ComboBox_Edition.DropDown, ComboBox_Edition.SelectionChangeCommitted,
            ComboBox_ShortcutScope.DropDown, ComboBox_ShortcutScope.SelectionChangeCommitted,
            CheckedListBox_Apps.Click, CheckedListBox_Apps.KeyDown,
            CheckBox_AddVisio.Click, CheckBox_AddProject.Click,
            CheckBox_CreateDesktopShortcuts.Click, CheckBox_RemoveMSI.Click
        CancelAutostart()
    End Sub

    Private Sub LinkLabel_AppsSelect_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_AppsSelect.LinkClicked
        CancelAutostart()
        For i As Integer = 0 To CheckedListBox_Apps.Items.Count - 1
            CheckedListBox_Apps.SetItemChecked(i, CBool(e.Link.LinkData))
        Next
    End Sub

    Private Sub ComboBox_Year_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Year.SelectedIndexChanged
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

    Private Sub CheckBox_CreateDesktopShortcuts_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_CreateDesktopShortcuts.CheckedChanged
        ComboBox_ShortcutScope.Enabled = CheckBox_CreateDesktopShortcuts.Checked
    End Sub

    Private Sub ComboBox_Edition_DropDown(sender As Object, e As EventArgs) Handles ComboBox_Edition.DropDown
        Button_Start.Focus()
    End Sub

    Private Sub ComboBox_Edition_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Edition.SelectedIndexChanged

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
        End If

        CanStartBeEnabled()

        Button_Start.Focus()
    End Sub

    Sub CanStartBeEnabled()
        If ComboBox_Edition.SelectedIndex > 0 And ComboBox_Year.SelectedIndex > 0 Then
            Button_Start.Enabled = True
        Else
            Button_Start.Enabled = False
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

        LinkLabel_CheckForUpdates.Text = "Check for updates"

        ' A failed check returns error text instead of a version, so parse instead of assuming
        Dim latest As Version = Nothing
        If latestVersion Is Nothing OrElse Not Version.TryParse(latestVersion.TrimStart("v"c), latest) Then
            MessageBox.Show("Could not check for updates. Check your internet connection and try again.", ApplicationName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If latest > New Version(Application.ProductVersion) Then
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

    Private Sub Timer_autostart_Tick(sender As Object, e As EventArgs) Handles Timer_autostart.Tick
        If Not autostart Then
            Timer_autostart.Enabled = False
            Return
        End If

        ' A mistyped setting falls back to a default, so never install unattended on top of one.
        ' Say so, and do not leave a hidden window hanging.
        If settingsHadErrors OrElse Not Button_Start.Enabled Then
            Timer_autostart.Enabled = False
            LogError("[Settings] Automatic start was skipped because of the settings problem above. Fix it, or pick the options here and click Install.")
            If autoclose Then Environment.Exit(1)
            Return
        End If

        ' Silent runs start right away. Visible runs count down so a person can cancel.
        If silent OrElse autostartSecondsLeft <= 0 Then
            Timer_autostart.Enabled = False
            Start()
            Return
        End If

        Timer_autostart.Interval = 1000
        Button_Start.Text = $"Installing in {autostartSecondsLeft} - click to cancel"
        autostartSecondsLeft -= 1
    End Sub
End Class