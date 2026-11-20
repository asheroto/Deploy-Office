<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.MainLabel = New System.Windows.Forms.Label()
        Me.Label_Product = New System.Windows.Forms.Label()
        Me.Label_Year = New System.Windows.Forms.Label()
        Me.ComboBox_Year = New System.Windows.Forms.ComboBox()
        Me.Label_Edition = New System.Windows.Forms.Label()
        Me.ComboBox_Edition = New System.Windows.Forms.ComboBox()
        Me.Label_Sep1 = New System.Windows.Forms.Label()
        Me.Label_Apps = New System.Windows.Forms.Label()
        Me.CheckedListBox_Apps = New System.Windows.Forms.CheckedListBox()
        Me.LinkLabel_AppsSelect = New System.Windows.Forms.LinkLabel()
        Me.Label_Sep2 = New System.Windows.Forms.Label()
        Me.Label_AddOns = New System.Windows.Forms.Label()
        Me.CheckBox_AddVisio = New System.Windows.Forms.CheckBox()
        Me.CheckBox_AddProject = New System.Windows.Forms.CheckBox()
        Me.Label_Sep3 = New System.Windows.Forms.Label()
        Me.Label_Options = New System.Windows.Forms.Label()
        Me.CheckBox_CreateDesktopShortcuts = New System.Windows.Forms.CheckBox()
        Me.ComboBox_ShortcutScope = New System.Windows.Forms.ComboBox()
        Me.CheckBox_RemoveMSI = New System.Windows.Forms.CheckBox()
        Me.Button_Start = New System.Windows.Forms.Button()
        Me.ProgressBar_Install = New System.Windows.Forms.ProgressBar()
        Me.LogTextbox = New System.Windows.Forms.TextBox()
        Me.LinkLabel_Repo = New System.Windows.Forms.LinkLabel()
        Me.LinkLabel_CheckForUpdates = New System.Windows.Forms.LinkLabel()
        Me.Timer_autostart = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'MainLabel
        '
        Me.MainLabel.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainLabel.Location = New System.Drawing.Point(20, 12)
        Me.MainLabel.Margin = New System.Windows.Forms.Padding(0)
        Me.MainLabel.Name = "MainLabel"
        Me.MainLabel.Size = New System.Drawing.Size(516, 36)
        Me.MainLabel.TabIndex = 0
        Me.MainLabel.Text = "Deploy Office"
        Me.MainLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label_Product
        '
        Me.Label_Product.AutoSize = True
        Me.Label_Product.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Product.Location = New System.Drawing.Point(22, 58)
        Me.Label_Product.Name = "Label_Product"
        Me.Label_Product.Size = New System.Drawing.Size(56, 17)
        Me.Label_Product.TabIndex = 1
        Me.Label_Product.Text = "Product"
        '
        'Label_Year
        '
        Me.Label_Year.AutoSize = True
        Me.Label_Year.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Year.ForeColor = System.Drawing.Color.FromArgb(CType(CType(96, Byte), Integer), CType(CType(96, Byte), Integer), CType(CType(96, Byte), Integer))
        Me.Label_Year.Location = New System.Drawing.Point(22, 82)
        Me.Label_Year.Name = "Label_Year"
        Me.Label_Year.Size = New System.Drawing.Size(29, 15)
        Me.Label_Year.TabIndex = 2
        Me.Label_Year.Text = "Year"
        '
        'ComboBox_Year
        '
        Me.ComboBox_Year.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Year.FormattingEnabled = True
        Me.ComboBox_Year.Location = New System.Drawing.Point(24, 101)
        Me.ComboBox_Year.Name = "ComboBox_Year"
        Me.ComboBox_Year.Size = New System.Drawing.Size(160, 25)
        Me.ComboBox_Year.TabIndex = 3
        '
        'Label_Edition
        '
        Me.Label_Edition.AutoSize = True
        Me.Label_Edition.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Edition.ForeColor = System.Drawing.Color.FromArgb(CType(CType(96, Byte), Integer), CType(CType(96, Byte), Integer), CType(CType(96, Byte), Integer))
        Me.Label_Edition.Location = New System.Drawing.Point(198, 82)
        Me.Label_Edition.Name = "Label_Edition"
        Me.Label_Edition.Size = New System.Drawing.Size(44, 15)
        Me.Label_Edition.TabIndex = 4
        Me.Label_Edition.Text = "Edition"
        '
        'ComboBox_Edition
        '
        Me.ComboBox_Edition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Edition.Enabled = False
        Me.ComboBox_Edition.FormattingEnabled = True
        Me.ComboBox_Edition.Location = New System.Drawing.Point(200, 101)
        Me.ComboBox_Edition.Name = "ComboBox_Edition"
        Me.ComboBox_Edition.Size = New System.Drawing.Size(336, 25)
        Me.ComboBox_Edition.TabIndex = 5
        '
        'Label_Sep1
        '
        Me.Label_Sep1.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Label_Sep1.Location = New System.Drawing.Point(24, 140)
        Me.Label_Sep1.Name = "Label_Sep1"
        Me.Label_Sep1.Size = New System.Drawing.Size(512, 1)
        Me.Label_Sep1.TabIndex = 20
        '
        'Label_Apps
        '
        Me.Label_Apps.AutoSize = True
        Me.Label_Apps.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Apps.Location = New System.Drawing.Point(22, 150)
        Me.Label_Apps.Name = "Label_Apps"
        Me.Label_Apps.Size = New System.Drawing.Size(97, 17)
        Me.Label_Apps.TabIndex = 6
        Me.Label_Apps.Text = "Apps to install"
        '
        'LinkLabel_AppsSelect
        '
        Me.LinkLabel_AppsSelect.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel_AppsSelect.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.LinkLabel_AppsSelect.LinkColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(103, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LinkLabel_AppsSelect.Location = New System.Drawing.Point(316, 152)
        Me.LinkLabel_AppsSelect.Name = "LinkLabel_AppsSelect"
        Me.LinkLabel_AppsSelect.Size = New System.Drawing.Size(220, 15)
        Me.LinkLabel_AppsSelect.TabIndex = 20
        Me.LinkLabel_AppsSelect.TabStop = True
        Me.LinkLabel_AppsSelect.Text = "Select all   Select none"
        Me.LinkLabel_AppsSelect.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CheckedListBox_Apps
        '
        Me.CheckedListBox_Apps.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.CheckedListBox_Apps.CheckOnClick = True
        Me.CheckedListBox_Apps.ColumnWidth = 127
        Me.CheckedListBox_Apps.FormattingEnabled = True
        Me.CheckedListBox_Apps.IntegralHeight = False
        Me.CheckedListBox_Apps.Location = New System.Drawing.Point(26, 174)
        Me.CheckedListBox_Apps.MultiColumn = True
        Me.CheckedListBox_Apps.Name = "CheckedListBox_Apps"
        Me.CheckedListBox_Apps.Size = New System.Drawing.Size(510, 72)
        Me.CheckedListBox_Apps.TabIndex = 7
        '
        'Label_Sep2
        '
        Me.Label_Sep2.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Label_Sep2.Location = New System.Drawing.Point(24, 254)
        Me.Label_Sep2.Name = "Label_Sep2"
        Me.Label_Sep2.Size = New System.Drawing.Size(512, 1)
        Me.Label_Sep2.TabIndex = 21
        '
        'Label_AddOns
        '
        Me.Label_AddOns.AutoSize = True
        Me.Label_AddOns.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_AddOns.Location = New System.Drawing.Point(22, 264)
        Me.Label_AddOns.Name = "Label_AddOns"
        Me.Label_AddOns.Size = New System.Drawing.Size(76, 17)
        Me.Label_AddOns.TabIndex = 8
        Me.Label_AddOns.Text = "Also install"
        '
        'CheckBox_AddVisio
        '
        Me.CheckBox_AddVisio.AutoSize = True
        Me.CheckBox_AddVisio.Location = New System.Drawing.Point(27, 288)
        Me.CheckBox_AddVisio.Name = "CheckBox_AddVisio"
        Me.CheckBox_AddVisio.Size = New System.Drawing.Size(130, 21)
        Me.CheckBox_AddVisio.TabIndex = 9
        Me.CheckBox_AddVisio.Text = "Visio Professional"
        Me.CheckBox_AddVisio.UseVisualStyleBackColor = True
        '
        'CheckBox_AddProject
        '
        Me.CheckBox_AddProject.AutoSize = True
        Me.CheckBox_AddProject.Location = New System.Drawing.Point(280, 288)
        Me.CheckBox_AddProject.Name = "CheckBox_AddProject"
        Me.CheckBox_AddProject.Size = New System.Drawing.Size(141, 21)
        Me.CheckBox_AddProject.TabIndex = 10
        Me.CheckBox_AddProject.Text = "Project Professional"
        Me.CheckBox_AddProject.UseVisualStyleBackColor = True
        '
        'Label_Sep3
        '
        Me.Label_Sep3.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.Label_Sep3.Location = New System.Drawing.Point(24, 320)
        Me.Label_Sep3.Name = "Label_Sep3"
        Me.Label_Sep3.Size = New System.Drawing.Size(512, 1)
        Me.Label_Sep3.TabIndex = 22
        '
        'Label_Options
        '
        Me.Label_Options.AutoSize = True
        Me.Label_Options.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Options.Location = New System.Drawing.Point(22, 330)
        Me.Label_Options.Name = "Label_Options"
        Me.Label_Options.Size = New System.Drawing.Size(56, 17)
        Me.Label_Options.TabIndex = 11
        Me.Label_Options.Text = "Options"
        '
        'CheckBox_CreateDesktopShortcuts
        '
        Me.CheckBox_CreateDesktopShortcuts.AutoSize = True
        Me.CheckBox_CreateDesktopShortcuts.Checked = True
        Me.CheckBox_CreateDesktopShortcuts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_CreateDesktopShortcuts.Location = New System.Drawing.Point(27, 356)
        Me.CheckBox_CreateDesktopShortcuts.Name = "CheckBox_CreateDesktopShortcuts"
        Me.CheckBox_CreateDesktopShortcuts.Size = New System.Drawing.Size(196, 21)
        Me.CheckBox_CreateDesktopShortcuts.TabIndex = 12
        Me.CheckBox_CreateDesktopShortcuts.Text = "Create desktop shortcuts for"
        Me.CheckBox_CreateDesktopShortcuts.UseVisualStyleBackColor = True
        '
        'ComboBox_ShortcutScope
        '
        Me.ComboBox_ShortcutScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ShortcutScope.FormattingEnabled = True
        Me.ComboBox_ShortcutScope.Items.AddRange(New Object() {"All users", "Current user only"})
        Me.ComboBox_ShortcutScope.Location = New System.Drawing.Point(280, 353)
        Me.ComboBox_ShortcutScope.Name = "ComboBox_ShortcutScope"
        Me.ComboBox_ShortcutScope.Size = New System.Drawing.Size(180, 25)
        Me.ComboBox_ShortcutScope.TabIndex = 13
        '
        'CheckBox_RemoveMSI
        '
        Me.CheckBox_RemoveMSI.AutoSize = True
        Me.CheckBox_RemoveMSI.Location = New System.Drawing.Point(27, 386)
        Me.CheckBox_RemoveMSI.Name = "CheckBox_RemoveMSI"
        Me.CheckBox_RemoveMSI.Size = New System.Drawing.Size(232, 21)
        Me.CheckBox_RemoveMSI.TabIndex = 14
        Me.CheckBox_RemoveMSI.Text = "Remove old MSI versions of Office"
        Me.CheckBox_RemoveMSI.UseVisualStyleBackColor = True
        '
        'Button_Start
        '
        Me.Button_Start.BackColor = System.Drawing.Color.FromArgb(CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.Button_Start.Enabled = False
        Me.Button_Start.FlatAppearance.BorderSize = 0
        Me.Button_Start.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Start.Font = New System.Drawing.Font("Segoe UI Semibold", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_Start.ForeColor = System.Drawing.Color.White
        Me.Button_Start.Location = New System.Drawing.Point(24, 424)
        Me.Button_Start.Name = "Button_Start"
        Me.Button_Start.Size = New System.Drawing.Size(512, 40)
        Me.Button_Start.TabIndex = 15
        Me.Button_Start.Text = "Install"
        Me.Button_Start.UseVisualStyleBackColor = False
        '
        'ProgressBar_Install
        '
        Me.ProgressBar_Install.Location = New System.Drawing.Point(24, 468)
        Me.ProgressBar_Install.MarqueeAnimationSpeed = 30
        Me.ProgressBar_Install.Name = "ProgressBar_Install"
        Me.ProgressBar_Install.Size = New System.Drawing.Size(512, 5)
        Me.ProgressBar_Install.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar_Install.TabIndex = 16
        Me.ProgressBar_Install.Visible = False
        '
        'LogTextbox
        '
        Me.LogTextbox.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.LogTextbox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.LogTextbox.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LogTextbox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.LogTextbox.Location = New System.Drawing.Point(24, 482)
        Me.LogTextbox.Multiline = True
        Me.LogTextbox.Name = "LogTextbox"
        Me.LogTextbox.ReadOnly = True
        Me.LogTextbox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.LogTextbox.Size = New System.Drawing.Size(512, 80)
        Me.LogTextbox.TabIndex = 17
        Me.LogTextbox.TabStop = False
        '
        'LinkLabel_Repo
        '
        Me.LinkLabel_Repo.AutoSize = True
        Me.LinkLabel_Repo.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel_Repo.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.LinkLabel_Repo.LinkColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(103, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LinkLabel_Repo.Location = New System.Drawing.Point(22, 572)
        Me.LinkLabel_Repo.Name = "LinkLabel_Repo"
        Me.LinkLabel_Repo.Size = New System.Drawing.Size(43, 13)
        Me.LinkLabel_Repo.TabIndex = 18
        Me.LinkLabel_Repo.TabStop = True
        Me.LinkLabel_Repo.Text = "GitHub"
        '
        'LinkLabel_CheckForUpdates
        '
        Me.LinkLabel_CheckForUpdates.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel_CheckForUpdates.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.LinkLabel_CheckForUpdates.LinkColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(103, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.LinkLabel_CheckForUpdates.Location = New System.Drawing.Point(376, 572)
        Me.LinkLabel_CheckForUpdates.Name = "LinkLabel_CheckForUpdates"
        Me.LinkLabel_CheckForUpdates.Size = New System.Drawing.Size(160, 16)
        Me.LinkLabel_CheckForUpdates.TabIndex = 19
        Me.LinkLabel_CheckForUpdates.TabStop = True
        Me.LinkLabel_CheckForUpdates.Text = "Check for updates"
        Me.LinkLabel_CheckForUpdates.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Timer_autostart
        '
        Me.Timer_autostart.Enabled = True
        Me.Timer_autostart.Interval = 1
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(560, 598)
        Me.Controls.Add(Me.LinkLabel_CheckForUpdates)
        Me.Controls.Add(Me.LinkLabel_Repo)
        Me.Controls.Add(Me.LogTextbox)
        Me.Controls.Add(Me.ProgressBar_Install)
        Me.Controls.Add(Me.Button_Start)
        Me.Controls.Add(Me.CheckBox_RemoveMSI)
        Me.Controls.Add(Me.ComboBox_ShortcutScope)
        Me.Controls.Add(Me.CheckBox_CreateDesktopShortcuts)
        Me.Controls.Add(Me.Label_Options)
        Me.Controls.Add(Me.Label_Sep3)
        Me.Controls.Add(Me.CheckBox_AddProject)
        Me.Controls.Add(Me.CheckBox_AddVisio)
        Me.Controls.Add(Me.Label_AddOns)
        Me.Controls.Add(Me.Label_Sep2)
        Me.Controls.Add(Me.LinkLabel_AppsSelect)
        Me.Controls.Add(Me.CheckedListBox_Apps)
        Me.Controls.Add(Me.Label_Apps)
        Me.Controls.Add(Me.Label_Sep1)
        Me.Controls.Add(Me.ComboBox_Edition)
        Me.Controls.Add(Me.Label_Edition)
        Me.Controls.Add(Me.ComboBox_Year)
        Me.Controls.Add(Me.Label_Year)
        Me.Controls.Add(Me.Label_Product)
        Me.Controls.Add(Me.MainLabel)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deploy Office"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MainLabel As Label
    Friend WithEvents Label_Product As Label
    Friend WithEvents Label_Year As Label
    Friend WithEvents ComboBox_Year As ComboBox
    Friend WithEvents Label_Edition As Label
    Friend WithEvents ComboBox_Edition As ComboBox
    Friend WithEvents Label_Sep1 As Label
    Friend WithEvents Label_Apps As Label
    Friend WithEvents CheckedListBox_Apps As CheckedListBox
    Friend WithEvents LinkLabel_AppsSelect As LinkLabel
    Friend WithEvents Label_Sep2 As Label
    Friend WithEvents Label_AddOns As Label
    Friend WithEvents CheckBox_AddVisio As CheckBox
    Friend WithEvents CheckBox_AddProject As CheckBox
    Friend WithEvents Label_Sep3 As Label
    Friend WithEvents Label_Options As Label
    Friend WithEvents CheckBox_CreateDesktopShortcuts As CheckBox
    Friend WithEvents ComboBox_ShortcutScope As ComboBox
    Friend WithEvents CheckBox_RemoveMSI As CheckBox
    Friend WithEvents Button_Start As Button
    Friend WithEvents ProgressBar_Install As ProgressBar
    Friend WithEvents LogTextbox As TextBox
    Friend WithEvents LinkLabel_Repo As LinkLabel
    Friend WithEvents LinkLabel_CheckForUpdates As LinkLabel
    Friend WithEvents Timer_autostart As Timer
End Class
