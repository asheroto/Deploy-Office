<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.CountdownTimer = New System.Windows.Forms.Timer(Me.components)
        Me.MainLabel = New System.Windows.Forms.Label()
        Me.CountdownLabel = New System.Windows.Forms.Label()
        Me.LogTextbox = New System.Windows.Forms.TextBox()
        Me.EditionSelector = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'CountdownTimer
        '
        Me.CountdownTimer.Enabled = True
        Me.CountdownTimer.Interval = 1000
        '
        'MainLabel
        '
        Me.MainLabel.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainLabel.Location = New System.Drawing.Point(12, 9)
        Me.MainLabel.Margin = New System.Windows.Forms.Padding(0)
        Me.MainLabel.Name = "MainLabel"
        Me.MainLabel.Size = New System.Drawing.Size(515, 48)
        Me.MainLabel.TabIndex = 0
        Me.MainLabel.Text = "Starting Office 2019 Installation"
        Me.MainLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CountdownLabel
        '
        Me.CountdownLabel.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CountdownLabel.Location = New System.Drawing.Point(12, 57)
        Me.CountdownLabel.Margin = New System.Windows.Forms.Padding(0)
        Me.CountdownLabel.Name = "CountdownLabel"
        Me.CountdownLabel.Size = New System.Drawing.Size(515, 31)
        Me.CountdownLabel.TabIndex = 1
        Me.CountdownLabel.Text = "10"
        Me.CountdownLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LogTextbox
        '
        Me.LogTextbox.Location = New System.Drawing.Point(12, 108)
        Me.LogTextbox.Multiline = True
        Me.LogTextbox.Name = "LogTextbox"
        Me.LogTextbox.ReadOnly = True
        Me.LogTextbox.Size = New System.Drawing.Size(515, 337)
        Me.LogTextbox.TabIndex = 2
        '
        'EditionSelector
        '
        Me.EditionSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.EditionSelector.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EditionSelector.FormattingEnabled = True
        Me.EditionSelector.Items.AddRange(New Object() {"Home & Business", "Home & Student", "Personal", "Professional", "Professional Plus", "Professional Plus - Volume", "Standard", "Standard - Volume", "Visio Standard", "Visio Standard - Volume", "Visio Professional", "Visio Professional - Volume"})
        Me.EditionSelector.Location = New System.Drawing.Point(337, 63)
        Me.EditionSelector.Name = "EditionSelector"
        Me.EditionSelector.Size = New System.Drawing.Size(190, 23)
        Me.EditionSelector.TabIndex = 3
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(539, 461)
        Me.Controls.Add(Me.EditionSelector)
        Me.Controls.Add(Me.LogTextbox)
        Me.Controls.Add(Me.CountdownLabel)
        Me.Controls.Add(Me.MainLabel)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Deploy Office 2019"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CountdownTimer As Timer
    Friend WithEvents MainLabel As Label
    Friend WithEvents CountdownLabel As Label
    Friend WithEvents LogTextbox As TextBox
    Friend WithEvents EditionSelector As ComboBox
End Class
