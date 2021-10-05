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
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Button_Start = New System.Windows.Forms.Button()
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
        Me.MainLabel.Text = "Starting Office Installation"
        Me.MainLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CountdownLabel
        '
        Me.CountdownLabel.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CountdownLabel.Location = New System.Drawing.Point(406, 62)
        Me.CountdownLabel.Margin = New System.Windows.Forms.Padding(0)
        Me.CountdownLabel.Name = "CountdownLabel"
        Me.CountdownLabel.Size = New System.Drawing.Size(121, 31)
        Me.CountdownLabel.TabIndex = 1
        Me.CountdownLabel.Text = "30"
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
        Me.EditionSelector.Location = New System.Drawing.Point(213, 66)
        Me.EditionSelector.Name = "EditionSelector"
        Me.EditionSelector.Size = New System.Drawing.Size(190, 23)
        Me.EditionSelector.TabIndex = 3
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(12, 67)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(54, 21)
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.Text = "2019"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Location = New System.Drawing.Point(72, 67)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(54, 21)
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "2022"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Button_Start
        '
        Me.Button_Start.Location = New System.Drawing.Point(132, 66)
        Me.Button_Start.Name = "Button_Start"
        Me.Button_Start.Size = New System.Drawing.Size(75, 23)
        Me.Button_Start.TabIndex = 6
        Me.Button_Start.Text = "Start"
        Me.Button_Start.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(539, 461)
        Me.Controls.Add(Me.Button_Start)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
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
        Me.Text = "Deploy Office"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CountdownTimer As Timer
    Friend WithEvents MainLabel As Label
    Friend WithEvents CountdownLabel As Label
    Friend WithEvents LogTextbox As TextBox
    Friend WithEvents EditionSelector As ComboBox
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents Button_Start As Button
End Class
