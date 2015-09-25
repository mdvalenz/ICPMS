<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form4))
        Me.headingLabel = New System.Windows.Forms.Label()
        Me.wetRadioButton = New System.Windows.Forms.RadioButton()
        Me.dryRadioButton = New System.Windows.Forms.RadioButton()
        Me.ppbRadioButton = New System.Windows.Forms.RadioButton()
        Me.ppmRadioButton = New System.Windows.Forms.RadioButton()
        Me.returnButton = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.diluteRadioButton = New System.Windows.Forms.RadioButton()
        Me.microwaveRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.juiceCheckBox = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ExitProgramToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.skipButton = New System.Windows.Forms.Button()
        Me.exportButton = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.sigFig3RadioButton = New System.Windows.Forms.RadioButton()
        Me.sigFig2RadioButton = New System.Windows.Forms.RadioButton()
        Me.consumerLeadCheckBox = New System.Windows.Forms.CheckBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.endDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.startDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.sampleNumberLabel = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'headingLabel
        '
        Me.headingLabel.AutoSize = True
        Me.headingLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headingLabel.Location = New System.Drawing.Point(66, 26)
        Me.headingLabel.Name = "headingLabel"
        Me.headingLabel.Size = New System.Drawing.Size(243, 20)
        Me.headingLabel.TabIndex = 0
        Me.headingLabel.Text = "Please select the options for:"
        '
        'wetRadioButton
        '
        Me.wetRadioButton.AutoSize = True
        Me.wetRadioButton.Checked = True
        Me.wetRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.wetRadioButton.Location = New System.Drawing.Point(21, 10)
        Me.wetRadioButton.Name = "wetRadioButton"
        Me.wetRadioButton.Size = New System.Drawing.Size(59, 24)
        Me.wetRadioButton.TabIndex = 0
        Me.wetRadioButton.TabStop = True
        Me.wetRadioButton.Text = "Wet"
        Me.wetRadioButton.UseVisualStyleBackColor = True
        '
        'dryRadioButton
        '
        Me.dryRadioButton.AutoSize = True
        Me.dryRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dryRadioButton.Location = New System.Drawing.Point(21, 40)
        Me.dryRadioButton.Name = "dryRadioButton"
        Me.dryRadioButton.Size = New System.Drawing.Size(54, 24)
        Me.dryRadioButton.TabIndex = 0
        Me.dryRadioButton.Text = "Dry"
        Me.dryRadioButton.UseVisualStyleBackColor = True
        '
        'ppbRadioButton
        '
        Me.ppbRadioButton.AutoSize = True
        Me.ppbRadioButton.Checked = True
        Me.ppbRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ppbRadioButton.Location = New System.Drawing.Point(9, 9)
        Me.ppbRadioButton.Name = "ppbRadioButton"
        Me.ppbRadioButton.Size = New System.Drawing.Size(57, 24)
        Me.ppbRadioButton.TabIndex = 0
        Me.ppbRadioButton.TabStop = True
        Me.ppbRadioButton.Text = "ppb"
        Me.ppbRadioButton.UseVisualStyleBackColor = True
        '
        'ppmRadioButton
        '
        Me.ppmRadioButton.AutoSize = True
        Me.ppmRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ppmRadioButton.Location = New System.Drawing.Point(9, 39)
        Me.ppmRadioButton.Name = "ppmRadioButton"
        Me.ppmRadioButton.Size = New System.Drawing.Size(61, 24)
        Me.ppmRadioButton.TabIndex = 0
        Me.ppmRadioButton.Text = "ppm"
        Me.ppmRadioButton.UseVisualStyleBackColor = True
        '
        'returnButton
        '
        Me.returnButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.returnButton.Location = New System.Drawing.Point(264, 321)
        Me.returnButton.Name = "returnButton"
        Me.returnButton.Size = New System.Drawing.Size(80, 30)
        Me.returnButton.TabIndex = 1
        Me.returnButton.Text = "OK"
        Me.returnButton.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dryRadioButton)
        Me.GroupBox1.Controls.Add(Me.wetRadioButton)
        Me.GroupBox1.Location = New System.Drawing.Point(26, 159)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(89, 73)
        Me.GroupBox1.TabIndex = 50
        Me.GroupBox1.TabStop = False
        '
        'diluteRadioButton
        '
        Me.diluteRadioButton.AutoSize = True
        Me.diluteRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.diluteRadioButton.Location = New System.Drawing.Point(6, 38)
        Me.diluteRadioButton.Name = "diluteRadioButton"
        Me.diluteRadioButton.Size = New System.Drawing.Size(74, 24)
        Me.diluteRadioButton.TabIndex = 0
        Me.diluteRadioButton.Text = "Dilute"
        Me.diluteRadioButton.UseVisualStyleBackColor = True
        '
        'microwaveRadioButton
        '
        Me.microwaveRadioButton.AutoSize = True
        Me.microwaveRadioButton.Checked = True
        Me.microwaveRadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.microwaveRadioButton.Location = New System.Drawing.Point(6, 8)
        Me.microwaveRadioButton.Name = "microwaveRadioButton"
        Me.microwaveRadioButton.Size = New System.Drawing.Size(110, 24)
        Me.microwaveRadioButton.TabIndex = 0
        Me.microwaveRadioButton.TabStop = True
        Me.microwaveRadioButton.Text = "Microwave"
        Me.microwaveRadioButton.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox2.Controls.Add(Me.microwaveRadioButton)
        Me.GroupBox2.Controls.Add(Me.diluteRadioButton)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.Location = New System.Drawing.Point(128, 161)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(130, 70)
        Me.GroupBox2.TabIndex = 51
        Me.GroupBox2.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox3.Controls.Add(Me.ppmRadioButton)
        Me.GroupBox3.Controls.Add(Me.ppbRadioButton)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox3.Location = New System.Drawing.Point(264, 160)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(80, 70)
        Me.GroupBox3.TabIndex = 52
        Me.GroupBox3.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(85, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Sample ID:"
        '
        'juiceCheckBox
        '
        Me.juiceCheckBox.AutoSize = True
        Me.juiceCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.juiceCheckBox.Location = New System.Drawing.Point(47, 245)
        Me.juiceCheckBox.Name = "juiceCheckBox"
        Me.juiceCheckBox.Size = New System.Drawing.Size(76, 24)
        Me.juiceCheckBox.TabIndex = 0
        Me.juiceCheckBox.TabStop = False
        Me.juiceCheckBox.Text = "Water"
        Me.juiceCheckBox.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(93, 125)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 20)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "End Date:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(85, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 20)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Start Date:"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitProgramToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1062, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ExitProgramToolStripMenuItem
        '
        Me.ExitProgramToolStripMenuItem.Name = "ExitProgramToolStripMenuItem"
        Me.ExitProgramToolStripMenuItem.Size = New System.Drawing.Size(86, 20)
        Me.ExitProgramToolStripMenuItem.Text = "Exit Program"
        '
        'skipButton
        '
        Me.skipButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.skipButton.Location = New System.Drawing.Point(150, 321)
        Me.skipButton.Name = "skipButton"
        Me.skipButton.Size = New System.Drawing.Size(80, 30)
        Me.skipButton.TabIndex = 0
        Me.skipButton.TabStop = False
        Me.skipButton.Text = "Skip"
        Me.skipButton.UseVisualStyleBackColor = True
        '
        'exportButton
        '
        Me.exportButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exportButton.Location = New System.Drawing.Point(35, 321)
        Me.exportButton.Name = "exportButton"
        Me.exportButton.Size = New System.Drawing.Size(80, 30)
        Me.exportButton.TabIndex = 0
        Me.exportButton.TabStop = False
        Me.exportButton.Text = "Export"
        Me.exportButton.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox4.Controls.Add(Me.sigFig3RadioButton)
        Me.GroupBox4.Controls.Add(Me.sigFig2RadioButton)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox4.Location = New System.Drawing.Point(228, 237)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(116, 70)
        Me.GroupBox4.TabIndex = 53
        Me.GroupBox4.TabStop = False
        '
        'sigFig3RadioButton
        '
        Me.sigFig3RadioButton.AutoSize = True
        Me.sigFig3RadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sigFig3RadioButton.Location = New System.Drawing.Point(9, 40)
        Me.sigFig3RadioButton.Name = "sigFig3RadioButton"
        Me.sigFig3RadioButton.Size = New System.Drawing.Size(102, 24)
        Me.sigFig3RadioButton.TabIndex = 0
        Me.sigFig3RadioButton.Text = "3 SigFigs"
        Me.sigFig3RadioButton.UseVisualStyleBackColor = True
        '
        'sigFig2RadioButton
        '
        Me.sigFig2RadioButton.AutoSize = True
        Me.sigFig2RadioButton.Checked = True
        Me.sigFig2RadioButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sigFig2RadioButton.Location = New System.Drawing.Point(9, 9)
        Me.sigFig2RadioButton.Name = "sigFig2RadioButton"
        Me.sigFig2RadioButton.Size = New System.Drawing.Size(102, 24)
        Me.sigFig2RadioButton.TabIndex = 0
        Me.sigFig2RadioButton.TabStop = True
        Me.sigFig2RadioButton.Text = "2 SigFigs"
        Me.sigFig2RadioButton.UseVisualStyleBackColor = True
        '
        'consumerLeadCheckBox
        '
        Me.consumerLeadCheckBox.AutoSize = True
        Me.consumerLeadCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.consumerLeadCheckBox.Location = New System.Drawing.Point(47, 275)
        Me.consumerLeadCheckBox.Name = "consumerLeadCheckBox"
        Me.consumerLeadCheckBox.Size = New System.Drawing.Size(154, 24)
        Me.consumerLeadCheckBox.TabIndex = 54
        Me.consumerLeadCheckBox.TabStop = False
        Me.consumerLeadCheckBox.Text = "Consumer Lead"
        Me.consumerLeadCheckBox.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(350, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(700, 339)
        Me.DataGridView1.TabIndex = 55
        '
        'endDateTimePicker
        '
        Me.endDateTimePicker.CustomFormat = "MMM, dd,yy"
        Me.endDateTimePicker.DataBindings.Add(New System.Windows.Forms.Binding("Value", Global.ICP_MS_Export.My.MySettings.Default, "endDate", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.endDateTimePicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.endDateTimePicker.Location = New System.Drawing.Point(184, 123)
        Me.endDateTimePicker.MinDate = New Date(2014, 1, 1, 0, 0, 0, 0)
        Me.endDateTimePicker.Name = "endDateTimePicker"
        Me.endDateTimePicker.Size = New System.Drawing.Size(107, 26)
        Me.endDateTimePicker.TabIndex = 0
        Me.endDateTimePicker.TabStop = False
        Me.endDateTimePicker.Value = Global.ICP_MS_Export.My.MySettings.Default.endDate
        '
        'startDateTimePicker
        '
        Me.startDateTimePicker.CustomFormat = "MMM, dd,yy"
        Me.startDateTimePicker.DataBindings.Add(New System.Windows.Forms.Binding("Value", Global.ICP_MS_Export.My.MySettings.Default, "startDate", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.startDateTimePicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.startDateTimePicker.Location = New System.Drawing.Point(184, 93)
        Me.startDateTimePicker.MinDate = New Date(2014, 1, 1, 0, 0, 0, 0)
        Me.startDateTimePicker.Name = "startDateTimePicker"
        Me.startDateTimePicker.Size = New System.Drawing.Size(107, 26)
        Me.startDateTimePicker.TabIndex = 0
        Me.startDateTimePicker.TabStop = False
        Me.startDateTimePicker.Value = Global.ICP_MS_Export.My.MySettings.Default.startDate
        '
        'sampleNumberLabel
        '
        Me.sampleNumberLabel.AutoSize = True
        Me.sampleNumberLabel.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.ICP_MS_Export.My.MySettings.Default, "sampleID", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.sampleNumberLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sampleNumberLabel.Location = New System.Drawing.Point(180, 64)
        Me.sampleNumberLabel.Name = "sampleNumberLabel"
        Me.sampleNumberLabel.Size = New System.Drawing.Size(0, 20)
        Me.sampleNumberLabel.TabIndex = 0
        Me.sampleNumberLabel.Text = Global.ICP_MS_Export.My.MySettings.Default.sampleID
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1062, 362)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.consumerLeadCheckBox)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.exportButton)
        Me.Controls.Add(Me.skipButton)
        Me.Controls.Add(Me.endDateTimePicker)
        Me.Controls.Add(Me.startDateTimePicker)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.juiceCheckBox)
        Me.Controls.Add(Me.sampleNumberLabel)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.returnButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.headingLabel)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form4"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sample Options"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents headingLabel As System.Windows.Forms.Label
    Friend WithEvents wetRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents dryRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents ppbRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents ppmRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents returnButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents diluteRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents microwaveRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents sampleNumberLabel As System.Windows.Forms.Label
    Friend WithEvents juiceCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents startDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents endDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ExitProgramToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents skipButton As System.Windows.Forms.Button
    Friend WithEvents exportButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents sigFig3RadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents sigFig2RadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents consumerLeadCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class
