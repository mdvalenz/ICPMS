<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.exportButton = New System.Windows.Forms.Button()
        Me.fileLocationTextBox = New System.Windows.Forms.TextBox()
        Me.fileLocationButton = New System.Windows.Forms.Button()
        Me.headingLabel = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.analystTextBox = New System.Windows.Forms.TextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.SelectMassesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SelectMassesToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChangeDefaultMDLsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.progressBarLabel = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'exportButton
        '
        Me.exportButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exportButton.Location = New System.Drawing.Point(447, 106)
        Me.exportButton.Name = "exportButton"
        Me.exportButton.Size = New System.Drawing.Size(80, 30)
        Me.exportButton.TabIndex = 4
        Me.exportButton.Text = "Export"
        Me.exportButton.UseVisualStyleBackColor = True
        '
        'fileLocationTextBox
        '
        Me.fileLocationTextBox.AllowDrop = True
        Me.fileLocationTextBox.Location = New System.Drawing.Point(22, 72)
        Me.fileLocationTextBox.Name = "fileLocationTextBox"
        Me.fileLocationTextBox.Size = New System.Drawing.Size(419, 20)
        Me.fileLocationTextBox.TabIndex = 1
        '
        'fileLocationButton
        '
        Me.fileLocationButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fileLocationButton.Location = New System.Drawing.Point(447, 65)
        Me.fileLocationButton.Name = "fileLocationButton"
        Me.fileLocationButton.Size = New System.Drawing.Size(80, 30)
        Me.fileLocationButton.TabIndex = 2
        Me.fileLocationButton.Text = "Browse"
        Me.fileLocationButton.UseVisualStyleBackColor = True
        '
        'headingLabel
        '
        Me.headingLabel.AutoSize = True
        Me.headingLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headingLabel.Location = New System.Drawing.Point(55, 28)
        Me.headingLabel.Name = "headingLabel"
        Me.headingLabel.Size = New System.Drawing.Size(431, 20)
        Me.headingLabel.TabIndex = 0
        Me.headingLabel.Text = "Please select the file you wish to export to LabWorks"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.InitialDirectory = """L:\eRecords\Chemistry\Instrument Data\ICP-2\Report data\"""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(18, 106)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(202, 20)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Analyst LabWorks Login"
        '
        'analystTextBox
        '
        Me.analystTextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.ICP_MS_Export.My.MySettings.Default, "metalAnalyst", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.analystTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.analystTextBox.Location = New System.Drawing.Point(226, 103)
        Me.analystTextBox.Name = "analystTextBox"
        Me.analystTextBox.Size = New System.Drawing.Size(70, 26)
        Me.analystTextBox.TabIndex = 3
        Me.analystTextBox.Text = Global.ICP_MS_Export.My.MySettings.Default.metalAnalyst
        Me.analystTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectMassesToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(547, 24)
        Me.MenuStrip1.TabIndex = 7
        Me.MenuStrip1.Text = "MenuStrip1"
        Me.MenuStrip1.Visible = False
        '
        'SelectMassesToolStripMenuItem
        '
        Me.SelectMassesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectMassesToolStripMenuItem1, Me.ChangeDefaultMDLsToolStripMenuItem1})
        Me.SelectMassesToolStripMenuItem.Name = "SelectMassesToolStripMenuItem"
        Me.SelectMassesToolStripMenuItem.Size = New System.Drawing.Size(61, 20)
        Me.SelectMassesToolStripMenuItem.Text = "Settings"
        '
        'SelectMassesToolStripMenuItem1
        '
        Me.SelectMassesToolStripMenuItem1.Name = "SelectMassesToolStripMenuItem1"
        Me.SelectMassesToolStripMenuItem1.Size = New System.Drawing.Size(189, 22)
        Me.SelectMassesToolStripMenuItem1.Text = "Select Masses"
        '
        'ChangeDefaultMDLsToolStripMenuItem1
        '
        Me.ChangeDefaultMDLsToolStripMenuItem1.Name = "ChangeDefaultMDLsToolStripMenuItem1"
        Me.ChangeDefaultMDLsToolStripMenuItem1.Size = New System.Drawing.Size(189, 22)
        Me.ChangeDefaultMDLsToolStripMenuItem1.Text = "Change Default MDLs"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(22, 162)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(419, 23)
        Me.ProgressBar1.TabIndex = 8
        Me.ProgressBar1.Visible = False
        '
        'progressBarLabel
        '
        Me.progressBarLabel.AutoSize = True
        Me.progressBarLabel.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.progressBarLabel.Location = New System.Drawing.Point(18, 140)
        Me.progressBarLabel.Name = "progressBarLabel"
        Me.progressBarLabel.Size = New System.Drawing.Size(52, 19)
        Me.progressBarLabel.TabIndex = 9
        Me.progressBarLabel.Text = "Label2"
        Me.progressBarLabel.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 197)
        Me.Controls.Add(Me.progressBarLabel)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.analystTextBox)
        Me.Controls.Add(Me.headingLabel)
        Me.Controls.Add(Me.fileLocationButton)
        Me.Controls.Add(Me.fileLocationTextBox)
        Me.Controls.Add(Me.exportButton)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ICP-MS Results Import"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents exportButton As System.Windows.Forms.Button
    Friend WithEvents fileLocationTextBox As System.Windows.Forms.TextBox
    Friend WithEvents fileLocationButton As System.Windows.Forms.Button
    Friend WithEvents headingLabel As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents analystTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents SelectMassesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectMassesToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChangeDefaultMDLsToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents progressBarLabel As System.Windows.Forms.Label

End Class
