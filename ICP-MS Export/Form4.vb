Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.wetCheck = True Then
            wetRadioButton.Checked = True
        Else
            dryRadioButton.Checked = True
        End If

        If My.Settings.microwaveCheck = True Then
            microwaveRadioButton.Checked = True
        Else
            diluteRadioButton.Checked = True
        End If

        If My.Settings.juiceCheck = True Then
            juiceCheckBox.Checked = True
        Else
            juiceCheckBox.Checked = False
        End If

        If My.Settings.ppbCheck = True Then
            ppbRadioButton.Checked = True
        Else
            ppmRadioButton.Checked = True
        End If

        If My.Settings.leadConsumer = True Then
            consumerLeadCheckBox.Checked = True
        Else
            consumerLeadCheckBox.Checked = False
        End If

        startDateTimePicker.Format = DateTimePickerFormat.Custom
        startDateTimePicker.CustomFormat = "MM/dd/yyyy"
        startDateTimePicker.Value = My.Settings.startDate

        endDateTimePicker.Format = DateTimePickerFormat.Custom
        endDateTimePicker.CustomFormat = "MM/dd/yyyy"
        endDateTimePicker.Value = My.Settings.endDate

        reviewCheckBox.Checked = My.Settings.reviewExport

        If My.Settings.setSigFigs = "3" Then
            sigFig3RadioButton.Checked = True
        Else
            sigFig2RadioButton.Checked = True
        End If

    End Sub

    Private Sub returnButton_Click(sender As Object, e As EventArgs) Handles returnButton.Click

        If wetRadioButton.Checked = True Then
            My.Settings.wetCheck = True
            My.Settings.dryCheck = False
        Else
            My.Settings.wetCheck = False
            My.Settings.dryCheck = True
        End If

        If microwaveRadioButton.Checked = True Then
            My.Settings.microwaveCheck = True
            My.Settings.diluteCheck = False
        Else
            My.Settings.microwaveCheck = False
            My.Settings.diluteCheck = True
        End If

        If juiceCheckBox.Checked = True Then
            My.Settings.juiceCheck = True
        Else
            My.Settings.juiceCheck = False
        End If

        If ppbRadioButton.Checked = True Then
            My.Settings.ppbCheck = True
            My.Settings.ppmCheck = False
        Else
            My.Settings.ppbCheck = False
            My.Settings.ppmCheck = True
        End If

        If sigFig3RadioButton.Checked = True Then
            My.Settings.setSigFigs = "3"
        Else
            My.Settings.setSigFigs = "2"
        End If

        If consumerLeadCheckBox.Checked = True Then
            My.Settings.leadConsumer = True
        Else
            My.Settings.leadConsumer = False
        End If

        My.Settings.skipSample = False
        Me.Close()
    End Sub

    Private Sub ExitProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitProgramToolStripMenuItem.Click
        My.Settings.exitProgram = True
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.skipSample = True
        Me.Close()
    End Sub

    Private Sub reviewCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles reviewCheckBox.CheckedChanged

        If reviewCheckBox.Checked = True Then
            My.Settings.reviewExport = True
        Else
            My.Settings.reviewExport = False
        End If

    End Sub

    Private Sub startDateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles startDateTimePicker.ValueChanged
        My.Settings.startDate = startDateTimePicker.Value

        My.Settings.endDate = endDateTimePicker.Value

    End Sub

    Private Sub exportButton_Click(sender As Object, e As EventArgs) Handles exportButton.Click
        My.Settings.exportNow = True
        Me.Close()
    End Sub

End Class