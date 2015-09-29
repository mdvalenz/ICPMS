Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Set radio button values, checkbox values, and dates
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

        If My.Settings.ppbCheck = True Then
            ppbRadioButton.Checked = True
            My.Settings.sampleUnits = "ppb"
        Else
            ppmRadioButton.Checked = True
            My.Settings.sampleUnits = "ppm"
        End If

        If My.Settings.juiceCheck = True Then
            juiceCheckBox.Checked = True
        Else
            juiceCheckBox.Checked = False
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

        If My.Settings.setSigFigs = "3" Then
            sigFig3RadioButton.Checked = True
        Else
            sigFig2RadioButton.Checked = True
        End If

        'Fill in datagridview1
        Call displayArray()

        'Detect changes to radio buttons and checkboxes
        AddHandler Me.wetRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.dryRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.microwaveRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.diluteRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.ppbRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.ppmRadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.sigFig2RadioButton.CheckedChanged, AddressOf OnChange
        AddHandler Me.sigFig3RadioButton.CheckedChanged, AddressOf OnChange

        AddHandler Me.juiceCheckBox.CheckedChanged, AddressOf OnChange
        AddHandler Me.consumerLeadCheckBox.CheckedChanged, AddressOf OnChange

    End Sub

    Private Sub displayArray()

        'Clear datagridview1
        DataGridView1.DataSource = Nothing

        'DataGridView1.Rows.Clear()

        Dim dataTable1 As DataTable
        dataTable1 = New DataTable("Sample Array")
        Try
            Dim colA As DataColumn = New DataColumn("Analyte")
            colA.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colA)
            Dim colB As DataColumn = New DataColumn("Concentration (ppb)")
            colB.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colB)
            Dim colC As DataColumn = New DataColumn("MDL (ppb)")
            colC.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colC)
            Dim colD As DataColumn = New DataColumn("Rounded Result " & My.Settings.sampleUnits)
            colD.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colD)
            Dim colE As DataColumn = New DataColumn("Reported MDL " & My.Settings.sampleUnits)
            colE.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colE)
            Dim colF As DataColumn = New DataColumn("Reported Result " & My.Settings.sampleUnits)
            colF.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colF)

            Dim j As Integer = 0
            For j = 0 To CInt(My.Settings.exportArrayCount)

                If exportArray(1, j) <> "" Then

                    Dim row As DataRow
                    row = dataTable1.NewRow()
                    row.Item("Analyte") = exportArray(2, j)
                    row.Item("Concentration (ppb)") = exportArray(3, j)
                    row.Item("MDL (ppb)") = exportArray(5, j)
                    row.Item("Rounded Result " & My.Settings.sampleUnits) = exportArray(7, j)
                    row.Item("Reported MDL " & My.Settings.sampleUnits) = exportArray(9, j)
                    row.Item("Reported Result " & My.Settings.sampleUnits) = exportArray(8, j)
                    dataTable1.Rows.Add(row)

                End If

            Next

        Catch
        End Try

        DataGridView1.DataSource = dataTable1
        DataGridView1.Refresh()

    End Sub

    Private Sub OnChange(sender As System.Object, e As System.EventArgs)
        'Dim rb = CType(sender, RadioButton)
        'Console.WriteLine(rb.Name + " " + rb.Checked.ToString)

        'Detect changes and update settings
        'Set radio button values, checkbox values, and dates
        If wetRadioButton.Checked = True Then
            My.Settings.wetCheck = True
            My.Settings.dryCheck = False
        Else
            My.Settings.dryCheck = True
            My.Settings.wetCheck = False
        End If

        If microwaveRadioButton.Checked = True Then
            My.Settings.microwaveCheck = True
            My.Settings.diluteCheck = False
        Else
            My.Settings.diluteCheck = True
            My.Settings.microwaveCheck = False
        End If

        If ppbRadioButton.Checked = True Then
            My.Settings.ppbCheck = True
            My.Settings.ppmCheck = False
            My.Settings.sampleUnits = "ppb"
        Else
            My.Settings.ppmCheck = True
            My.Settings.ppbCheck = False
            My.Settings.sampleUnits = "ppm"
        End If

        If juiceCheckBox.Checked = True Then
            My.Settings.juiceCheck = True
        Else
            My.Settings.juiceCheck = False
        End If

        If consumerLeadCheckBox.Checked = True Then
            My.Settings.leadConsumer = True
        Else
            My.Settings.leadConsumer = False
        End If

        If sigFig3RadioButton.Checked = True Then
            My.Settings.setSigFigs = "3"
        Else
            My.Settings.setSigFigs = "2"
        End If

        'Recalculate values in array
        Call recalculateValues()

    End Sub

    Private Sub recalculateValues()

        'Work through array and alter values based on changes to radiobuttons and checkboxes
        For rowNum = 0 To CInt(My.Settings.exportArrayCount)
            If exportArray(1, rowNum) <> "" Then

                'MDLs
                Dim mdlType As String = Nothing
                If My.Settings.wetCheck = True Then
                    If My.Settings.diluteCheck = True Then
                        If My.Settings.juiceCheck = True Then
                            mdlType = "wj"
                        Else
                            mdlType = "wd"
                        End If
                    Else
                        If My.Settings.juiceCheck = True Then
                            mdlType = "wj"
                        Else
                            mdlType = "wm"
                        End If
                    End If
                Else
                    If My.Settings.diluteCheck = True Then
                        If My.Settings.juiceCheck = True Then
                            mdlType = "dj"
                        Else
                            mdlType = "dd"
                        End If
                    Else
                        If My.Settings.juiceCheck = True Then
                            mdlType = "dj"
                        Else
                            mdlType = "dm"
                        End If
                    End If
                End If

                'Units
                Dim metalUnits As String = Nothing
                If My.Settings.ppbCheck = True Then metalUnits = "ppb" Else metalUnits = "ppm"

                'Results
                Dim metalName As String = Nothing
                Dim metalTest As Double = Nothing
                Dim metalResult As String = Nothing

                metalName = exportArray(2, rowNum)
                metalTest = exportArray(3, rowNum)

                Dim finalMDL As Double = Nothing
                Dim reportedMDL As String = Nothing

                'Get Metal mdl
                Select Case metalName
                    Case "Ag"
                        If mdlType = "wm" Then finalMDL = Form3.wmAgTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmAgTextBox.Text
                    Case "Al"
                        If mdlType = "wm" Then finalMDL = Form3.wmAlTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmAlTextBox.Text
                    Case "As"
                        If mdlType = "wm" Then finalMDL = Form3.wmAsTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmAsTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjAsTextBox.Text
                    Case "B"
                        If mdlType = "wm" Then finalMDL = Form3.wmBTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmBTextBox.Text
                    Case "Ba"
                        If mdlType = "wm" Then finalMDL = Form3.wmBaTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmBaTextBox.Text
                    Case "Be"
                        If mdlType = "wm" Then finalMDL = Form3.wmBeTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmBeTextBox.Text
                    Case "Cd"
                        If mdlType = "wm" Then finalMDL = Form3.wmCdTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmCdTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjCdTextBox.Text
                    Case "Co"
                        If mdlType = "wm" Then finalMDL = Form3.wmCoTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmCoTextBox.Text
                    Case "Cr"
                        If mdlType = "wm" Then finalMDL = Form3.wmCrTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmCrTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjCrTextBox.Text
                    Case "Cu"
                        If mdlType = "wm" Then finalMDL = Form3.wmCuTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmCuTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjCuTextBox.Text
                    Case "Mn"
                        If mdlType = "wm" Then finalMDL = Form3.wmMnTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmMnTextBox.Text
                    Case "Mo"
                        If mdlType = "wm" Then finalMDL = Form3.wmMoTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmMoTextBox.Text
                    Case "Ni"
                        If mdlType = "wm" Then finalMDL = Form3.wmNiTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmNiTextBox.Text
                    Case "P"
                        If mdlType = "wm" Then finalMDL = Form3.wmPTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmPTextBox.Text
                    Case "Pb"
                        If mdlType = "wd" Then finalMDL = Form3.wdPbTextBox.Text
                        If mdlType = "dd" Then finalMDL = Form3.ddPbTextBox.Text
                        If mdlType = "wm" Then finalMDL = Form3.wmPbTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmPbTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjPbTextBox.Text
                    Case "Pd"
                        If mdlType = "wm" Then finalMDL = Form3.wmPdTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmPdTextBox.Text
                    Case "Sb"
                        If mdlType = "wm" Then finalMDL = Form3.wmSbTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmSbTextBox.Text
                    Case "Se"
                        If mdlType = "wm" Then finalMDL = Form3.wmSeTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmSeTextBox.Text
                        If mdlType = "wj" Then finalMDL = Form3.wjSeTextBox.Text
                    Case "Sn"
                        If mdlType = "wd" Then finalMDL = Form3.wdSnTextBox.Text
                        If mdlType = "dd" Then finalMDL = Form3.ddSnTextBox.Text
                        If mdlType = "wm" Then finalMDL = Form3.wmSnTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmSnTextBox.Text
                    Case "Ti"
                        If mdlType = "wm" Then finalMDL = Form3.wmTiTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmTiTextBox.Text
                    Case "Tl"
                        If mdlType = "wm" Then finalMDL = Form3.wmTlTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmTlTextBox.Text
                    Case "V"
                        If mdlType = "wm" Then finalMDL = Form3.wmVTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmVTextBox.Text
                    Case "Zn"
                        If mdlType = "wm" Then finalMDL = Form3.wmZnTextBox.Text
                        If mdlType = "dm" Then finalMDL = Form3.dmZnTextBox.Text
                End Select

                'Conversion if units is ppm
                If metalUnits = "ppm" Then
                    metalTest = (metalTest / 1000)
                    finalMDL = (finalMDL / 1000)
                End If

                'Dim textResult As String = 0
                Dim roundResult As Decimal = Nothing
                Dim testRound As Integer = metalTest
                Dim N As Integer = CStr(testRound).Length
                Dim sigFig, sigFigMax As Integer

                'set sigfigs
                sigFig = N

                If My.Settings.setSigFigs = "3" Then
                    sigFigMax = 3
                Else
                    sigFigMax = 2
                End If

                roundResult = Math.Round(metalTest / 10 ^ sigFig, sigFigMax) * 10 ^ sigFig

                If roundResult < finalMDL Then
                    metalResult = "Not Detected"
                    'textResult = 1
                Else
                    metalResult = roundResult
                    'If roundResult < 1000000 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "000000"))
                    'End If

                    'If roundResult < 100000 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "00000"))
                    'End If

                    'If roundResult < 10000 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "0000"))
                    'End If

                    'If roundResult < 1000 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "000"))
                    'End If

                    'If roundResult < 100 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "00.0"))
                    'End If

                    'If roundResult < 10 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "0.00"))
                    'End If

                    'If roundResult < 1 Then
                    '    metalResult = CStr(FormatNumber(roundResult, "0.000"))
                    'End If

                End If

                If metalUnits = "ppm" Then

                    If finalMDL < 1000000 Then
                        reportedMDL = CStr(Format(finalMDL, "000000"))
                    End If

                    If finalMDL < 100000 Then
                        reportedMDL = CStr(Format(finalMDL, "00000"))
                    End If

                    If finalMDL < 10000 Then
                        reportedMDL = CStr(Format(finalMDL, "0000"))
                    End If

                    If finalMDL < 1000 Then
                        reportedMDL = CStr(Format(finalMDL, "000"))
                    End If

                    If finalMDL < 100 Then
                        reportedMDL = CStr(Format(finalMDL, "00.0"))
                    End If

                    If finalMDL < 10 Then
                        reportedMDL = CStr(Format(finalMDL, "0.00"))
                    End If

                    If finalMDL < 1 Then
                        reportedMDL = CStr(Format(finalMDL, "0.000"))
                    End If

                Else
                    reportedMDL = finalMDL
                End If

                'Put revised data into export array
                exportArray(0, rowNum) = My.Settings.leadConsumer
                'exportArray(1, rowNum) = sampleID
                'exportArray(2, rowNum) = metalName
                'exportArray(3, rowNum) = metalTest
                exportArray(4, rowNum) = finalMDL
                exportArray(5, rowNum) = metalUnits
                exportArray(6, rowNum) = sigFigMax
                exportArray(7, rowNum) = roundResult
                exportArray(8, rowNum) = metalResult
                exportArray(9, rowNum) = reportedMDL

            End If
        Next

        'Update datagridview1
        Call displayArray()

    End Sub

    Private Sub returnButton_Click(sender As Object, e As EventArgs) Handles returnButton.Click

        'Put datagridview1 data back into export array
        Dim j As Integer = 0
        For j = 0 To CInt(My.Settings.exportArrayCount)

            If exportArray(1, j) <> "" Then

                If exportArray(2, j) = DataGridView1.Rows(j).Cells(0).ToString Then
                    'Transfer information from sample array to the export array
                    If My.Settings.leadConsumer = True Then
                        exportArray(0, j) = True
                    Else
                        exportArray(0, j) = False
                    End If

                    exportArray(7, j) = DataGridView1.Rows(j).Cells(3).ToString
                    exportArray(9, j) = DataGridView1.Rows(j).Cells(4).ToString
                    exportArray(8, j) = DataGridView1.Rows(j).Cells(5).ToString

                End If
            End If
        Next

        My.Settings.skipSample = False

        'Update Sample Array
        Call transferData()

        Me.Close()
    End Sub

    Private Sub transferData()

        'Update Sample Array with new Export Array data
        Dim j As Integer = 0
        For j = 0 To CInt(My.Settings.sampleArrayCount)

            If SampleArray(1, j) <> "" And SampleArray(1, j) = My.Settings.sampleID Then

                'Transfer information from export array to the sample array
                SampleArray(0, j) = exportArray(0, j)
                SampleArray(1, j) = exportArray(1, j)
                SampleArray(7, j) = exportArray(7, j)
                SampleArray(8, j) = exportArray(8, j)
                SampleArray(9, j) = exportArray(9, j)

            End If
        Next

    End Sub

    Private Sub ExitProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitProgramToolStripMenuItem.Click
        My.Settings.exitProgram = True
        Me.Close()
    End Sub

    Private Sub SkipButton_Click(sender As Object, e As EventArgs) Handles skipButton.Click
        My.Settings.skipSample = True

        'Clear Export Array
        Dim j As Integer = 0
        For j = 0 To CInt(My.Settings.exportArrayCount)

            exportArray(1, j) = ""

        Next

        'Update Sample Array
        Call transferData()

        Me.Close()
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