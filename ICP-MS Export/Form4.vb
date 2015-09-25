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
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

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
            Dim colE As DataColumn = New DataColumn("Reported Result " & My.Settings.sampleUnits)
            colE.DataType = System.Type.GetType("System.String")
            dataTable1.Columns.Add(colE)
            Dim colF As DataColumn = New DataColumn("Reported MDL " & My.Settings.sampleUnits)
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
                    row.Item("Reported Result " & My.Settings.sampleUnits) = exportArray(8, j)
                    row.Item("Reported MDL " & My.Settings.sampleUnits) = exportArray(9, j)
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
        Else
            My.Settings.dryCheck = True
        End If

        If microwaveRadioButton.Checked = True Then
            My.Settings.microwaveCheck = True
        Else
            My.Settings.diluteCheck = True
        End If

        If ppbRadioButton.Checked = True Then
            My.Settings.ppbCheck = True
            My.Settings.sampleUnits = "ppb"
        Else
            My.Settings.ppmCheck = True
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

                '*****************************************************************************************
                'MDLs********************************************************
                If callSampleSettingsForm.wetRadioButton.Checked = True Then
                    If callSampleSettingsForm.diluteRadioButton.Checked = True Then
                        If callSampleSettingsForm.juiceCheckBox.Checked = True Then
                            mdlType = "wj"
                        Else
                            mdlType = "wd"
                        End If
                    Else
                        If callSampleSettingsForm.juiceCheckBox.Checked = True Then
                            mdlType = "wj"
                        Else
                            mdlType = "wm"
                        End If
                    End If
                Else
                    If callSampleSettingsForm.diluteRadioButton.Checked = True Then
                        If callSampleSettingsForm.juiceCheckBox.Checked = True Then
                            mdlType = "dj"
                        Else
                            mdlType = "dd"
                        End If
                    Else
                        If callSampleSettingsForm.juiceCheckBox.Checked = True Then
                            mdlType = "dj"
                        Else
                            mdlType = "dm"
                        End If
                    End If
                End If

                'Units
                If callSampleSettingsForm.ppbRadioButton.Checked = True Then metalUnits = "ppb" Else metalUnits = "ppm"
                cellValue = ExcelApp.Cells(rowNumber, 2).Value

                Do While cellValue = "" And rowNumber < lastRow + 1
                    leadCount = 0
                    rowNumber = rowNumber + 1
                    cellValue = ExcelApp.Cells(rowNumber, 2).Value
                Loop

                Do While cellValue <> "" And rowNumber < lastRow + 1

                    Dim metalTest As Double = ExcelWorkbook.Range("G" & rowNumber).value

                    If metalTest <> 0.0 Then

                        metalName = ExcelWorkbook.Range("B" & rowNumber).Value
                        metalMass = ExcelWorkbook.Range("C" & rowNumber).Value
                        My.Settings.metalName = metalName
                        Dim finalMDL As Double = Nothing
                        Dim reportedMDL As String = Nothing

                        'Get Metal Result
                        Select Case metalName
                            Case "Ag"
                                If metalMass = My.Settings.AgMass Then
                                    analysisCode = "SILVER-ICP-MS"
                                    metalAnalyte = "Silver"
                                    If mdlType = "wm" Then finalMDL = Form3.wmAgTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmAgTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Al"
                                If metalMass = My.Settings.AlMass Then
                                    analysisCode = "ALUMINUM-ICP-MS"
                                    metalAnalyte = "Aluminum"
                                    If mdlType = "wm" Then finalMDL = Form3.wmAlTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmAlTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "As"
                                If metalMass = My.Settings.AsMass Then
                                    analysisCode = "ARSENIC-ICP-MS"
                                    metalAnalyte = "Arsenic"
                                    If mdlType = "wm" Then finalMDL = Form3.wmAsTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmAsTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjAsTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "B"
                                If metalMass = My.Settings.BMass Then
                                    analysisCode = "BORON-ICP-MS"
                                    metalAnalyte = "Boron"
                                    If mdlType = "wm" Then finalMDL = Form3.wmBTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmBTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Ba"
                                If metalMass = My.Settings.BaMass Then
                                    analysisCode = "BARIUM-ICP-MS"
                                    metalAnalyte = "Barium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmBaTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmBaTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Be"
                                If metalMass = My.Settings.BeMass Then
                                    analysisCode = "BERYLLIUM-ICP-MS"
                                    metalAnalyte = "Beryllium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmBeTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmBeTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Cd"
                                If metalMass = My.Settings.CdMass Then
                                    analysisCode = "CADMIUM-ICP-MS"
                                    metalAnalyte = "Cadmium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmCdTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmCdTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjCdTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Co"
                                If metalMass = My.Settings.CoMass Then
                                    analysisCode = "COBALT-ICP-MS"
                                    metalAnalyte = "Cobalt"
                                    If mdlType = "wm" Then finalMDL = Form3.wmCoTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmCoTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Cr"
                                If metalMass = My.Settings.CrMass Then
                                    analysisCode = "CHROMIUM-ICP-MS"
                                    metalAnalyte = "Chromium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmCrTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmCrTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjCrTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Cu"
                                If metalMass = My.Settings.CuMass Then
                                    analysisCode = "COPPER-ICP-MS"
                                    metalAnalyte = "Copper"
                                    If mdlType = "wm" Then finalMDL = Form3.wmCuTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmCuTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjCuTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Mn"
                                If metalMass = My.Settings.MnMass Then
                                    analysisCode = "MANGANESE-ICP-MS"
                                    metalAnalyte = "Manganese"
                                    If mdlType = "wm" Then finalMDL = Form3.wmMnTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmMnTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Mo"
                                If metalMass = My.Settings.MoMass Then
                                    analysisCode = "MOLYBDENUM-ICP-MS"
                                    metalAnalyte = "Molybdenum"
                                    If mdlType = "wm" Then finalMDL = Form3.wmMoTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmMoTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Ni"
                                If metalMass = My.Settings.NiMass Then
                                    analysisCode = "NICKEL-ICP-MS"
                                    metalAnalyte = "Nickel"
                                    If mdlType = "wm" Then finalMDL = Form3.wmNiTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmNiTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "P"
                                If metalMass = My.Settings.PMass Then
                                    analysisCode = "PHOSPHOR-ICP-MS"
                                    metalAnalyte = "Phosphorus"
                                    If mdlType = "wm" Then finalMDL = Form3.wmPTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmPTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Pb"
                                If metalMass = My.Settings.PbMass Then
                                    If My.Settings.leadConsumer = True Then
                                        analysisCode = "LEAD_CONSUM"
                                        metalAnalyte = "Lead"
                                    Else
                                        analysisCode = "LEAD-ICP-MS"
                                        metalAnalyte = "Lead"
                                    End If
                                    If mdlType = "wd" Then finalMDL = Form3.wdPbTextBox.Text
                                    If mdlType = "dd" Then finalMDL = Form3.ddPbTextBox.Text
                                    If mdlType = "wm" Then finalMDL = Form3.wmPbTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmPbTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjPbTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Pd"
                                If metalMass = My.Settings.PdMass Then
                                    analysisCode = "PALLADIUM-ICP-MS"
                                    metalAnalyte = "Palladium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmPdTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmPdTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Sb"
                                If metalMass = My.Settings.SbMass Then
                                    analysisCode = "ANTIMONY-ICP-MS"
                                    metalAnalyte = "Antimony"
                                    If mdlType = "wm" Then finalMDL = Form3.wmSbTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmSbTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Se"
                                If metalMass = My.Settings.SeMass Then
                                    analysisCode = "SELENIUM-ICP-MS"
                                    metalAnalyte = "Selenium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmSeTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmSeTextBox.Text
                                    If mdlType = "wj" Then finalMDL = Form3.wjSeTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Sn"
                                If metalMass = My.Settings.SnMass Then
                                    analysisCode = "TIN-ICP-MS"
                                    metalAnalyte = "Tin"
                                    If mdlType = "wd" Then finalMDL = Form3.wdSnTextBox.Text
                                    If mdlType = "dd" Then finalMDL = Form3.ddSnTextBox.Text
                                    If mdlType = "wm" Then finalMDL = Form3.wmSnTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmSnTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Ti"
                                If metalMass = My.Settings.TiMass Then
                                    analysisCode = "TITANIUM-ICP-MS"
                                    metalAnalyte = "Titanium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmTiTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmTiTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Tl"
                                If metalMass = My.Settings.TlMass Then
                                    analysisCode = "THALLIUM-ICP-MS"
                                    metalAnalyte = "Thallium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmTlTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmTlTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "V"
                                If metalMass = My.Settings.VMass Then
                                    analysisCode = "VANADIUM-ICP-MS"
                                    metalAnalyte = "Vanadium"
                                    If mdlType = "wm" Then finalMDL = Form3.wmVTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmVTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case "Zn"
                                If metalMass = My.Settings.ZnMass Then
                                    analysisCode = "ZINC-ICP-MS"
                                    metalAnalyte = "Zinc"
                                    If mdlType = "wm" Then finalMDL = Form3.wmZnTextBox.Text
                                    If mdlType = "dm" Then finalMDL = Form3.dmZnTextBox.Text
                                Else
                                    GoTo nextLine
                                End If
                            Case Else
                                GoTo nextLine
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

                            If roundResult < 1000000 Then
                                metalResult = CStr(FormatNumber(roundResult, "000000"))
                            End If

                            If roundResult < 100000 Then
                                metalResult = CStr(FormatNumber(roundResult, "00000"))
                            End If

                            If roundResult < 10000 Then
                                metalResult = CStr(FormatNumber(roundResult, "0000"))
                            End If

                            If roundResult < 1000 Then
                                metalResult = CStr(FormatNumber(roundResult, "000"))
                            End If

                            If roundResult < 100 Then
                                metalResult = CStr(FormatNumber(roundResult, "00.0"))
                            End If

                            If roundResult < 10 Then
                                metalResult = CStr(FormatNumber(roundResult, "0.00"))
                            End If

                            If roundResult < 1 Then
                                metalResult = CStr(FormatNumber(roundResult, "0.000"))
                            End If

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

                        '******************************************************************************************

                        'Put revised data into export array




                    End If

            End If 
        Next

        'Update datagridview1
        Call displayArray()

    End Sub

    Private Sub returnButton_Click(sender As Object, e As EventArgs) Handles returnButton.Click

        'Put datagridview1 data back into export array (or sample array if it's easier)**************************************

        My.Settings.skipSample = False
        Me.Close()
    End Sub

    Private Sub ExitProgramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitProgramToolStripMenuItem.Click
        My.Settings.exitProgram = True
        Me.Close()
    End Sub

    Private Sub SkipButton_Click(sender As Object, e As EventArgs) Handles skipButton.Click
        My.Settings.skipSample = True
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