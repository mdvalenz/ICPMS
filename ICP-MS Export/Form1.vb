Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1

    Private Sub fileLocationButton_Click(sender As Object, e As EventArgs) Handles fileLocationButton.Click
        OpenFileDialog1.Title = "Please Select your file"
        OpenFileDialog1.InitialDirectory = "L:\eRecords\Chemistry\Instrument Data\ICP-2\Report data\"
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub importButton_Click(sender As Object, e As EventArgs) Handles exportButton.Click

        If fileLocationTextBox.Text = "" Then
            MsgBox("A run file was not chosen, please choose one and try again.", vbMsgBoxSetForeground)
            GoTo endit
        End If

        If analystTextBox.Text = "" Then
            MsgBox("Analyst initials are missing, please enter them and try again.", vbMsgBoxSetForeground)
            GoTo endit
        End If

        'Declare Variables
        Dim ExcelApp As Excel.Application
        Dim ExcelWorkBooks As Excel.Workbooks
        Dim ExcelWorkSheet As Excel.Worksheet = Nothing

        'Open the run file
        ExcelApp = CreateObject("Excel.Application")
        ExcelWorkBooks = ExcelApp.Workbooks
        ExcelWorkBooks.OpenText(Filename:="""" & fileLocationTextBox.Text & """", DataType:=Excel.XlTextParsingType.xlDelimited, Comma:=True)
        ExcelApp.Visible = True

        'Declare More Variables
        Dim sampleID As String = Nothing
        Dim sampleID2 As String = Nothing
        Dim sampleID3 As String = Nothing
        Dim sampleID4 As String = Nothing
        Dim sampleID5 As String = Nothing

        Dim analysisCode As String = Nothing
        Dim endDate As String = Nothing
        Dim ExcelWorkbook As Object
        Dim metalAnalyst As String = Nothing

        'Assign Main Variables
        ExcelWorkbook = ExcelApp.Application
        endDate = Date.Today
        metalAnalyst = analystTextBox.Text

        'Get Metals Results
        Dim lastRow As String
        lastRow = ExcelApp.ActiveSheet.UsedRange.Rows.Count

        Dim rowNumber As Integer = 1
        Dim colNumber As Integer = 1
        Dim metalName As String = Nothing
        Dim metalMass As String = Nothing
        Dim metalResult As String = Nothing
        Dim metalAnalyte As String = Nothing

        'Delete the temporary Worksheets, if they exist
        For Each ExcelWorkSheet In ExcelApp.Worksheets
            If ExcelWorkSheet.Name = "temp" Or ExcelWorkSheet.Name = "temp2" Then
                ExcelApp.DisplayAlerts = False
                ExcelWorkSheet.Delete()
                ExcelApp.DisplayAlerts = True
            End If
        Next ExcelWorkSheet

        'Create temp worksheet
        Dim newWorksheet As Excel.Worksheet
        newWorksheet = CType(ExcelApp.Worksheets.Add(), Excel.Worksheet)
        newWorksheet.Name = "temp"

        'Create temp2 worksheet
        newWorksheet = CType(ExcelApp.Worksheets.Add(), Excel.Worksheet)
        newWorksheet.Name = "temp2"

        ExcelApp.Sheets(My.Settings.fileName).Activate()

        Dim IDCheck As String = Nothing
        Dim letterCheck As String = Nothing
        Dim leadCount As Integer = 0
        Dim sampleArrayCount As Integer = 0
        Dim callSampleSettingsForm As New Form4

        For rowNumber = 1 To lastRow + 1

            'Show progress bar
            Me.Show()
            progressBarLabel.Visible = True
            progressBarLabel.Text = "Reading data from file"
            ProgressBar1.Visible = True
            ProgressBar1.Value = CInt((rowNumber / lastRow) * 100)
            ProgressBar1.Refresh()
            ProgressBar1.Update()

            If rowNumber <> 1 Then rowNumber = rowNumber - 2
            Dim cellValue As String = ExcelApp.Cells(rowNumber, 2).Value
            If cellValue = "" Then
                sampleID = UCase(ExcelApp.Cells(rowNumber, 1).value)

                Dim mdlType As String = Nothing
                Dim metalUnits As String = Nothing

                IDCheck = Microsoft.VisualBasic.Right(sampleID, 1)
                letterCheck = Microsoft.VisualBasic.Left(sampleID, 1)

                If IsNumeric(letterCheck) = False Then
                    If IsNumeric(IDCheck) = True Then

                        If sampleID Like "ALT*" Then GoTo duplicateContinue
                        If sampleID.Length > 7 Then GoTo duplicateContinue

                        Me.TopMost = True
                        Me.Focus()
                        Me.TopMost = False

                        'Set settings
                        My.Settings.ppbCheck = True
                        My.Settings.ppmCheck = False
                        My.Settings.setSigFigs = "2"
                        My.Settings.reviewExport = True
                        My.Settings.leadConsumer = False
                        My.Settings.wetCheck = True
                        My.Settings.dryCheck = False
                        My.Settings.microwaveCheck = True
                        My.Settings.diluteCheck = False
                        My.Settings.juiceCheck = False

                        'MDLs
                        mdlType = "wm"

                        'Units
                        metalUnits = "ppb"

                        'Get analytes and results
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
                                            finalMDL = Form3.wmAgTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Al"
                                        If metalMass = My.Settings.AlMass Then
                                            finalMDL = Form3.wmAlTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "As"
                                        If metalMass = My.Settings.AsMass Then
                                            finalMDL = Form3.wmAsTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "B"
                                        If metalMass = My.Settings.BMass Then
                                            finalMDL = Form3.wmBTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Ba"
                                        If metalMass = My.Settings.BaMass Then
                                            finalMDL = Form3.wmBaTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Be"
                                        If metalMass = My.Settings.BeMass Then
                                            finalMDL = Form3.wmBeTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Cd"
                                        If metalMass = My.Settings.CdMass Then
                                            finalMDL = Form3.wmCdTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Co"
                                        If metalMass = My.Settings.CoMass Then
                                            finalMDL = Form3.wmCoTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Cr"
                                        If metalMass = My.Settings.CrMass Then
                                            finalMDL = Form3.wmCrTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Cu"
                                        If metalMass = My.Settings.CuMass Then
                                            finalMDL = Form3.wmCuTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Mn"
                                        If metalMass = My.Settings.MnMass Then
                                            finalMDL = Form3.wmMnTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Mo"
                                        If metalMass = My.Settings.MoMass Then
                                            finalMDL = Form3.wmMoTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Ni"
                                        If metalMass = My.Settings.NiMass Then
                                            finalMDL = Form3.wmNiTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "P"
                                        If metalMass = My.Settings.PMass Then
                                            finalMDL = Form3.wmPTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Pb"
                                        If metalMass = My.Settings.PbMass Then
                                            finalMDL = Form3.wmPbTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Pd"
                                        If metalMass = My.Settings.PdMass Then
                                            finalMDL = Form3.wmPdTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Sb"
                                        If metalMass = My.Settings.SbMass Then
                                            finalMDL = Form3.wmSbTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Se"
                                        If metalMass = My.Settings.SeMass Then
                                            finalMDL = Form3.wmSeTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Sn"
                                        If metalMass = My.Settings.SnMass Then
                                            finalMDL = Form3.wmSnTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Ti"
                                        If metalMass = My.Settings.TiMass Then
                                            finalMDL = Form3.wmTiTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Tl"
                                        If metalMass = My.Settings.TlMass Then
                                            finalMDL = Form3.wmTlTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "V"
                                        If metalMass = My.Settings.VMass Then
                                            finalMDL = Form3.wmVTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case "Zn"
                                        If metalMass = My.Settings.ZnMass Then
                                            finalMDL = Form3.wmZnTextBox.Text
                                        Else
                                            GoTo nextLine
                                        End If
                                    Case Else
                                        GoTo nextLine
                                End Select

                                Dim roundResult As Double = Nothing
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
                                Else
                                    metalResult = roundResult

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
                                        metalResult = CStr(FormatNumber(roundResult, "00"))
                                    End If

                                    If roundResult < 10 Then
                                        metalResult = CStr(FormatNumber(roundResult, "0.0"))
                                    End If

                                    If roundResult < 1 Then
                                        metalResult = CStr(FormatNumber(roundResult, "0.00"))
                                    End If

                                End If

                                reportedMDL = finalMDL

                                'Add data to sample array
                                sampleArrayCount = sampleArrayCount + 1
                                My.Settings.sampleArrayCount = sampleArrayCount
                                ReDim Preserve SampleArray(9, sampleArrayCount)

                                SampleArray(0, sampleArrayCount) = False
                                SampleArray(1, sampleArrayCount) = sampleID
                                SampleArray(2, sampleArrayCount) = metalName
                                SampleArray(3, sampleArrayCount) = metalTest
                                SampleArray(4, sampleArrayCount) = finalMDL
                                SampleArray(5, sampleArrayCount) = "ppb"
                                SampleArray(6, sampleArrayCount) = sigFigMax
                                SampleArray(7, sampleArrayCount) = roundResult
                                SampleArray(8, sampleArrayCount) = metalResult
                                SampleArray(9, sampleArrayCount) = reportedMDL

                            End If

nextLine:
                            rowNumber = rowNumber + 1
                            ExcelApp.Sheets(My.Settings.fileName).Activate()
                            cellValue = ExcelApp.Cells(rowNumber, 2).Value

                            analysisCode = Nothing
                            metalAnalyte = Nothing

                        Loop

                    Else

                        Do While cellValue = "" And rowNumber < lastRow + 1
                            rowNumber = rowNumber + 1
                            cellValue = ExcelApp.Cells(rowNumber, 2).Value
                        Loop

                        Do While cellValue <> "" And rowNumber < lastRow + 1
                            rowNumber = rowNumber + 1
                            cellValue = ExcelApp.Cells(rowNumber, 2).Value
                        Loop

                    End If
                Else

duplicateContinue:

                    Do While cellValue = "" And rowNumber < lastRow + 1
                        rowNumber = rowNumber + 1
                        cellValue = ExcelApp.Cells(rowNumber, 2).Value
                    Loop

                    Do While cellValue <> "" And rowNumber < lastRow + 1
                        rowNumber = rowNumber + 1
                        cellValue = ExcelApp.Cells(rowNumber, 2).Value
                    Loop

                End If
            End If
            rowNumber = rowNumber + 1
        Next

        'Resetting progress bar
        Me.Show()
        progressBarLabel.Text = ""
        progressBarLabel.Text = "Processing Samples"
        ProgressBar1.Value = 0
        ProgressBar1.Refresh()
        ProgressBar1.Update()

        'Get one sample number into the export array and go to form 4 for changes
        Dim exportArrayCount As Integer = 0
        Dim analyteCheck As String = Nothing

        For rowNum = 0 To CInt(My.Settings.sampleArrayCount)

            ProgressBar1.Value = CInt((rowNum / CInt(My.Settings.sampleArrayCount)) * 100)
            ProgressBar1.Refresh()
            ProgressBar1.Update()

            'Check for duplicates and remove first entry if found
            For rowNumDup = rowNum + 1 To CInt(My.Settings.sampleArrayCount)
                If SampleArray(1, rowNum) = SampleArray(1, rowNumDup) And SampleArray(2, rowNum) = SampleArray(2, rowNumDup) Then
                    SampleArray(1, rowNum) = ""
                End If
            Next

            If SampleArray(1, rowNum) <> "" Then

                exportArrayCount = exportArrayCount + 1
                My.Settings.exportArrayCount = exportArrayCount
                ReDim Preserve exportArray(9, exportArrayCount)

                'Transfer information from sample array to the export array
                exportArray(0, exportArrayCount) = SampleArray(0, rowNum)
                exportArray(1, exportArrayCount) = SampleArray(1, rowNum)
                exportArray(2, exportArrayCount) = SampleArray(2, rowNum)
                exportArray(3, exportArrayCount) = SampleArray(3, rowNum)
                exportArray(4, exportArrayCount) = SampleArray(4, rowNum)
                exportArray(5, exportArrayCount) = SampleArray(5, rowNum)
                exportArray(6, exportArrayCount) = SampleArray(6, rowNum)
                exportArray(7, exportArrayCount) = SampleArray(7, rowNum)
                exportArray(8, exportArrayCount) = SampleArray(8, rowNum)
                exportArray(9, exportArrayCount) = SampleArray(9, rowNum)

                'Open Form 4 for one sample
                If rowNum = CInt(My.Settings.sampleArrayCount) Then
                    My.Settings.sampleID = exportArray(1, exportArrayCount)

                    'Go to Form 4 and show results and allow changes
                    callSampleSettingsForm.ShowDialog(Me)

                    If My.Settings.exportNow = True Then

                        'Clear the rest of the Sample Array
                        For rowNumExportNow = rowNum + 1 To CInt(My.Settings.sampleArrayCount)
                            SampleArray(1, rowNumExportNow) = ""
                        Next

                        GoTo exportNow

                    End If

                    If My.Settings.exitProgram = True Then
                        MsgBox("The program has been canceled." & vbCrLf & "Nothing was exported to LabWorks.", vbMsgBoxSetForeground)
                        'Close Run File
                        ExcelApp.Sheets(My.Settings.fileName).Activate()
                        ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

                        'Release Excel Objects and close Excel
                        ExcelApp.Visible = True
                        If Not ExcelWorkSheet Is Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet)
                            ExcelWorkSheet = Nothing
                        End If

                        If Not ExcelWorkbook Is Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook)
                            ExcelWorkbook = Nothing
                        End If

                        If Not ExcelWorkBooks Is Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBooks)
                            ExcelWorkBooks = Nothing
                        End If

                        ExcelApp.Quit()
                        If Not ExcelApp Is Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp)
                            ExcelApp = Nothing
                        End If
                        GoTo endit
                    End If

                    exportArrayCount = 0
                    My.Settings.exportArrayCount = exportArrayCount
                    ReDim exportArray(9, exportArrayCount)

                Else

                    If SampleArray(1, rowNum) <> SampleArray(1, rowNum + 1) Then
                        My.Settings.sampleID = exportArray(1, exportArrayCount)

                        'Go to Form 4 and show results and allow changes
                        callSampleSettingsForm.ShowDialog(Me)

                        If My.Settings.exportNow = True Then

                            'Clear the rest of the Sample Array
                            For rowNumExportNow = rowNum + 1 To CInt(My.Settings.sampleArrayCount)
                                SampleArray(1, rowNumExportNow) = ""
                            Next

                            GoTo exportNow

                        End If

                        If My.Settings.exitProgram = True Then
                            MsgBox("The program has been canceled." & vbCrLf & "Nothing was exported to LabWorks.", vbMsgBoxSetForeground)
                            'Close Run File
                            ExcelApp.Sheets(My.Settings.fileName).Activate()
                            ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

                            'Release Excel Objects and close Excel
                            ExcelApp.Visible = True
                            If Not ExcelWorkSheet Is Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet)
                                ExcelWorkSheet = Nothing
                            End If

                            If Not ExcelWorkbook Is Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook)
                                ExcelWorkbook = Nothing
                            End If

                            If Not ExcelWorkBooks Is Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBooks)
                                ExcelWorkBooks = Nothing
                            End If

                            ExcelApp.Quit()
                            If Not ExcelApp Is Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp)
                                ExcelApp = Nothing
                            End If
                            GoTo endit
                        End If

                        exportArrayCount = 0
                        My.Settings.exportArrayCount = exportArrayCount
                        ReDim exportArray(9, exportArrayCount)

                    End If
                End If
            End If
nextSample:
        Next

        'Work through revised sample array and add to text or numeric temp sheet

        'Enter headers for temp sheet
        ExcelApp.Sheets("temp").Activate()
        ExcelApp.Range("A1").Select()

        'Add Headers
        ExcelApp.Range("A1").Value = "SIDN"
        ExcelApp.Range("B1").Value = "ACODE"
        ExcelApp.Range("C1").Value = "ANLNAME"
        ExcelApp.Range("D1").Value = "RLT"
        ExcelApp.Range("E1").Value = "RMDL"
        ExcelApp.Range("F1").Value = "RUNT"
        ExcelApp.Range("G1").Value = "ASTD"
        ExcelApp.Range("H1").Value = "AEND"
        ExcelApp.Range("I1").Value = "AANALYST"

        'change format of rlt and rmdl columns to text
        ExcelApp.Columns(4).EntireColumn.NumberFormat = "@"
        'ExcelApp.Columns(5).EntireColumn.NumberFormat = "@"

        'Enter headers for temp2 sheet
        ExcelApp.Sheets("temp2").Activate()
        ExcelApp.Range("A1").Select()

        'Add Headers
        ExcelApp.Range("A1").Value = "SIDN"
        ExcelApp.Range("B1").Value = "ACODE"
        ExcelApp.Range("C1").Value = "ANLNAME"
        ExcelApp.Range("D1").Value = "RLT"
        ExcelApp.Range("E1").Value = "RMDL"
        ExcelApp.Range("F1").Value = "RUNT"
        ExcelApp.Range("G1").Value = "ASTD"
        ExcelApp.Range("H1").Value = "AEND"
        ExcelApp.Range("I1").Value = "AANALYST"

        'change format of rlt and rmdl columns to text
        ExcelApp.Columns(4).EntireColumn.NumberFormat = "@"
        'ExcelApp.Columns(5).EntireColumn.NumberFormat = "@"

        'Work through array and enter results
        Dim nextRow As String = Nothing
        For rowNum = 0 To CInt(My.Settings.sampleArrayCount)
            If SampleArray(1, rowNum) <> "" Then

                'Select correct temp sheet
                If IsNumeric(SampleArray(8, rowNum)) Then
                    ExcelApp.Sheets("temp").Activate()
                    ExcelApp.Range("A1").Select()
                Else
                    ExcelApp.Sheets("temp2").Activate()
                    ExcelApp.Range("A1").Select()
                End If

                'Select next row
                nextRow = ExcelApp.Cells(ExcelApp.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1

                'Get analysis code and analyte name
                Select Case SampleArray(2, rowNum)
                    Case "Ag"
                        analysisCode = "SILVER-ICP-MS"
                        metalAnalyte = "Silver"
                    Case "Al"
                        analysisCode = "ALUMINUM-ICP-MS"
                        metalAnalyte = "Aluminum"
                    Case "As"
                        analysisCode = "ARSENIC-ICP-MS"
                        metalAnalyte = "Arsenic"
                    Case "B"
                        analysisCode = "BORON-ICP-MS"
                        metalAnalyte = "Boron"
                    Case "Ba"
                        analysisCode = "BARIUM-ICP-MS"
                        metalAnalyte = "Barium"
                    Case "Be"
                        analysisCode = "BERYLLIUM-ICP-MS"
                        metalAnalyte = "Beryllium"
                    Case "Cd"
                        analysisCode = "CADMIUM-ICP-MS"
                        metalAnalyte = "Cadmium"
                    Case "Co"
                        analysisCode = "COBALT-ICP-MS"
                        metalAnalyte = "Cobalt"
                    Case "Cr"
                        analysisCode = "CHROMIUM-ICP-MS"
                        metalAnalyte = "Chromium"
                    Case "Cu"
                        analysisCode = "COPPER-ICP-MS"
                        metalAnalyte = "Copper"
                    Case "Mn"
                        analysisCode = "MANGANESE-ICP-MS"
                        metalAnalyte = "Manganese"
                    Case "Mo"
                        analysisCode = "MOLYBDENUM-ICP-MS"
                        metalAnalyte = "Molybdenum"
                    Case "Ni"
                        analysisCode = "NICKEL-ICP-MS"
                        metalAnalyte = "Nickel"
                    Case "P"
                        analysisCode = "PHOSPHOR-ICP-MS"
                        metalAnalyte = "Phosphorus"
                    Case "Pb"
                        If SampleArray(0, rowNum) = True Then
                            analysisCode = "LEAD_CONSUM"
                            metalAnalyte = "Lead"
                        Else
                            analysisCode = "LEAD-ICP-MS"
                            metalAnalyte = "Lead"
                        End If
                    Case "Pd"
                        analysisCode = "PALLADIUM-ICP-MS"
                        metalAnalyte = "Palladium"
                    Case "Sb"
                        analysisCode = "ANTIMONY-ICP-MS"
                        metalAnalyte = "Antimony"
                    Case "Se"
                        analysisCode = "SELENIUM-ICP-MS"
                        metalAnalyte = "Selenium"
                    Case "Sn"
                        analysisCode = "TIN-ICP-MS"
                        metalAnalyte = "Tin"
                    Case "Ti"
                        analysisCode = "TITANIUM-ICP-MS"
                        metalAnalyte = "Titanium"
                    Case "Tl"
                        analysisCode = "THALLIUM-ICP-MS"
                        metalAnalyte = "Thallium"
                    Case "V"
                        analysisCode = "VANADIUM-ICP-MS"
                        metalAnalyte = "Vanadium"
                    Case "Zn"
                        analysisCode = "ZINC-ICP-MS"
                        metalAnalyte = "Zinc"
                End Select

                'Enter data
                ExcelApp.Range("A" & nextRow).Value = SampleArray(1, rowNum)
                ExcelApp.Range("B" & nextRow).Value = analysisCode
                ExcelApp.Range("C" & nextRow).Value = metalAnalyte
                ExcelApp.Range("D" & nextRow).Value = SampleArray(8, rowNum)
                ExcelApp.Range("E" & nextRow).Value = SampleArray(9, rowNum)
                ExcelApp.Range("F" & nextRow).Value = SampleArray(5, rowNum)
                ExcelApp.Range("G" & nextRow).Value = My.Settings.startDate
                ExcelApp.Range("H" & nextRow).Value = My.Settings.endDate
                ExcelApp.Range("I" & nextRow).Value = metalAnalyst
            End If
        Next

exportNow:
        ExcelApp.Sheets("temp").Activate()
        ExcelApp.Range("A1").Select()
        ExcelApp.Sheets("temp").Columns("A:I").AutoFit()

        Dim continueReview As String = Nothing
        continueReview = MsgBox("Please review the data." & vbCrLf & "Then click 'OK' to export the data to LabWorks", vbOKCancel + vbMsgBoxSetForeground, "Final Review")

        If continueReview = vbOK Then
        Else
            MsgBox("The program has been canceled." & vbCrLf & "Nothing was exported to LabWorks.", vbMsgBoxSetForeground)

            'Close Run File
            ExcelApp.Sheets(My.Settings.fileName).Activate()
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

            'Release Excel Objects and close Excel
            ExcelApp.Visible = True
            If Not ExcelWorkSheet Is Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet)
                ExcelWorkSheet = Nothing
            End If

            If Not ExcelWorkbook Is Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook)
                ExcelWorkbook = Nothing
            End If

            If Not ExcelWorkBooks Is Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBooks)
                ExcelWorkBooks = Nothing
            End If

            ExcelApp.Quit()
            If Not ExcelApp Is Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp)
                ExcelApp = Nothing
            End If
            GoTo endit
        End If

        'Copy temp worksheet
        ExcelApp.Sheets("temp").Copy()

        'Check if someone else is running the macro
        Dim LoopCount As String = 1

CheckLoop:
        If Not Dir("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv", vbDirectory) = vbNullString Then
            MsgBox("Someone else is running a macro, please wait..." & vbCrLf & "Attempt: " & LoopCount & " of 3", vbInformation + vbMsgBoxSetForeground)
            System.Threading.Thread.Sleep(3000)
            LoopCount = LoopCount + 1
            If LoopCount < 3 Then
                GoTo CheckLoop
            Else
                MsgBox("Error: Temp file already exists. Notify QA.", vbMsgBoxSetForeground)
                GoTo endit
            End If
        End If

        'Create temp CSV file and close
        ExcelApp.ActiveWorkbook.SaveAs(Filename:="L:\eRecords\Chemistry\CSV_Temp\Temp_CSV", FileFormat:=Excel.XlFileFormat.xlCSV)
        ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

        'Import Temp CSV File
        'Check for Windows version (32-bit or 64-bit)
        ChDir("L:\eRecords\Chemistry\CSV_Temp\")

        If FileFolderExists("C:\Program Files\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
            Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel.bat", vbNormalFocus)
        Else
            If FileFolderExists("C:\Program Files (x86)\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
                Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel_x64.bat", vbNormalFocus)
            Else
                MsgBox("Labworks not found, please contact QA.", vbMsgBoxSetForeground)
                GoTo endit
            End If
        End If

        'Delete the temp file in CSV Temp
        MsgBox("Click OK after Labworks has imported the file.", vbMsgBoxSetForeground)
        Kill("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv")



        'Copy temp2 worksheet
        ExcelApp.Sheets("temp2").Activate()
        ExcelApp.Range("A1").Select()
        ExcelApp.Sheets("temp2").Columns("A:I").AutoFit()
        ExcelApp.Sheets("temp2").Copy()

        'Check if someone else is running the macro
        Dim LoopCount2 As String = 1

CheckLoop2:
        If Not Dir("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv", vbDirectory) = vbNullString Then
            MsgBox("Someone else is running a macro, please wait..." & vbCrLf & "Attempt: " & LoopCount2 & " of 3", vbInformation + vbMsgBoxSetForeground)
            System.Threading.Thread.Sleep(3000)
            LoopCount2 = LoopCount2 + 1
            If LoopCount2 < 3 Then
                GoTo CheckLoop2
            Else
                MsgBox("Error: Temp file already exists. Notify QA.", vbMsgBoxSetForeground)
                GoTo endit
            End If
        End If

        'Create temp CSV file and close
        ExcelApp.ActiveWorkbook.SaveAs(Filename:="L:\eRecords\Chemistry\CSV_Temp\Temp_CSV", FileFormat:=Excel.XlFileFormat.xlCSV)
        ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

        'Import Temp CSV File
        'Check for Windows version (32-bit or 64-bit)
        ChDir("L:\eRecords\Chemistry\CSV_Temp\")

        If FileFolderExists("C:\Program Files\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
            Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel.bat", vbNormalFocus)
        Else
            If FileFolderExists("C:\Program Files (x86)\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
                Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel_x64.bat", vbNormalFocus)
            Else
                MsgBox("Labworks not found, please contact QA.", vbMsgBoxSetForeground)
                GoTo endit
            End If
        End If

        'Delete the temp file in CSV Temp
        MsgBox("Click OK after Labworks has imported the file.", vbMsgBoxSetForeground)
        Kill("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv")

        'Close Run File
        ExcelApp.Sheets(My.Settings.fileName).Activate()
        ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

        'Release Excel Objects and close Excel
        ExcelApp.Visible = True
        If Not ExcelWorkSheet Is Nothing Then
            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet)
            ExcelWorkSheet = Nothing
        End If

        If Not ExcelWorkbook Is Nothing Then
            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook)
            ExcelWorkbook = Nothing
        End If

        If Not ExcelWorkBooks Is Nothing Then
            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBooks)
            ExcelWorkBooks = Nothing
        End If

        ExcelApp.Quit()
        If Not ExcelApp Is Nothing Then
            Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp)
            ExcelApp = Nothing
        End If
endit:
        Me.Close()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim strm As System.IO.Stream = Nothing

        Try
            strm = OpenFileDialog1.OpenFile()
            fileLocationTextBox.Text = OpenFileDialog1.FileName.ToString()
            My.Settings.fileName = OpenFileDialog1.SafeFileName.ToString()

        Catch ex As System.IO.IOException
            MsgBox("Please close the file and try selecting again", vbMsgBoxSetForeground)
        End Try

        If Not (strm Is Nothing) Then
            'insert code to read the file data
            strm.Close()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Settings.exitProgram = False
        My.Settings.reviewExport = False
        My.Settings.exportNow = False
        My.Settings.endDate = DateTime.Now

        Form4.startDateTimePicker.Value = DateTime.Now.AddDays(-1)

    End Sub

    Private Function FileFolderExists(strFullPath As String) As Boolean
        'Macro Purpose: Check if a file or folder exists

        On Error GoTo EarlyExit

        If Not Dir(strFullPath, vbDirectory) = vbNullString Then
            FileFolderExists = True
        End If

EarlyExit:
        On Error GoTo 0

    End Function

    Private Sub SelectMassesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SelectMassesToolStripMenuItem1.Click
        Form2.ShowDialog()
    End Sub

    Private Sub ChangeDefaultMDLsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ChangeDefaultMDLsToolStripMenuItem1.Click
        Form3.ShowDialog()
    End Sub

End Class
