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

        Dim analysisCode As String
        Dim endDate As String
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
        My.Settings.ppbCheck = True
        My.Settings.ppmCheck = False
        My.Settings.setSigFigs = "2"
        My.Settings.reviewExport = True
        My.Settings.leadConsumer = False

        'Delete the temporary Worksheets, if they exist
        For Each ExcelWorkSheet In ExcelApp.Worksheets
            If ExcelWorkSheet.Name = "temp" Then
                ExcelApp.DisplayAlerts = False
                ExcelWorkSheet.Delete()
                ExcelApp.DisplayAlerts = True
            End If
        Next ExcelWorkSheet

        'Create temp worksheet
        Dim newWorksheet As Excel.Worksheet
        newWorksheet = CType(ExcelApp.Worksheets.Add(), Excel.Worksheet)
        newWorksheet.Name = "temp"

        ''Create temp worksheet
        'Dim newWorksheet2 As Excel.Worksheet
        'newWorksheet2 = CType(ExcelApp.Worksheets.Add(), Excel.Worksheet)
        'newWorksheet2.Name = "temp2"

        ExcelApp.Sheets(My.Settings.fileName).Activate()

        Dim count As String = 0
        Dim IDCheck As String = Nothing
        Dim letterCheck As String = Nothing
        Dim leadCount As Integer = 0

        For rowNumber = 1 To lastRow + 1
            If rowNumber <> 1 Then rowNumber = rowNumber - 2
            Dim cellValue As String = ExcelApp.Cells(rowNumber, 2).Value
            If cellValue = "" Then
                sampleID = Microsoft.VisualBasic.Left(ExcelApp.Cells(rowNumber, 1).value, 7)
                sampleID = UCase(sampleID)

                If sampleID = My.Settings.sampleID Then
                    IDCheck = Microsoft.VisualBasic.Right(sampleID, 1)
                    letterCheck = Microsoft.VisualBasic.Left(sampleID, 1)
                    If IsNumeric(letterCheck) = False Then
                        If IsNumeric(IDCheck) = True Then
                            MsgBox(sampleID & " is a duplicate sample." & vbCrLf & "Please enter your results into LabWorks manually.", vbMsgBoxSetForeground)
                        End If
                    End If

                    GoTo duplicateContinue
                End If

                IDCheck = Microsoft.VisualBasic.Right(sampleID, 1)
                letterCheck = Microsoft.VisualBasic.Left(sampleID, 1)

                If IsNumeric(letterCheck) = False Then
                    If IsNumeric(IDCheck) = True Then
                        'Call Sample Options form to determine MDLs and Units
                        My.Settings.sampleID = sampleID
                    End If
                End If

                Dim callSampleSettingsForm As New Form4
                Dim mdlType As String = Nothing
                Dim metalUnits As String = Nothing

                IDCheck = Microsoft.VisualBasic.Right(sampleID, 1)
                letterCheck = Microsoft.VisualBasic.Left(sampleID, 1)

                If IsNumeric(letterCheck) = False Then
                    If IsNumeric(IDCheck) = True Then

                        If sampleID Like "ALT*" Then GoTo duplicateContinue

                        Me.TopMost = True
                        Me.Focus()
                        Me.TopMost = False

                        callSampleSettingsForm.ShowDialog(Me)

                        If My.Settings.exportNow = True Then
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

                        If My.Settings.skipSample = True Then
                            GoTo duplicateContinue
                        End If

                        'Me.BringToFront()
                        ExcelApp.Sheets(My.Settings.fileName).Activate()

                        'MDLs
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

                                Dim nextRow As String = Nothing

                                'Enter Data
                                If count = 0 Then

                                    'Enter the data into the temp worksheet
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
                                    ExcelApp.Columns(5).EntireColumn.NumberFormat = "@"

                                    ''Add headers to temp2
                                    'ExcelApp.Sheets("temp2").Activate()
                                    'ExcelApp.Range("A1").Select()

                                    ''Add Headers
                                    'ExcelApp.Range("A1").Value = "SIDN"
                                    'ExcelApp.Range("B1").Value = "ACODE"
                                    'ExcelApp.Range("C1").Value = "ANLNAME"
                                    'ExcelApp.Range("D1").Value = "RLT"
                                    'ExcelApp.Range("E1").Value = "RMDL"
                                    'ExcelApp.Range("F1").Value = "RUNT"
                                    'ExcelApp.Range("G1").Value = "ASTD"
                                    'ExcelApp.Range("H1").Value = "AEND"
                                    'ExcelApp.Range("I1").Value = "AANALYST"

                                    'change format of rlt and rmdl columns to text
                                    ExcelApp.Columns(4).EntireColumn.NumberFormat = "@"
                                    'ExcelApp.Columns(5).EntireColumn.NumberFormat = "@"

                                End If

                                'If textResult = 1 Then
                                'Enter text results onto temp2 sheet
                                'ExcelApp.Sheets("temp2").Activate()
                                'ExcelApp.Range("A1").Select()
                                'Else
                                'Enter numerical results onto temp sheet
                                ExcelApp.Sheets("temp").Activate()
                                ExcelApp.Range("A1").Select()
                                'End If

                                'Select next row
                                nextRow = ExcelApp.Cells(ExcelApp.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1

                                'Enter data
                                ExcelApp.Range("A" & nextRow).Value = sampleID
                                ExcelApp.Range("B" & nextRow).Value = analysisCode
                                ExcelApp.Range("C" & nextRow).Value = metalAnalyte
                                ExcelApp.Range("D" & nextRow).Value = metalResult
                                ExcelApp.Range("E" & nextRow).Value = reportedMDL
                                ExcelApp.Range("F" & nextRow).Value = metalUnits
                                ExcelApp.Range("G" & nextRow).Value = My.Settings.startDate
                                ExcelApp.Range("H" & nextRow).Value = My.Settings.endDate
                                ExcelApp.Range("I" & nextRow).Value = metalAnalyst

                                count = count + 1

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

exportNow:
        'Check if analyst wants to see the CSV file
        If My.Settings.reviewExport = True Then

            'show file
            'ExcelApp.Sheets("temp2").Activate()
            'ExcelApp.Range("A1").Select()
            'ExcelApp.Sheets("temp2").Columns("A:I").AutoFit()

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
        Else
            Dim finalReview As String = Nothing
            finalReview = MsgBox("Would you like to review before exporting?", MsgBoxStyle.YesNoCancel + vbMsgBoxSetForeground, "Final Review")
            If finalReview = vbYes Then

                'show file
                'ExcelApp.Sheets("temp2").Activate()
                'ExcelApp.Range("A1").Select()
                'ExcelApp.Sheets("temp2").Columns("A:I").AutoFit()

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
            Else
                If finalReview = vbCancel Then
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
            End If
        End If

        'Go to the temp worksheet for numerical results
        ExcelApp.Sheets("temp").Activate()
        ExcelApp.Range("A1").Select()

        'Copy worksheet
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

        '        'Go to the temp2 worksheet for text results
        '        ExcelApp.Sheets("temp2").Activate()
        '        ExcelApp.Range("A1").Select()

        '        'Copy worksheet
        '        ExcelApp.Sheets("temp2").Copy()

        '        'Check if someone else is running the macro
        '        Dim LoopCount2 As String = 1

        'CheckLoop2:
        '        If Not Dir("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv", vbDirectory) = vbNullString Then
        '            MsgBox("Someone else is running a macro, please wait..." & vbCrLf & "Attempt: " & LoopCount2 & " of 3", vbInformation + vbMsgBoxSetForeground)
        '            System.Threading.Thread.Sleep(3000)
        '            LoopCount2 = LoopCount2 + 1
        '            If LoopCount2 < 3 Then
        '                GoTo CheckLoop2
        '            Else
        '                MsgBox("Error: Temp file already exists. Notify QA.", vbMsgBoxSetForeground)
        '                GoTo endit
        '            End If
        '        End If

        '        'Create temp2 CSV file and close
        '        ExcelApp.ActiveWorkbook.SaveAs(Filename:="L:\eRecords\Chemistry\CSV_Temp\Temp_CSV", FileFormat:=Excel.XlFileFormat.xlCSV)
        '        ExcelApp.ActiveWorkbook.Close(SaveChanges:=False)

        '        'Import Temp2 CSV File
        '        'Check for Windows version (32-bit or 64-bit)
        '        ChDir("L:\eRecords\Chemistry\CSV_Temp\")

        '        If FileFolderExists("C:\Program Files\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
        '            Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel.bat", vbNormalFocus)
        '        Else
        '            If FileFolderExists("C:\Program Files (x86)\PerkinElmer\LABWORKS64\POSTRESULTS6.exe") Then
        '                Call Shell(Environ$("COMSPEC") & " /c L:\eRecords\Chemistry\CSV_Temp\MCexcel_x64.bat", vbNormalFocus)
        '            Else
        '                MsgBox("Labworks not found, please contact QA.", vbMsgBoxSetForeground)
        '                GoTo endit
        '            End If
        '        End If

        '        'Delete the temp2 file in CSV Temp
        '        MsgBox("Click OK after Labworks has imported the file.", vbMsgBoxSetForeground)
        '        Kill("L:\eRecords\Chemistry\CSV_Temp\Temp_CSV.csv")

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
        Form2.Show()
    End Sub

    Private Sub ChangeDefaultMDLsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ChangeDefaultMDLsToolStripMenuItem1.Click
        Form3.Show()
    End Sub

End Class
