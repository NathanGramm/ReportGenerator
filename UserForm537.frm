'====================================================================================
'
'Author: Nathan Gramm
'Start Date: 6/23/2023
'Latest Edit Date: 9/18/2023
'For the data analysis reports of The Pennsylvania State University's Environmental
'Contaminants Laboratory (ECAL)
'
'====================================================================================

Private referenceWorkbookName, LCSLCSDWorkbookName As String

Private Sub Userform_Initialize()
    With MatrixComboBox
        .AddItem ("Water")
        .AddItem ("POCIS")
        .AddItem ("Sediment")
        .AddItem ("Tissue")
    End With
    SampleLabel.Visible = False
    SampleTextBox.Visible = False
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SelectFile(ByRef WorkbookName, Optional ByRef DisplayTextBox)
    With Application.FileDialog(msoFileDialogFilePicker)
      
      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Excel", "*.xlsx"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show Then
        WorkbookName = .SelectedItems(1)
        If Not IsMissing(DisplayTextBox) Then DisplayTextBox.Value = .SelectedItems(1)
      End If
   End With
End Sub

Private Sub SelectLCSLCSDFileButton_Click()
    Call SelectFile(LCSLCSDWorkbookName, LCSLCSDWorkbookTextBox)
End Sub

Private Sub SelectReferenceFileButton_Click()
    Call SelectFile(referenceWorkbookName, ReferenceWorkbookTextBox)
End Sub

Private Sub ReferenceWorkbookTextBox_Change()
    Call TextBoxsFilled
End Sub

Private Sub TextBoxsFilled()
    GenerateReport.Enabled = Len(ReferenceWorkbookTextBox.Text) > 0
End Sub

Private Sub MatrixComboBox_Change()
    Dim MatrixComboBoxValue As String
    MatrixComboBoxValue = MatrixComboBox.Value
    'Changes the text an visibility of the sample quantity label and textbox based on which matrix is selected
    SampleLabel.Visible = True
    SampleTextBox.Visible = True
    If MatrixComboBoxValue = "Water" Then
        SampleLabel.Caption = "Sample Volume (mL)"
    ElseIf MatrixComboBoxValue = "POCIS" Then
        SampleLabel.Caption = "Sorbent Weight (g)"
    ElseIf MatrixComboBoxValue = "Sediment" Or MatrixComboBoxValue = "Tissue" Then
        SampleLabel.Caption = "Sample Weight (g)"
    Else
        SampleLabel.Visible = False
        SampleTextBox.Visible = False
    End If
End Sub

Private Sub GenerateReport_Click()
    '==================  Initial Workbook Setup  ===================
    'Set the current workbook to our control file
    Dim controlFile As Workbook
    Set controlFile = Application.ActiveWorkbook
    Application.ScreenUpdating = False
    
    'Allow the search parameter to be used as the sheet name
    'so must replace invalid characters
    Dim sheetName As String
    sheetName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(SearchParameterTextBox.Text, ":", ""), "\", "."), "/", "."), "?", ""), "*", ""), "[", ""), "]", "")

    controlFile.Activate
    Dim ws As Worksheet, foundEmptyWorksheet, foundCoverPage, foundGlossary, foundDuplicateReport, foundLCSLCSD As Boolean
    For Each ws In controlFile.Worksheets
        If ws.Name = "Cover Page" Then
            foundCoverPage = True
        ElseIf ws.Name = "Glossary" Then
            foundGlossary = True
        ElseIf ws.Name = "PFAS " + sheetName + " Report" Then
            foundDuplicateReport = True
        ElseIf ws.Name = "LCSLCSD " + sheetName + " Report" Then
            foundLCSLCSD = True
        End If
        If ws.UsedRange.Address = "$A$1" And ws.Range("A1") = "" And ws.Name <> "Cover Page" And ws.Name <> "Glossary" Then
            foundEmptyWorksheet = True
            For j = 1 To controlFile.Worksheets.Count
                If controlFile.Worksheets(j).Name = "PFAS " + sheetName + " Report" Then
                    foundDuplicateReport = True
                    Exit For
                End If
            Next j
            If Not foundDuplicateReport Then
                ws.Name = "PFAS " + sheetName + " Report"
            Else
                foundEmptyWorksheet = False
            End If
        End If
    Next ws
    'If we found a sheet with the same name then handle it
    If foundDuplicateReport Then
        'If the duplicate report is the only file in the workbook we cannot delete it so we rename the report sheet, add a new sheet,
        'rename it the report's name, then delete the old sheet
        If controlFile.Worksheets.Count = 1 Then
            controlFile.Worksheets("PFAS " + sheetName + " Report").Name = "Null"
            Sheets.Add.Name = "PFAS " + sheetName + " Report"
            Application.DisplayAlerts = False
            controlFile.Worksheets("Null").Delete
            Application.DisplayAlerts = True
            
            'Since we created a new worksheet to avoid having no worksheet in the workbook, we set the boolean to true
            foundEmptyWorksheet = True
        Else
            'Otherwise we just delete the report with the same name as the report we want to add
            Application.DisplayAlerts = False
            controlFile.Worksheets("PFAS " + sheetName + " Report").Delete
            Application.DisplayAlerts = True
        End If
    End If
    'If we never found a sheet that could be used as our report sheet make one
    If Not foundEmptyWorksheet Then
        If foundLCSLCSD Then
            Sheets.Add(Before:=controlFile.Worksheets("LCSLCSD " + sheetName + " Report")).Name = "PFAS " + sheetName + " Report"
        Else
            Sheets.Add(Before:=controlFile.Worksheets(controlFile.Worksheets.Count)).Name = "PFAS " + sheetName + " Report"
        End If
    End If
    
    'Need to turn of Display Alerts so if we delete a worksheet no popup occurs
    Application.DisplayAlerts = False
    
    'If we needed to remake the Cover Page sheet, delete it then add a new one
    If foundCoverPage Then controlFile.Worksheets("Cover Page").Delete
    Sheets.Add(Before:=controlFile.Worksheets(1)).Name = "Cover Page"
    
    'If we needed to remake the Glossary sheet, delete it then add a new one
    If foundGlossary Then controlFile.Worksheets("Glossary").Delete
    Sheets.Add(After:=controlFile.Worksheets("Cover Page")).Name = "Glossary"
    
    'If we needed to remake the LCSLCSD Report sheet, delete it then add a new one
    If foundLCSLCSD Then controlFile.Worksheets("LCSLCSD " + sheetName + " Report").Delete
    If Len(LCSLCSDWorkbookName) > 0 Then Sheets.Add(After:=controlFile.Worksheets(controlFile.Worksheets.Count)).Name = "LCSLCSD " + sheetName + " Report"
    
    'Turn back on Display Alerts
    Application.DisplayAlerts = True
    '==============  End of Initial Workbook Setup  ================
    
    
    '==================  Cover Page Management  ====================
    'CoverPage Image
    Dim img As Picture
    Set img = controlFile.Worksheets("Cover Page").Pictures.Insert(AddIns("Report Generator").Path & "\Images\ReportCoverPage.png")
    With img
        .ShapeRange.LockAspectRatio = msoFalse
        .Height = Application.InchesToPoints(10.79) '11.6 for Hlengilizwe's computer
        .Width = Application.InchesToPoints(7.99)
        .TopLeftCell = controlFile.Worksheets("Cover Page").Cells(1, 1)
    End With
    
    'Cover Page TextBoxes
    Dim TextBoxPositionList As Variant, TextBoxPosition As Variant
    TextBoxPositionList = Array("B14", "B36") 'B14 and B37 for Hlengilizwe's computer
    For Each TextBoxPosition In TextBoxPositionList
        With controlFile.Worksheets("Cover Page").Range(TextBoxPosition)
            Dim ProjectInfoTextBox As TextBox
            Set ProjectInfoTextBox = .Parent.TextBoxes.Add(Top:=.Top, Left:=.Left, Width:=Application.InchesToPoints(3.73), Height:=Application.InchesToPoints(1.4))
            With ProjectInfoTextBox
                .ShapeRange.Fill.Visible = msoFalse
                .ShapeRange.Line.Visible = msoFalse
                If TextBoxPosition = "B14" Then
                    .Caption = "Client: " & ClientTextBox.Text & vbCrLf & _
                                "Project: " & ProjectTextBox.Text & vbCrLf & _
                                "Project Number: " & ProjectNumberTextBox.Text & vbCrLf & _
                                "Method: EPA 537.1"
                    .Font.Size = 14
                Else
                    .Caption = "Collection Date: " & CollectionDateTextBox.Text & vbCrLf & _
                                "Receipt Date: " & ReceiptDateTextBox.Text & vbCrLf & _
                                "Extraction Date: " & ExtractionDateTextBox.Text & vbCrLf & _
                                "Report Date: " & Format(Date, "m/dd/yyyy")
                    .Font.Size = 12
                End If
                With .Font
                    .Color = RGB(29, 65, 125)
                    .Name = "Arial Black"
                    .Bold = True
                End With
            End With
        End With
    Next TextBoxPosition
    '================  End of Cover Page Management  ===============
    
    
    '=====================  Glossary Management  ===================
    'Header Initialization
    With controlFile.Worksheets("Glossary")
        With .PageSetup
            .CenterHeader = "&G"
            .CenterFooter = "&G"
            .RightFooter = "&Kffffff&P-1   "
            With .CenterHeaderPicture
                .Filename = AddIns("Report Generator").Path & "\Images\ReportHeader.png"
                .Height = .Height * 0.95
                .Width = .Width * 0.95
            End With
            With .CenterFooterPicture
                .Filename = AddIns("Report Generator").Path & "\Images\ReportFooter.png"
                .Height = .Height * 0.95
                .Width = .Width * 0.95
            End With
        End With
        
        .Range("A1", "K1").Merge Across:=False
        .Range("A1").Font.Size = 18
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Value = "Glossary"
        .Range("A1").Font.Bold = True
        .Range("A1", "K1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
        .Range("A3").Value = "Abbreviation"
        .Range("A3").ColumnWidth = 10.56
        .Range("A3").Font.Bold = True
        .Range("A3").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .Range("D3").Value = "Definition of abbreviations that may or may not be present in the report"
        .Range("D3").Font.Bold = True
        .Range("D3", "K3").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .Range("A4").Value = "†"
        .Range("A6").Value = "% Rec"
        .Range("A7").Value = "CAL"
        .Range("A8").Value = "CCC"
        .Range("A9").Value = "FD"
        .Range("A10").Value = "FRB"
        .Range("A11").Value = "HR"
        .Range("A13").Value = "LCS"
        .Range("A14").Value = "LCSD"
        .Range("A15").Value = "LFB"
        .Range("A16").Value = "LRB"
        .Range("A17").Value = "m/z"
        .Range("A18").Value = "MAX RPD"
        .Range("A19").Value = "MDL"
        .Range("A22").Value = "MW"
        .Range("A23").Value = "ND"
        .Range("A24").Value = "PDS"
        .Range("A25").Value = "PFAS"
        .Range("A26").Value = "POCIS"
        .Range("A27").Value = "RL"
        .Range("A29").Value = "RPD"
        .Range("A30").Value = "RT"
        .Range("A31").Value = "S/N"
        
        .Range("D4", "K5").Merge Across:=False
        .Range("D4").Value = "Indicates that while compounds have been detected, their concentration is lower than the RL, and therefore represents an approximate value."
        .Range("D4").WrapText = True
        .Range("D6").Value = "Percent Recovery"
        .Range("D7").Value = "Calibration Standard"
        .Range("D8").Value = "Continuing Calibration Check"
        .Range("D9").Value = "Field Duplicate"
        .Range("D10").Value = "Field Reagent Blank"
        .Range("D11", "K12").Merge Across:=False
        .Range("D11").Value = "Half Range, computed according to Section 9.2.6 of EPA Method 537.1 (recovery criteria set at 50% - 150%)"
        .Range("D11").WrapText = True
        .Range("D13").Value = "Laboratory Control Samples"
        .Range("D14").Value = "Laboratory Control Sample Duplicates"
        .Range("D15").Value = "Laboratory Fortified Blank"
        .Range("D16").Value = "Laboratory Reagent Blank"
        .Range("D17").Value = "Mass to Charge Ratio"
        .Range("D18").Value = "Maximum Relative Percent Difference"
        .Range("D19", "K21").Merge Across:=False
        .Range("D19").Value = "Method Detection Limit, determined by fortifying, extracting, and analyzing seven replicate LFBs. The standard deviation was then multiplied by the t-value with 99% confidence"
        .Range("D19").WrapText = True
        .Range("D22").Value = "Molecular Weight"
        .Range("D23").Value = "Not Detected"
        .Range("D24").Value = "Primary Dilution Standard"
        .Range("D25").Value = "Per- or Poly- Fluorinated Alkyl Substances"
        .Range("D26").Value = "Polar Organic Chemical Integrative Sampler"
        .Range("D27", "K28").Merge Across:=False
        .Range("D27").Value = "Reporting Limit, used as reporting cut-off concentration, representing the lowest detectable concentration with acceptable reproducibility"
        .Range("D27").WrapText = True
        .Range("D29").Value = "Relative Percent Difference, a measure of relative difference between two points"
        .Range("D30").Value = "Retention Times"
        .Range("D31").Value = "Signal to Noise Ratio"
    End With
    '==================  End Glossary Management  ==================
    
    
    '===================  Report File Management  ==================
    With controlFile.Worksheets("PFAS " + sheetName + " Report")
        'Header Initialization
        With .PageSetup
            .CenterHeader = "&G"
            .CenterFooter = "&G"
            .RightFooter = "&Kffffff&P-1   "
            With .CenterHeaderPicture
                .Filename = AddIns("Report Generator").Path & "\Images\ReportHeader.png"
                .Height = .Height * 0.95
                .Width = .Width * 0.95
            End With
            With .CenterFooterPicture
                .Filename = AddIns("Report Generator").Path & "\Images\ReportFooter.png"
                .Height = .Height * 0.95
                .Width = .Width * 0.95
            End With
            .RightHeader = "Lab ID: " & LabIDTextBox.Text & vbCrLf & vbCrLf & vbCrLf & _
                           "Matrix: " & MatrixComboBox.Value
        End With
        
        'Freeze columns
        controlFile.Worksheets("PFAS " + sheetName + " Report").Activate
        With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
            .SplitColumn = 3
            .SplitRow = 0
            .FreezePanes = True
        End With
        
        'Static Column Widths
        .Columns("A").ColumnWidth = 25.11
        .Columns("B:C").ColumnWidth = 4.56
    
        'Static Row Heights
        .Rows("1").AutoFit
    
        'Static Borders
        .Range("A1", "C1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    
        'Formatting First Row
        .Rows("1").WrapText = True
        .Rows("1").HorizontalAlignment = xlCenter
        .Range("A1").HorizontalAlignment = xlLeft
        .Rows("1").VerticalAlignment = xlBottom
        .Rows("1").Font.Bold = True
        .Range("A1").Value = "Analyte Name"
        .Range("B1").Value = "Units"
        .Range("C1").Value = "RL"
    
        'Write out the list of analytes specified in the array below
        analyteList = Array("11Cl-PF3OUds", "9Cl-PF3ONS", "ADONA", "HFPO-DA", "N-EtFOSAA", "N-MeFOSAA", "Perfluorobutanoic acid", "Perfluorobutanesulfonic acid", "Perfluorodecanoic acid", "Perfluorododecanoic acid", "Perfluoroheptanoic acid", "Perfluorohexanoic acid", "Perfluorohexanesulfonic acid", "Perfluorononanoic acid", "Perfluorooctanoic acid", "Perfluorooctanesulfonic acid", "Perfluoropentanoic acid", "Perfluorotetradecanoic acid", "Perfluorotridecanoic acid", "Perfluoroundecanoic acid")
        RLList = Array(0.1, 0.2, 0.2, 0.4, 0.4, 0.2, 0.1, 0.4, 0.02, 0.2, 0.2, 0.02, 0.2, 0.2, 0.1, 0.4, 0.2, 0.2, 0.1, 0.04)
        'IDLList = Array(0.02, 0.02, 0.04, 0.19, 0.02, 0.03, 0.05, 0.01, 0.003, 0.03, 0.16, 0.11, 0.03, 0.22, 0.01, 0.04, 0.18, 0.01, 0.02, 0.16)
        'MDLList = Array(0.125, 0.125, 0.125, 0.25, 0.125, 0.15, 0.5, 0.125, 0.125, 0.125, 0.125, 0.225, 0.125, 0.125, 0.125, 0.25, 0.125, 0.125, 0.125, 0.125)
        Dim analyteListStartingRow, analyteListLength, LastRowOfAnalytesPosition As Integer
        analyteListStartingRow = 2
        analyteListLength = UBound(analyteList) + 1
        LastRowOfAnalytesPosition = analyteListStartingRow + analyteListLength - 1
        For Counter = 0 To analyteListLength - 1
            If Counter Mod 2 = 0 Then
                .Range(.Cells(analyteListStartingRow + Counter, 1), .Cells(analyteListStartingRow + Counter, 3)).Interior.Color = RGB(245, 245, 245)
            End If
            .Cells(analyteListStartingRow + Counter, 1).Value = analyteList(Counter)
            .Cells(analyteListStartingRow + Counter, 1).Font.Bold = True
            .Cells(analyteListStartingRow + Counter, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(analyteListStartingRow + Counter, 2).Value = IIf(MatrixComboBox.Value = "Water", "ng/L", "ng/g")
            .Cells(analyteListStartingRow + Counter, 2).HorizontalAlignment = xlCenter
            .Cells(analyteListStartingRow + Counter, 2).VerticalAlignment = xlCenter
            .Cells(analyteListStartingRow + Counter, 3).HorizontalAlignment = xlCenter
            .Cells(analyteListStartingRow + Counter, 3).VerticalAlignment = xlCenter
            If MatrixComboBox.Value <> "" And SampleTextBox.Value <> "" Then
                .Cells(analyteListStartingRow + Counter, 3).Value = Round(RLList(Counter) * IIf(MatrixComboBox.Value = "Water", 1000 / (SampleTextBox.Value / 1), 1 / SampleTextBox.Value), 2)
                .Cells(analyteListStartingRow + Counter, 3).NumberFormat = IIf(.Cells(analyteListStartingRow + Counter, 3).Value >= 10, "0", IIf(.Cells(analyteListStartingRow + Counter, 3).Value >= 1, "0.0", IIf(.Cells(analyteListStartingRow + Counter, 3).Value >= 0.1, "0.00", "0.000")))
                '.Cells(analyteListStartingRow + counter, 4).Value = Round(IIf(IDLList(counter) < RLList(counter), IDLList(counter), RLList(counter)) * IIf(MatrixComboBox.Value = "Water", 1000 / (SampleTextBox.Value / 1), 1 / SampleTextBox.Value), 2)
                '.Cells(analyteListStartingRow + counter, 5).Value = Round(MDLList(counter) * IIf(MatrixComboBox.Value = "Water", 1000 / (SampleTextBox.Value / 1), 1 / SampleTextBox.Value), 2)
            End If
            .Cells(analyteListStartingRow + Counter, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(analyteListStartingRow + Counter, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next Counter
    
        .Range(.Cells(LastRowOfAnalytesPosition, 1), .Cells(LastRowOfAnalytesPosition, 5)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
    
        'Set the position of the first sample column
        currentSampleCellColumn = 4
    
        'Open the reference Workbook
        If Len(referenceWorkbookName) > 0 Then
            Dim ReferenceFile, wb As Workbook, referenceFileAlreadyOpen, referenceFileProtected As Boolean
            referenceFileAlreadyOpen = False
            For Each wb In Workbooks
                If wb.Path = referenceWorkbookName Then
                    ReferenceFile = wb
                    referenceFileAlreadyOpen = True
                    Exit For
                End If
            Next wb
            If ReferenceFile = Empty Then Set ReferenceFile = Workbooks.Open(referenceWorkbookName)
        
            'Reset the analyteList to match the file output in the LCMS data report
            analyteList = Array("11Cl-PF3OUds", "9Cl-PF3ONS", "ADONA", "HFPO-DA", "N-EtFOSAA", "N-MeFOSAA", "PFBA", "PFBS", "PFDA", "PFDoA", "PFHpA", "PFHxA", "PFHxS", "PFNA", "PFOA", "PFOS", "PFPeA", "PFTeDA", "PFTrDA", "PFUdA")
            'Start the loop through the reference file sheets
            For Each ws In ReferenceFile.Worksheets
                If InStr(1, ws.Range("I9").Value, SearchParameterTextBox.Text) > 0 Then
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Borders(xlEdgeBottom).LineStyle = xlDouble
                    .Columns(currentSampleCellColumn).ColumnWidth = 16
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Value = ws.Range("I9").Value
                    Dim rowNum As Long
                    For Counter = 0 To analyteListLength - 1
                        With .Cells(analyteListStartingRow + Counter, currentSampleCellColumn)
                            If Counter Mod 2 = 0 Then
                                .Interior.Color = RGB(245, 245, 245)
                            End If
                            rowNum = ws.Columns(1).Find(What:=analyteList(Counter), LookIn:=xlValues, LookAt:=xlWhole).Row
                            .Value = IIf(ws.Cells(rowNum, 13) = "N/F" Or ws.Cells(rowNum, 13) = "N/A" Or ws.Cells(rowNum, 13) = "-" Or ws.Cells(rowNum, 13) < .Cells(analyteListStartingRow + Counter, 4), "ND", ws.Cells(rowNum, 13))
                            If InStr(1, .Value, "<") > 0 Then
                                .Value = Right(.Value, Len(.Value) - 1)
                                .NumberFormat = IIf(.Value >= 10, "0""†""", IIf(.Value >= 1, "0.0""†""", IIf(.Value >= 0.1, "0.00""†""", "0.000""†""")))
                            ElseIf IsNumeric(.Value) And .Value <> "" Then
                                .Value = Application.WorksheetFunction.RoundUp(.Value, IIf(.Value >= 1, 3 - InStr(1, .Value, "."), 4))
                                .NumberFormat = IIf(.Value >= 10, "0", IIf(.Value >= 1, "0.0", IIf(.Value >= 0.1, "0.00", "0.000")))
                                If .Value < controlFile.Worksheets("PFAS " + sheetName + " Report").Cells(analyteListStartingRow + Counter, 3).Value Then
                                    .NumberFormat = IIf(.Value >= 10, "0""†""", IIf(.Value >= 1, "0.0""†""", IIf(.Value >= 0.1, "0.00""†""", "0.000""†""")))
                                End If
                            End If
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                            If Counter = analyteListLength - 1 Then
                                .Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                            End If
                        End With
                    Next Counter
                    currentSampleCellColumn = currentSampleCellColumn + 1
                End If
            Next ws
            If Not referenceFileAlreadyOpen Then ReferenceFile.Close SaveChanges:=False
        End If
    End With
    '===============  End of Report File Management  ===============
    
    
    '==================  LCSLCSDFile Management  ===================
    If Len(LCSLCSDWorkbookName) > 0 Then
        Dim LCSLCSDFile, LCSLCSDFileAlreadyOpen, LCSLCSDFileProtected As Boolean
            LCSLCSDFileAlreadyOpen = False
            LCSLCSDFileProtected = False
            For Each wb In Workbooks
                If wb.Path = LCSLCSDWorkbookName Then
                    LCSLCSDFile = wb
                    LCSLCSDFileAlreadyOpen = True
                    Exit For
                End If
            Next wb
            If LCSLCSDFile = Empty Then Set LCSLCSDFile = Workbooks.Open(LCSLCSDWorkbookName)
        
        With controlFile.Worksheets("LCSLCSD " + sheetName + " Report")
            'Header Initialization
            With .PageSetup
                .CenterHeader = "&G"
                .CenterFooter = "&G"
                .RightFooter = "&Kffffff&P-1   "
                With .CenterHeaderPicture
                    .Filename = AddIns("Report Generator").Path & "\Images\ReportHeader.png"
                    .Height = .Height * 0.95
                    .Width = .Width * 0.95
                End With
                With .CenterFooterPicture
                    .Filename = AddIns("Report Generator").Path & "\Images\ReportFooter.png"
                    .Height = .Height * 0.95
                    .Width = .Width * 0.95
                End With
                .RightHeader = "Lab ID: " & LabIDTextBox.Text & vbCrLf & vbCrLf & vbCrLf & _
                               "Matrix: " & MatrixComboBox.Value
            End With
            .Range(.Cells(1, 1), .Cells(1, 13)).Merge Across:=False
            .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeBottom).LineStyle = xlDouble
            .Cells(1, 1).Value = "LCS LCSD Report"
            .Cells(1, 1).Font.Size = 18
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 1).HorizontalAlignment = xlCenter
            analyteListStartingRow = 13
            For Rw = 0 To analyteListLength
                For Col = 0 To 12
                    With .Cells(2 + Rw, Col + 1)
                        If Col = 0 Then
                            .Range(.Cells(2 + Rw, Col + 1), .Cells(2 + Rw, Col + 2)).Merge Across:=False
                        End If
                        .Value = LCSLCSDFile.Worksheets(1).Cells(analyteListStartingRow + Rw, Col + 1).Value
                        If IsNumeric(.Value) And .Value <> "" Then
                            .NumberFormat = IIf(.Value >= 10, "0", IIf(.Value >= 1, "0.0", IIf(.Value >= 0.1, "0.00", "0.000")))
                        End If
                        If Rw = 0 Then
                            .Font.Bold = True
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Columns(Col + 1).AutoFit
                        End If
                        If Rw Mod 2 = 0 And .Value <> "" Then
                            .Interior.Color = RGB(245, 245, 245)
                        End If
                    End With
                Next Col
            Next Rw
        End With
        If Not LCSLCSDFileAlreadyOpen Then LCSLCSDFile.Close SaveChanges:=False
    End If
    '================ End of LCSLCSDFile Management ================
    
    
    '======================  Print Settings  =======================
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        If ws.Name = "Cover Page" Then
            Application.ExecuteExcel4Macro ("Page.Setup(,,0.3,0.3,0.3,0.3,,,,,,,TRUE,,,,,0.3,0.3,,)")
            ws.PageSetup.CenterHorizontally = True
            ws.PageSetup.CenterVertically = True
        ElseIf InStr(1, ws.Name, "LCSLCSD") > 0 Or ws.Name = "Glossary" Then
            ws.PageSetup.PrintTitleRows = ws.Rows("1:2").Address
            ws.PageSetup.ScaleWithDocHeaderFooter = False
            Application.ExecuteExcel4Macro ("Page.Setup(,,0.3,0.3,2.15,0.3,,,,,,,{1,#N/A},,,,,0.3,0.3,,)")
        Else
            ws.PageSetup.PrintTitleColumns = ws.Columns("$A:$C").Address
            ws.PageSetup.ScaleWithDocHeaderFooter = False
            Application.ExecuteExcel4Macro ("Page.Setup(,,0.3,0.3,2.15,0.3,,,,,,,{#N/A,1},,,,,0.3,0.3,,)")
        End If
    Next ws
    '===================  End of Print Settings  ===================
    
    'Reset Focus onto the first worksheet in the workbook
    controlFile.Worksheets(1).Activate
    
    'Allow the screen to update and unload the userform
    Application.ScreenUpdating = True
    Unload Me
    
End Sub


