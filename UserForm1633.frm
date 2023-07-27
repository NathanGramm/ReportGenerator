VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1633 
   Caption         =   "Report Generation"
   ClientHeight    =   6588
   ClientLeft      =   -252
   ClientTop       =   -1080
   ClientWidth     =   10188
   OleObjectBlob   =   "UserForm1633.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1633"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'
'Author: Nathan Gramm
'Start Date: 6/23/2023
'Latest Edit Date: 7/27/2023
'For the data analysis reports of The Pennsylvania State University's IEE Department
'
'====================================================================================

Private referenceWorkbookName, LCSLCSDWorkbookName As String

Private Sub UserForm_Initialize()
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
    Dim ws As Worksheet, foundEmptyWorksheet, foundCoverPage, foundGlossary, foundDuplicateReport, foundCalibrationSolutions, foundLCSLCSD As Boolean
    For Each ws In controlFile.Worksheets
        If ws.Name = "Cover Page" Then
            foundCoverPage = True
        ElseIf ws.Name = "Glossary" Then
            foundGlossary = True
        ElseIf ws.Name = sheetName + " Report" Then
            foundDuplicateReport = True
        ElseIf ws.Name = "Calibration Solutions" Then
            foundCalibrationSolutions = True
        ElseIf ws.Name = "LCSLCSD " + sheetName + " Report" Then
            foundLCSLCSD = True
        End If
        If ws.UsedRange.Address = "$A$1" And ws.Range("A1") = "" And ws.Name <> "Cover Page" And ws.Name <> "Glossary" Then
            foundEmptyWorksheet = True
            For j = 1 To controlFile.Worksheets.Count
                If controlFile.Worksheets(j).Name = sheetName + " Report" Then
                    foundDuplicateReport = True
                    Exit For
                End If
            Next j
            If Not foundDuplicateReport Then
                controlFile.Worksheets(i).Name = sheetName + " Report"
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
            controlFile.Worksheets(sheetName + " Report").Name = "Null"
            Sheets.Add.Name = sheetName + " Report"
            Application.DisplayAlerts = False
            controlFile.Worksheets("Null").Delete
            Application.DisplayAlerts = True
            
            'Since we created a new worksheet to avoid having no worksheet in the workbook, we set the boolean to true
            foundEmptyWorksheet = True
        Else
            'Otherwise we just delete the report with the same name as the report we want to add
            Application.DisplayAlerts = False
            controlFile.Worksheets(sheetName + " Report").Delete
            Application.DisplayAlerts = True
        End If
    End If
    'If we never found a sheet that could be used as our report sheet make one
    If Not foundEmptyWorksheet Then
        If foundCalibrationSolutions Then
            Sheets.Add(Before:=controlFile.Worksheets("Calibration Solutions")).Name = sheetName + " Report"
        Else
            Sheets.Add(Before:=controlFile.Worksheets(controlFile.Worksheets.Count)).Name = sheetName + " Report"
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
    
    'If we needed to remake the Calibration Solutions sheet, delete it then add a new one
    If foundGlossary Then controlFile.Worksheets("Calibration Solutions").Delete
    Sheets.Add(After:=controlFile.Worksheets(sheetName + " Report")).Name = "Calibration Solutions"
    
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
        .Height = Application.InchesToPoints(10.78)
        .Width = Application.InchesToPoints(7.99)
        .TopLeftCell = controlFile.Worksheets("Cover Page").Cells(1, 1)
    End With
    
    'Cover Page TextBoxes
    Dim TextBoxPositionList As Variant, TextBoxPosition As Variant
    TextBoxPositionList = Array("B14", "B36")
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
                                "Method: CEC Quantification"
                    .Font.Size = 14
                Else
                    .Caption = "Collection Date: " & CollectionDateTextBox.Text & vbCrLf & _
                                "Receipt Date: " & ReceiptDateTextBox.Text & vbCrLf & _
                                "Analysis Date: " & AnalysisDateTextBox.Text
                    .Font.Size = 13
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
            .RightHeader = "Lab ID: " & LabIDTextBox.Text & vbCrLf & vbCrLf & vbCrLf
        End With
        
        .Range("A1", "J1").Merge Across:=False
        .Range("A1").Font.Size = 18
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Value = "Glossary"
        .Range("A1").Font.Bold = True
        .Range("A1", "J1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
        .Range("A4").Value = "Abbreviation"
        .Range("A4").ColumnWidth = 10.56
        .Range("A4").Font.Bold = True
        .Range("A4").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .Range("C4").Value = "Definition of abbreviations that may or may not be present in the report"
        .Range("C4").Font.Bold = True
        .Range("C4", "J4").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        .Range("A6").Value = "RL"
        .Range("A7").Value = "PFAS"
        .Range("A8").Value = "POCIS"
        .Range("A9").Value = "CEC"
        .Range("A10").Value = "SVOC"
        .Range("A11").Value = "LCS"
        .Range("A12").Value = "LCSD"
        .Range("A13").Value = "% Rec"
        .Range("A14").Value = "RPD"
        .Range("A15").Value = "MAX RPD"
        .Range("A16").Value = "LFB"
        .Range("A17").Value = "HR"
        .Range("A18").Value = "MDL"
        
        .Range("C6").Value = "Reporting Limit"
        .Range("C7").Value = "Per- or Poly- Fluorinated Alkyl Substances"
        .Range("C8").Value = "Polar Organic Chemical Integrative Sampler"
        .Range("C9").Value = "Contaminants of Emerging Concern"
        .Range("C10").Value = "Semi-Volatile Organic Compounds"
        .Range("C11").Value = "Laboratory Control Samples"
        .Range("C12").Value = "Laboratory Control Sample Duplicates"
        .Range("C13").Value = "Percent Recovery"
        .Range("C14").Value = "Relative Percent Difference, a measure of relative difference between two points"
        .Range("C15").Value = "Maximum Relative Percent Difference"
        .Range("C16").Value = "Laboratory Fortified Blank"
        .Range("C17").Value = "Half Range"
        .Range("C18").Value = "Method Detection Limit"
        
    End With
    '==================  End Glossary Management  ==================
    
    
    '===================  Report File Management  ==================
    With controlFile.Worksheets(sheetName + " Report")
        'Header Initialization
        With .PageSetup
            .CenterHeader = "&G"
            .CenterFooter = "&G"
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
            .LeftFooter = "*RL detertermined by fortifying, extracting and analyzing seven replicate LFBs. Mean and the Half Range (HR) computed according to Section 9.2.6 of EPA Method 537.1 (recovery criteria set at 50% - 150%)" & vbCrLf & vbCrLf
        End With
        
        'Freeze columns
        controlFile.Worksheets(sheetName + " Report").Activate
        With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
            .SplitColumn = 3
            .SplitRow = 0
            .FreezePanes = True
        End With
        
        'Static Column Widths
        .Columns("A").ColumnWidth = 24.33
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 7.3
    
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
        .Range("C1").Value = "*RL"
    
        'Write out the list of analytes specified in the array below
        analyteList = Array("11-Chloroeicosafluoro-3-oxaundecane-1-sulfonic acid", "9-Chlorohexadecafluoro-3-oxanonane-1-sulfonic acid", "3-Perfluoropropyl propanoic acid", "3-Perfluoropentyl propanoic acid", "1H,1H, 2H, 2H-Perfluorohexane sulfonic acid", "1H,1H, 2H, 2H-Perfluorooctane sulfonic acid", "3-Perfluoroheptyl propanoic acid", "1H,1H, 2H, 2H-Perfluorodecane sulfonic acid", "Hexafluoropropylene oxide dimer acid", "Perfluorobutanesulfonic acid", "Perfluorododecanesulfonic acid", "Perfluorodecanesulfonic acid", "Perfluoroheptanesulfonic acid", "Perfluorononanesulfonic acid", "Perfluoropentansulfonic acid", "4,8-Dioxa-3H-perfluorononanoic acid", "N-ethyl perfluorooctanesulfonamide", "N-ethyl perfluorooctanesulfonamidoacetic acid", "N-ethyl perfluorooctanesulfonamidoacetic acid (Branched)", "N-ethyl perfluorooctanesulfonamidoethanol", "Nonafluoro-3,6-dioxaheptanoic acid", "N-methyl perfluorooctanesulfonamide", "N-methyl perfluorooctanesulfonamidoacetic acid", _
        "N-methyl perfluorooctanesulfonamidoacetic acid (Branched)", "N-methyl perfluorooctanesulfonamidoethanol", "Perfluoro-4-oxapentanoic acid", "Perfluoro-5-oxahexanoic acid", "Perfluorobutanoic acid", "Perfluorodecanoic acid", "Perfluorododecanoic acid", "Perfluoro(2-ethoxyethane)sulfonic acid", "Perfluoroheptanoic acid", "Perfluorohexanoic acid", "Perfluorohexanesulfonic acid", "Perfluorohexanesulfonic acid (Branched)", "Perfluorononanoic acid", "Perfluorooctanoic acid", "Perfluorooctanesulfonic acid", "Perfluorooctanesulfonic acid (Branched)", "Perfluorooctanesulfonamide", "Perfluoropentanoic acid", "Perfluorotetradecanoic acid", "Perfluorotridecanoic acid", "Perfluoroundecanoic acid")
        RLList = Array(0.04, 0.04, 0.08, 0.4, 0.08, 0.16, 0.4, 0.08, 0.2, 0.02, 0.04, 0.02, 0.04, 0.04, 0.04, 0.08, 0.04, 0.078, 0.023, 0.2, 0.04, 0.04, 0.076, 0.024, 0.2, 0.04, 0.08, 0.4, 0.04, 0.04, 0.04, 0.04, 0.02, 0.016, 0.008, 0.04, 0.02, 0.016, 0.008, 0.02, 0.04, 0.1, 0.1, 0.1)
        Dim analyteListStartingRow, analyteListLength, LastRowOfAnalytesPosition As Integer
        analyteListStartingRow = 2
        analyteListLength = UBound(analyteList) + 1
        LastRowOfAnalytesPosition = analyteListStartingRow + analyteListLength - 1
        For counter = 0 To analyteListLength - 1
            .Cells(analyteListStartingRow + counter, 1).Value = analyteList(counter)
            .Cells(analyteListStartingRow + counter, 1).Font.Bold = True
            .Cells(analyteListStartingRow + counter, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(analyteListStartingRow + counter, 2).Value = IIf(MatrixComboBox.Value = "Water", "ng/L", "ng/g")
            .Cells(analyteListStartingRow + counter, 2).HorizontalAlignment = xlCenter
            .Cells(analyteListStartingRow + counter, 2).VerticalAlignment = xlCenter
            .Cells(analyteListStartingRow + counter, 3).HorizontalAlignment = xlCenter
            .Cells(analyteListStartingRow + counter, 3).VerticalAlignment = xlCenter
            If MatrixComboBox.Value <> "" And SampleTextBox.Value <> "" Then
                .Cells(analyteListStartingRow + counter, 3).Value = Round(RLList(counter) * IIf(MatrixComboBox.Value = "Water", 1000 / (SampleTextBox.Value / 1), 1 / SampleTextBox.Value), 2)
            End If
            .Cells(analyteListStartingRow + counter, 2).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(analyteListStartingRow + counter, 3).Borders(xlEdgeRight).LineStyle = xlContinuous
        Next counter
    
        .Range(.Cells(LastRowOfAnalytesPosition, 1), .Cells(LastRowOfAnalytesPosition, 3)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
    
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
            analyteList = Array("11Cl-PF3OUds", "9Cl-PF3ONS", "3:3FTCA/FPrPA", "FPePA", "4:2FTS", "6:2FTS", "7:3FTCA/FHpPA", "8:2FTS", "HFPO-DA", "L-PFBS", "L-PFDoS", "L-PFDS", "L-PFHpS", "L-PFNS", "L-PFPeS", "NaDONA", "NEtFOSA", "N-EtFOSAA", "N-EtFOSAA Branched", "NEtFOSE", "NFDHA/3,6-OPFHpA", "NMeFOSA", "N-MeFOSAA", "N-MeFOSAA Branched", "NMeFOSE", "PF4OPeA", "PF5OHxA", "PFBA", "PFDA", "PFDoA", "PFEESA", "PFHpA", "PFHxA", "PFHxS", "PFHxS Branched", "PFNA", "PFOA", "PFOS", "PFOS Branched", "PFOSA", "PFPeA", "PFTeDA", "PFTrDA", "PFUdA")
            'Start the loop through the reference file sheets
            For Each ws In ReferenceFile.Worksheets
                analyteNameList = Array(analyteListLength)
                analyteConcentrationList = Array(analyteListLength)
                If InStr(1, ws.Range("I9").Value, SearchParameterTextBox.Text) > 0 Then
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Borders(xlEdgeBottom).LineStyle = xlDouble
                    .Columns(currentSampleCellColumn).ColumnWidth = 16
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Value = ws.Range("I9").Value
                    Dim rowNum As Long
                    For counter = 0 To analyteListLength - 1
                        rowNum = ws.Columns(1).Find(What:=analyteList(counter), LookIn:=xlValues, LookAt:=xlWhole).Row
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Value = IIf(ws.Cells(rowNum, 13) = "N/F" Or ws.Cells(rowNum, 13) = "N/A" Or ws.Cells(rowNum, 13) = "-", "ND", ws.Cells(rowNum, 13))
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).HorizontalAlignment = xlCenter
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).VerticalAlignment = xlCenter
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        If counter = analyteListLength - 1 Then
                            .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        End If
                    Next counter
                    currentSampleCellColumn = currentSampleCellColumn + 1
                End If
            Next ws
            If Not referenceFileAlreadyOpen Then ReferenceFile.Close SaveChanges:=False
        End If
    End With
    '===============  End of Report File Management  ===============
        
        
    '===============  Calibration Solution Worksheet  ==============
    With controlFile.Worksheets("Calibration Solutions")
        'Header Initialization
        With .PageSetup
            .CenterHeader = "&G"
            .CenterFooter = "&G"
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
            .RightHeader = "Lab ID: " & LabIDTextBox.Text & vbCrLf & vbCrLf & vbCrLf
        End With
        Dim StandardSolutionsArray As Variant
        StandardSolutionsArray = Array( _
                                    Array("EPA 1633 Standard Calibration Solutions - QE (ng/mL)"), _
                                    Array("Compound", "Stock Conc.", "CS1 (LOQ)", "CS2", "CS3", "CS4", "CS5", "CS6", "CS7", "CS8", "CS9", "CS10"), _
                                    Array("Perfluoroalkyl carboxylic acids"), Array("PFBA", 4000, 0.08, 0.16, 0.4, 0.8, 1.6, 4, 8, 16, 40, 80), Array("PFPeA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("PFHxA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFHpA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFOA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFNA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFDA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFUnA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFDoA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFTrDA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFTeDA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Perfluoroalkyl sulfonic acids"), Array("PFBS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFPeS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFHxS Linear", 811, 0.01622, 0.03244, 0.0811, 0.1622, 0.3244, 0.811, 1.622, 3.244, 8.11, 16.22), Array("PFHxS Branched", 189, 0.00378, 0.00756, 0.0189, 0.0378, 0.0756, 0.189, 0.378, 0.756, 1.89, 3.78), Array("PFHpS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFOS Linear", 788, 0.01576, 0.03152, 0.0788, 0.1576, 0.3152, 0.788, 1.576, 3.152, 7.88, 15.76), Array("PFOS Branched", 211, 0.00422, 0.00844, 0.0211, 0.0422, 0.0844, 0.211, 0.422, 0.844, 2.11, 4.22), Array("PFNS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFDS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("PFDoS", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Fluorotelomer sulfonic acids"), Array("4:2 FTS", 4000, 0.08, 0.16, 0.4, 0.8, 1.6, 4, 8, 16, 40, 80), Array("6:2 FTS", 4000, 0.08, 0.16, 0.4, 0.8, 1.6, 4, 8, 16, 40, 80), Array("8:2 FTS", 4000, 0.08, 0.16, 0.4, 0.8, 1.6, 4, 8, 16, 40, 80), _
                                    Array("Perfluorooctane sulfonamides"), Array("PFOSA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("NMeFOSA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("NEtFOSA", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Perfluorooctane sulfonamidoacetic acids"), Array("NMeFOSAA Linear", 760, 0.0152, 0.0304, 0.076, 0.152, 0.304, 0.76, 1.52, 3.04, 7.6, 15.2), Array("NMeFOSAA Branched", 240, 0.0048, 0.0096, 0.024, 0.048, 0.096, 0.24, 0.48, 0.96, 2.4, 4.8), Array("NEtFOSAA Linear", 775, 0.0155, 0.031, 0.0775, 0.155, 0.31, 0.775, 1.55, 3.1, 7.75, 15.5), Array("NEtFOSAA Branched", 225, 0.0045, 0.009, 0.0225, 0.045, 0.09, 0.225, 0.45, 0.9, 2.25, 4.5), _
                                    Array("Perfluorooctane sulfonamide ethanols"), Array("NMeFOSE", 10000, 0.2, 0.4, 1, 2, 4, 10, 20, 40, 100, 200), Array("NEtFOSE", 10000, 0.2, 0.4, 1, 2, 4, 10, 20, 40, 100, 200), _
                                    Array("Per- and polyfluoroether carboxylic acids"), Array("HFPO-DA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("ADONA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("NFDHA (3,6-OPFHpA)", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("FPePA", 20000, 0.4, 0.8, 2, 4, 8, 20, 40, 80, 200, 400), Array("PF4OPeA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("PF5OHxA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), _
                                    Array("Ether sulfonic acids"), Array("9Cl-PF3ONS", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("11Cl-PF3OUdS", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), Array("PFEESA", 2000, 0.04, 0.08, 0.2, 0.4, 0.8, 2, 4, 8, 20, 40), _
                                    Array("Fluorotelomer carboxylic Acids"), Array("FPrPA (3:3FTCA)", 4000, 0.08, 0.16, 0.4, 0.8, 1.6, 4, 8, 16, 40, 80), Array("FHpPA (7:3FTCA)", 20000, 0.4, 0.8, 2, 4, 8, 20, 40, 80, 200, 400), _
                                    Array("Extracted Internal Standard (EIS) Analytes"), Array("M4PFBA", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("M5PFPeA", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("M5PFHxA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M4PFHpA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M8PFOA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M9PFNA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M6PFDA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M7PFUnA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M2PFDoA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M2PFTeDA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M3PFBS", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M3PFHxS", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M8PFOS", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M2-4:2FTS", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), _
                                    Array("M2-6:2FTS", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("M2-8:2FTS", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("M8FOSA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("d-N-MeFOSA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("d-N-EtFOSA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("d3-N-MeFOSAA", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("d5-N-EtFOSAA", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("d7-N-MeFOSE", 5000, 50, 50, 50, 50, 50, 50, 50, 50, 50, 50), Array("d9-N-EtFOSE", 5000, 50, 50, 50, 50, 50, 50, 50, 50, 50, 50), Array("M3HFPO-DA", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), _
                                    Array("Non-extracted Internal Standard (NIS) Analytes"), Array("M3PFBA", 1000, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10), Array("M2PFHxA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M4PFOA", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M5PFNA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("M2PFDA", 250, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5), Array("O2PFHxS", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5), Array("M4PFOS", 500, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5))
        For Rw = LBound(StandardSolutionsArray, 1) To UBound(StandardSolutionsArray, 1)
            For Col = LBound(StandardSolutionsArray(Rw), 1) To UBound(StandardSolutionsArray(Rw), 1)
                .Cells(Rw + 1, Col + 1).Value = StandardSolutionsArray(Rw)(Col)
                If (Rw = 0 Or Rw = 2 Or Rw = 14 Or Rw = 25 Or Rw = 29 Or Rw = 33 Or Rw = 38 Or Rw = 41 Or Rw = 48 Or Rw = 52 Or Rw = 55 Or Rw = 80) And Col = 0 Then
                    .Cells(Rw + 1, Col + 1).Font.Bold = True
                    .Range(.Cells(Rw + 1, LBound(StandardSolutionsArray(Rw + 1), 1) + 1), .Cells(Rw + 1, UBound(StandardSolutionsArray(Rw + 1), 1) + 1)).Merge Across:=False
                    .Range(.Cells(Rw + 1, LBound(StandardSolutionsArray(Rw + 1), 1) + 1), .Cells(Rw + 1, UBound(StandardSolutionsArray(Rw + 1), 1) + 1)).BorderAround LineStyle:=XlLineStyle.xlContinuous
                Else
                    If Col <> 0 And Rw <> 1 Then
                        .Cells(Rw + 1, Col + 1).HorizontalAlignment = xlCenter
                        .Cells(Rw + 1, Col + 1).VerticalAlignment = xlCenter
                    End If
                End If
                If Rw = 1 Then
                    .Cells(Rw + 1, Col + 1).Font.Bold = True
                End If
                .Cells(Rw + 1, Col + 1).BorderAround LineStyle:=XlLineStyle.xlContinuous
            Next Col
        Next Rw
        .Columns("A:C").AutoFit
    End With
    '===========  End of Calibration Solution Worksheet  ===========
    
    
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
                .RightHeader = "Lab ID: " & LabIDTextBox.Text & vbCrLf & vbCrLf & vbCrLf
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
                    .Cells(2 + Rw, Col + 1).Value = LCSLCSDFile.Worksheets(1).Cells(analyteListStartingRow + Rw, Col + 1).Value
                    If Rw = 0 Then
                        .Cells(2 + Rw, Col + 1).Font.Bold = True
                        .Cells(2 + Rw, Col + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Columns(Col + 1).AutoFit
                    End If
                Next Col
            Next Rw
        End With
        If Not LCSLCSDFileAlreadyOpen Then LCSLCSDFile.Close SaveChanges:=False
    End If
    '================ End of LCSLCSDFile Management ================
    
    
    '======================  Print Settings  =======================
    For Each ws In ActiveWorkbook.Worksheets
        With ws.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperLetter
            .Zoom = False
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .LeftMargin = Application.InchesToPoints(0.3)
            .RightMargin = Application.InchesToPoints(0.3)
            .TopMargin = Application.InchesToPoints(0.3)
            .BottomMargin = Application.InchesToPoints(0.3)
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            If ws.Name = "Cover Page" Then
                .CenterVertically = True
                .CenterHorizontally = True
                .HeaderMargin = Application.InchesToPoints(0.3)
                .FooterMargin = Application.InchesToPoints(0.3)
            ElseIf ws.Name = "LCSLCSD " + sheetName + " Report" Or ws.Name = "Glossary" Or ws.Name = "Calibration Solutions" Then
                .ScaleWithDocHeaderFooter = False
                .CenterHorizontally = True
                .TopMargin = Application.InchesToPoints(2.15)
                .BottomMargin = Application.InchesToPoints(2.15)
                .FitToPagesTall = False
                .PrintTitleRows = ws.Rows("1:2").Address
            Else
                .ScaleWithDocHeaderFooter = False
                .PrintTitleColumns = ws.Columns("$A:$C").Address
                .PrintTitleRows = ws.Rows(LastAnalytesPosition + 1).Address
                .Order = xlOverThenDown
                .TopMargin = Application.InchesToPoints(2.15)
                .BottomMargin = Application.InchesToPoints(2.15)
                .FitToPagesWide = False
            End If
        End With
    Next ws
    '===================  End of Print Settings  ===================
    
    'Reset Focus onto the first worksheet in the workbook
    controlFile.Worksheets(1).Activate
    
    'Allow the screen to update and unload the userform
    Application.ScreenUpdating = True
    Unload Me
    
End Sub


