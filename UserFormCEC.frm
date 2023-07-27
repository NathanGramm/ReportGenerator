VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormCEC 
   Caption         =   "Report Generation"
   ClientHeight    =   6588
   ClientLeft      =   -252
   ClientTop       =   -1080
   ClientWidth     =   10188
   OleObjectBlob   =   "UserFormCEC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormCEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private referenceWorkbookName, LCSLCSDWorkbookName As String


Private Sub CommandButton2_Click()
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Excel", "*.xlsx"
      .Filters.Add "Excel, Macro-Enabled Workbook", "*.xlsm"
      .Filters.Add "Excel 2003", "*.xls"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show Then
        LCSLCSDWorkbookTextBox = .SelectedItems(1)
        LCSLCSDWorkbookName = .SelectedItems(1)
      End If
   End With
End Sub

Private Sub CommandButton1_Click()
    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Excel", "*.xlsx"
      .Filters.Add "Excel, Macro-Enabled Workbook", "*.xlsm"
      .Filters.Add "Excel 2003", "*.xls"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show Then
        ReferenceWorkbookTextBox = .SelectedItems(1)
        referenceWorkbookName = .SelectedItems(1)
      End If
   End With
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub ReferenceWorkbookTextBox_Change()
    TextBoxsFilled
End Sub

Private Sub TextBoxsFilled()
    GenerateReport.Enabled = Len(ReferenceWorkbookTextBox.Text) > 0
End Sub

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

Private Sub MatrixComboBox_Change()
    If MatrixComboBox.Value = "Water" Then
        SampleLabel.Caption = "Sample Volume (mL)"
        SampleTextBox.Enabled = True
        SampleLabel.Visible = True
        SampleTextBox.Visible = True
    ElseIf MatrixComboBox.Value = "POCIS" Then
        SampleLabel.Caption = "Sorbent Weight (g)"
        SampleTextBox.Enabled = True
        SampleLabel.Visible = True
        SampleTextBox.Visible = True
    ElseIf MatrixComboBox.Value = "Sediment" Or MatrixComboBox.Value = "Tissue" Then
        SampleLabel.Caption = "Sample Weight (g)"
        SampleTextBox.Enabled = True
        SampleLabel.Visible = True
        SampleTextBox.Visible = True
    Else
        SampleLabel.Visible = False
        SampleTextBox.Visible = False
    End If
End Sub

Private Sub GenerateReport_Click()
    'Author: Nathan Gramm
    'Start Date: 6/23/2023
    'Latest Edit Date: 7/14/2023
    'For the data analysis reports of The Pennsylvania State University's IEE Department
    
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
    Dim ws As Worksheet, foundEmptyWorksheet, deleteCoverPage, deleteDisclaimer, foundDuplicateReport As Boolean
    'Loop through all worksheets in the workbook and save whether files need to be replaced
'    Loop through all worksheets
'       If there is one named Cover Page then change bool
'       If there is one named Glossary then change bool
'       If there is one named SheetName + Report then change bool
'       If there is one named Calibration Solutions then change bool
'       If there is one named LCSLCSD Report then change bool
'       If there is an empty sheet save that index to an array
'    Finish Loop
'
'   If index array > 4 then delete any worksheet after 4
'
'   If Cover Page is found then delete if not the only one
'       If it is the only one then rename found sheet to NULL,
'       Check if first index of empty sheet array exists if so use that sheet
'           If not Add a new sheet before Null and name it coverpage
'       Delete Null

    For i = 1 To controlFile.Worksheets.Count
        Set ws = controlFile.Worksheets(i)
        If ws.Name = "Cover Page" Then
            deleteCoverPage = True
        End If
        If ws.Name = "Glossary" Then
            deleteDisclaimer = True
        End If
        If ws.Name = sheetName + " Report" Then
            foundDuplicateReport = True
        ElseIf ws.UsedRange.Address = "$A$1" And ws.Range("A1") = "" And ws.Name <> "Cover Page" And ws.Name <> "Glossary" Then
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
            Exit For
        End If
    Next i
    'If we found a sheet with the same name then handle it
    If foundDuplicateReport Then
        If controlFile.Worksheets.Count = 1 Then
            Dim tempSheet As Worksheet
            controlFile.Worksheets(sheetName + " Report").Name = "Null"
            Sheets.Add.Name = sheetName + " Report"
            Application.DisplayAlerts = False
            controlFile.Worksheets("Null").Delete
            Application.DisplayAlerts = True
            foundEmptyWorksheet = True
        Else
            Application.DisplayAlerts = False
            controlFile.Worksheets(sheetName + " Report").Delete
            Application.DisplayAlerts = True
        End If
    End If
    'If we never found a sheet that could be used as our report sheet make one
    If Not foundEmptyWorksheet Then
        Sheets.Add After:=controlFile.Worksheets(controlFile.Worksheets.Count)
        controlFile.Worksheets(controlFile.Worksheets.Count).Name = sheetName + " Report"
    End If
    'If we needed to remake the cover page, delete it then add a new one
    If deleteCoverPage Then
        Application.DisplayAlerts = False
        controlFile.Worksheets("Cover Page").Delete
        Application.DisplayAlerts = True
    End If
    If deleteDisclaimer Then
        Application.DisplayAlerts = False
        controlFile.Worksheets("Glossary").Delete
        Application.DisplayAlerts = True
    End If
    Sheets.Add(Before:=controlFile.Worksheets(1)).Name = "Cover Page"
    Sheets.Add(After:=controlFile.Worksheets("Cover Page")).Name = "Glossary"
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
    
    'Top Cover Page TextBox
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
        analyteList = Array("Acetaminophen", "Amoxicillin", "Ampicillin", "Atenolol", "Atrazine", "Benzophenone", "Benzophenone-3", "Caffeine", "Carbamazepine", "Carbaryl", "Chlorpyrifos", "Chlortetracycline", "Citalopram", "Clarithromycin", "Clothianidin", "Erythromycin", "Imidacloprid", "Ketoprofen", "Metformin", "Metoprolol", "Naproxen", "Ofloxacin", "Oxytetracycline", "Simazine", "Sulfadiazine", "Sulfadimethoxine", "Sulfamethazine", "Sulfamethoxazole", "Tetracycline", "Theobromine", "Thiacloprid", "Thiamethoxam", "Trimethoprim")
        RLList = Array(0.02, 0.1, 0.04, 1, 0.4, 1, 0.2, 1, 0.2, 0.04, 0.4, 0.4, 2, 2, 1, 2, 1, 1, 0.02, 0.4, 0.04, 2, 0.1, 0.4, 0.4, 1, 0.4, 0.4, 0.1, 2, 0.4, 1, 0.4)
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
            Dim ReferenceFile As Workbook
            Set ReferenceFile = Workbooks.Open(referenceWorkbookName)
        
            'Reset the analyteList to match the file output in the autogenerated data report
            analyteList = Array("Acetaminophen", "Amoxicillin", "Ampicillin", "Atenolol", "Atrazine", "Benzophenone", "Benzophenone 1", "Cafeine", "Carbamazepine", "Carbaryl", "Chlorpyrifos", "Chlortetracycline", "Citalopram", "Clarithromycin", "Clothianidin", "Erythromycine", "Imidacloprid", "Ketoprofen", "Metformin", "Metoprolol", "Naproxene", "Ofloxacin", "Oxytetracycline", "Simazine", "Sulfadiazine", "Sulfadimethoxine", "Sulfamethazine", "Sulfamethoxazole", "Tetracycline", "Theobromine", "Thiacloprid", "Thiamethoxam", "Trimethoprim")
            'Start the loop through the reference file sheets
            For i = 1 To ReferenceFile.Worksheets.Count
                analyteNameList = Array(analyteListLength)
                analyteConcentrationList = Array(analyteListLength)
                If InStr(1, ReferenceFile.Worksheets(i).Range("I9").Value, SearchParameterTextBox.Text) > 0 Then
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Borders(xlEdgeBottom).LineStyle = xlDouble
                    .Columns(currentSampleCellColumn).ColumnWidth = 16
                    .Cells(analyteListStartingRow - 1, currentSampleCellColumn).Value = ReferenceFile.Worksheets(i).Range("I9").Value
                    Dim rowNum As Long
                    For counter = 0 To analyteListLength - 1
                        rowNum = ReferenceFile.Worksheets(i).Columns(1).Find(What:=analyteList(counter), LookIn:=xlValues, LookAt:=xlWhole).Row
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Value = IIf(ReferenceFile.Worksheets(i).Cells(rowNum, 13) = "N/F" Or ReferenceFile.Worksheets(i).Cells(rowNum, 13) = "N/A", "ND", ReferenceFile.Worksheets(i).Cells(rowNum, 13))
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).HorizontalAlignment = xlCenter
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).VerticalAlignment = xlCenter
                        .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
                        If counter = analyteListLength - 1 Then
                            .Cells(analyteListStartingRow + counter, currentSampleCellColumn).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                        End If
                    Next counter
                    currentSampleCellColumn = currentSampleCellColumn + 1
                End If
            Next i
            ReferenceFile.Close SaveChanges:=False
        End If
        
    End With
    '===============  End of Report File Management  ===============
        
        
    '===============  Calibration Solution Worksheet  ==============
    Dim currentWS As Worksheet, AddedSheet As Boolean
    AddedSheet = False
    For Each currentWS In controlFile.Worksheets
        If currentWS.Name = "Calibration Solutions" Then
            currentWS.Name = "NULL"
            Sheets.Add After:=controlFile.Worksheets(sheetName + " Report")
            ActiveSheet.Name = "Calibration Solutions"
            Application.DisplayAlerts = False
            currentWS.Delete
            Application.DisplayAlerts = True
            AddedSheet = True
            ActiveSheet.Name = "Calibration Solutions"
            Exit For
        End If
    Next currentWS
    If Not AddedSheet Then
        Sheets.Add After:=controlFile.Worksheets(sheetName + " Report")
        ActiveSheet.Name = "Calibration Solutions"
    End If
              
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
                                    Array("CEC Quantification Standard Calibration Solutions - QE (ng/mL)"), _
                                    Array("Compound", "Stock Conc.", "CS1 (LOQ)", "CS2", "CS3", "CS4", "CS5", "CS6", "CS7", "CS8", "CS9", "CS10"), _
                                    Array("Herbicides"), Array("Atrazine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Simazine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Neonicotinoid insecticides"), Array("Clothianidin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Imidacloprid", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Thiacloprid", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Thiamethoxam", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Carbamate insecticides"), Array("Carbaryl", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Organophosphate pesticides"), Array("Chlorpyrifos", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Personal care products"), Array("Benzophenone", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Benzophenone-3", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Caffeine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Theobromine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Analgesic pharmaceuticals"), Array("Acetaminophen", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Ketoprofen", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Naproxen", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Aminopenicillin antibiotic pharmaceuticals"), Array("Amoxicillin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Ampicillin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Macrolide antibiotic pharmaceuticals"), Array("Clarithromycin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Erythromycin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Sulfonamide antibiotic pharmaceuticals"), Array("Sulfadiazine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Sulfadimethoxine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Sulfamethazine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Sulfamethoxazole", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Tetracycline antibiotic pharmaceuticals"), Array("Chlortetracycline", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Oxytetracycline", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Tetracycline", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Other antibiotic pharmaceuticals"), Array("Trimethoprim", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Ofloxacin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Anticonvulsant pharmaceuticals"), Array("Carbamazepine", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("SSRI antidepressant pharmaceuticals"), Array("Citalopram", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Beta blocker pharmaceuticals"), Array("Atenolol", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), Array("Metoprolol", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Biguanidine pharmaceuticals"), Array("Metformin", 1000, 0.02, 0.04, 0.1, 0.2, 0.4, 1, 2, 4, 10, 20), _
                                    Array("Extracted Internal Standard (EIS) Analytes"), Array("Acetamiprid-d3", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("Dinotefuran-d3", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("Nitenpyram-d3", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), _
                                    Array("Non-extracted Internal Standard (NIS) Analytes"), Array("Clothianidin-d3", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("Imidacloprid-d4", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("Thiacloprid-d4", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20), Array("Thiamethoxam-d3", 2000, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20))
        For Rw = LBound(StandardSolutionsArray, 1) To UBound(StandardSolutionsArray, 1)
            For Col = LBound(StandardSolutionsArray(Rw), 1) To UBound(StandardSolutionsArray(Rw), 1)
                .Cells(Rw + 1, Col + 1).Value = StandardSolutionsArray(Rw)(Col)
                If (Rw = 0 Or Rw = 2 Or Rw = 5 Or Rw = 10 Or Rw = 12 Or Rw = 14 Or Rw = 19 Or Rw = 23 Or Rw = 26 Or Rw = 29 Or Rw = 34 Or Rw = 38 Or Rw = 41 Or Rw = 43 Or Rw = 45 Or Rw = 48 Or Rw = 50 Or Rw = 54) And Col = 0 Then
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
        Dim LCSLCSDFile As Workbook
        Set LCSLCSDFile = Workbooks.Open(LCSLCSDWorkbookName)
        controlFile.Activate
        AddedSheet = False
        For Each currentWS In controlFile.Worksheets
            If currentWS.Name = LCSLCSDFile.Worksheets(1).Name Or currentWS.Name = "LCSLCSD Report" Then
                currentWS.Name = "NULL"
                Sheets.Add After:=controlFile.Worksheets(controlFile.Worksheets.Count)
                Application.DisplayAlerts = False
                currentWS.Delete
                Application.DisplayAlerts = True
                AddedSheet = True
                ActiveSheet.Name = "LCSLCSD Report"
                Exit For
            End If
        Next currentWS
        If Not AddedSheet Then
            Sheets.Add After:=controlFile.Worksheets(controlFile.Worksheets.Count)
            ActiveSheet.Name = "LCSLCSD Report"
        End If
    Else
        For Each currentWS In controlFile.Worksheets
            If currentWS.Name = "LCSLCSD Report" Then
                If controlFile.Worksheets.Count = 1 Then
                    Sheets.Add
                End If
                Application.DisplayAlerts = False
                currentWS.Delete
                Application.DisplayAlerts = True
            End If
        Next currentWS
    End If
    
    With controlFile.Worksheets("LCSLCSD Report")
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
                .Cells(2 + Rw, Col + 1).Value = IIf(LCSLCSDFile.Worksheets(1).Cells(analyteListStartingRow + Rw, Col + 1).Value = "Benzophenone 1", "Benzophenone-3", LCSLCSDFile.Worksheets(1).Cells(analyteListStartingRow + Rw, Col + 1).Value)
                If Rw = 0 Then
                    .Cells(2 + Rw, Col + 1).Font.Bold = True
                    .Cells(2 + Rw, Col + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Columns(Col + 1).AutoFit
                End If
            Next Col
        Next Rw
    End With
    LCSLCSDFile.Close SaveChanges:=False
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
            ElseIf ws.Name = "LCSLCSD Report" Or ws.Name = "Glossary" Or ws.Name = "Calibration Solutions" Then
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
