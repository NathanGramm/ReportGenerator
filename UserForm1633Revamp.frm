Option Explicit

'Public Variables
Dim i       As Integer '<-- An Enumerator
Dim FileNumber As Integer '<-- Used as the File IO Number for Compound List .txt file
Dim Path As String '<-- The Path of the Add-In to find Compound List
Dim AddingCompound As Boolean '<-- Fixes a phenomenon that when editing a listbox item it calls the click event for the listbox

'====================================================
'UserForm Handling
'====================================================

'/**
'*Initialization of the userform
'*/
Private Sub Userform_Initialize()
    'Compound List Initialization
    Path = AddIns("Report Generator").Path & "\CompoundList.txt"
    With Me.CompoundListBox
        .ColumnCount = 3
        .ColumnWidths = "140;90;30"
    End With
    Call ReadCompoundList
    
    'Setting Sidebar to default minimized settings
    Me.SideBarImage.Width = 48
    Me.NewReportInfoLabel.Width = 0
    Me.CoverPageSetupLabel.Width = 0
    Me.GlossaryPageSetupLabel.Width = 0
    Me.ReportPageSetupLabel.Width = 0
    Me.FinalizeAndPrintLabel.Width = 0
    
    'Sets the Title of the page
    Me.TitleLabel.Caption = "Create New 1633 Report"
    
    'Makes sure AddingCompound starts at false
    AddingCompound = False
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call WriteCompoundList
End Sub

'====================================================
'Sidebar Section
'====================================================
Private Sub SideBarImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
End Sub

Private Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub CompoundNameTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub CompoundAbbreviationTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub CompoundLORTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub TitleUnderline_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub TitleLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub CompoundListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MinimizeSideBar
    Call Highlight(Me.PennStateIcon)
End Sub

Private Sub NewReportInfoTab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.NewReportInfoHighlightImage)
End Sub

Private Sub NewReportInfoImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.NewReportInfoHighlightImage)
End Sub

Private Sub CoverPageSetupTab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.CoverPageSetupHighlightImage)
End Sub

Private Sub CoverPageSetupImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.CoverPageSetupHighlightImage)
End Sub

Private Sub GlossaryPageSetupTab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.GlossaryPageSetupHighlightImage)
End Sub

Private Sub GlossaryPageSetupImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.GlossaryPageSetupHighlightImage)
End Sub

Private Sub ReportPageSetupTab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.ReportPageSetupHighlightImage)
End Sub

Private Sub ReportPageSetupImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.ReportPageSetupHighlightImage)
End Sub

Private Sub FinalizeAndPrintTab_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.FinalizeAndPrintHighlightImage)
End Sub

Private Sub FinalizeAndPrintImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExpandSideBar
    Call Highlight(Me.FinalizeAndPrintHighlightImage)
End Sub

Private Sub NewReportInfoHighlight_Click()
    Call NewReportInfoPageEnabled(True)
End Sub

Private Sub NewReportInfoHighlightImage_Click()
    Call NewReportInfoPageEnabled(True)
End Sub

Private Sub NewReportInfoLabel_Click()
    Call NewReportInfoPageEnabled(True)
End Sub


Private Sub NewReportInfoPageEnabled(Enabled As Boolean)
    
End Sub

'/**
'*Plays an animation in steps of the sidebar and all of its elements expanding
'*/
Private Sub ExpandSideBar()
    If Me.SideBarImage.Width = 48 Then
        For i = 48 To 180 Step 1
            DoEvents
            Me.SideBarImage.Width = i
            Me.NewReportInfoLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.CoverPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.GlossaryPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.ReportPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.FinalizeAndPrintLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.NewReportInfoHighlight.Width = i
            Me.NewReportInfoTab.Width = i
            Me.CoverPageSetupHighlight.Width = i
            Me.CoverPageSetupTab.Width = i
            Me.GlossaryPageSetupHighlight.Width = i
            Me.GlossaryPageSetupTab.Width = i
            Me.ReportPageSetupHighlight.Width = i
            Me.ReportPageSetupTab.Width = i
            Me.FinalizeAndPrintHighlight.Width = i
            Me.FinalizeAndPrintTab.Width = i
        Next i
    End If
    If Me.SideBarImage.Width > 180 Then Me.SideBarImage.Width = 180
End Sub

'/**
'*Plays an animation in steps of the sidebar and all of its elements minimizing
'*/
Private Sub MinimizeSideBar()
    If Me.SideBarImage.Width = 180 Then
        For i = 180 To 48 Step -1
            DoEvents
            Me.SideBarImage.Width = i
            Me.NewReportInfoLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.CoverPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.GlossaryPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.ReportPageSetupLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.FinalizeAndPrintLabel.Width = IIf(i - 48 < 0, 0, i - 48)
            Me.NewReportInfoHighlight.Width = i
            Me.NewReportInfoTab.Width = i
            Me.CoverPageSetupHighlight.Width = i
            Me.CoverPageSetupTab.Width = i
            Me.GlossaryPageSetupHighlight.Width = i
            Me.GlossaryPageSetupTab.Width = i
            Me.ReportPageSetupHighlight.Width = i
            Me.ReportPageSetupTab.Width = i
            Me.FinalizeAndPrintHighlight.Width = i
            Me.FinalizeAndPrintTab.Width = i
        Next i
    End If
    If Me.SideBarImage.Width < 48 Then Me.SideBarImage.Width = 48
End Sub

Private Sub Highlight(ByVal HighlightImage As Image)
    Dim NewReport, CoverPage, Glossary, ReportPage, Finalize, BackColor As Long
    NewReport = 12632319
    CoverPage = 8421631
    Glossary = 255
    ReportPage = 192
    Finalize = 128
    BackColor = HighlightImage.BorderColor
    Me.NewReportInfoImage.Visible = Not (BackColor = NewReport)
    Me.NewReportInfoTab.Visible = Not (BackColor = NewReport)
    Me.CoverPageSetupImage.Visible = Not (BackColor = CoverPage)
    Me.CoverPageSetupTab.Visible = Not (BackColor = CoverPage)
    Me.GlossaryPageSetupImage.Visible = Not (BackColor = Glossary)
    Me.GlossaryPageSetupTab.Visible = Not (BackColor = Glossary)
    Me.ReportPageSetupImage.Visible = Not (BackColor = ReportPage)
    Me.ReportPageSetupTab.Visible = Not (BackColor = ReportPage)
    Me.FinalizeAndPrintImage.Visible = Not (BackColor = Finalize)
    Me.FinalizeAndPrintTab.Visible = Not (BackColor = Finalize)
End Sub

'====================================================
'End of Sidebar Section
'====================================================



'====================================================
'Report Page Setup
'====================================================

'/**
'*Adds a compound to the listbox of compounds
'*/
Private Sub AddCompound_Click()
    AddingCompound = True
    Call AddCompoundToListBox
    AddingCompound = False
End Sub

'/**
'*Checks the listbox for the current compound data in
'*  the text boxes and edits the item if it exists,
'*  otherwise adds it to the end of the listbox
'*/
Private Sub AddCompoundToListBox()
    If CompoundNameTextBox.Text = "" Or CompoundAbbreviationTextBox.Text = "" Or CompoundLORTextBox.Text = "" Then Exit Sub
    Dim lRow, lCol As Long
    With Me.CompoundListBox
        For lRow = 0 To .ListCount - 1
            If .List(lRow, 0) = CompoundNameTextBox.Text Or .List(lRow, 1) = CompoundAbbreviationTextBox.Text Then
                .List(lRow, 0) = CompoundNameTextBox.Text
                .List(lRow, 1) = CompoundAbbreviationTextBox.Text
                .List(lRow, 2) = CompoundLORTextBox.Text
                MsgBox ("Updated Existing Compound: " & CompoundNameTextBox.Text)
                Exit Sub
            End If
        Next lRow
    End With
    With Me.CompoundListBox
        .AddItem
        .List(.ListCount - 1, 0) = CompoundNameTextBox.Text
        .List(.ListCount - 1, 1) = CompoundAbbreviationTextBox.Text
        .List(.ListCount - 1, 2) = CompoundLORTextBox.Text
    End With
End Sub

'/**
'*Sets the textboxes next to the compound list to
'*   the selected item
'*/
Private Sub CompoundListBox_Click()
    'Check to see if the list is empty, the selected item is the current index, and that we are not adding a compound
    If Me.CompoundListBox.ListIndex > -1 And Me.CompoundListBox.Selected(Me.CompoundListBox.ListIndex) And Not AddingCompound Then
        Me.CompoundNameTextBox.Text = Me.CompoundListBox.List(Me.CompoundListBox.ListIndex, 0)
        Me.CompoundAbbreviationTextBox.Text = Me.CompoundListBox.List(Me.CompoundListBox.ListIndex, 1)
        Me.CompoundLORTextBox.Text = Me.CompoundListBox.List(Me.CompoundListBox.ListIndex, 2)
    End If
End Sub

'/**
'*Reads the compoundlist.txt and inputs the data
'*  into the listbox
'*/
Private Sub ReadCompoundList()
    Dim DataLine As String
    Dim temp() As String
    FileNumber = FreeFile
    Open Path For Input As FileNumber
        While Not EOF(FileNumber)
            Line Input #FileNumber, DataLine
                temp = Split(DataLine, " | ")
                If UBound(temp) = 2 Then
                    CompoundNameTextBox.Text = temp(0)
                    CompoundAbbreviationTextBox.Text = temp(1)
                    CompoundLORTextBox.Text = temp(2)
                    AddCompoundToListBox
                End If
        Wend
    Close FileNumber
    Me.CompoundNameTextBox.Text = ""
    Me.CompoundAbbreviationTextBox.Text = ""
    Me.CompoundLORTextBox.Text = ""
End Sub

'/**
'*Writes the current state of the listbox to the
'*  compoundlist.txt
'*/
Private Sub WriteCompoundList()
    Dim compound As String
    Dim lRow, lCol As Long
    FileNumber = FreeFile
    Open Path For Output As FileNumber
        With Me.CompoundListBox
            For lRow = 0 To .ListCount - 1
                compound = .List(lRow, 0) & " | " & .List(lRow, 1) & " | " & .List(lRow, 2)
                Print #FileNumber, compound
            Next lRow
        End With
        Close FileNumber
End Sub

'/**
'*Deletes the currently selected item from the listbox
'*/
Private Sub DeleteCompoundFromListBox()
    If Me.CompoundListBox.ListIndex > -1 And Me.CompoundListBox.Selected(Me.CompoundListBox.ListIndex) Then
        Me.CompoundListBox.RemoveItem (Me.CompoundListBox.ListIndex)
    End If
End Sub

'====================================================
'End of Report Page Setup
'====================================================
