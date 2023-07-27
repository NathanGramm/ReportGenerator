VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Generate537Report(ByRef control As Office.IRibbonControl)
    UserForm537.Show vbModeless
End Sub

Sub Generate1633Report(ByRef control As Office.IRibbonControl)
    UserForm1633.Show vbModeless
End Sub

Sub GenerateCECReport(ByRef control As Office.IRibbonControl)
    UserFormCEC.Show vbModeless
End Sub

Sub PrintWorkbook(ByRef control As Office.IRibbonControl)
    ActiveWorkbook.PrintOut Preview:=True
End Sub


