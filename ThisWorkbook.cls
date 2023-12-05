Sub Generate537Report(ByRef control As Office.IRibbonControl)
    UserForm537.Show vbModeless
End Sub

Sub Generate1633Report(ByRef control As Office.IRibbonControl)
    UserForm1633.Show vbModeless
End Sub

Sub GenerateCECReport(ByRef control As Office.IRibbonControl)
    UserFormCEC.Show vbModeless
End Sub

Sub GenerateMESAReport(ByRef control As Office.IRibbonControl)
    UserFormMESA.Show vbModeless
End Sub

Sub PrintWorkbook(ByRef control As Office.IRibbonControl)
    ActiveWorkbook.PrintOut Preview:=True
End Sub

Sub Userform1633Test()
    UserForm1633Revamp.Show vbModeless
End Sub


