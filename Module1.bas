Attribute VB_Name = "Module1"
Function AutoSave() 'method to autosave every specified amount of time. Continuously runs as long as the workbook is open.
    ThisWorkbook.Save 'save the workbook
Application.OnTime Now + TimeSerial(0, 30, 0), "AutoSave" 'time value- hours, minutes, seconds 'recur the AutoSave method every specified amount of time.
End Function


