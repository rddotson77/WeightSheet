VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open() 'this code fires when the workbook is opened

    Application.DisplayAlerts = False 'turn off alerts

        If ActiveWorkbook.Name <> "Kettles & Line 10.xltm" Xor ActiveWorkbook.Name <> "Kettles & Line 101" Then 'if the workbook name is the name of the template (aka it is NOT an already saved weekly sheet), then execute following code

            Dim Week  As String 'create variable for the string "Week", used for week number
            Week = Sheet10.Range("h1").Value 'this finds the current week of the year from the calculation in cell H1 of sheet 10.

            If Dir("Y:\Scale Weights\Kettles and Line 10\Line 6 & 10 Week " & Week & ".xlsm") <> "" Then
                Application.Quit 'if the weekly sheet already exists, quit Excel and skip all following code.
                Exit Sub 'AKA CURRENT SHEET WILL NOT BE OVERWRITTEN.
            Else
                ActiveWorkbook.SaveAs Filename:="Y:\Scale Weights\Kettles and Line 10\Line 6 & 10 Week " & Week, _
                FileFormat:=xlOpenXMLWorkbookMacroEnabled, Password:="", WriteResPassword:="", _
                ReadOnlyRecommended:=False, CreateBackup:=False

                '^ if weekly sheet does not exist, save in the specified path with the name "Week __" where the blank is the week number of the year. save as macro workbook, no passwords.
            End If

            Application.DisplayAlerts = True 'reenable alerts.

            Application.OnTime Now + TimeValue("00:30:00"), "AutoSave" 'time value- hours:minutes:seconds 'start autosave code

        Else 'otherwise, it is not the template (it's the already created weekly sheet). Start autosave code.

            Application.OnTime Now + TimeValue("00:30:00"), "AutoSave" 'if not the template, start the autosave code and do nothing else.

        End If
End Sub
