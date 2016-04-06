VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Code for Worksheet "Cells as Checkboxes"
 
Option Explicit
 
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ActiveSheet.Unprotect "TBRADY"
     'Limit Target count to 1
    If Target.Count > 1 Then
        ActiveSheet.Protect "TBRADY"
        Exit Sub
    End If
     'Isolate Target to a specific range
    If Intersect(Target, Range("Bulk")) Is Nothing Then
        ActiveSheet.Protect "TBRADY"
        Exit Sub
    End If
     'set Target font tp "marlett"
   Target.Font.Name = "marlett"
     'Check value of target
    If Target.Value <> "a" Then
        Target.Value = "a" 'Sets target Value = "a"
        Cancel = True
        ActiveSheet.Protect "TBRADY"
        Exit Sub
    End If
    If Target.Value = "a" Then
        Target.ClearContents 'Sets Target Value = ""
        Cancel = True
        ActiveSheet.Protect "TBRADY"
        Exit Sub
    End If
    ActiveSheet.Protect "TBRADY"
End Sub

 
