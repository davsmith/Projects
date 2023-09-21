Attribute VB_Name = "Tests"
Option Explicit

Sub testRenameSheet()
    RenameSheet "Sheet9", "DisplaceMe", True
End Sub

Sub testPurgeTempSheets()
    PurgeTempSheets ("Sheet")
End Sub

Sub testSetPerformanceEnvironment()
    SetPerformanceEnvironment
End Sub
