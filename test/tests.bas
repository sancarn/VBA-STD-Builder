Attribute VB_Name = "tests"

Sub MainTestAll()
    Application.StatusBar = "Tested " & now()
    Application.OnTime Now() + TimeValue("00:00:02"), "FlushStatusBar", , True
End Sub
Sub FlushStatusBar()
    Application.StatusBar = false
End Sub