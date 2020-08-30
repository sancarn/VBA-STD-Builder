VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public root As String
Public UpdatedOnDict As Object

Public Function Create(Optional ByVal sRoot As String = "") As clsCodeBuilder
    Set Create = New clsCodeBuilder
    Call Create.Init(IIf(sRoot = "", ThisWorkbook.Path, sRoot))
End Function
Public Sub Init(ByVal sRoot As String)
    
    root = sRoot
End Sub



Public Sub Rebuild()
    
End Sub
Public Sub Build()
    
End Sub
Public Sub Destroy()
    
End Sub



Private Function getVBProjectReference() As Boolean
    'ThisWorkbook.VBProject.References
End Function








'#If VBA7 Then
'    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#Else
'    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#End If

Sub CheckTrustAccess()
Dim strStatus, strOpp, strCheck As String
Dim bEnabled As Boolean
If Not VBAIsTrusted Then
'ask the user if they want me to try to programatically toggle trust access. If I fail, give them directions.
    bEnabled = False
    strStatus = "DISABLE"
    strOpp = "ENABLE"
    v = MsgBox("Trust Access to the VBA Project Object Model is " & strStatus & "D." & Chr(10) & Chr(10) & _
     "Would you like me to attempt to " & strOpp & " it?", vbYesNo, strOpp & " Trust Access?")
Else
    bEnabled = True
    strStatus = "ENABLE"
    strOpp = "DISABLE"
    v = MsgBox("Trust Access to the VBA Project Object Model is " & strStatus & "D." & Chr(10) & Chr(10) & _
     "Would you like me to attempt to " & strOpp & " it?", vbYesNo, strOpp & " Trust Access?")
End If

    If v = 6 Then
        'try to toggle trust
        Call ToggleTrust(bEnabled)
        If VBAIsTrusted = bEnabled Then
            'if ToggleTrust fails to programatically toggle trust
            MsgBox "I failed to " & strOpp & " Trust Access." & Chr(10) & Chr(10) & _
            "To " & strOpp & " this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Toggle the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "Auto " & strOpp & " Failed"
            End
        Else
            MsgBox "I successfully " & strOpp & "D Trust Access." & Chr(10) & Chr(10) & _
            "To " & strStatus & " this setting yourself:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Toggle the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "Auto " & strOpp & " Failed"
        End If
    Else
        MsgBox "To manually " & strOpp & " Trust Access:" & Chr(10) & Chr(10) & _
            Space(5) & "1) Click " & Chr(145) & "File-> Options-> Trust Center-> Trust Center Settings" & Chr(146) & Chr(10) & _
            Space(5) & "2) Click Macro Settings" & Chr(10) & _
            Space(5) & "3) Toggle the box next to ""Trust Access to the VBA project object model""", vbOKOnly, "How to " & strOpp & " Trust Access"
        End
    End If

If VBAIsTrusted Then
    'if you want to write your own macro, do it here. You only get here if access is trusted
End If

End Sub

Private Function VBAIsTrusted() As Boolean
Dim a1 As Integer
On Error GoTo Label1
a1 = ActiveWorkbook.VBProject.VBComponents.Count
VBAIsTrusted = True
Exit Function
Label1:
VBAIsTrusted = False
End Function

Private Sub ToggleTrust(bEnabled As Boolean)
Dim b1 As Integer, i As Integer
Dim strkeys As String
On Error Resume Next
    Do While i <= 2 'try to sendkeys 3 times
        'Sleep 100
        DoEvents
    strkeys = "%tms%v{ENTER}"
        Call SendKeys(Trim(strkeys)) 'ST%V{ENTER}")
        DoEvents
        If VBAIsTrusted <> bEnabled Then Exit Do 'successfully toggled trust
        DoEvents
        i = i + 1
    Loop
End Sub
