VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlsProjectBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1
Public WithEvents FW As xlsFileWatcher
Attribute FW.VB_VarHelpID = -1
Private bRunning As Boolean

Friend Sub WatchReset()
    If Not bRunning Then
        Call DestroyAll
        Exit Sub
    End If
    Call Application.OnTime(Now() + TimeValue("00:00:01"), "xlsProjectBuilder.WatchReset", , True)
End Sub



Public Sub StartWatching()
    'Ensure trust in VBA Object model
    If VBEIsTrusted() Then
        'Pre-requesites exist, start building and testing
        
        'Get roots:
        Dim vRoots As Variant: vRoots = xlsProjectBuilder.ListObjects("Roots").DataBodyRange.Value
        
        'Start watching files
        Set FW = xlsFileWatcher
        Call FW.fwStartWatching(ThisWorkbook.Path & "\" & vRoots(1, 2), vRoots(1, 3))
        
        Dim i As Long
        For i = 2 To UBound(vRoots, 1)
            Call FW.fwAddRoot(ThisWorkbook.Path & "\" & vRoots(i, 2), vRoots(i, 3))
        Next
        
        'initialise workbook events
        Set wb = ThisWorkbook
        
        'Watch for VBA reset
        bRunning = True
        Call WatchReset
        
        
    Else
        MsgBox "Please enable trust in VBA Extensibility.", vbCritical
    End If
End Sub

Public Sub RunTests()
    'Get test sub name
    Dim sTestSubName As String
    sTestSubName = GetSetting("TestMain")
    
    'Run test sub
    If sTestSubName <> "" Then Call Application.Run(sTestSubName)
End Sub

Public Sub DestroyAll()
    'Remove all rows from physical table
    Dim data As Variant
    data = getProjectData()
    
    'Ensure not empty
    If Not IsEmpty(data) Then
        Dim iNameIndex As Long: iNameIndex = getProjectHeaderIndex("Name")
        
        Dim i As Long
        For i = UBound(data, 1) To 1 Step -1
            'Get VB Component name
            Dim sItemName As String: sItemName = data(i, iNameIndex)
            
            'Remove VB Component
            Call RemoveVBComponent(sItemName)
            
            'Remove list object row
            Me.ListObjects("VBProject").ListRows.Item(i).Delete
        Next
    End If
End Sub



Private Sub FW_FileInit(ByVal sFileName As String)
    'Add to physical table
    If isVBFile(sFileName) Then
        If getVBAttributeName(sFileName) <> "" Then
            Dim vbc As VBComponent
            Set vbc = ImportVBComponent(sFileName)
            If Not vbc Is Nothing Then
                With Me.ListObjects("VBProject").ListRows.Add()
                    With .Range
                        .Cells(1, getProjectHeaderIndex("File")).Value = sFileName
                        .Cells(1, getProjectHeaderIndex("Name")).Value = vbc.Name
                    End With
                End With
            Else
                Debug.Print "Error in Importing VB Component"
            End If
        Else
            Debug.Print "File created but no Attribute VB_Name " & sFileName
        End If
    End If
End Sub

Private Sub FW_FileChanged(ByVal sFileName As String)
    'Update physical table date
    If isVBFile(sFileName) Then
        'Try to remove VB Component
        Call RemoveVBComponent(getVBProjectNameFromPath(sFileName))
        
        'Now reimport the file
        Call ImportVBComponent(sFileName)
        
        'Run tests
        Call RunTests
    End If
End Sub

Private Sub FW_FileCreated(ByVal sFileName As String)
    'Add new row to physical table
    If isVBFile(sFileName) Then
        If getVBAttributeName(sFileName) <> "" Then
            Dim vbc As VBComponent
            Set vbc = ImportVBComponent(sFileName)
            If Not vbc Is Nothing Then
                With Me.ListObjects("VBProject").ListRows.Add()
                    With .Range
                        .Cells(1, getProjectHeaderIndex("File")).Value = sFileName
                        .Cells(1, getProjectHeaderIndex("Name")).Value = vbc.Name
                    End With
                End With
            Else
                Debug.Print "Error in Importing VB Component"
            End If
        Else
            Debug.Print "File created but no Attribute VB_Name " & sFileName
        End If
        
        'Run tests
        Call RunTests
    End If
End Sub

Private Sub FW_FileRemoved(ByVal sFileName As String)
    'Remove row from physical table
    If isVBFile(sFileName) Then
        'Try to remove VB Component
        Call RemoveVBComponent(getVBProjectNameFromPath(sFileName))
                
        'Delete row from list object
        Me.ListObjects("VBProject").ListRows.Item(getVBProjectRowFromPath(sFileName)).Delete
        
        'Run tests
        Call RunTests
    End If
End Sub
Private Sub FW_FileRenamed(ByVal sOldFilePath As String, ByVal sNewFilePath As String)
    'Not currently identified, instead we get around this by ensuring Remove() is called prior to Create()
End Sub

'Helper functions
Private Function getVBProjectRowFromPath(ByVal sFileName As String)
    Dim data As Variant
    data = getProjectData()
    'Find Class/Module/Form row
    Dim iFileIndex As Long: iFileIndex = getProjectHeaderIndex("File")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        'If item path is the path which was removed then, find name and remove component
        If data(i, iFileIndex) = sFileName Then
            getVBProjectRowFromPath = i
            Exit Function
        End If
    Next
End Function

Private Function getVBProjectNameFromPath(ByVal sFileName As String) As String
    Dim data As Variant
    data = getProjectData()
    
    'Find Class/Module/Form name
    Dim iFileIndex As Long: iFileIndex = getProjectHeaderIndex("File")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        'If item path is the path which was removed then, find name and remove component
        If data(i, iFileIndex) = sFileName Then
            'Get item name
            Dim sItemName As String
            getVBProjectNameFromPath = data(i, getProjectHeaderIndex("Name"))
            Exit Function
        End If
    Next
End Function

Private Function getProjectData() As Variant
    If Me.ListObjects("VBProject").ListRows.Count > 0 Then
        getProjectData = Me.ListObjects("VBProject").DataBodyRange.Value
    End If
End Function
Private Function getProjectHeaderIndex(ByVal sHeaderName As String) As Long
    On Error GoTo ErrorOccurred
    getProjectHeaderIndex = Me.ListObjects("VBProject").ListColumns(sHeaderName).Index
    Exit Function
ErrorOccurred:
    getProjectHeaderIndex = 0
End Function



Private Function GetSetting(ByVal sKey As String) As String
    Dim vData As Variant: vData = Me.ListObjects("Settings").DataBodyRange.Value
    Dim i As Long
    For i = 1 To UBound(vData, 1)
        If sKey = vData(i, 1) Then
            GetSetting = vData(i, 2)
            Exit Function
        End If
    Next
End Function


Private Function getVBAttributeName(ByVal sFileName As String) As String
    Dim ff As Long
    ff = FreeFile
    
    Dim textRow As String
    Open sFileName For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, textRow
        If textRow Like "Attribute VB_Name = *" Then
            getVBAttributeName = Mid(textRow, 22, Len(textRow) - 22)
            Exit Do
        End If
        
        textRow = UCase(textRow)
        If textRow Like "*FUNCTION*" Or textRow Like "*SUB*" Then
            Exit Do
        End If
    Loop
    Close #ff
End Function


Private Function getVBComponent(ByVal sName As String) As VBComponent
    On Error GoTo ErrorOccurred
    Set getVBComponent = ThisWorkbook.VBProject.VBComponents.Item(sName)
    Exit Function
ErrorOccurred:
    Set getVBComponent = Nothing
End Function
Private Sub RemoveVBComponent(ByVal sName As String)
    Set vbc = getVBComponent(sName)
    If Not vbc Is Nothing Then
        Call ThisWorkbook.VBProject.VBComponents.Remove(vbc)
    End If
End Sub
Private Function ImportVBComponent(ByVal sPath As String) As VBComponent
    On Error GoTo ErrorOccurred
    Set ImportVBComponent = ThisWorkbook.VBProject.VBComponents.Import(sPath)
    Exit Function
ErrorOccurred:
    Set ImportVBComponent = Nothing
End Function
Private Function isVBFile(ByVal sPath As String) As Boolean
    Select Case LCase(Right(sPath, 4))
        Case ".bas", ".cls", ".frm"
            isVBFile = True
        Case Else
            isVBFile = False
    End Select
End Function




Private Function VBEIsTrusted() As Boolean
    Dim a1 As Integer
    On Error GoTo Label1
    a1 = ActiveWorkbook.VBProject.VBComponents.Count
    VBEIsTrusted = True
    Exit Function
Label1:
    VBEIsTrusted = False
End Function

Private Sub wb_BeforeClose(Cancel As Boolean)
    Call DestroyAll
End Sub
