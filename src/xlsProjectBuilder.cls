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
        
        'Start watching files
        Set FW = xlsFileWatcher
        Call FW.fwStartWatching
        
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
    sTestSubName = getIniValue(ThisWorkbook & "\..\config.ini", "TestMain")
    
    'Run test sub
    On Error Resume Next
    Application.Run sTestSubName
End Sub

Public Sub DestroyAll()
    'Remove all rows from physical table
    Dim data As Variant
    data = getProjectData()
    
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
        'We have the file, so we can attempt to remove it
        Call RemoveVBComponent(getVBAttributeName(sFileName))
        
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
                sItemName = data(i, getProjectHeaderIndex("Name"))
                
                'Try to remove VB Component
                Call RemoveVBComponent(sItemName)
                
                'Delete row from list object
                Me.ListObjects("VBProject").ListRows.Item(i).Delete
                
                Exit For
            End If
        Next
        
        'Run tests
        Call RunTests
    End If
End Sub
Private Sub FW_FileRenamed(ByVal sOldFilePath As String, ByVal sNewFilePath As String)
    'Not currently identified, instead we get around this by ensuring Remove() is called prior to Create()
End Sub





Private Function getProjectData() As Variant
    getProjectData = Me.ListObjects("VBProject").DataBodyRange.Value
End Function
Private Function getProjectHeaderIndex(ByVal sHeaderName As String) As Long
    On Error GoTo ErrorOccurred
    getProjectHeaderIndex = Me.ListObjects("VBProject").ListColumns(sHeaderName).Index
    Exit Function
ErrorOccurred:
    getProjectHeaderIndex = 0
End Function



Private Function getIniValue(ByVal sIniFile As String, ByVal sIniKey As String) As String
    Dim ff As Long
    ff = FreeFile
    
    Dim textRow As String
    Open sIniFile For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, textRow
        If textRow Like sIniKey & "=*" Then
            getIniValue = Mid(sIniKey, Len(sIniKey) + 2)
            Exit Function
        End If
    Loop
    Close #ff
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
