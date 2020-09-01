VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlsFileWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Why a sheet object?
'We want events - so we need a class or a object
'We want to use Application.OnTime therefore we can't use a class
'Therefore we have to use an Excel object.
'Sheet vs Workbook - Sheets are more maluable for multi-instance workers.

Public Event FileInit(ByVal sFileName As String)
Public Event FileCreated(ByVal sFileName As String)
Public Event FileRemoved(ByVal sFileName As String)
Public Event FileChanged(ByVal sFileName As String)

Public oFileManager As Object
Public cRoots As Collection  'allow for multiple roots
Public oInitialisedRoots As Object
Public bTrackDeletes As Boolean

Public cScheduledInit As Collection
Public cScheduledCreated As Collection
Public cScheduledRemoved As Collection
Public cScheduledChanged As Collection


Public Sub fwStartWatching(Optional ByVal root As String = "", Optional ByVal TrackDeletes As Boolean = True)
    'Initialise file manager and initialisation dictionary
    Set oFileManager = CreateObject("Scripting.Dictionary")
    Set oInitialisedRoots = CreateObject("Scripting.Dictionary")
    Set cRoots = New Collection
    
    'Add root and initialise root as false
    Dim sRoot As String: sRoot = IIf(root = "", ThisWorkbook.Path, root)
    cRoots.Add sRoot
    oInitialisedRoots(sRoot) = False
    bTrackDeletes = TrackDeletes
    
    'Schedule watch
    Application.OnTime Now() + TimeValue("00:00:01"), "xlsFileWatcher.Watch", Schedule:=True
End Sub
Public Sub fwStopWatching()
    'Deschedule watch
    Application.OnTime Now() + TimeValue("00:00:01"), "xlsFileWatcher.Watch", Schedule:=False
End Sub
Public Sub fwAddRoot(ByVal sRoot As String)
    'Schedule watch of new root
    cRoots.Add sRoot
    oInitialisedRoots(sRoot) = False
End Sub


Friend Sub Watch()
    'cRoots shouldn't ever be nothing unless a scheduled call was made when VBA has stopped, so just exit if it occurs
    'It's interesting this is called, in theory this means we can create a VBA reset detector
    If cRoots Is Nothing Then Exit Sub
    
    'Set scheduled events to new collections
    Set cScheduledInit = New Collection
    Set cScheduledCreated = New Collection
    Set cScheduledRemoved = New Collection
    Set cScheduledChanged = New Collection
    
    'If we are tracking deletes set all files exists status to false
    If bTrackDeletes Then
        Dim DeletedKeys As Variant
        Dim iDeletedKey As Long
        
        DeletedKeys = oFileManager.keys()
        For iDeletedKey = LBound(DeletedKeys) To UBound(DeletedKeys)
            oFileManager(DeletedKeys(iDeletedKey))("Exists") = False
        Next
    End If
    
    'Loop through roots and handle folders/files
    Dim vRoot As Variant
    For Each vRoot In cRoots
        Dim sRoot As String: sRoot = vRoot
        Dim bRootInitialised As Boolean: bRootInitialised = oInitialisedRoots(sRoot)
        
        If VBA.FileSystem.GetAttr(sRoot) = vbDirectory Then
            Call HandleFolder(bRootInitialised, sRoot)
        Else
            Call HandleFile(bRootInitialised, sRoot)
        End If
        
        'Initialise roots
        If Not bRootInitialised Then oInitialisedRoots(sRoot) = True
        
        DoEvents
    Next
    
    'If track deletes then look for files which have been deleted
    If bTrackDeletes Then
        'Get files to be checked
        DeletedKeys = oFileManager.keys()
        For iDeletedKey = LBound(DeletedKeys) To UBound(DeletedKeys)
            Dim sDeletedKey As String: sDeletedKey = DeletedKeys(iDeletedKey)
            
            'If exists is false, then file removed
            If Not oFileManager(sDeletedKey)("Exists") Then
                cScheduledRemoved.Add sDeletedKey
                Call oFileManager.Remove(sDeletedKey)
            End If
        Next
    End If
    
    ''Potential for determining file move events here
    'set cScheduledRenamed = DetectFileRenames(cScheduledCreated, cScheduledRemoved)
    
    'Dispatch all events in order: Init, Change, Remove, Create
    Call DispatchEvents
    
    'Loop watch
    Application.OnTime Now() + TimeValue("00:00:01"), "xlsFileWatcher.Watch", Schedule:=True
End Sub

Private Sub DispatchEvents()
    Dim vFile As Variant
    
    'Raise file initialise events
    For Each vFile In cScheduledInit
        RaiseEvent FileInit(vFile)
    Next
    
    'Raise file change events
    For Each vFile In cScheduledChanged
        RaiseEvent FileChanged(vFile)
    Next
    
    'Raise file deletion events
    For Each vFile In cScheduledRemoved
        RaiseEvent FileRemoved(vFile)
    Next
    
    'Raise file creation events
    For Each vFile In cScheduledCreated
        RaiseEvent FileCreated(vFile)
    Next
End Sub

'Watches the root specified
Private Sub HandleFolder(ByVal bRootInitialised As Boolean, ByVal sFolder As String)
    'Ensure folder has a "\" on the back
    sFolder = sFolder & IIf(Right(sFolder, 1) <> "\", "\", "")
    
    'Dir() has a limit of 1 (i.e. can't recurse... Therefore we must briefly store the folders for later handling)
    Dim cFolders As Collection
    Set cFolders = New Collection
    
    'Find all folders:
    Dim sItem As String
    sItem = Dir(sFolder, vbDirectory) 'Include directories
    Do Until sItem = ""
        If sItem <> "." And sItem <> ".." Then
            'Schedule for handling of directories, and handle files
            If GetAttr(sFolder & sItem) = vbDirectory Then
                cFolders.Add sFolder & sItem
            Else
                Call HandleFile(bRootInitialised, sFolder & sItem)
            End If
        End If
        DoEvents
        
        'Get next item
        sItem = Dir
    Loop
    
    'Loop over folders and handle them
    Dim vFolder As Variant
    For Each vFolder In cFolders
        Call HandleFolder(bRootInitialised, vFolder)
    Next
End Sub
Private Sub HandleFile(ByVal bRootInitialised As Boolean, ByVal sFile As String)
    'Get modified date
    Dim DateModified As Date: DateModified = FileDateTime(sFile)
    
    'If root is initialising, initialise file
    If Not bRootInitialised Then
        cScheduledInit.Add sFile
        Set oFileManager(sFile) = CreateFileObject(sFile, DateModified, True)
    Else
        'If file already exists, then not created
        If oFileManager.Exists(sFile) Then
            'If date modified now later than cached then
            If DateModified > oFileManager(sFile)("DateModified") Then
                cScheduledChanged.Add sFile
                oFileManager(sFile)("DateModified") = DateModified
                oFileManager(sFile)("Exists") = True
            Else
                oFileManager(sFile)("Exists") = True
            End If
        Else
            'New file created
            cScheduledCreated.Add sFile
            Set oFileManager(sFile) = CreateFileObject(sFile, DateModified, True)
        End If
    End If
End Sub

Private Function CreateFileObject(ByVal sFile As String, ByVal DateModified As Date, ByVal bExists As Boolean) As Object
    Dim ret As Object
    Set ret = CreateObject("Scripting.Dictionary")
    ret.CompareMode = 1
    ret("Exists") = bExists
    ret("Path") = sFile
    ret("DateModified") = DateModified
    Set CreateFileObject = ret
End Function

