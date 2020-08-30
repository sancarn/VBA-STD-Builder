VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents FileWatcher As xlsFileWatcher
Attribute FileWatcher.VB_VarHelpID = -1



Sub StartTimer()
    Set FileWatcher = xlsFileWatcher
    FileWatcher.fwStartWatching
End Sub
Sub StopTimer()
    FileWatcher.fwStopWatching
End Sub



Private Sub FileWatcher_FileInit(ByVal sFileName As String)
    Debug.Print "FileWatcher_FileInit: " & sFileName
End Sub

Private Sub FileWatcher_FileChanged(ByVal sFileName As String)
    Debug.Print "FileWatcher_FileChanged: " & sFileName
End Sub

Private Sub FileWatcher_FileCreated(ByVal sFileName As String)
    Debug.Print "FileWatcher_FileCreated: " & sFileName
End Sub

Private Sub FileWatcher_FileRemoved(ByVal sFileName As String)
    Debug.Print "FileWatcher_FileRemoved: " & sFileName
End Sub
