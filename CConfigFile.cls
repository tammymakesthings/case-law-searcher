VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mdictConfig As Scripting.Dictionary

Private Sub ReadConfigFile()
    Dim objFSO As Scripting.FileSystemObject
    Dim objTS As Scripting.TextStream
    Dim strLine As String
    Dim varFields As Variant
    
    Set objFSO = New Scripting.FileSystemObject
    Set objTS = objFSO.OpenTextFile(App.Path & "\Sources.cfg", ForReading)
    
    Do Until objTS.AtEndOfStream
        strLine = objTS.ReadLine
        If InStr(1, strLine, "|") > 0 Then
            varFields = Split(strLine, "|", 2)
            mdictConfig.Add varFields(0), varFields(1)
        End If
    Loop
    objTS.Close
    Set objTS = Nothing
    Set objFSO = Nothing
End Sub

Public Function GetSourceList() As Variant
    Dim varIter As Variant
    Dim avarKeys() As Variant
    Dim blnRedimmed As Boolean
    
    blnRedimmed = False
    For Each varIter In mdictConfig.Keys
        If blnRedimmed = False Then
            ReDim avarKeys(0 To 0) As Variant
            blnRedimmed = True
        Else
            ReDim Preserve avarKeys(LBound(avarKeys) To UBound(avarKeys) + 1) As Variant
        End If
        avarKeys(UBound(avarKeys)) = varIter
    Next varIter
    GetSourceList = avarKeys
    
End Function

Public Function GetSourceURL(ByVal strSourceName As String) As String
    If mdictConfig.Exists(strSourceName) = True Then
        GetSourceURL = mdictConfig.Item(strSourceName)
    Else
        GetSourceURL = ""
    End If
End Function

Private Sub Class_Initialize()
    Set mdictConfig = New Scripting.Dictionary
    ReadConfigFile
End Sub

Private Sub Class_Terminate()
    Set mdictConfig = Nothing
End Sub
