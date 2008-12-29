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

Private Sub WriteConfigFile()
    Dim objFSO As Scripting.FileSystemObject
    Dim objTS As Scripting.TextStream
    Dim varKey As Variant
    
    Set objFSO = New Scripting.FileSystemObject
    Set objTS = objFSO.OpenTextFile(App.Path & "\cls.cfg", ForWriting, True)
    For Each varKey In mdictConfig.Keys
        objTS.WriteLine varKey & ": " & mdictConfig(varKey)
    Next varKey
    objTS.Close
    Set objTS = Nothing
    Set objFSO = Nothing
End Sub

Private Sub ReadConfigFile()
    Dim objFSO As Scripting.FileSystemObject
    Dim objTS As Scripting.TextStream
    Dim strLine As String
    Dim varFields As Variant
    Dim strKey As String
    
    Set objFSO = New Scripting.FileSystemObject
    Set objTS = objFSO.OpenTextFile(App.Path & "\cls.cfg", ForReading)
    
    Do Until objTS.AtEndOfStream
        strLine = objTS.ReadLine
        If InStr(1, strLine, ":") > 0 Then
            varFields = Split(strLine, ":", 2)
            mdictConfig.Add Trim(varFields(0)), Trim(varFields(1))
        End If
    Loop
    objTS.Close
    Set objTS = Nothing
    Set objFSO = Nothing
End Sub

Public Function GetConfigValue(ByVal strKey As String) As String
    If mdictConfig.Exists(strKey) = True Then
        GetConfigValue = CStr(mdictConfig(strKey))
    Else
        GetConfigValue = ""
    End If
End Function

Public Sub SetConfigValue(ByVal strKey As String, ByVal strValue As String)
    mdictConfig(strKey) = strValue
End Sub

Private Sub Class_Initialize()
    Set mdictConfig = New Scripting.Dictionary
    ReadConfigFile
End Sub

Private Sub Class_Terminate()
    WriteConfigFile
    Set mdictConfig = Nothing
End Sub

