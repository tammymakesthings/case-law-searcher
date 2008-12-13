Attribute VB_Name = "modLiveUpdate"
Dim objHTTP As MabryHttpClient.HttpXCom
Dim mstrLiveUpdateHost As String
Dim mstrLiveUpdatePath As String

Public Sub FetchLatestSourcesList()
    Dim strResult As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objTS As Scripting.TextStream
    Dim strLine As String
    Dim avarLineParts As Variant
    
    Set objFSO = New Scripting.FileSystemObject
    strResult = ""
    
    ' Defaults for Live Update
    mstrLiveUpdateHost = "www.tammycravit.us"
    mstrLiveUpdatePath = "/caselawsrch/Sources.txt"
    If objFSO.FileExists(App.Path & "\lu.cfg") Then
        Set objTS = objFSO.OpenTextFile(App.Path & "\lu.cfg", ForReading, False)
        Do Until objTS.AtEndOfStream
            strLine = Trim(objTS.ReadLine)
            avarLineParts = Split(strLine, ":", 2)
            If LCase(Trim(avarLineParts(0))) = "host" Then
                mstrLiveUpdateHost = Trim(avarLineParts(1))
            ElseIf LCase(Trim(avarLineParts(0))) = "path" Then
                mstrLiveUpdatePath = Trim(avarLineParts(1))
            End If
        Loop
        objTS.Close
        Set objTS = Nothing
    End If
    
    Set objHTTP = New MabryHttpClient.HttpXCom
    With objHTTP
        .LicenseKey = "XXXXXXXX" ' License key removed for anti-piracy
								 ' You'll need your own key to compile this.
        .Host = mstrLiveUpdateHost
        .BlockingMode = HttpXTrueBlocking
        .Connect
        .Get mstrLiveUpdatePath
        
        If .Response.Status = "HTTP/1.1 200 OK" Then
            strResult = .Response.Body
        End If
        .Disconnect
    End With
    
    If Len(strResult) > 0 Then
        
        Set objTS = objFSO.OpenTextFile(App.Path & "\Sources.cfg", ForWriting, True)
        objTS.WriteLine (strResult)
        objTS.Close
        Set objTS = Nothing
        Set objFSO = Nothing
    End If
    
    Set objHTTP = Nothing
    Set objFSO = New Scripting.FileSystemObject
End Sub
