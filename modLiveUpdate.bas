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
    Dim blnSuppressUpdate As Boolean
    
    On Error GoTo err_flsl
    blnSuppressUpdate = False
    gblnConfirmQuit = True
    
    Set objFSO = New Scripting.FileSystemObject
    strResult = ""
    
    ' Defaults for Live Update
    mstrLiveUpdateHost = "www.tammycravit.us"
    mstrLiveUpdatePath = "/caselawsrch/Sources.txt"
    
    If Len(gAppConfig.GetConfigValue("liveupdate.host")) > 0 Then
        mstrLiveUpdateHost = gAppConfig.GetConfigValue("liveupdate.host")
    End If
    If Len(gAppConfig.GetConfigValue("liveupdate.path")) > 0 Then
        mstrLiveUpdatePath = gAppConfig.GetConfigValue("liveupdate.path")
    End If
    If gAppConfig.GetConfigValue("liveupdate.enabled") <> "yes" Then
        blnSuppressUpdate = True
    End If
    
    If blnSuppressUpdate = False Then
        Set objHTTP = New MabryHttpClient.HttpXCom
        With objHTTP
            .LicenseKey = "ENTER MABRY INTERNET CLIENT PACK SERIAL HERE"
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
    End If
    
exit_flsl:
    Set objHTTP = Nothing
    Set objFSO = New Scripting.FileSystemObject
    Exit Sub
    
err_flsl:
    Resume exit_flsl
End Sub
