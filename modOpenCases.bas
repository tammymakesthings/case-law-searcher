Attribute VB_Name = "modOpenCases"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub SearchKeywordHelp()
    LaunchUrl "http://www.findlaw.com/info/helpers/searchhelp.html"
End Sub

Public Sub LawDictionary(ByVal aWord As String)
    Dim strURL As String
    
    strURL = "http://dictionary.law.com/default2.asp?typed=#w&type=1&submit1.x=0&submit1.y=0&submit1=Look+up"
    strURL = Replace(strURL, "#w", URLEncode(aWord))
    LaunchUrl strURL
End Sub

Public Sub SearchByKeyword(ByVal strKeyword As String)
    Dim strURL As String
    
    strURL = "http://caselaw.lp.findlaw.com/scripts/callawft.pl?CiRestriction=#s"
    strURL = Replace(strURL, "#s", URLEncode(strKeyword))
    LaunchUrl strURL
End Sub

Public Sub SearchByPartyName(ByVal strPartyName As String)
    Dim strURL As String
    
    strURL = "http://caselaw.lp.findlaw.com/scripts/callawps.pl?pname=#s"
    strURL = Replace(strURL, "#s", URLEncode(strPartyName))
    LaunchUrl strURL
End Sub
Public Sub Shepardize(ByVal strSeries As String, ByVal strReporter As String, ByVal strPage As String)
    Dim strURL As String
    strURL = "http://caselaw.lp.findlaw.com/scripts/callawci2.pl?vol=#v&reporter=#r&page=#p"
    
    strURL = Replace(strURL, "#r", strReporter)
    strURL = Replace(strURL, "#v", strSeries)
    strURL = Replace(strURL, "#p", strPage)
    
    LaunchUrl strURL
End Sub
Public Sub OpenCase(ByVal strSeries As String, ByVal strReporter As String, ByVal strPage As String)

    Dim strURL As String
    strURL = GetBaseURL(strReporter)
    
    strURL = Replace(strURL, "#v", strSeries)
    strURL = Replace(strURL, "#p", strPage)
    
    LaunchUrl strURL
End Sub

Private Function GetBaseURL(strSelectedReporter As String) As String
    GetBaseURL = modGlobals.gSourcesList.GetSourceURL(strSelectedReporter)
End Function

Public Function URLEncode(ByVal aString As String) As String
    Dim i As Integer
    Dim acode As Integer
    Dim char As String
    
    URLEncode = aString
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
End Function

Private Sub LaunchUrl(ByVal strURL As String)
    ShellExecute 0, "open", strURL, vbNullString, vbNullString, 1
End Sub
