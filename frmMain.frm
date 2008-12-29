VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Case Law Searcher"
   ClientHeight    =   4404
   ClientLeft      =   120
   ClientTop       =   804
   ClientWidth     =   6732
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4404
   ScaleWidth      =   6732
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame2 
      Caption         =   "Other Searches"
      Height          =   972
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   6492
      Begin VB.CommandButton cmdLawDictionary 
         Caption         =   "Define"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5520
         TabIndex        =   19
         Top             =   360
         Width           =   852
      End
      Begin VB.TextBox txtDefineWord 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Width           =   3372
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Define legal term:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1812
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Case Law Search"
      Height          =   2772
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6492
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   5760
         TabIndex        =   15
         Top             =   1920
         Width           =   612
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4800
         TabIndex        =   13
         Top             =   1920
         Width           =   852
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4800
         TabIndex        =   12
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox txtKeyword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   3252
      End
      Begin VB.TextBox txtPartyName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   3252
      End
      Begin VB.CommandButton cmdShepardize 
         Caption         =   "Shepardize(ish)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   1812
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   852
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Go!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   852
      End
      Begin VB.TextBox txtPage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   852
      End
      Begin VB.ComboBox cboReporter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1932
      End
      Begin VB.TextBox txtSeries 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Note: Party name and keyword/advanced searches only search California law."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   6252
      End
      Begin VB.Label Label3 
         Caption         =   "Keyword:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1092
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Party Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1332
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Case law searcher 0.01, Copyright (c) 2008, Tammy Cravit. All rights reserved."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   6492
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mitExit 
         Caption         =   "E&xit..."
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuUtils 
      Caption         =   "&Utilities"
      Begin VB.Menu mitBatchMode 
         Caption         =   "&Batch Mode"
         Shortcut        =   ^B
      End
      Begin VB.Menu mitUtilsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mitLiveUpdatePrefs 
         Caption         =   "Preferences..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mitAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mitHelpSearchKeyword 
         Caption         =   "&Keyword Search Help"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mitPopupFindCaselaw 
         Caption         =   "&Find Case Law..."
      End
      Begin VB.Menu mitPopupShepardize 
         Caption         =   "&Shepardize(ish) Case..."
      End
      Begin VB.Menu mitSysTraySep2 
         Caption         =   "-"
      End
      Begin VB.Menu mitPopUpSearchPartyName 
         Caption         =   "Search by &Party Name..."
      End
      Begin VB.Menu mitPopUpSearchKeyword 
         Caption         =   "Search by &Keyword..."
      End
      Begin VB.Menu mitPopUpBatchMode 
         Caption         =   "&Batch Mode..."
      End
      Begin VB.Menu mitDefineLegalTerm 
         Caption         =   "&Define Legal Term..."
      End
      Begin VB.Menu mitSysTraySep3 
         Caption         =   "-"
      End
      Begin VB.Menu mitPopupRestore 
         Caption         =   "&Restore Main Window"
      End
      Begin VB.Menu mitPopUpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mitPopupAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mitPopUpKeywordSearchHelp 
         Caption         =   "Keyword  Search Help..."
      End
      Begin VB.Menu mitPopUpSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mitPopupExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExitCount As Integer
Dim blnExplicitClose As Boolean

Private Sub cmdLawDictionary_Click()
    modOpenCases.LawDictionary txtDefineWord.Text
End Sub

Private Sub Command1_Click()
        MaybeDoSearch
End Sub

Private Sub Command2_Click()
    intExitCount = intExitCount + 1
    If intExitCount > 1 Then Exit Sub
    If gblnConfirmQuit = True Then
        If MsgBox("Are you sure you want to quit?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Exit Case Law Searcher?") = vbYes Then
            Unload Me
            End
        End If
    Else
        Unload Me
        End
    End If
End Sub

Private Sub Command3_Click()
    modOpenCases.SearchByPartyName txtPartyName.Text
End Sub

Private Sub Command4_Click()
    modOpenCases.SearchByKeyword txtKeyword.Text
End Sub

Private Sub Command5_Click()
    frmKeywordSearchHelp.Show
End Sub

Private Sub Form_Load()
    intExitCount = 0
    
    Dim varFoo As Variant
    Dim varIter As Variant
    varFoo = modGlobals.gSourcesList.GetSourceList
    cboReporter.Clear
    For Each varIter In varFoo
        cboReporter.AddItem varIter
    Next varIter
    
    'txtSeries.SetFocus
    
    Dim strFoo As String
    strFoo = "Case law searcher #m.#n.#r, Copyright (c) 2008, Tammy Cravit."
    strFoo = Replace(strFoo, "#m", App.Major)
    strFoo = Replace(strFoo, "#n", App.Minor)
    strFoo = Replace(strFoo, "#r", App.Revision)
    Label1.Caption = strFoo
    
    ' SysTray support
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Case Law Searcher" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub MaybeDoSearch()

    If Len(Trim(txtSeries.Text)) = 0 Then
        MsgBox "You must specify a full citation.", vbOKOnly + vbExclamation, "Missing Series"
        Exit Sub
    End If
    If Len(Trim(cboReporter.Text)) = 0 Then
        MsgBox "You must specify a full citation.", vbOKOnly + vbExclamation, "Missing Reporter"
        Exit Sub
    End If
    If Len(Trim(txtPage.Text)) = 0 Then
        MsgBox "You must specify a full citation.", vbOKOnly + vbExclamation, "Missing Page"
        Exit Sub
    End If

    modOpenCases.OpenCase txtSeries.Text, cboReporter.Text, txtPage.Text
    ResetForm
End Sub

Private Sub cmdShepardize_Click()
    modOpenCases.Shepardize txtSeries.Text, cboReporter.Text, txtPage.Text
    ResetForm
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    intExitCount = intExitCount + 1
    If intExitCount = 1 And blnExplicitClose = False And gblnConfirmQuit = True Then
        If MsgBox("Are you sure you want to quit?", vbYesNo + vbDefaultButton2 + vbQuestion + vbApplicationModal, "Exit Case Law Searcher?") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub


Private Sub mitAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mitBatchMode_Click()
    frmBatch.Show 1, Me
End Sub

Private Sub mitDefineLegalTerm_Click()
    Dim strKeyword As String
    strKeyword = InputBox("Enter the term you'd like to define:", "Define Legal Term", "")
    If Len(Trim(strKeyword)) = 0 Then
        MsgBox "No input provided. Search cancelled."
        Exit Sub
    End If
    modOpenCases.LawDictionary strKeyword
End Sub

Private Sub mitExit_Click()
    Call Command2_Click
End Sub

Private Sub mitGo_Click()
    Call Command1_Click
End Sub

Private Sub mitHelpSearchKeyword_Click()
    frmKeywordSearchHelp.Show
End Sub

Private Sub mitLiveUpdatePrefs_Click()
    frmPreferences.Show 1, Me
End Sub

Private Sub mitPopupAbout_Click()
    Dim X As frmAbout
    Set X = New frmAbout
    X.Show
End Sub

Private Sub mitPopUpBatchMode_Click()
    Dim X As frmBatch
    Set X = New frmBatch
    X.Show
End Sub

Private Sub mitPopupExit_Click()
    Call Command2_Click
End Sub

Private Sub mitPopupFindCaselaw_Click()
    Dim X As frmPopupCaselaw
    Set X = New frmPopupCaselaw
    X.CitationFormMode = eFindCaselaw
    X.Show
End Sub

Private Sub mitPopUpKeywordSearchHelp_Click()
    frmKeywordSearchHelp.Show
End Sub

Private Sub mitPopupRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub mitPopUpSearchKeyword_Click()
    Dim strKeyword As String
    strKeyword = InputBox("Enter the keyword you'd like to search for:", "Search by Keyword", "")
    If Len(Trim(strKeyword)) = 0 Then
        MsgBox "No input provided. Search cancelled."
        Exit Sub
    End If
    modOpenCases.SearchByKeyword strKeyword
End Sub

Private Sub mitPopUpSearchPartyName_Click()
        Dim strKeyword As String
    strKeyword = InputBox("Enter the party name you'd like to search for:", "Search by Party Name", "")
    If Len(Trim(strKeyword)) = 0 Then
        MsgBox "No input provided. Search cancelled."
        Exit Sub
    End If
    modOpenCases.SearchByPartyName strKeyword
End Sub

Private Sub mitPopupShepardize_Click()
    Dim X As frmPopupCaselaw
    Set X = New frmPopupCaselaw
    X.CitationFormMode = eShepardize
    X.Show
End Sub

Private Sub mitShepardize_Click()
    Call cmdShepardize_Click
End Sub

Private Sub ResetForm()
    txtPage.Text = ""
    txtSeries.Text = ""
    txtSeries.SetFocus
End Sub


Private Sub txtPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        MaybeDoSearch
    End If
End Sub

Private Sub txtSeries_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        If Len(txtSeries.Text) > 0 And Len(txtPage.Text) > 0 And Len(cboReporter.Text) > 0 Then
            MaybeDoSearch
        Else
            cboReporter.SetFocus
        End If
    End If
End Sub

Private Sub cboReporter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        If Len(txtSeries.Text) > 0 And Len(txtPage.Text) > 0 And Len(cboReporter.Text) > 0 Then
            MaybeDoSearch
        Else
            txtPage.SetFocus
        End If
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' System Tray support
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'this procedure receives the callbacks from the System Tray icon.
   Dim Result As Long
   Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
     msg = X
    Else
     msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
     Case WM_LBUTTONUP        '514 restore form window
      Me.WindowState = vbNormal
      Result = SetForegroundWindow(Me.hwnd)
      Me.Show
     Case WM_LBUTTONDBLCLK    '515 restore form window
      Me.WindowState = vbNormal
      Result = SetForegroundWindow(Me.hwnd)
      Me.Show
     Case WM_RBUTTONUP        '517 display popup menu
      Result = SetForegroundWindow(Me.hwnd)
      Me.PopupMenu Me.mPopupSys
    End Select
   End Sub

   Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
   End Sub

   Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
    
    ' Force the config file to be written out to disk
    Set gAppConfig = Nothing
   End Sub
