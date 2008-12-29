VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4236
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4236
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Loading...when this screen disappears, Case Law Searcher will be available in your system tray."
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   6852
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Searcher"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   756
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   1320
         Width           =   2772
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright (c) 2008, Tammy Cravit <tammy@tammycravit.us>. All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   2520
         TabIndex        =   1
         Top             =   2760
         Width           =   4452
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2520
         TabIndex        =   2
         Top             =   2280
         Width           =   888
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Case Law"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   756
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   600
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Dim strMsg As String
    
    If App.PrevInstance = True Then
        strMsg = "Another instance of " & App.ProductName & " is already running." & vbCrLf & vbCrLf & _
            "You can only run one instance at a time." & vbCrLf & _
            "Please click on the icon in the system tray to access " & App.ProductName
        MsgBox strMsg, vbOKOnly + vbExclamation + vbSystemModal, "Previous Instance Detected"
        Unload Me
        End
    End If
    
    ' Load the app config file
    Set modGlobals.gAppConfig = New CConfigFile
    
    ' Run Live Update
    modLiveUpdate.FetchLatestSourcesList
    
    ' Load the sources file
    Set modGlobals.gSourcesList = New CSourcesFile
    
    frmMain.Show
    Unload Me

End Sub
