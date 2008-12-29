VERSION 5.00
Begin VB.Form frmPreferences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Case Law Searcher Preferences"
   ClientHeight    =   3720
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   492
      Left            =   3360
      TabIndex        =   4
      Top             =   3120
      Width           =   1332
   End
   Begin VB.CommandButton cmdSavePrefs 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1332
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1092
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4572
      Begin VB.CheckBox chkConfirmBeforeExit 
         Caption         =   "C&onfirm before application exit"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4212
      End
      Begin VB.CheckBox chkSuppressLiveUpdate 
         Caption         =   "&Disable Live Updates to Sources file."
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4332
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update Server"
      Height          =   1692
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4572
      Begin VB.CommandButton cmdForceLiveUpdate 
         Caption         =   "&Force Live Update Now"
         Height          =   372
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   3732
      End
      Begin VB.TextBox txtUpdatePath 
         Height          =   288
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   3732
      End
      Begin VB.TextBox txtUpdateHost 
         Height          =   288
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   3732
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Path:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   738
         Width           =   492
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Host:"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   378
         Width           =   492
      End
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForceLiveUpdate_Click()
    modLiveUpdate.FetchLatestSourcesList
    MsgBox "The sources file was updated." & vbCrLf & vbCrLf & _
        "You must exit and restart the program to activate the latest sources list.", vbOKOnly + vbApplicationModal + vbInformation
End Sub

Private Sub cmdSavePrefs_Click()

    gAppConfig.SetConfigValue "liveupdate.host", txtUpdateHost.Text
    gAppConfig.SetConfigValue "liveupdate.path", txtUpdatePath.Text
    gAppConfig.SetConfigValue "liveupdate.enabled", IIf(chkSuppressLiveUpdate.Value > 0, "no", "yes")
    gAppConfig.SetConfigValue "core.confirmquit", IIf(chkConfirmBeforeExit.Value > 0, "yes", "no")
    
    gblnConfirmQuit = IIf(chkConfirmBeforeExit.Value > 0, True, False)
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txtUpdateHost.Text = gAppConfig.GetConfigValue("liveupdate.host")
    txtUpdatePath.Text = gAppConfig.GetConfigValue("liveupdate.path")
    chkConfirmBeforeExit.Value = IIf(gAppConfig.GetConfigValue("core.confirmquit") = "yes", 1, 0)
    chkSuppressLiveUpdate.Value = IIf(gAppConfig.GetConfigValue("liveupdate.enabled") = "yes", 0, 1)
    
End Sub
