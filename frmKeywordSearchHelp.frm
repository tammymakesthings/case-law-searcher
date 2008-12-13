VERSION 5.00
Begin VB.Form frmKeywordSearchHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyword Search Help"
   ClientHeight    =   4548
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6504
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4548
   ScaleWidth      =   6504
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4800
      TabIndex        =   3
      Top             =   3840
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Full Guide On-Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2892
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Searches"
      Height          =   2052
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   6252
      Begin VB.Label Label15 
         Caption         =   "a and (b and not  c)"
         Height          =   252
         Left            =   3000
         TabIndex        =   18
         Top             =   1680
         Width           =   2892
      End
      Begin VB.Label Label14 
         Caption         =   "a and (b or c)"
         Height          =   252
         Left            =   3000
         TabIndex        =   17
         Top             =   1440
         Width           =   2892
      End
      Begin VB.Label Label13 
         Caption         =   "Grouping of search terms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   2412
      End
      Begin VB.Label Label12 
         Caption         =   "fly**"
         Height          =   252
         Left            =   3000
         TabIndex        =   15
         Top             =   1080
         Width           =   2892
      End
      Begin VB.Label Label11 
         Caption         =   "Words from same stem word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2412
      End
      Begin VB.Label Label10 
         Caption         =   "comput*"
         Height          =   252
         Left            =   3000
         TabIndex        =   13
         Top             =   720
         Width           =   2892
      End
      Begin VB.Label Label9 
         Caption         =   "Words with same prefix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label Label8 
         Caption         =   "toxic near torts"
         Height          =   252
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   2892
      End
      Begin VB.Label Label7 
         Caption         =   "Proximity search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2412
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Basics"
      Height          =   1452
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6252
      Begin VB.Label Label6 
         Caption         =   "chicago and not economics"
         Height          =   252
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   2892
      End
      Begin VB.Label Label5 
         Caption         =   "NOT search (exclude words)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2412
      End
      Begin VB.Label Label4 
         Caption         =   "chicago or economics"
         Height          =   252
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   2892
      End
      Begin VB.Label Label3 
         Caption         =   "OR search (any keyword)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2412
      End
      Begin VB.Label Label2 
         Caption         =   "chicago and economics"
         Height          =   252
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   2892
      End
      Begin VB.Label Label1 
         Caption         =   "AND search (all keywords)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2412
      End
   End
End
Attribute VB_Name = "frmKeywordSearchHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    modOpenCases.SearchKeywordHelp
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
