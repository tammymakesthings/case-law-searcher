VERSION 5.00
Begin VB.Form frmBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Case Law Searcher - Batch Mode"
   ClientHeight    =   4320
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5892
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5892
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3840
      Width           =   1812
   End
   Begin VB.ListBox lstCitations 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3048
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   5652
   End
   Begin VB.CommandButton cmdAddCite 
      Caption         =   "&Add"
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
      TabIndex        =   3
      Top             =   120
      Width           =   852
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
      Left            =   1200
      TabIndex        =   0
      Top             =   96
      Width           =   612
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   108
      Width           =   1932
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
      Height          =   420
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   3840
      Width           =   1572
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
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
      TabIndex        =   4
      Top             =   3840
      Width           =   1812
   End
   Begin VB.Label Label2 
      Caption         =   "Add Cite:"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExitCount As Integer

Private Sub cmdAddCite_Click()
    MaybeAddCite
End Sub

Private Sub MaybeAddCite()
    Dim strBuff As String
    
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
    
    strBuff = txtSeries.Text & " " & cboReporter.Text & " " & txtPage.Text & vbCrLf
    
    lstCitations.AddItem Left(strBuff, Len(strBuff) - 2) ' Strip the vbCrLf
    
    txtPage.Text = ""
    txtSeries.Text = ""
    txtSeries.SetFocus
    
    If lstCitations.ListCount > 0 Then
        cmdRemove.Visible = True
    Else
        cmdRemove.Visible = False
    End If
    
End Sub
Private Sub cmdClose_Click()
    intExitCount = intExitCount + 1
    If intExitCount = 1 Then
        If MsgBox("Exit batch mode?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdGo_Click()
    Dim avarCites As Variant
    Dim varLine As Variant
    Dim avarLineParts As Variant
    Dim intLoop As Integer
    
    If lstCitations.ListCount > 15 Then
        If MsgBox("Opening many citations may take a long time. Are you sure?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirm Opening of Many Cites") = vbNo Then
            Exit Sub
        End If
    End If
            
    For intLoop = 0 To lstCitations.ListCount - 1
        varLine = lstCitations.List(intLoop)
        If Len(Trim(varLine)) > 0 Then
            avarLineParts = Split(varLine, " ")
            DoEvents
            modOpenCases.OpenCase avarLineParts(0), avarLineParts(1), avarLineParts(2)
            DoEvents
        End If
    Next intLoop
    lstCitations.Clear
    txtSeries.SetFocus
End Sub

Private Sub cmdRemove_Click()

    lstCitations.RemoveItem lstCitations.ListIndex
    
    If lstCitations.ListCount > 0 Then
        cmdRemove.Visible = True
    Else
        cmdRemove.Visible = False
    End If

End Sub

Private Sub Form_Load()
    Dim varFoo As Variant
    Dim varIter As Variant
    varFoo = modGlobals.gSourcesList.GetSourceList
    cboReporter.Clear
    For Each varIter In varFoo
        cboReporter.AddItem varIter
    Next varIter
    lstCitations.Clear
    cmdRemove.Visible = False
    
    intExitCount = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    intExitCount = intExitCount + 1
    If intExitCount = 1 Then
        If MsgBox("Exit batch mode?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Or KeyAscii = 13 Then
        MaybeAddCite
    End If
End Sub

Private Sub txtSeries_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Or KeyAscii = 13 Then
        If Len(txtSeries.Text) > 0 And Len(txtPage.Text) > 0 And Len(cboReporter.Text) > 0 Then
            MaybeAddCite
        Else
            cboReporter.SetFocus
        End If
    End If
End Sub
Private Sub cboReporter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Or KeyAscii = 13 Then
        If Len(txtSeries.Text) > 0 And Len(txtPage.Text) > 0 And Len(cboReporter.Text) > 0 Then
            MaybeAddCite
        Else
            txtPage.SetFocus
        End If
    End If
End Sub

