VERSION 5.00
Begin VB.Form frmPopupCaselaw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Caselaw..."
   ClientHeight    =   588
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6228
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   588
   ScaleWidth      =   6228
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   1092
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
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
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go!"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmPopupCaselaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mtypFormMode As ePopupFormMode
Public Property Let CitationFormMode(ByVal e As ePopupFormMode)
    mtypFormMode = e
    Select Case e
        Case ePopupFormMode.eFindCaselaw
            Me.Caption = "CaseLawSearcher - Find Case"
        Case ePopupFormMode.eShepardize
            Me.Caption = "CaseLawSearcher - Shepardize(ish) Case"
    End Select
End Property

Public Property Get CitationFormMode() As ePopupFormMode
    CitationFormMode = mtypFormMode
End Property

Private Sub MaybeDoSearch()
    If Len(txtSeries.Text) > 0 And Len(txtPage.Text) > 0 And Len(cboReporter.Text) > 0 Then
        Select Case mtypFormMode
            Case eFindCaselaw
                modOpenCases.OpenCase txtSeries.Text, cboReporter.Text, txtPage.Text
            Case eShepardize
                modOpenCases.Shepardize txtSeries.Text, cboReporter.Text, txtPage.Text
        End Select
        Unload Me
    Else
        MsgBox "A full citation is required."
    End If
End Sub

Private Sub Command1_Click()
    MaybeDoSearch
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim varFoo As Variant
    Dim varIter As Variant
    varFoo = modGlobals.gAppConfig.GetSourceList
    cboReporter.Clear
    For Each varIter In varFoo
        cboReporter.AddItem varIter
    Next varIter
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
