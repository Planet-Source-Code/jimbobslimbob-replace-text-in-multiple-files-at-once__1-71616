VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple File Text Replacer"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   135
      TabIndex        =   11
      Top             =   6570
      Width           =   8295
      Begin VB.CheckBox chkFind 
         Caption         =   "Find (not replace) only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   18
         Top             =   405
         Width           =   2265
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2925
         Picture         =   "frmMain.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   135
      TabIndex        =   6
      Top             =   3330
      Width           =   8295
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear all text"
         Height          =   375
         Left            =   3330
         TabIndex        =   15
         Top             =   2655
         Width           =   1635
      End
      Begin VB.CheckBox chkCaseSensitive 
         Caption         =   "Case Sensitive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   14
         Top             =   2700
         Width           =   1950
      End
      Begin VB.TextBox txtReplace 
         Height          =   2040
         Left            =   4230
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   540
         Width           =   3750
      End
      Begin VB.TextBox txtFind 
         Height          =   2040
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   540
         Width           =   3750
      End
      Begin VB.Label Label2 
         Caption         =   "Replace With:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4230
         TabIndex        =   10
         Top             =   315
         Width           =   3660
      End
      Begin VB.Label Label1 
         Caption         =   "Find:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   9
         Top             =   315
         Width           =   3660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files to search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   8295
      Begin VB.CheckBox chkSelected 
         Caption         =   "Process Selected Files Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   17
         ToolTipText     =   "Unselecting this will process every file in the file list, regardless of whether it has been selected."
         Top             =   2610
         Width           =   2715
      End
      Begin VB.FileListBox filFolder 
         Height          =   2235
         Left            =   4635
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   360
         Width           =   3390
      End
      Begin VB.PictureBox picProgress 
         Height          =   330
         Left            =   4770
         ScaleHeight     =   270
         ScaleWidth      =   3195
         TabIndex        =   16
         Top             =   1620
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtFilePattern 
         Height          =   285
         Left            =   4680
         TabIndex        =   4
         Text            =   "*.txt;*.doc;*.htm;*.asp;*.html;*.xml"
         Top             =   2655
         Width           =   3345
      End
      Begin VB.DirListBox dirFolder 
         Height          =   1890
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   4470
      End
      Begin VB.DriveListBox drvFolder 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   4470
      End
      Begin VB.Label lblProcessing 
         Alignment       =   2  'Center
         Caption         =   "Processing. Please wait..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4815
         TabIndex        =   13
         Top             =   1125
         Width           =   3165
      End
      Begin VB.Label lblFileFilter 
         Caption         =   "File Filter:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3600
         TabIndex        =   5
         Top             =   2655
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ReplaceAll(ByRef theFileList As FileListBox, CaseSensitive As Integer, SelectedFiles As Integer, FindTextOnly As Integer)
    On Error Resume Next
    Dim X As Long
    Dim fileNum As Integer
    Dim theText As String
    Dim alteredText As String

    theFileList.Visible = False
    
    fileChangeCount = 0
    fileCount = 0
    fileList = vbNewLine
    fileNum = FreeFile
    DoEvents
    Debug.Print vbNewLine
    For X = 0 To theFileList.ListCount - 1
        theFileList.ListIndex = X
        Debug.Print "ITEM " & X & "  SELECTED=" & theFileList.Selected(X) & " (" & theFileList.List(X) & ")"
        If SelectedFiles = 1 And theFileList.Selected(X) = False Then GoTo SkipIT
            'Open file
            Open ProperPath(theFileList.Path) & theFileList.List(X) For Input As #fileNum
                theText = Input(LOF(fileNum), #fileNum)
            Close #fileNum
            'Replace text
            If CaseSensitive = 0 Then
                alteredText = Replace(theText, txtFind.Text, txtReplace.Text, , , vbTextCompare)
            Else
                alteredText = Replace(theText, txtFind.Text, txtReplace.Text, , , vbBinaryCompare)
            End If
            'Write file or just find files?
            If FindTextOnly = 0 Then
                Open ProperPath(theFileList.Path) & theFileList.List(X) For Output As #fileNum
                    Print #fileNum, alteredText
                Close #fileNum
                If theText <> alteredText Then
                    fileChangeCount = fileChangeCount + 1
                    fileList = fileList & theFileList.List(X) & vbNewLine
                    Debug.Print X & " (" & theFileList.List(X) & ") altered"
                End If
            Else
                If theText <> alteredText Then
                    fileChangeCount = fileChangeCount + 1
                    fileList = fileList & theFileList.List(X) & vbNewLine
                    Debug.Print X & " (" & theFileList.List(X) & ") found (not replaced)"
                End If
            End If
            fileCount = fileCount + 1
            DoEvents
            PercentBar picProgress, X, theFileList.ListCount - 1
            DoEvents
SkipIT:
    Next X
    theFileList.Visible = True
    Exit Sub
Error:
    theFileList.Visible = True
    Close #fileNum
    MsgBox "An error occurred trying to process the files: " & Err.Description, vbCritical
End Sub

Private Sub cmdClear_Click()
    On Error Resume Next
    txtFind.Text = ""
    txtReplace.Text = ""
End Sub

Private Sub cmdProcess_Click()
    On Error GoTo Error
    
    Dim ans As VbMsgBoxResult
    Dim txtStart, txtEnd, txtPos As Integer
    Dim theText As String
    
    'Error checking
    If txtFind.Text = "" Then
        MsgBox "Please enter text to search for!", vbExclamation
        Exit Sub
    End If
    If chkFind.Value = 0 Then
        ans = MsgBox("Are you sure you wish to perform this text replace? Please remember that it is NOT REVERSABLE. If you screw up some important files it's tough luck!" & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbExclamation + vbDefaultButton2, "Are you sure?")
        If ans = vbNo Then Exit Sub
    End If
    
'    If SelectedFiles = 1 And filFolder.Selected(X) = False Then
'        MsgBox "Please select one or more files to process" & vbCrLf & "OR" & vbCrLf & "Uncheck 'Replace Selected Files Only'", vbExclamation
'        Exit Sub
'    End If
    DoEvents
    
    cmdProcess.Enabled = False
    drvFolder.Enabled = False
    dirFolder.Enabled = False
    txtFilePattern.Enabled = False
    txtFind.Enabled = False
    txtReplace.Enabled = False
    
    picProgress.Visible = True

    Call ReplaceAll(filFolder, chkCaseSensitive.Value, chkSelected.Value, chkFind.Value)  'Process!
    
    If fileChangeCount > 0 Then
        'Highlight found files?
        ans = MsgBox("Do you want to highlight files found in the file list?", vbYesNo + vbQuestion + vbDefaultButton2, "Highlight?")
        If ans = vbYes Then
            txtPos = 1
            filFolder.Refresh
            While X < fileChangeCount
                txtStart = InStr(txtPos, fileList, vbNewLine) + 2
                txtEnd = InStr(txtPos + 1, fileList, vbNewLine)
                theText = Mid(fileList, txtStart, txtEnd - txtStart)
                txtPos = txtEnd
                filFolder.Selected(ListFind(filFolder, theText)) = True
                'Debug.Print theText & " " & ListFind(filFolder, theText)
                X = X + 1
            Wend
        End If
    End If
    
    cmdProcess.Enabled = True
    drvFolder.Enabled = True
    dirFolder.Enabled = True
    txtFilePattern.Enabled = True
    txtFind.Enabled = True
    txtReplace.Enabled = True
    
    picProgress.Visible = False
    
    If chkFind.Value = 0 Then
        MsgBox "Text replace complete!" & vbCrLf & vbCrLf & fileChangeCount & " of " & fileCount & " selected files were altered:" & fileList, vbInformation
    Else
        MsgBox "Text find complete!" & vbCrLf & vbCrLf & fileChangeCount & " of " & fileCount & " selected files were found:" & fileList, vbInformation
    End If
    Exit Sub
    
Error:
    cmdProcess.Enabled = True
    drvFolder.Enabled = True
    dirFolder.Enabled = True
    txtFilePattern.Enabled = True
    txtFind.Enabled = True
    txtReplace.Enabled = True
    
    picProgress.Visible = False
    MsgBox "An error occurred trying to process the files.", vbCritical
End Sub

Private Sub dirFolder_Change()
    On Error GoTo Error
    
    filFolder.Path = dirFolder.Path
    Exit Sub
    
Error:
    MsgBox "An error occurred trying to change the folder: " & Err.Description, vbCritical
    dirFolder.Path = filFolder.Path
End Sub

Private Sub drvFolder_Change()
    On Error GoTo Error
    
    dirFolder.Path = drvFolder.Drive
    Exit Sub
    
Error:
    MsgBox "An error occurred trying to change the drive: " & Err.Description, vbCritical
    drvFolder.Drive = dirFolder.Path
End Sub

Private Sub Form_Load()
    'Fix nasty file list display bug that puts a black border on it (Windows XP)
    'Before you ask; NO you can't just change the property at design time....does not work.
    filFolder.Appearance = 1
    txtFilePattern_Change
    frmMain.Caption = "Multiple File Text Replacer v" & ProgVer & " by James Compton 2009"
End Sub


Private Sub txtFilePattern_Change()
    On Error GoTo Error
    
    filFolder.Pattern = txtFilePattern.Text
    Exit Sub
    
Error:
    'MsgBox "An error occurred trying to set the file filter: " & Err.Description & vbCrLf & vbCrLf & "Please ensure you have entered a valid extension. eg. *.pdf or *.pdf;*.doc", vbCritical
End Sub
