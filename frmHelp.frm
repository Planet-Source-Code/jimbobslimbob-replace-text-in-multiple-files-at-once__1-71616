VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2550
      TabIndex        =   1
      Top             =   4500
      Width           =   1815
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6915
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload frmHelp
End Sub

Private Sub Form_Load()
    Dim FileNum As Integer
    'Load text
    On Error GoTo Error
    Screen.MousePointer = 11
    FileNum = FreeFile
    Open ProgPath & "Readme.txt" For Input As #FileNum
        If Err Then Close #FileNum: GoTo Error
        txtHelp.Text = ""
        txtHelp.Text = Input(LOF(1), 1)
    Close #FileNum
    Screen.MousePointer = 0
    Exit Sub

Error:
    MsgBox "Error loading readme file : " & Chr$(10) & Chr$(10) & Err.Description, vbCritical, "Readme file error"
    Screen.MousePointer = 0
    Close #FileNum

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtHelp.Width = frmHelp.Width - 100
    txtHelp.Height = frmHelp.Height - 1100
    cmdClose.Top = frmHelp.Height - 1000
    cmdClose.Left = frmHelp.Width / 2 - 1000
End Sub
