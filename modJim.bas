Attribute VB_Name = "modJim"
'------------------------------------------
'Jim's Commonly used Subroutine's Module
'------------------------------------------
Public Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" _
        Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long
Option Explicit

'Tray Icon stuff...
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
        Alias "Shell_NotifyIconA" ( _
        ByVal dwMessage As Long, _
        pnid As NOTIFYICONDATA) As Boolean

Public Enum sndConst
    SND_ASYNC = &H1 ' play asynchronously
    SND_LOOP = &H8 ' loop the sound until Next sndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points To a memory file
    SND_NODEFAULT = &H2 ' silence Not default, If sound not found
    SND_NOSTOP = &H10 ' don't stop any currently playing sound
    SND_SYNC = &H0 ' play synchronously (default), halts prog use till done playing
End Enum

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const HWND_TOPMOST = -1

Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Function FileExists(FileName As String) As Boolean

    'Returns true if a file exists
    On Error GoTo errHandler
    FileExists = False

    If Dir(FileName) <> "" Then
        If (GetAttr(FileName) And vbDirectory) = 0 Then
            FileExists = True
         Else
            FileExists = False
            Exit Function
        End If

     Else
        FileExists = False
        Exit Function
    End If

errHandler:

End Function

Public Function FindString(theString As String, SearchText As String) As Boolean

    'Find a string in a string
    Dim i As Long

    FindString = False

    For i = 1 To Len(theString)

        If StrComp((Mid(theString, i, Len(SearchText))), SearchText, vbTextCompare) = 0 Then
            FindString = True
        End If

    Next i

End Function

Function GetFilename(Path As String, Optional DontIncludeExtension As Boolean = False) As String

    On Error GoTo Error
    
    If DontIncludeExtension = True Then
        Dim tmp As String
        'Get filename from full path
        tmp = Mid(Path, InStrRev(Path, "\") + 1)
        GetFilename = Left(tmp, InStrRev(tmp, ".") - 1)
    Else
        'Get filename from full path
        GetFilename = Mid(Path, InStrRev(Path, "\") + 1)
    End If
    Exit Function

Error:
    GetFilename = ""

End Function

Function GetFileExtension(Path As String) As String

    On Error GoTo Error
    
    'Get filename extension from full path
    GetFileExtension = Mid(Path, InStrRev(Path, ".") + 1)
    Exit Function

Error:
    GetFileExtension = ""

End Function

Function GetPath(FullFilePath As String) As String

    On Error GoTo Error
    'Get path from full path and filename
    GetPath = Mid(FullFilePath, 1, InStrRev(FullFilePath, "\"))
    Exit Function

Error:
    GetPath = ""
    
End Function

Public Function INIRead(Section As String, _
                        Key As String, _
                        Optional Directory As String, _
                        Optional Default As String) As String

    'Read from an ini file
    On Error Resume Next
     Dim strBuffer As String

    strBuffer = String(750, Chr(0))
    Key$ = LCase$(Key$)
    
    If Directory <> "" Then
        INIRead$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, _
            Len(strBuffer), Directory$))
    
        If INIRead$ = "" Then
            INIRead$ = Default
            Call WritePrivateProfileString(Section$, UCase$(Key$), Default$, Directory$)
        End If
    Else
        INIRead$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, _
            Len(strBuffer), ProgPath & "Settings.ini"))
    
        If INIRead$ = "" Then
            INIRead$ = Default
            Call WritePrivateProfileString(Section$, UCase$(Key$), Default$, ProgPath & "Settings.ini")
        End If
    End If

End Function

Public Sub INIWrite(Section As String, Key As String, KeyValue As String, Optional Directory As String)

    'Write to an ini file
    On Error Resume Next
    If Directory <> "" Then
        Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
    Else
        Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, ProgPath & "Settings.ini")
    End If

End Sub

Public Sub ListLoad(Directory As String, TheList As ListBox)

    'Load a file into a listbox
    Dim MyString As String

    On Error GoTo Error
    Open Directory$ For Input As #1

    While Not EOF(1)
        Line Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend

    Close #1
    Exit Sub
Error:

End Sub

Public Sub ListSave(Directory As String, TheList As ListBox)

    'Save a listbox to a file
    Dim SaveList As Long

    On Error Resume Next
    Open Directory$ For Output As #1

    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&

    Close #1

End Sub

Public Sub LoadFormPosition(theForm As Form)

    Dim DefaultPosLeft As String
    Dim DefaultPosTop As String

    'Default position to load if no info in the ini file
    DefaultPosTop = (Screen.Height / 2) - theForm.Height / 2
    DefaultPosLeft = (Screen.Width / 2) - theForm.Width / 2

    'Load Window position settings
    theForm.Top = INIRead(theForm.Name, "Window Top", ProgPath & "Settings.ini", DefaultPosTop)
    theForm.Left = INIRead(theForm.Name, "Window Left", ProgPath & "Settings.ini", DefaultPosLeft)

    'Stop it dissappearing off screen
    If theForm.Left < 0 Then theForm.Left = 0
    If theForm.Left + theForm.Width > Screen.Width Then theForm.Left = Screen.Width - theForm.Width
    If theForm.Top < 0 Then theForm.Top = 0
    If theForm.Top + theForm.Height > Screen.Height Then theForm.Top = Screen.Height - _
        theForm.Height

End Sub

Sub PercentBar(Shape As Control, Done As Long, Total As Long)

    Dim X As Long

    'Makes a picture object into a percentage bar
    On Error Resume Next
    Shape.AutoRedraw = True
    Shape.FillStyle = 0
    Shape.DrawStyle = 0
    Shape.FontName = "MS Sans Serif"
    Shape.FontSize = 8.25
    Shape.FontBold = False
    X = Done / Total * Shape.Width
    Shape.Line (0, 0)-(Shape.Width, Shape.Height), vbButtonFace, BF
    Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 127), BF
    Shape.CurrentX = (Shape.Width / 2) - 100
    Shape.CurrentY = (Shape.Height / 2) - 125
    Shape.ForeColor = vbButtonFace

End Sub

Public Sub Playsound(SoundName As String)

    'Play a wav sound file from \Sounds
    On Error GoTo Error
    Dim X As Integer
    Dim Var As Long
    Dim SoundName2 As String

    DoEvents
    SoundName2$ = ProgPath & "Sounds\" & SoundName
    Var& = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundName2$, Var&)
    Exit Sub

Error:
    MsgBox "Error playing sound : " & SoundName & Chr$(10) & Chr$(10) & Err.Description, vbCritical, _
        "Error playing sound"

End Sub

Public Function ProgPath() As String

    'Returns the program's directory

    If Mid$(App.Path, Len(App.Path)) = "\" Then
        ProgPath$ = App.Path
     Else
        ProgPath$ = App.Path & "\"
    End If

End Function

Public Function ProgVer() As String

    'Returns the program's version
    ProgVer$ = App.Major & "." & App.Minor & App.Revision

End Function

Public Function ProperPath(thePath As String) As String

    'Returns a path with a "\" if it has not got it

    If Mid$(thePath, Len(thePath)) = "\" Then
        ProperPath$ = thePath
     Else
        ProperPath$ = thePath & "\"
    End If

End Function

Public Sub SaveFormPosition(theForm As Form)

    'Save window position settings
    If theForm.WindowState = 0 Then INIWrite theForm.Name, "Window Left", theForm.Left, ProgPath & "Settings.ini"
    If theForm.WindowState = 0 Then INIWrite theForm.Name, "Window Top", theForm.Top, ProgPath & "Settings.ini"

End Sub

Public Sub SnapForm(theForm As Form) ' If window is close to screen edge, snaps it to the edge

    Dim SnapDist As Integer

    SnapDist = 150 'distance the windows has to be from edge to snap on
    'Snap on
    If theForm.Left < SnapDist And theForm.Left > 0 Then theForm.Left = 0
    If theForm.Left > Screen.Width - (SnapDist + theForm.Width) And theForm.Left < (Screen.Width - _
        theForm.Width) Then theForm.Left = Screen.Width - theForm.Width
    If theForm.Top < SnapDist And theForm.Top > 0 Then theForm.Top = 0
    If theForm.Top > Screen.Height - (SnapDist + theForm.Height) And theForm.Top < (Screen.Height - _
        theForm.Height) Then theForm.Top = Screen.Height - theForm.Height

End Sub

Sub TextLoad(txtLoad As TextBox, Path As String)

    'Load text into a textbox
    Dim TextString As String

    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$

End Sub

Sub TextSave(txtSave As TextBox, Path As String)

    'Save text from textbox to a file
    Dim TextString As String

    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1

End Sub

Public Sub TrayCreate(theForm As Form, Optional Tooltip As String)

    Dim TooltipCaption As String

    If Tooltip = "" Then
        TooltipCaption = theForm.Caption
     Else
        TooltipCaption = Tooltip
    End If

    With nid
        .cbSize = Len(nid)
        .hwnd = theForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = theForm.Icon
        .szTip = TooltipCaption & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, nid

End Sub

Public Sub TrayRemove()

    Shell_NotifyIcon NIM_DELETE, nid

End Sub

Public Sub TrayTooltipChange(theForm As Form, Tooltip As String)

    With nid
        .cbSize = Len(nid)
        .hwnd = theForm.hwnd
        .uId = vbNull
        .szTip = Tooltip & vbNullChar
    End With

    Shell_NotifyIcon NIM_MODIFY, nid

End Sub
Public Sub TrayDetectButtons(theForm As Form, Button As Integer, X As Single, Optional theMenu As Menu)

    Dim lngMsg As Long
    Dim result As Long

    'Call this from the mousemove event of the form - send it Button and X (which are part of
    '   mousemove)
    lngMsg = X / Screen.TwipsPerPixelX

    Select Case lngMsg
     Case WM_RBUTTONUP ' right button
        SetForegroundWindow theForm.hwnd
        If Not theMenu Is Nothing Then theForm.PopupMenu theMenu

     Case WM_LBUTTONUP ' left button
        SetForegroundWindow theForm.hwnd
        'theForm.PopupMenu theMenu
    End Select

End Sub

