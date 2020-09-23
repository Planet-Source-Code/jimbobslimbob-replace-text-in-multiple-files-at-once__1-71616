Attribute VB_Name = "modTextReplacer"
Global fileChangeCount As Long
Global fileList As String
Global fileCount As Long
Public Function ListFind(ListBox As Control, sSearchText As String) As Integer

  Dim i As Integer

    'Loop through the contents of the listbo
    '     x
    ListFind = -1

    For i = 0 To ListBox.ListCount - 1
        Dim sCurrentItem As String
        'Place the current item in a variable
        sCurrentItem = ListBox.List(i)
        'Compare the current item with the strin
        '     g specified

        If LCase(sCurrentItem) = LCase(sSearchText) Then
            'If we find it remove it
            ListFind = i
        End If

    Next i

End Function
