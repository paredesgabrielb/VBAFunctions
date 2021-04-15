Function ContainsSpecialCharacters(str As String) As Boolean
    For I = 1 To Len(str)
        ch = Mid(str, I, 1)
        Select Case ch
            Case "0" To "9", "A" To "Z", "a" To "z", " "
                ContainsSpecialCharacters = False
            Case Else
                ContainsSpecialCharacters = True
                Exit For
        End Select
    Next
End Function
Function CleanSpecialCharacters(str As String)
    Dim nstr As String
    nstr = str
    For I = 1 To Len(str)
        ch = Mid(str, I, 1)
        Select Case ch
            Case "0" To "9", "A" To "Z", "a" To "z", " "
                'Do nothing
            Case Else
                nstr = Replace(nstr, ch, "")
        End Select
    Next
    CleanSpecialCharacters = nstr
End Function

Function CountSpecialCharacters(str As String) As Integer
    Dim count As Integer
    count = 0
    For I = 1 To Len(str)
        ch = Mid(str, I, 1)
        Select Case ch
            Case "0" To "9", "A" To "Z", "a" To "z", " ", " ", ".", ",", "-", "'", ")", "("
                'Do nothing
            Case Else
                count = count + 1
        End Select
    Next
    CountSpecialCharacters = count
End Function

Function GetSpecialCharacters(str As String)
    Dim nstr As String
    nstr = str
    For I = 1 To Len(str)
        ch = Mid(str, I, 1)
        Select Case ch
            Case "0" To "9", "A" To "Z", "a" To "z", " "
                nstr = Replace(nstr, ch, "")
            Case Else
                'Do nothing
        End Select
    Next
    GetSpecialCharacters = nstr
End Function

