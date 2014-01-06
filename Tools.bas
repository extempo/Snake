'Snake v1.0, 2014-01-02
'By Matt Carleton

'Module "Tools" - useful little subs used in other parts of the codebase


'Pass it a column number and it will give you the corresponding letter(s) for it
Function ColumnLetter(ColumnNumber As Integer) As String
    Dim n As Integer
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function


Public Sub unlockSheet(targetSheet As Worksheet, pass As String)
'unlock the worksheet

    On Error Resume Next
    targetSheet.Unprotect Password:=pass
    On Error GoTo 0
    
End Sub

Public Sub lockSheet(targetSheet As Worksheet, pass As String)
'lock the worksheet

    On Error Resume Next
    targetSheet.Protect Password:=pass
    On Error GoTo 0

End Sub
