Option Explicit

Function MatchesPattern(text As String, pattern As String) As Boolean
    MatchesPattern = text Like pattern
End Function

Function ExtractValue(line As String, key As String) As String
    Dim startPos As Long
    startPos = InStr(line, key) + Len(key)
    ExtractValue = Trim(Mid(line, startPos))
End Function

Function ExtractValues(line As String, keyStringsArray As Variant) As Variant
    Dim values() As String
    ReDim values(UBound(keyStringsArray))
    Dim startPos As Long, endPos As Long, i As Long

    For i = LBound(keyStringsArray) To UBound(keyStringsArray) - 1
        startPos = InStr(line, keyStringsArray(i)) + Len(keyStringsArray(i))
        endPos = InStr(line, keyStringsArray(i + 1)) - 1
        values(i) = Trim(Mid(line, startPos, endPos - startPos))
    Next i

    startPos = InStr(line, keyStringsArray(UBound(keyStringsArray))) + Len(keyStringsArray(UBound(keyStringsArray)))
    values(UBound(keyStringsArray)) = Trim(Mid(line, startPos))

    ExtractValues = values
End Function


Function GetRegExpMatch(ByVal text As String, ByVal pattern As String) As Object
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")

    With RegExp
        .Global = False
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With

    Dim Matches As Object
    Set Matches = RegExp.Execute(text)

    If Matches.Count > 0 Then
        Set GetRegExpMatch = Matches(0)
    Else
        Set GetRegExpMatch = Nothing
    End If
End Function

