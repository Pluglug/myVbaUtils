Option Explicit

' テキストもしくは日付のデータを扱う必要があったときにがんばって作ったやつ
Public Function FormatDateOrText(ByVal value As Variant) As String
    Const DATE_FORMAT As String = "yyyy/mm/dd(aaa)"
    Const TIME_FORMAT As String = "hh:mm"
    Const DATE_TIME_FORMAT As String = "yyyy/mm/dd(aaa) hh:mm"
    Const BASE_DATE As String = "1899/12/30"
    Const SEC_IN_DAY As Long = 86400

    ' Validation
    Select Case VarType(value)
        Case vbString, vbDate, vbDouble, vbInteger, vbLong
            ' OK
        Case vbEmpty
            FormatDateOrText = ""
            Exit Function
        Case vbBoolean
            FormatDateOrText = IIf(value, "True", "False")
            Exit Function
    End Select
    
    If IsDate(value) Then
        If CDate(value) = Int(CDate(value)) Then
            FormatDateOrText = Format(CDate(value), DATE_FORMAT)
        ElseIf CDate(value) = TimeValue(CDate(value)) Then
            FormatDateOrText = Format(CDate(value), TIME_FORMAT)
        Else
            FormatDateOrText = Format(CDate(value), DATE_TIME_FORMAT)
        End If
    
    ElseIf IsNumeric(value) Then
        If IsDateSerialInRange(value) Or value < 1 Then  ' 日時もしくは時間を示す数値
            If Int(value) = value Then
                FormatDateOrText = Format(DateAdd("d", value, BASE_DATE), DATE_FORMAT)
            ElseIf Int(value) = 0 Then
                FormatDateOrText = Format(DateAdd("s", value * SEC_IN_DAY, BASE_DATE), TIME_FORMAT)
            Else
                FormatDateOrText = Format(DateAdd("s", value * SEC_IN_DAY, BASE_DATE), DATE_TIME_FORMAT)
            End If
        Else
            FormatDateOrText = CStr(value)
        End If
    Else
        FormatDateOrText = CStr(value)
    End If
End Function

Private Function IsDateSerialInRange(ByVal serialValue As Double) As Boolean
    Const BASE_DATE As String = "1899/12/30"
    Const SEC_IN_DAY As Long = 86400
    Dim today As Date
    today = Date
    
    Dim minDate As Date
    minDate = DateAdd("yyyy", -1, today)
    
    Dim maxDate As Date
    maxDate = DateAdd("yyyy", 10, today)
    
    Dim dateValue As Date
    dateValue = DateAdd("s", serialValue * SEC_IN_DAY, BASE_DATE)
    
    IsDateSerialInRange = (dateValue >= minDate) And (dateValue <= maxDate)
End Function