Module WxWork2LxIIS
    Public goIISMain As fIISMain

    Public Function SQLStr(strValue As String, Optional IsNotNull As Boolean = False) As String
        strValue = Replace(strValue, "'", "''")
        If UCase(strValue) = "NULL" And IsNotNull = False Then
            SQLStr = "NULL"
        Else
            SQLStr = "'" & strValue & "'"
        End If
    End Function


    Public Function SQLNum(NumValue As String) As String
        If UCase(NumValue) = "NULL" Then
            SQLNum = "NULL"
        ElseIf IsNumeric(NumValue) = True Then
            SQLNum = NumValue
        Else
            SQLNum = "Not Numeric"
        End If
    End Function

    Public Function SQLBool(BoolValue As Boolean) As String
        Dim lngValue As Long
        lngValue = Math.Abs(CLng(BoolValue))
        SQLBool = CStr(lngValue)
    End Function

    Public Function SQLDate(DateValue As Date) As String
        SQLDate = "'" & Format(DateValue, "yyyy-MM-dd HH:mm:ss.fff") & "'"
    End Function

End Module
