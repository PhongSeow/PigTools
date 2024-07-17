'**********************************
'* Name: PigFuncLite
'* Author: Seow Phong
'* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Lightweight PigFunc, code from PigFunc|轻量的PigFunc，代码来自PigFunc
'* Home Url: https://en.seowphong.com
'* Version: 1.69
'**********************************

Imports PigToolsLiteLib

Friend Class PigFuncLite
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "69" & "." & "6"

    Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <remarks>截取字符串</remarks>
    Public Function GetStr(ByRef SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
        Try
            Dim lngBegin As Long
            Dim lngEnd As Long
            Dim lngBeginLen As Long
            Dim lngEndLen As Long
            lngBeginLen = Len(strBegin)
            lngBegin = InStr(SourceStr, strBegin, CompareMethod.Text)
            lngEndLen = Len(strEnd)
            If lngEndLen = 0 Then
                lngEnd = Len(SourceStr) + 1
            Else
                lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd, CompareMethod.Text)
                If lngBegin = 0 Then Return "" 'Throw New Exception("lngBegin=0")
            End If
            If lngEnd <= lngBegin Then Return "" ' Throw New Exception("lngEnd <= lngBegin")
            If lngBegin = 0 Then Return "" 'Throw New Exception("lngBegin=0[2]")
            GetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
            If IsCut = True Then
                SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
            End If
        Catch ex As Exception
            Return ""
            Me.SetSubErrInf("GetStr", ex)
        End Try
    End Function

    Public Function StrSpaceMulti2One(SrcStr As String, ByRef OutStr As String, Optional IsTrimConvert As Boolean = True) As String
        Try
            Const SPACE_2 As String = "  "
            Const SPACE_1 As String = " "
            OutStr = SrcStr
            Do While InStr(OutStr, SPACE_2) > 0
                OutStr = Replace(OutStr, SPACE_2, SPACE_1)
            Loop
            If IsTrimConvert = True Then
                OutStr = Trim(OutStr)
                OutStr = "<" & Replace(OutStr, SPACE_1, "><") & ">"
            End If
            Return "OK"
        Catch ex As Exception
            OutStr = ""
            Return Me.GetSubErrInf("StrSpaceMulti2One", ex)
        End Try
    End Function


End Class
