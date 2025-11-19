'**********************************
'* Name: WebLogicAccessLogLine
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic one line log, only supports default configuration.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 7/11/2023
'* 1.1  9/11/2023   Add DayOfWeek,DayOfMonth,DayOfYear
'* 1.2  9/11/2023   Modify New
'* 1.3  28/7/2024   Modify PigStepLog to StruStepLog
'************************************
Imports PigToolsLiteLib
Public Class WebLogicAccessLogLine
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "8"

    Public ReadOnly Property ClientIp As String
    Public ReadOnly Property ClientUser As String
    Public ReadOnly Property AccessTime As Date
    Public ReadOnly Property HttpMethod As String
    Public ReadOnly Property RequestUri As String
    Public ReadOnly Property HttpStatusCode As String
    Public ReadOnly Property AccessBytes As Long
    Public ReadOnly Property ClientReferer As String
    Public ReadOnly Property ClientUserAgent As String
    Public ReadOnly Property ClientCookie As String
    Public ReadOnly Property IsAccessTimeErr As Boolean = False


    Friend Sub New(Line As String)
        MyBase.New(CLS_VERSION)
        Dim LOG As New StruStepLog : LOG.SubName = "New"
        Try
            Line &= " "
            Dim strItem As String = "", strSubItem As String = ""
            LOG.StepName = "Get Time"
            '[25/Jul/2023:18:59:39 +0800] 
            strItem = Me.mGetStr(Line, "[", "] ")
            strSubItem = Me.mGetStr(strItem, "", ":")
            strItem = strSubItem & " " & strItem
            LOG.StepName = "Get AccessTime"
            If IsDate(strItem) = False Then
                Me.IsAccessTimeErr = True
                LOG.AddStepNameInf(strItem)
                Throw New Exception("Invalid AccessTime")
            End If
            Me.AccessTime = CDate(strItem)
            LOG.StepName = "Get Http"
            '"GET /ShjSl97/techlevelsyssh/bigData/js/bubble.js HTTP/1.0"
            LOG.StepName = "Get Other"
            strItem = Me.mGetStr(Line, """", """ ")
            Me.HttpMethod = Me.mGetStr(strItem, "", " ")
            Me.RequestUri = Me.mGetStr(strItem, "", " ")
            Me.ClientIp = Me.mGetStr(Line, "", " ")
            strItem = Me.mGetStr(Line, "", " ")
            strItem = Me.mGetStr(Line, "", " ")
            Me.HttpStatusCode = Me.mGetStr(Line, "", " ")
            LOG.StepName = "Get AccessBytes"
            strItem = Me.mGetStr(Line, "", " ")
            If IsNumeric(strItem) = False Then
                LOG.AddStepNameInf(strItem)
                Throw New Exception("Invalid AccessBytes")
            End If
            Me.AccessBytes = CLng(strItem)
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Private Function mGetStr(ByRef SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
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
            mGetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
            If IsCut = True Then
                SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
            End If
        Catch ex As Exception
            Return ""
            Me.SetSubErrInf("mGetStr", ex)
        End Try
    End Function

    Public ReadOnly Property DayOfWeek() As Integer
        Get
            Return DatePart(DateInterval.Weekday, Me.AccessTime)
        End Get
    End Property

    Public ReadOnly Property DayOfMonth() As Integer
        Get
            Return DatePart(DateInterval.Day, Me.AccessTime)
        End Get
    End Property

    Public ReadOnly Property DayOfYear() As Integer
        Get
            Return DatePart(DateInterval.DayOfYear, Me.AccessTime)
        End Get
    End Property

    Public ReadOnly Property IsDateIn(DateList As String) As Boolean
        Get
            Try
                Dim strDate As String = "<" & Format(Me.AccessTime, "yyyy-MM-dd") & ">"
                If InStr(DateList, strDate) > 0 Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IsDateIn.Get", ex)
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property FmtDate() As String
        Get
            Return Format(Me.AccessTime, "yyyy-MM-dd")
        End Get
    End Property

End Class
