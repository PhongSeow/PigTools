'**********************************
'* Name: WebLogicAccessLogDayCnt
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Statistical indicators for recording WebLogic access logs for a day
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 8/11/2023
'* 1.1  9/11/2023   Add TotalDays
'* 1.2  10/11/2023  Modify AddOneDayCnt
'************************************
Imports PigToolsLiteLib
Friend Class WebLogicAccessLogDayCnt
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "2" & "." & "18"

    Public Enum EnmDayType
        Unknow = 0
        DayOfWeek = 1
        DayOfMonth = 2
        DayOfYear = 3
    End Enum

    Private ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property DayNo As String
    Public ReadOnly Property DayType As EnmDayType
    Public Property OkAccessCnt As Long
    Public Property OkAccessMB As Decimal
    Public Property ErrAccessCnt As Long
    Public Property InvalidAccessCnt As Long
    Private Property mDateList As String
    Public ReadOnly Property TotalDays As Integer
        Get
            Try
                Dim strDateList As String = Me.mDateList
                TotalDays = 0
                Do While True
                    Dim strItem As String = mPigFunc.GetStr(strDateList, "<", ">")
                    If strItem = "" Then Exit Do
                    TotalDays += 1
                Loop
            Catch ex As Exception
                TotalDays = -1
                Me.SetSubErrInf("TotalDays.Get", ex)
            End Try
        End Get
    End Property



    Public Sub New(DayNo As String, DayType As EnmDayType)
        MyBase.New(CLS_VERSION)
        Try
            Select Case DayType
                Case EnmDayType.DayOfWeek
                    Select Case DayNo
                        Case "1" To "7"
                            Me.DayNo = DayNo
                        Case Else
                            Throw New Exception("Invalid Week DayNo : " & DayNo)
                    End Select
                Case EnmDayType.DayOfMonth
                    Select Case DayNo
                        Case "1" To "31"
                            Me.DayNo = DayNo
                        Case Else
                            Throw New Exception("Invalid Month DayNo : " & DayNo)
                    End Select
                Case EnmDayType.DayOfYear
                    Select Case DayNo
                        Case "1" To "366"
                            Me.DayNo = DayNo
                        Case Else
                            Throw New Exception("Invalid Year DayNo : " & DayNo)
                    End Select
                Case Else
                    Throw New Exception("Invalid DayType : " & DayType)
            End Select
            Me.DayType = DayType
            Me.Reset()
        Catch ex As Exception
            Me.DayType = EnmDayType.Unknow
            Me.DayNo = -1
            Me.SetSubErrInf("", ex)
        End Try
        Me.DayNo = DayNo
    End Sub

    Public Sub Reset()
        With Me
            .OkAccessMB = 0
            .OkAccessCnt = 0
            .ErrAccessCnt = 0
            .InvalidAccessCnt = 0
            .mDateList = ""
        End With
    End Sub

    Public Function AddOneDayCnt(ByRef WebLogicAccessLogLine As WebLogicAccessLogLine) As String
        Try
            With WebLogicAccessLogLine
                Select Case Left(.HttpStatusCode, 1)
                    Case "1", "2", "3"
                        Me.OkAccessCnt += 1
                        Me.OkAccessMB += CDec(.AccessBytes) / 1024 / 1024
                    Case "4"
                        Me.InvalidAccessCnt += 1
                    Case "5"
                        Me.ErrAccessCnt += 1
                End Select
                If .IsDateIn(Me.mDateList) = False Then
                    Me.mDateList &= "<" & .FmtDate & ">"
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddOneDayCnt", ex)
        End Try
    End Function

    ''' <summary>
    ''' Total Visits|访问次数
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AccessCnt() As Long
        Get
            Return Me.OkAccessCnt + Me.InvalidAccessCnt + Me.ErrAccessCnt
        End Get
    End Property



    ''' <summary>
    ''' Average daily visits|一天平均访问次数
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AvgAccessCntOneDay() As Decimal
        Get
            Return Me.AccessCnt / Me.TotalDays
        End Get
    End Property

    ''' <summary>
    ''' Average number of MB accessed per day|一天平均访问MB数
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AvgAccessMBOneDay() As Decimal
        Get
            Return Me.OkAccessMB / Me.TotalDays
        End Get
    End Property

    ''' <summary>
    ''' Average daily access success rate|平均每天访问成功率
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property AvgAccessOKRateOneDay() As Decimal
        Get
            Return Me.OkAccessCnt / Me.AccessCnt
        End Get
    End Property

End Class
