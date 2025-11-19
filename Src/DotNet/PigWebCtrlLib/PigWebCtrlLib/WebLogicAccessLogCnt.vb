'**********************************
'* Name: WebLogicAccessLogCnt
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic access log statistics , only supports default configuration.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 7/11/2023
'************************************
Imports PigToolsLiteLib
Friend Class WebLogicAccessLogCnt
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.0.18"
    Public Enum EnmtStatisticsType
        LastWeek = 0
        LastMonth = 1
        CurrentWeek = 10
        CurrentMonth = 10
    End Enum


    Public ReadOnly Property LogDirPath As String


    Public Sub New(LogDirPath As String)
        MyBase.New(CLS_VERSION)
        Me.LogDirPath = LogDirPath
    End Sub

    Public Function StartStatistics(StatisticsType As EnmtStatisticsType) As String
        Dim LOG As New PigStepLog("StartStatistics")
        Try
            Dim dteBegin As Date
            Dim dteEnd As Date
            Select Case StatisticsType
                Case EnmtStatisticsType.CurrentMonth

                Case EnmtStatisticsType.LastWeek
                    dteBegin = Now.AddDays(-7)
                    dteEnd = Now.AddDays(-1)
                Case Else
                    Throw New Exception("Unsupported StatisticsType : " & StatisticsType.ToString)
            End Select

            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
