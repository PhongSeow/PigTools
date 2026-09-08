'**********************************
'* Name: PigLFtp
'* Author: Seow Phong
'* License: Copyright (c) 2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 一个调用 lftp 进行处理的类|A class that calls lftp for processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 22/7/2024
'* 1.1  23/7/2024   Add Run,AsyncRun
'* 1.2  28/7/2024   Modify PigStepLog to StruStepLog
'**********************************
Imports PigToolsLiteLib
Friend Class PigLFtp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "2" & "." & "2"
    Private ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property LFtpPath As String
    Public ReadOnly Property SFtpServer As String
    Public ReadOnly Property UserName As String
    Public ReadOnly Property Password As String
    Public Property IsPassiveMode As Boolean



    Public Sub New(LFtpPath As String, SFtpServer As String, UserName As String, Password As String)
        MyBase.New(CLS_VERSION)
        Try
            Me.LFtpPath = LFtpPath
            Me.SFtpServer = SFtpServer
            Me.UserName = UserName
            Me.Password = Password
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Private ReadOnly Property LftpLoginHead As String
        Get
            Return "ftp://" & Me.UserName & ":" & Me.Password & "@" & Me.SFtpServer
        End Get
    End Property

    Private Function mRunCmd(Cmd As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mRunCmd"
        Try
            LOG.StepName = "IsFileExists"
            If Me.mPigFunc.IsFileExists(Me.LFtpPath) = False Then
                LOG.AddStepNameInf(Me.LFtpPath)
                LOG.Ret = "The lftp execution file does not exist."
            End If
            Dim oPigCmdApp As New PigCmdApp, strPara As String = Me.LftpLoginHead & " -e """ & Cmd & """"
            With oPigCmdApp
                LOG.StepName = "CallFile"
                LOG.Ret = .CallFile(Me.LFtpPath, strPara, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If .StandardError <> "" Then Throw New Exception(.StandardError)
            End With
            oPigCmdApp = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(Me.LFtpPath)
            LOG.AddStepNameInf(Cmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Test() As String
        Dim LOG As New PigStepLog("Test")
        Try

            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
