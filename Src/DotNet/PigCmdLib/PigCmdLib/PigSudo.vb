'**********************************
'* Name: PigSudo
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: About the Processing Class of sudo Command|关于sudo命令的处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 22/5/2023
'* 1.1  23/5/2023   Add Run,AsyncRun
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' About the Processing Class of sudo Command|关于sudo命令的处理类
''' </summary>
Public Class PigSudo
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.1.28"

    Private ReadOnly Property mExePath As String = "/bin/sudo"

    Public ReadOnly Property Cmd As String

    Public ReadOnly Property SudoUser As String

    Public ReadOnly Property IsBackRun As Boolean

    Private WithEvents mPigCmdApp As New PigCmdApp

    Public Event AsyncRet_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String)


    Public Sub New(Cmd As String, SudoUser As String, Optional IsBackRun As Boolean = True)
        MyBase.New(CLS_VERSION)
        Me.Cmd = Cmd
        Me.SudoUser = SudoUser
        Me.IsBackRun = IsBackRun
    End Sub

    Private mStandardOutput As String
    Public Property StandardOutput As String
        Get
            Return mStandardOutput
        End Get
        Friend Set(value As String)
            mStandardOutput = value
        End Set
    End Property

    Private mStandardError As String
    Public Property StandardError As String
        Get
            Return mStandardError
        End Get
        Friend Set(value As String)
            mStandardError = value
        End Set
    End Property


    Private ReadOnly Property mPara As String
        Get
            mPara = "-u " & Me.SudoUser
            If Me.IsBackRun = True Then mPara &= " -b"
            mPara &= " " & Me.Cmd
        End Get
    End Property


    Public Function AsyncRun() As String
        Dim LOG As New PigStepLog("AsyncRun")
        Try
            If Me.IsWindows = True Then Throw New Exception("Cannot execute on Windows")
            If Me.IsDebug = True Then Me.PrintDebugLog(LOG.SubName, Me.mPara)
            Dim intThreadID As Integer
            LOG.StepName = "AsyncCallFile"
            LOG.Ret = Me.mPigCmdApp.AsyncCallFile(Me.mExePath, Me.mPara, intThreadID)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(intThreadID)
                Throw New Exception(LOG.Ret)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(mPara)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Run() As String
        Dim LOG As New PigStepLog("Run")
        Try
            If Me.IsWindows = True Then Throw New Exception("Cannot execute on Windows")
            If Me.IsDebug = True Then Me.PrintDebugLog(LOG.SubName, Me.mPara)
            LOG.StepName = "CallFile"
            LOG.Ret = Me.mPigCmdApp.CallFile(Me.mExePath, Me.mPara)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
            Me.StandardOutput = Me.mPigCmdApp.StandardOutput
            Me.StandardError = Me.mPigCmdApp.StandardError
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(mPara)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Sub mPigCmdApp_AsyncRet_CallFile_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles mPigCmdApp.AsyncRet_CallFile_FullString
        RaiseEvent AsyncRet_FullString(AsyncRet, StandardOutput, StandardError)
    End Sub
End Class
