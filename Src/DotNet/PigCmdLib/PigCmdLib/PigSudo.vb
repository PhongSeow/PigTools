'**********************************
'* Name: PigSudo
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: About the Processing Class of sudo Command|关于sudo命令的处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 22/5/2023
'* 1.1  23/5/2023   Add Run,AsyncRun
'* 1.2  20/11/2023  Modify mPara,New
'* 1.3  13/8/2024  Modify Run,StandardOutput,StandardError, add StringArrayToSpaceMulti2OneStr
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' About the Processing Class of sudo Command|关于sudo命令的处理类
''' </summary>
Public Class PigSudo
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "6"

    Private ReadOnly Property mExePath As String

    Public ReadOnly Property Cmd As String
    Private ReadOnly Property mIsCmd As Boolean
    Public ReadOnly Property CallFilePath As String
    Public ReadOnly Property CallFilePara As String

    Public ReadOnly Property SudoUser As String

    Public ReadOnly Property IsBackRun As Boolean

    Private WithEvents mPigCmdApp As New PigCmdApp

    Public Event AsyncRet_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String)


    Public Sub New(Cmd As String, SudoUser As String, Optional IsBackRun As Boolean = True)
        MyBase.New(CLS_VERSION)
        Me.mExePath = "/bin/sudo"
        Me.Cmd = Cmd
        Me.mIsCmd = True
        Me.SudoUser = SudoUser
        Me.IsBackRun = IsBackRun
    End Sub

    Public Sub New(FilePath As String, Para As String, SudoUser As String, Optional IsBackRun As Boolean = True)
        MyBase.New(CLS_VERSION)
        Me.mExePath = "/bin/sudo"
        Me.CallFilePath = FilePath
        Me.CallFilePara = Para
        Me.mIsCmd = False
        Me.SudoUser = SudoUser
        Me.IsBackRun = IsBackRun
    End Sub

    Public ReadOnly Property StandardOutput As String
        Get
            Try
                If Me.mPigCmdApp IsNot Nothing Then
                    Return Me.mPigCmdApp.StandardOutput
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("StandardOutput", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property StringArrayToSpaceMulti2OneStr(Optional IsTrimConvert As Boolean = True) As String
        Get
            Try
                If Me.mPigCmdApp IsNot Nothing Then
                    Dim strRet As String = "", strOutStr As String = ""
                    strRet = Me.mPigCmdApp.StringArrayToSpaceMulti2OneStr(strOutStr, IsTrimConvert)
                    If strRet <> "OK" Then Throw New Exception(strRet)
                    Return strOutStr
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("StringArrayToSpaceMulti2OneStr", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property StandardError As String
        Get
            Try
                If Me.mPigCmdApp IsNot Nothing Then
                    Return Me.mPigCmdApp.StandardError
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("StandardError", ex)
                Return ""
            End Try
        End Get
    End Property


    Private ReadOnly Property mPara As String
        Get
            mPara = "-u " & Me.SudoUser
            If Me.IsBackRun = True Then mPara &= " -b"
            If Me.mIsCmd = True Then
                mPara &= " /bin/sh -c """ & Me.Cmd & """"
            Else
                mPara &= " " & Me.CallFilePath & " " & Me.CallFilePara
            End If
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

    Public Function Run(Optional StandardOutputReadType As PigCmdApp.EnmStandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.FullString) As String
        Dim LOG As New PigStepLog("Run")
        Try
            If Me.IsWindows = True Then Throw New Exception("Cannot execute on Windows")
            If Me.IsDebug = True Then Me.PrintDebugLog(LOG.SubName, Me.mPara)
            LOG.StepName = "CallFile"
            LOG.Ret = Me.mPigCmdApp.CallFile(Me.mExePath, Me.mPara)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
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
