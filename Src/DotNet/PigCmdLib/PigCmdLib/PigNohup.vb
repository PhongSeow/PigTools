'**********************************
'* Name: PigNohup
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: About the Processing Class of nohup Command|关于nohup命令的处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 24/5/2023
'* 1.1  25/5/2023   Add Run,AsyncRun
'* 1.2  21/7/2024  Modify PigFunc to PigFuncLite
'**********************************
Imports PigToolsLiteLib
Public Class PigNohup
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "2" & "." & "30"

    Public ReadOnly Property Cmd As String
    Public ReadOnly Property OutFilePath As String = ""

    Private Property mPigFunc As PigFunc

    Private WithEvents mPigCmdApp As New PigCmdApp



    Public Sub New(Cmd As String, OutFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.Cmd = Cmd
        Me.OutFilePath = OutFilePath
    End Sub

    Public Sub New(Cmd As String)
        MyBase.New(CLS_VERSION)
        Me.Cmd = Cmd
    End Sub

    Private ReadOnly Property mCmd As String
        Get
            mCmd = "nohup " & Me.Cmd
            If Me.OutFilePath <> "" Then
                mCmd &= " > " & Me.OutFilePath
            Else
                mCmd &= ">/dev/null"
            End If
            mCmd &= " 2>&1 &"
        End Get
    End Property



    Public Function Run() As String
        Dim LOG As New StruStepLog : LOG.SubName = "Run"
        Try
            If Me.IsWindows = True Then Throw New Exception("Cannot execute on Windows")
            If Me.IsDebug = True Then Me.PrintDebugLog(LOG.SubName, Me.mCmd)
            If Me.OutFilePath <> "" Then
                LOG.StepName = "Check out file"
                If Me.mPigFunc Is Nothing Then Me.mPigFunc = New PigFunc
                Dim strPath As String = Me.mPigFunc.GetFilePart(Me.OutFilePath, PigFunc.EnmFilePart.Path)
                If Me.mPigFunc.IsFolderExists(strPath) = False Then
                    LOG.AddStepNameInf(Me.OutFilePath)
                    Throw New Exception("The directory for the out file does not exist.")
                End If
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(Me.mCmd)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(mCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
