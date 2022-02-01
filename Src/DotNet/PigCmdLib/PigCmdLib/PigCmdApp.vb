'**********************************
'* Name: PigCmdApp
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 调用操作系统命令的应用|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 15/1/2022
'*1.1  31/1/2022   Add CallFile, modify mWinHideShell,mLinuxHideShell
'*1.2  1/2/2022   Add CmdShell, modify CallFile
'**********************************
Imports PigToolsLiteLib
Imports System.IO
Public Class PigCmdApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.8"
    Public LinuxShPath As String = "/bin/sh"
    Public WindowsCmdPath As String
    Private moPigFunc As New PigFunc
    Public Sub New()
        MyBase.New(CLS_VERSION)
        If Me.IsWindows = True Then
            Me.WindowsCmdPath = moPigFunc.GetEnvVar("windir") & "\System32\cmd.exe"
        End If
    End Sub

    Private mintCmdWaitForExitTime As Integer = 10
    Public Property CmdWaitForExitTime As Long
        Get
            Return mintCmdWaitForExitTime
        End Get
        Set(value As Long)
            If value < 0 Then value = 10
            mintCmdWaitForExitTime = value
        End Set
    End Property

    Private mlngPID As Long
    Public Property PID As Long
        Get
            Return mlngPID
        End Get
        Friend Set(value As Long)
            mlngPID = value
        End Set
    End Property


    Private mstrStandardOutput As String = ""
    Public Property StandardOutput As String
        Get
            Return mstrStandardOutput
        End Get
        Friend Set(value As String)
            mstrStandardOutput = value
        End Set
    End Property

    Private mstrStandardError As String = ""
    Public Property StandardError As String
        Get
            Return mstrStandardError
        End Get
        Friend Set(value As String)
            mstrStandardError = value
        End Set
    End Property

    ''' <summary>
    ''' 执行操作系统命令|Execute operating system commands
    ''' </summary>
    ''' <param name="Cmd">命令语句|Command statement</param>
    ''' <returns></returns>
    Public Function CmdShell(Cmd As String) As String
        Dim LOG As New PigStepLog("CmdShell")
        Try
            Dim strShellPath As String
            If Me.IsWindows = True Then
                strShellPath = Me.WindowsCmdPath
            Else
                strShellPath = Me.LinuxShPath
            End If
            Dim strCmd As String
            Cmd = Replace(Cmd, """", """""")
            If Me.IsWindows = True Then
                strCmd = " /c "
            Else
                strCmd = " -c "
            End If
            strCmd &= """" & Cmd & """"
            LOG.StepName = "CallFile"
            LOG.Ret = Me.CallFile(strShellPath, strCmd)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(Cmd)
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' 类似VB6的隐藏窗口的Shell命令|Shell command of hidden window similar to VB6
    ''' </summary>
    ''' <param name="CmdFilePath">命令文件路径|Command file path</param>
    ''' <returns>执行文件的进程号|Shell command of hidden window similar to VB6</returns>
    Public Function HideShell(CmdFilePath As String) As Long
        If Me.IsWindows = True Then
            Return Me.mWinHideShell(CmdFilePath)
        Else
            Return Me.mLinuxHideShell(CmdFilePath)
        End If
    End Function

    Private Function mWinHideShell(CmdFilePath As String) As Long
        Dim LOG As New PigStepLog("mWinHideShell")
        Try
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(CmdFilePath)
            With moProcessStartInfo
                .UseShellExecute = True
                .CreateNoWindow = True
            End With
            LOG.StepName = "Process.Start"
            Dim oProcess As Process = Process.Start(moProcessStartInfo)
            mWinHideShell = oProcess.Id
            Me.PID = oProcess.Id
            oProcess = Nothing
            moProcessStartInfo = Nothing
            Me.ClearErr()
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(CmdFilePath)
            End If
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return -1
        End Try
    End Function

    Private Function mLinuxHideShell(CmdFilePath As String) As Long
        Dim LOG As New PigStepLog("mLinuxHideShell")
        Try
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(Me.LinuxShPath)
            With moProcessStartInfo
                .UseShellExecute = False
                .CreateNoWindow = True
                .RedirectStandardInput = True
            End With
            LOG.StepName = "Process.Start"
            Dim oProcess As Process = Process.Start(moProcessStartInfo)
            LOG.StepName = "Process.StandardInput"
            oProcess.StandardInput.WriteLine(CmdFilePath)
            mLinuxHideShell = oProcess.Id
            Me.PID = oProcess.Id
            oProcess = Nothing
            moProcessStartInfo = Nothing
            Me.ClearErr()
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(CmdFilePath)
            End If
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' 调用文件|Call file
    ''' </summary>
    ''' <param name="FilePath">调用文件的路径|Path to the calling file</param>
    ''' <param name="Para">调用文件的参数|Call file parameters</param>
    ''' <returns></returns>
    Public Function CallFile(FilePath As String, Para As String) As String
        Dim LOG As New PigStepLog("CallFile")
        Try
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(FilePath)
            With moProcessStartInfo
                .UseShellExecute = False
                .CreateNoWindow = True
                .RedirectStandardError = True
                .RedirectStandardOutput = True
                .Arguments = Para
            End With
            LOG.StepName = "Process.Start"
            Dim oProcess As Process = Process.Start(moProcessStartInfo)
            Me.PID = oProcess.Id
            LOG.StepName = "Process.StandardOutput"
            Dim srStandardOutput As StreamReader = oProcess.StandardOutput
            LOG.StepName = "StandardOutput.ReadToEnd"
            Me.StandardOutput = srStandardOutput.ReadToEnd
            LOG.StepName = "Process.StandardOutput.WaitForExit"
            oProcess.WaitForExit(Me.CmdWaitForExitTime)
            srStandardOutput = Nothing
            LOG.StepName = "Process.StandardError"
            Dim srStandardError As StreamReader = oProcess.StandardError
            LOG.StepName = "srStandardError.ReadToEnd"
            Me.StandardError = srStandardError.ReadToEnd
            LOG.StepName = "Process.StandardError.WaitForExit"
            oProcess.WaitForExit(Me.CmdWaitForExitTime)
            srStandardError = Nothing
            LOG.StepName = "Process.Close"
            oProcess.Close()
            oProcess = Nothing
            moProcessStartInfo = Nothing
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(FilePath)
                LOG.AddStepNameInf(Para)
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
