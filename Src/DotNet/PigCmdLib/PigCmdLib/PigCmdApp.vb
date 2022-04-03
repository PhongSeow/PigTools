'**********************************
'* Name: PigCmdApp
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 调用操作系统命令的应用|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 15/1/2022
'*1.1  31/1/2022   Add CallFile, modify mWinHideShell,mLinuxHideShell
'*1.2  1/2/2022   Add CmdShell, modify CallFile
'*1.3  31/3/2022  Add GetParentProc
'*1.4  1/4/2022   Modify GetParentProc
'*1.5  3/4/2022   Add EnmStandardOutputReadType,mCallFile, and modify CallFile
'**********************************
Imports PigToolsLiteLib
Imports System.IO
Public Class PigCmdApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.9"
    Public LinuxShPath As String = "/bin/sh"
    Public WindowsCmdPath As String
    Private moPigFunc As New PigFunc
    Private moPigProcApp As PigProcApp

    Public Enum EnmStandardOutputReadType
        FullString = 0
        StringArray = 1
        StreamReader = 2
    End Enum

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



    Private mintStandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString
    Public Property StandardOutputReadType As EnmStandardOutputReadType
        Get
            Return mintStandardOutputReadType
        End Get
        Friend Set(value As EnmStandardOutputReadType)
            mintStandardOutputReadType = value
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

    Private moStandardOutputStreamReader As StreamReader
    Public Property StandardOutputStreamReader As StreamReader
        Get
            Return moStandardOutputStreamReader
        End Get
        Friend Set(value As StreamReader)
            moStandardOutputStreamReader = value
        End Set
    End Property

    Private mabStandardOutputArray As String()
    Public ReadOnly Property StandardOutputArray As String()
        Get
            Return mabStandardOutputArray
        End Get
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
            Me.StandardOutputStreamReader = oProcess.StandardOutput
            Select Case Me.StandardOutputReadType
                Case EnmStandardOutputReadType.FullString
                    LOG.StepName = "StandardOutputStreamReader.ReadToEnd"
                    Me.StandardOutput = Me.StandardOutputStreamReader.ReadToEnd
                    LOG.StepName = "StandardOutputStreamReader.Close"
                    Me.StandardOutputStreamReader.Close()
                Case EnmStandardOutputReadType.StreamReader
                Case EnmStandardOutputReadType.StringArray
                    Dim i As Integer = 0
                    ReDim mabStandardOutputArray(i)
                    Do While Not Me.StandardOutputStreamReader.EndOfStream
                        ReDim Preserve mabStandardOutputArray(i)
                        LOG.StepName = "StandardOutputStreamReader.ReadLine(" & i & ")"
                        mabStandardOutputArray(i) = Me.StandardOutputStreamReader.ReadLine
                        i += 1
                    Loop
                    LOG.StepName = "StandardOutputStreamReader.Close"
                    Me.StandardOutputStreamReader.Close()
                Case Else
                    Throw New Exception("Invalid StandardOutputReadType")
            End Select
            LOG.StepName = "Process.StandardOutput.WaitForExit"
            oProcess.WaitForExit(Me.CmdWaitForExitTime)
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

    ''' <summary>
    ''' 获取指定进程号的父进程|Gets the parent process of the specified process number
    ''' </summary>
    ''' <param name="PID">进程号|Process number</param>
    ''' <returns></returns>
    Public Function GetParentProc(PID As Integer) As PigProc
        Dim LOG As New PigStepLog("GetParentProc")
        Try
            If moPigProcApp Is Nothing Then
                LOG.StepName = "New PigProcApp"
                moPigProcApp = New PigProcApp
                If moPigProcApp.LastErr <> "" Then Throw New Exception(moPigProcApp.LastErr)
            End If
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "wmic process where ProcessId=" & PID.ToString & " get ParentProcessId"
            Else
                strCmd = "ps -ef|awk '{if($2==""" & PID.ToString & """) print $3}'"
            End If
            Me.StandardOutputReadType = EnmStandardOutputReadType.StringArray
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.CmdShell(strCmd)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            Dim lngParentPID As Integer = -1
            For i = 0 To Me.StandardOutputArray.Length - 1
                Dim strLine As String = Trim(Me.StandardOutputArray(i))
                If IsNumeric(strLine) = True Then
                    lngParentPID = CInt(strLine)
                    Exit For
                End If
            Next
            If lngParentPID = -1 Then
                If Me.IsDebug = True Then LOG.AddStepNameInf(strCmd)
                Throw New Exception("Cannot get parent process number")
            End If
            LOG.StepName = "New PigProc"
            GetParentProc = New PigProc(lngParentPID)
            If GetParentProc.LastErr <> "" Then
                LOG.AddStepNameInf("PID=" & lngParentPID.ToString)
                Throw New Exception(GetParentProc.LastErr)
            End If

        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


End Class
