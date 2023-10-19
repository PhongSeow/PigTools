'**********************************
'* Name: PigCmdApp
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 调用操作系统命令的应用|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.18
'* Create Time: 15/1/2022
'* 1.1  31/1/2022  Add CallFile, modify mWinHideShell,mLinuxHideShell
'* 1.2  1/2/2022   Add CmdShell, modify CallFile
'* 1.3  31/3/2022  Add GetParentProc
'* 1.4  1/4/2022   Modify GetParentProc
'* 1.5  3/4/2022   Add EnmStandardOutputReadType,mCallFile, and modify CallFile
'* 1.6  5/4/2022   Add GetSubProcs
'* 1.7  17/5/2022  Add ASyncRet_CmdShell,mCmdShell,ASyncCmdShell, modify CmdShell
'* 1.8  18/5/2022  Modify mCallFile, add AsyncRet_CallFile_FullString,AsyncRet_CallFile_StringArray,AsyncRet_CmdShell_FullString,AsyncRet_CmdShell_StringArray,AsyncCallFile
'* 1.9  26/5/2022  Modify mCallFile
'* 1.10 26/7/2022  Modify Imports
'* 1.11 29/7/2022  Modify Imports
'* 1.12 22/5/2023  Add Nohup,Sudo
'* 1.13 5/6/2023   Modify EnmStandardOutputReadType,CmdShell,CallFile,AsyncCallFile,AsyncCmdShell,mCallFile
'* 1.15 7/8/2023   Add CallFileWaitForExit
'* 1.16 23/8/2023  Modify GetSubProcs,GetParentProc
'* 1.17 9/10/2023  Modify CallFileWaitForExit
'* 1.18 19/10/2023  Modify CallFileWaitForExit
'**********************************
Imports PigToolsLiteLib
Imports System.IO
Imports System.Threading
''' <summary>
''' Operating system command or execution file call processing class|操作系统命令或执行文件调用处理类
''' </summary>
Public Class PigCmdApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.18.8"
    Public LinuxShPath As String = "/bin/sh"
    Public WindowsCmdPath As String
    Private WithEvents mPigFunc As New PigFunc
    Private moPigProcApp As PigProcApp


    Public Enum EnmStandardOutputReadType
        FullString = 0
        StringArray = 1
        'StreamReader = 2
    End Enum

    Public Sub New()
        MyBase.New(CLS_VERSION)
        If Me.IsWindows = True Then
            Me.WindowsCmdPath = mPigFunc.GetEnvVar("windir") & "\System32\cmd.exe"
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

    'Private moStandardOutputStreamReader As StreamReader
    'Public Property StandardOutputStreamReader As StreamReader
    '    Get
    '        Return moStandardOutputStreamReader
    '    End Get
    '    Friend Set(value As StreamReader)
    '        moStandardOutputStreamReader = value
    '    End Set
    'End Property

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
    ''' <param name="StandardOutputReadType">标准输出读类型|Standard output read type</param>
    ''' <returns></returns>
    Public Function CmdShell(Cmd As String, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As String
        Me.StandardOutputReadType = StandardOutputReadType
        Dim struMain As mStruCmdShell
        With struMain
            .Cmd = Cmd
            .IsAsync = False
        End With
        Return Me.mCmdShell(struMain)
    End Function

    ''' <summary>
    ''' 执行操作系统命令|Execute operating system commands
    ''' </summary>
    ''' <param name="Cmd">命令语句|Command statement</param>
    ''' <param name="OutThreadID">输出的线程标识|Output thread ID</param>
    ''' <param name="StandardOutputReadType">标准输出读类型|Standard output read type</param>
    ''' <returns></returns>
    Public Function AsyncCmdShell(Cmd As String, ByRef OutThreadID As Integer, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As String
        Try
            Me.StandardOutputReadType = StandardOutputReadType
            Dim struMain As mStruCmdShell
            With struMain
                .Cmd = Cmd
                .IsAsync = True
            End With
            Dim oThread As New Thread(AddressOf mCmdShell)
            oThread.Start(struMain)
            OutThreadID = oThread.ManagedThreadId
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ASyncCmdShell", ex)
        End Try
    End Function

    ''' <summary>
    ''' 调用文件|Call file
    ''' </summary>
    ''' <param name="FilePath">调用文件的路径|Path to the calling file</param>
    ''' <param name="Para">调用文件的参数|Call file parameters</param>
    ''' <param name="StandardOutputReadType">标准输出读类型|Standard output read type</param>
    ''' <returns></returns>
    Public Function AsyncCallFile(FilePath As String, Para As String, ByRef OutThreadID As Integer, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As String
        Try
            Me.StandardOutputReadType = StandardOutputReadType
            Dim struMain As mStruCallFile
            With struMain
                .FilePath = FilePath
                .Para = Para
                .IsAsync = True
                .IsCmdShell = False
            End With
            Dim oThread As New Thread(AddressOf mCallFile)
            oThread.Start(struMain)
            OutThreadID = oThread.ManagedThreadId
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AsyncCallFile", ex)
        End Try
    End Function


    Private Structure mStruCmdShell
        Public Cmd As String
        Public IsAsync As Boolean
    End Structure
    Private Function mCmdShell(StruMain As mStruCmdShell) As String
        Dim LOG As New PigStepLog("mCmdShell")
        Try
            Dim strShellPath As String
            If Me.IsWindows = True Then
                strShellPath = Me.WindowsCmdPath
            Else
                strShellPath = Me.LinuxShPath
            End If
            Dim strCmd As String
            StruMain.Cmd = Replace(StruMain.Cmd, """", """""")
            If Me.IsWindows = True Then
                strCmd = " /c "
            Else
                strCmd = " -c "
            End If
            strCmd &= """" & StruMain.Cmd & """"
            Dim StruCallFile As mStruCallFile
            With StruCallFile
                .FilePath = strShellPath
                .Para = strCmd
                .IsAsync = StruMain.IsAsync
                .IsCmdShell = True
            End With
            LOG.StepName = "mCallFile"
            LOG.Ret = Me.mCallFile(StruCallFile)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(StruMain.Cmd)
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

    Private Function mWinHideShell(CmdFilePath As String, Optional IsRunAsAdmin As Boolean = False) As Long
        Dim LOG As New PigStepLog("mWinHideShell")
        Try
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(CmdFilePath)
            With moProcessStartInfo
                If IsRunAsAdmin = True Then
                    .Verb = "runas"
                Else
                    .UseShellExecute = False
                End If
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
    ''' <param name="StandardOutputReadType">标准输出读类型|Standard output read type</param>
    ''' <returns></returns>
    Public Function CallFile(FilePath As String, Para As String, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As String
        Me.StandardOutputReadType = StandardOutputReadType
        Dim struMain As mStruCallFile
        With struMain
            .FilePath = FilePath
            .Para = Para
            .IsAsync = False
            .IsCmdShell = False
        End With
        Return Me.mCallFile(struMain)
    End Function


    Private Structure mStruCallFile
        Public FilePath As String
        Public Para As String
        Public IsAsync As Boolean
        Public IsCmdShell As Boolean
    End Structure

    Public Event AsyncRet_CallFile_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String)
    Public Event AsyncRet_CallFile_StringArray(AsyncRet As PigAsync, StandardOutput As String(), StandardError As String)
    Public Event AsyncRet_CmdShell_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String)
    Public Event AsyncRet_CmdShell_StringArray(AsyncRet As PigAsync, StandardOutput As String(), StandardError As String)

    Private Function mCallFile(StruMain As mStruCallFile) As String
        Dim LOG As New PigStepLog("mCallFile")
        Dim oPigAsync As New PigAsync
        Try
            If StruMain.IsAsync = True Then
                Select Case Me.StandardOutputReadType
                    Case EnmStandardOutputReadType.FullString, EnmStandardOutputReadType.StringArray
                    Case Else
                        Throw New Exception("The asynchronous processing mode StandardOutputReadType does not support " & Me.StandardOutputReadType.ToString)
                End Select
                oPigAsync.AsyncBegin()
            End If
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(StruMain.FilePath)
            With moProcessStartInfo
                .UseShellExecute = False
                .CreateNoWindow = True
                .RedirectStandardError = True
                .RedirectStandardOutput = True
                .Arguments = StruMain.Para
            End With
            LOG.StepName = "Process.Start"
            Dim oProcess As Process = Process.Start(moProcessStartInfo)
            Me.PID = oProcess.Id
            oPigAsync.AsyncCmdPID = oProcess.Id
            LOG.StepName = "Process.StandardOutput"
            Dim oStreamReader As StreamReader = oProcess.StandardOutput
            Dim strStandardOutput As String = ""
            Dim abStandardOutputArray(0) As String
            Select Case Me.StandardOutputReadType
                Case EnmStandardOutputReadType.FullString
                    LOG.StepName = "Process.StandardOutput.WaitForExit"
                    oProcess.WaitForExit(Me.CmdWaitForExitTime)
                    LOG.StepName = "StreamReader.ReadToEnd"
                    strStandardOutput = oStreamReader.ReadToEnd
                    LOG.StepName = "StreamReader.Close"
                    oStreamReader.Close()
                    If StruMain.IsAsync = False Then Me.StandardOutput = strStandardOutput
'                Case EnmStandardOutputReadType.StreamReader
                Case EnmStandardOutputReadType.StringArray
                    Dim i As Integer = 0
                    If StruMain.IsAsync = True Then
                        Do While Not oStreamReader.EndOfStream
                            ReDim Preserve abStandardOutputArray(i)
                            'LOG.StepName = "StreamReader.ReadLine(" & i & ")"
                            abStandardOutputArray(i) = oStreamReader.ReadLine
                            i += 1
                        Loop
                    Else
                        ReDim mabStandardOutputArray(i)
                        Do While Not oStreamReader.EndOfStream
                            ReDim Preserve mabStandardOutputArray(i)
                            'LOG.StepName = "StreamReader.ReadLine(" & i & ")"
                            mabStandardOutputArray(i) = oStreamReader.ReadLine
                            i += 1
                        Loop
                    End If
                    LOG.StepName = "StreamReader.Close"
                    oStreamReader.Close()
                Case Else
                    Throw New Exception("Invalid StandardOutputReadType")
            End Select
            'LOG.StepName = "Process.StandardOutput.WaitForExit"
            'oProcess.WaitForExit(Me.CmdWaitForExitTime)
            'LOG.StepName = "Process.StandardError.WaitForExit"
            'oProcess.WaitForExit(Me.CmdWaitForExitTime)
            LOG.StepName = "Process.StandardError"
            Dim srStandardError As StreamReader = oProcess.StandardError
            LOG.StepName = "srStandardError.ReadToEnd"
            Dim strStandardError = srStandardError.ReadToEnd
            'LOG.StepName = "Process.StandardError.WaitForExit"
            'oProcess.WaitForExit(Me.CmdWaitForExitTime)
            srStandardError = Nothing
            LOG.StepName = "Process.Close"
            oProcess.Close()
            oProcess = Nothing
            moProcessStartInfo = Nothing
            If StruMain.IsAsync = True Then
                oPigAsync.AsyncSucc()
                Select Case Me.StandardOutputReadType
                    Case EnmStandardOutputReadType.FullString
                        If StruMain.IsCmdShell = True Then
                            RaiseEvent AsyncRet_CmdShell_FullString(oPigAsync, strStandardOutput, strStandardError)
                        Else
                            RaiseEvent AsyncRet_CallFile_FullString(oPigAsync, strStandardOutput, strStandardError)
                        End If
                    Case EnmStandardOutputReadType.StringArray
                        If StruMain.IsCmdShell = True Then
                            RaiseEvent AsyncRet_CmdShell_StringArray(oPigAsync, abStandardOutputArray, strStandardError)
                        Else
                            RaiseEvent AsyncRet_CallFile_StringArray(oPigAsync, abStandardOutputArray, strStandardError)
                        End If
                    Case Else
                        Throw New Exception("Invalid StandardOutputReadType")
                End Select
            Else
                Me.StandardOutput = strStandardOutput
                Me.StandardError = strStandardError
            End If
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(StruMain.FilePath)
                LOG.AddStepNameInf(StruMain.Para)
            End If
            Dim strError As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            If StruMain.IsAsync = True Then oPigAsync.AsyncError(strError)
            Return strError
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
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.CmdShell(strCmd, EnmStandardOutputReadType.StringArray)
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

    ''' <summary>
    ''' 结束指定进程号的子进程|Kill the child process with the specified process number
    ''' </summary>
    ''' <param name="PID">进程号|Process number</param>
    ''' <returns></returns>
    Public Function KillSubProcs(PID As Integer) As String
        Dim LOG As New PigStepLog("KillSubProcs")
        Try
            LOG.StepName = "GetSubProcs"
            Dim oPigProcs As PigProcs = Me.GetSubProcs(PID)
            If oPigProcs Is Nothing Then Throw New Exception("Unable to get child process")
            Dim strErr As String = ""
            LOG.StepName = "GetSubProcs"
            For Each oPigProc As PigProc In oPigProcs
                Dim intPID As String = oPigProc.ProcessID
                LOG.Ret = oPigProc.Close()
                If LOG.Ret <> "OK" Then strErr &= LOG.Ret & "[" & intPID.ToString & "]"
            Next
            oPigProcs = Nothing
            If strErr <> "OK" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 获取指定进程号的子进程|Gets the child process of the specified process number
    ''' </summary>
    ''' <param name="PID">进程号|Process number</param>
    ''' <returns></returns>
    Public Function GetSubProcs(PID As Integer) As PigProcs
        Dim LOG As New PigStepLog("GetSubProcs")
        Try
            If moPigProcApp Is Nothing Then
                LOG.StepName = "New PigProcApp"
                moPigProcApp = New PigProcApp
                If moPigProcApp.LastErr <> "" Then Throw New Exception(moPigProcApp.LastErr)
            End If
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "wmic process where ParentProcessId=" & PID.ToString & " get ProcessId"
            Else
                strCmd = "ps -ef|awk '{if($3==""" & PID.ToString & """) print $2}'"
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.CmdShell(strCmd, EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "New PigProcs"
            LOG.Ret = ""
            GetSubProcs = New PigProcs
            For Each strSubPID As String In Me.StandardOutputArray
                If IsNumeric(strSubPID) = True Then
                    GetSubProcs.Add(strSubPID)
                    If GetSubProcs.LastErr <> "" Then
                        If LOG.Ret = "" Then

                        End If
                        LOG.AddStepNameInf(strSubPID)
                        LOG.AddStepNameInf(GetSubProcs.LastErr)
                    End If
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Call a file and wait for return|调用一个文件并等待返回
    ''' </summary>
    ''' <param name="FilePath">Path to execute file|执行文件的路径</param>
    ''' <param name="IsRunAsAdmin">Run as administrator|是否以管理员身份运行</param>
    ''' <returns></returns>
    Public Function CallFileWaitForExit(FilePath As String, Optional IsRunAsAdmin As Boolean = False) As String
        Return Me.mCallFileWaitForExit(FilePath, "", IsRunAsAdmin)
    End Function

    ''' <summary>
    ''' Call a file and wait for return|调用一个文件并等待返回
    ''' </summary>
    ''' <param name="FilePath">Path to execute file|执行文件的路径</param>
    ''' <param name="Para">Parameters passed to the execution file|传给执行文件的参数</param>
    ''' <param name="IsRunAsAdmin">Run as administrator|是否以管理员身份运行</param>
    ''' <returns></returns>
    Public Function CallFileWaitForExit(FilePath As String, Para As String, Optional IsRunAsAdmin As Boolean = False) As String
        Return Me.mCallFileWaitForExit(FilePath, Para, IsRunAsAdmin)
    End Function

    Private Function mCallFileWaitForExit(FilePath As String, Optional Para As String = "", Optional IsRunAsAdmin As Boolean = False) As String
        Dim LOG As New PigStepLog("mCallFileWaitForExit")
        Try
            If Me.mPigFunc.IsFileExists(FilePath) = False Then Throw New Exception("The execution file does not exist.")
            LOG.StepName = "New ProcessStartInfo"
            Dim oProcessStartInfo As ProcessStartInfo
            oProcessStartInfo = New ProcessStartInfo(FilePath)
            With oProcessStartInfo
                If IsRunAsAdmin = True Then
                    .Verb = "runas"
                Else
                    .UseShellExecute = False
                End If
                .Arguments = Para
            End With
            Dim oProcess As New Process
            LOG.StepName = "Start and WaitForExit"
            oProcess.StartInfo = oProcessStartInfo
            oProcess.Start()
            oProcess.WaitForExit()
            oProcess = Nothing
            oProcessStartInfo = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
