'**********************************
'* Name: PigCmdAsync
'* Author: Seow Phong
'* License: Copyright (c) 2022-2026 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 异步处理调用操作系统命令的应用|Application of asynchronously calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 7/9/2026
'* 1.1  6/9/2026  Add CallFileAsync,CmdShellAsync,GetPsEfCmdInfAsync
'* 1.2  7/9/2026  Add GetSubProcsAsync,KillSubProcsAsync
'**********************************
Imports System.IO
Imports PigToolsLiteLib

Public Class PigCmdAsync
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "2" & "." & "38"
    Public Property LinuxShPath As String
    Public Property WindowsCmdPath As String
    Private WithEvents mPigFunc As New PigFunc
    Private moPigProcApp As PigProcApp

    Public Enum EnmStandardOutputReadType
        FullString = 0
        StringArray = 1
    End Enum

    Private Structure mStruCallFile
        Public FilePath As String
        Public Para As String
        Public IsCmdShell As Boolean
    End Structure

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Dim strCont1 As String, strCont2 As String
        If Me.IsWindows = True Then
            strCont1 = "windir"
            strCont2 = "\System32\cmd.exe"
            Me.WindowsCmdPath = mPigFunc.GetEnvVar(strCont1) & strCont2
        Else
            strCont1 = "/bin/sh"
            Me.LinuxShPath = strCont1
        End If
    End Sub

    Private mabStandardOutputArray(-1) As String
    Public Property StandardOutputArray As String()
        Get
            Return mabStandardOutputArray
        End Get
        Friend Set(value As String())
            mabStandardOutputArray = value
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

    Public Function CallFileWaitForExit(FilePath As String, Optional IsRunAsAdmin As Boolean = False) As String
        Return Me.mCallFileWaitForExit(FilePath, "", IsRunAsAdmin)
    End Function

    Private Function mCallFileWaitForExit(FilePath As String, Optional Para As String = "", Optional IsRunAsAdmin As Boolean = False) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mCallFileWaitForExit"
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
        Dim LOG As New StruStepLog : LOG.SubName = "mWinHideShell"
        Try
            LOG.StepName = "New ProcessStartInfo"
            Dim moProcessStartInfo As New ProcessStartInfo(CmdFilePath)
            With moProcessStartInfo
                If IsRunAsAdmin = True Then
                    Dim strCont1 As String
                    strCont1 = "runas"
                    .Verb = strCont1
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
        Dim LOG As New StruStepLog : LOG.SubName = "mLinuxHideShell"
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

    Public Function StringArrayToSpaceMulti2OneStr(ByRef OutStr As String, Optional IsTrimConvert As Boolean = True) As String
        Dim LOG As New PigStepLog("StringArrayToSpaceMulti2OneStr")
        Try
            OutStr = ""
            LOG.StepName = "StrSpaceMulti2One"
            Dim intLineNo As Integer = 1
            For Each strLine As String In Me.StandardOutputArray
                Dim strLineOut As String = ""
                LOG.Ret = Me.mPigFunc.StrSpaceMulti2One(strLine, strLineOut, IsTrimConvert)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(intLineNo.ToString)
                    LOG.AddStepNameInf(strLine)
                    Throw New Exception(LOG.Ret)
                End If
                If IsTrimConvert = False Then
                    strLineOut = "<" & Trim(strLineOut) & ">"
                End If
                OutStr &= strLineOut & Me.OsCrLf
                intLineNo += 1
            Next
            Return "OK"
        Catch ex As Exception
            OutStr = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

#If NETCOREAPP Or NET451_OR_GREATER Then
    Private Async Function mCallFileAsync(StruMain As mStruCallFile) As Task(Of String)
        Dim LOG As New StruStepLog : LOG.SubName = "mCallFileAsync"
        Dim oPigAsync As New PigAsync
        Try
            ReDim Me.StandardOutputArray(-1)
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
            Dim abStandardOutputArray(-1) As String
            Select Case Me.StandardOutputReadType
                Case EnmStandardOutputReadType.FullString
                    LOG.StepName = "Process.StandardOutput.WaitForExit"
                    oProcess.WaitForExit(Me.CmdWaitForExitTime)
                    LOG.StepName = "StreamReader.ReadToEndAsync"
                    strStandardOutput = Await oStreamReader.ReadToEndAsync
                    LOG.StepName = "StreamReader.Close"
                    oStreamReader.Close()
                    Me.StandardOutput = strStandardOutput
                Case EnmStandardOutputReadType.StringArray
                    Dim i As Integer = 0
                    Do While Not oStreamReader.EndOfStream
                        ReDim Preserve abStandardOutputArray(i)
                        abStandardOutputArray(i) = Await oStreamReader.ReadLineAsync
                        i += 1
                    Loop
                    LOG.StepName = "StreamReader.Close"
                    oStreamReader.Close()
                Case Else
                    Throw New Exception("Invalid StandardOutputReadType")
            End Select
            LOG.StepName = "Process.StandardError"
            Dim srStandardError As StreamReader = oProcess.StandardError
            LOG.StepName = "srStandardError.ReadToEndAsync"
            Dim strStandardError = Await srStandardError.ReadToEndAsync
            srStandardError = Nothing
            LOG.StepName = "Process.Close"
            oProcess.Close()
            oProcess = Nothing
            moProcessStartInfo = Nothing
            Me.StandardOutput = strStandardOutput
            Me.StandardError = strStandardError
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(StruMain.FilePath)
                LOG.AddStepNameInf(StruMain.Para)
            End If
            Dim strError As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            oPigAsync.AsyncError(strError)
            Return strError
        End Try
    End Function

    Public Async Function CallFileAsync(FilePath As String, Para As String, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As Task(Of String)
        Me.StandardOutputReadType = StandardOutputReadType
        Dim struMain As mStruCallFile
        With struMain
            .FilePath = FilePath
            .Para = Para
            .IsCmdShell = False
        End With
        Return Await Me.mCallFileAsync(struMain)
    End Function

    Private Async Function mCmdShellAsync(Cmd As String) As Task(Of String)
        Dim LOG As New StruStepLog : LOG.SubName = "mCmdShellAsync"
        Try
            Dim strShellPath As String
            If Me.IsWindows = True Then
                strShellPath = Me.WindowsCmdPath
            Else
                strShellPath = Me.LinuxShPath
            End If
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = " /C "
                If InStr(Cmd, """") > 0 Then
                    strCmd &= """"
                    strCmd &= Cmd
                    strCmd &= """"
                Else
                    strCmd &= Cmd
                End If
            Else
                If InStr(Cmd, """") > 0 Then
                    Cmd = Replace(Cmd, """", "\""")
                End If
                strCmd = " -c """
                strCmd &= Cmd
                strCmd &= """"
            End If
            Dim StruCallFile As mStruCallFile
            With StruCallFile
                .FilePath = strShellPath
                .Para = strCmd
                .IsCmdShell = True
            End With
            LOG.StepName = "mCallFileAsync"
            LOG.Ret = Await Me.mCallFileAsync(StruCallFile)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            If Me.IsDebug = True Then
                LOG.AddStepNameInf(Cmd)
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Async Function CmdShellAsync(Cmd As String, Optional StandardOutputReadType As EnmStandardOutputReadType = EnmStandardOutputReadType.FullString) As Task(Of String)
        Me.StandardOutputReadType = StandardOutputReadType
        Return Await Me.mCmdShellAsync(Cmd)
    End Function

        Public Async Function GetParentProcAsync(PID As Integer) As Task(Of PigProc)
        Dim LOG As New StruStepLog : LOG.SubName = "GetParentProcAsync"
        Try
            If moPigProcApp Is Nothing Then
                LOG.StepName = "New PigProcApp"
                moPigProcApp = New PigProcApp
                If moPigProcApp.LastErr <> "" Then Throw New Exception(moPigProcApp.LastErr)
            End If
            Dim strCmd As String = ""
            If Me.IsWindows = True Then
                strCmd &= "wmic process where ProcessId="
                strCmd &= PID.ToString
                strCmd &= " get ParentProcessId"
            Else
                strCmd &= "ps -ef|awk '{if($2=="""
                strCmd &= PID.ToString
                strCmd &= """) print $3}'"
            End If
            LOG.StepName = "CmdShellAsync"
            LOG.Ret = Await Me.CmdShellAsync(strCmd, EnmStandardOutputReadType.StringArray)
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
            Dim oPigProc As PigProc = New PigProc(lngParentPID)
            If oPigProc.LastErr <> "" Then
                Dim strCont1 As String = ""
                strCont1 &= "PID="
                strCont1 &= lngParentPID.ToString
                LOG.AddStepNameInf(strCont1)
                Throw New Exception(oPigProc.LastErr)
            End If
            Return oPigProc
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Async Function GetPsEfCmdInfAsync(PID As Integer) As Task(Of String)
        Dim LOG As New StruStepLog : LOG.SubName = "GetPsEfCmdInfAsync"
        Try
            If Me.IsWindows = True Then Throw New Exception("Can only run on Linux.")
            Dim strCmd As String = "ps -ef|awk '{if($2==""" & PID.ToString & """){for(i=8;i<=NF;i++)printf $i"" "";print """"}}'"
            LOG.StepName = "CmdShell"
            LOG.Ret = Await Me.CmdShellAsync(strCmd, EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            Return Trim(Me.StandardOutputArray(0))
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

        Public Async Function GetSubProcsAsync(PID As Integer) As Task(Of PigProcs)
        Dim LOG As New StruStepLog : LOG.SubName = "GetSubProcsAsync"
        Try
            If moPigProcApp Is Nothing Then
                LOG.StepName = "New PigProcApp"
                moPigProcApp = New PigProcApp
                If moPigProcApp.LastErr <> "" Then Throw New Exception(moPigProcApp.LastErr)
            End If
            Dim strCmd As String = ""
            If Me.IsWindows = True Then
                strCmd &= "wmic process where ParentProcessId="
                strCmd &= PID.ToString
                strCmd &= " get ProcessId"
            Else
                strCmd &= "ps -ef|awk '{if($3=="""
                strCmd &= PID.ToString
                strCmd &= """) print $2}'"
            End If
            LOG.StepName = "CmdShellAsync"
            LOG.Ret = Await Me.CmdShellAsync(strCmd, EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf Me.StandardError <> "" Then
                LOG.Ret = Me.StandardError
                If Me.IsDebug = True Then
                    LOG.AddStepNameInf(strCmd)
                End If
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "New PigProcs"
            Dim oPigProcs As PigProcs = New PigProcs
            LOG.StepName = "For Each StandardOutputArray"
            For Each strSubPID As String In Me.StandardOutputArray
                LOG.AddStepNameInf("SubPID=" & strSubPID)
                If IsNumeric(strSubPID) = True Then
                    oPigProcs.Add(strSubPID)
                    If oPigProcs.LastErr <> "" Then
                        LOG.Ret = oPigProcs.LastErr
                        Throw New Exception(LOG.Ret)
                    End If
                End If
            Next
            Me.ClearErr()
            Return oPigProcs
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Async Function KillSubProcsAsync(PID As Integer) As Task(Of String)
        Dim LOG As New StruStepLog : LOG.SubName = "KillSubProcsAsync"
        Try
            LOG.StepName = "GetSubProcs"
            Dim oPigProcs As PigProcs = Await Me.GetSubProcsAsync(PID)
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


#End If



End Class
