'**********************************
'* Name: WeblogicServer
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic Server
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 30/9/2024
'* 1.1  30/9/2024   Code initialization
'* 1.2  8/10/2024   Modify SecurityDirPath
'* 1.3  10/10/2024  Add RunStatus,RefRunStatus,SaveSecurityBoot,StartServer,StopServer
'* 1.5  12/10/2024  Add ManagementServer, Modify RefRunStatus,mPigCmdApp_AsyncRet_CmdShell_FullString
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib

Public Class WeblogicServer
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "5" & "." & "28"

    Private WithEvents mPigCmdApp As New PigCmdApp
    Private mPigSysCmd As New PigSysCmd
    Public ReadOnly Property Parent As WebLogicDomain
    Public ReadOnly Property ServerName As String
    Public ReadOnly Property HomeDirPath As String
    Public ReadOnly Property ListenPort As Integer
    Private ReadOnly Property mPigFunc As New PigFunc
    Private mPigProcApp As New PigProcApp
    Private mFS As New PigFileSystem

    Private Property mStopDomainThreadID As Integer
    Private Property mStopDomainBeginTime As Date
    Private Property mStartDomainThreadID As Integer
    Private Property mStartDomainBeginTime

    Public Event StopDomainSucc(ResInf As String)

    Public Event StopDomainFail(ErrInf As String)

    Private mManagementServer As String
    Public Property ManagementServer As String
        Get
            Return mManagementServer
        End Get
        Friend Set(value As String)
            mManagementServer = value
        End Set
    End Property

    Private mstrStartDomainRes As String
    Public Property StartDomainRes As String
        Get
            Return mstrStartDomainRes
        End Get
        Friend Set(value As String)
            mstrStartDomainRes = value
        End Set
    End Property

    Private mtsJavaCpuTime As TimeSpan = TimeSpan.Zero
    Public Property JavaCpuTime As TimeSpan
        Get
            Return mtsJavaCpuTime
        End Get
        Friend Set(value As TimeSpan)
            mtsJavaCpuTime = value
        End Set
    End Property

    Private mdecJavaMemoryUse As Decimal = 0

    ''' <summary>
    ''' JAVA进程使用的内存，单位：MB|Memory used by java process, unit: MB
    ''' </summary>
    ''' <returns></returns>
    Public Property JavaMemoryUse As Decimal
        Get
            Return mdecJavaMemoryUse
        End Get
        Friend Set(value As Decimal)
            mdecJavaMemoryUse = value
        End Set
    End Property

    Private mStopDomainRes As String
    Public Property StopDomainRes As String
        Get
            Return mStopDomainRes
        End Get
        Friend Set(value As String)
            mStopDomainRes = value
        End Set
    End Property

    Public Sub New(ServerName As String, ListenPort As Integer, Parent As WebLogicDomain)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
        Me.ServerName = ServerName
        Me.ListenPort = ListenPort
        Me.HomeDirPath = Me.Parent.HomeDirPath & Me.OsPathSep & "servers" & Me.OsPathSep & Me.ServerName
    End Sub

    Public ReadOnly Property SecurityDirPath() As String
        Get
            Return Me.HomeDirPath & Me.OsPathSep & "security"
        End Get
    End Property

    Public ReadOnly Property SecurityBootPath() As String
        Get
            Return Me.SecurityDirPath & Me.OsPathSep & "boot.properties"
        End Get
    End Property

    Private mintJavaPID As Integer = -1
    Public Property JavaPID As Integer
        Get
            Return mintJavaPID
        End Get
        Friend Set(value As Integer)
            mintJavaPID = value
        End Set
    End Property

    Private mdteJavaStartTime As DateTime = TEMP_DATE
    Public Property JavaStartTime As DateTime
        Get
            Return mdteJavaStartTime
        End Get
        Friend Set(value As DateTime)
            mdteJavaStartTime = value
        End Set
    End Property

    Private mintRunStatus As WebLogicDomain.EnmDomainRunStatus
    Public Property RunStatus As WebLogicDomain.EnmDomainRunStatus
        Get
            Return mintRunStatus
        End Get
        Friend Set(value As WebLogicDomain.EnmDomainRunStatus)
            mintRunStatus = value
        End Set
    End Property


    ''' <summary>
    ''' Refresh deployment status|刷新部署状态
    ''' </summary>
    ''' <returns></returns>
    Public Function RefDeployStatus() As String
        Dim LOG As New StruStepLog : LOG.SubName = "RefDeployStatus"
        Try
            If Me.mIsFolderExists(Me.HomeDirPath) = False Then
                Me.DeployStatus = WebLogicDomain.EnmDomainDeployStatus.NotCreate
            ElseIf Me.mIsFileExists(Me.startManagedWebLogicpath) = False Or Me.mIsFileExists(Me.stopManagedWebLogicpath) = False Then
                Me.DeployStatus = WebLogicDomain.EnmDomainDeployStatus.IncompleteCreation
            ElseIf Me.IsProdMode = False Then
                Me.DeployStatus = WebLogicDomain.EnmDomainDeployStatus.CreateDevMode
            ElseIf Me.mIsFileExists(Me.SecurityBootPath) = False Then
                Me.DeployStatus = WebLogicDomain.EnmDomainDeployStatus.CreateProdModeNotSecurity
            Else
                Me.DeployStatus = WebLogicDomain.EnmDomainDeployStatus.CreateProdMode
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mIsFolderExists(FolderPath As String) As Boolean
        Try
            Return Me.mPigFunc.IsFolderExists(FolderPath)
        Catch ex As Exception
            Me.SetSubErrInf("mIsFolderExists", ex)
            Return Nothing
        End Try
    End Function

    Private mintDeployStatus As WebLogicDomain.EnmDomainDeployStatus
    Public Property DeployStatus As WebLogicDomain.EnmDomainDeployStatus
        Get
            Return mintDeployStatus
        End Get
        Friend Set(value As WebLogicDomain.EnmDomainDeployStatus)
            mintDeployStatus = value
        End Set
    End Property

    Private Function mIsFileExists(FilePath As String) As Boolean
        Try
            Return Me.mPigFunc.IsFileExists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("mIsFileExists", ex)
            Return Nothing
        End Try
    End Function

    Public ReadOnly Property IsProdMode As Boolean
        Get
            Return Me.Parent.IsProdMode
        End Get
    End Property

    Public ReadOnly Property startManagedWebLogicPath() As String
        Get
            startManagedWebLogicPath = Me.Parent.HomeDirPath & Me.OsPathSep & "bin" & Me.OsPathSep & "startManagedWebLogic."
            If Me.IsWindows = True Then
                startManagedWebLogicPath &= "cmd"
            Else
                startManagedWebLogicPath &= "sh"
            End If
        End Get
    End Property

    Public ReadOnly Property stopManagedWebLogicPath() As String
        Get
            stopManagedWebLogicPath = Me.Parent.HomeDirPath & Me.OsPathSep & "bin" & Me.OsPathSep & "stopManagedWebLogic."
            If Me.IsWindows = True Then
                stopManagedWebLogicPath &= "cmd"
            Else
                stopManagedWebLogicPath &= "sh"
            End If
        End Get
    End Property

    Public ReadOnly Property LogDirPath() As String
        Get
            Return Me.HomeDirPath & Me.OsPathSep & "logs"
        End Get
    End Property

    Public ReadOnly Property ConsolePath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "Console.log"
        End Get
    End Property

    Public ReadOnly Property AccessLogTime() As Date
        Get
            Try
                Me.mPigFunc.GetFileUpdateTime(Me.AccessLogPath, AccessLogTime)
            Catch ex As Exception
                Me.SetSubErrInf("AccessLogTime", ex)
                Return TEMP_DATE
            End Try
        End Get
    End Property

    Public ReadOnly Property ConsoleLogTime() As Date
        Get
            Try
                Me.mPigFunc.GetFileUpdateTime(Me.ConsolePath, ConsoleLogTime)
            Catch ex As Exception
                Me.SetSubErrInf("ConsoleLogTime", ex)
                Return TEMP_DATE
            End Try
        End Get
    End Property

    Public ReadOnly Property DomainLogPath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "Domain" & Me.ListenPort.ToString & ".log"
        End Get
    End Property

    Public ReadOnly Property AccessLogPath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "access.log"
        End Get
    End Property

    ''' <summary>
    ''' Refresh running status|刷新运行状态
    ''' </summary>
    ''' <returns></returns>
    Public Function RefRunStatus() As String
        Dim LOG As New StruStepLog : LOG.SubName = "RefRunStatus"
        Try
            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.ExecutingWLST
                    'If Math.Abs(DateDiff(DateInterval.Second, Me.mCallWlstBeginTime, Now)) > Me.fParent.CallWlstTimeout Then
                    '    Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.ExecWLSTFail
                    '    Me.CallWlstRes = "WLST execution timeout."
                    'End If
                Case Else
                    Select Case Me.DeployStatus
                        Case WebLogicDomain.EnmDomainDeployStatus.CreateProdModeNotSecurity, WebLogicDomain.EnmDomainDeployStatus.IncompleteCreation, WebLogicDomain.EnmDomainDeployStatus.NotCreate
                            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.DeployNoReady
                        Case Else
                            Select Case Me.RunStatus
                                Case WebLogicDomain.EnmDomainRunStatus.Stopping
                                    If Math.Abs(DateDiff(DateInterval.Second, Me.mStopDomainBeginTime, Now)) > Me.Parent.fParent.StartOrStopTimeout Then
                                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.StopFail
                                        Me.StopDomainRes = "Stop server timeout."
                                    End If
                                Case Else
                                    Dim bolIsGetListenPortProcID As Boolean = False
                                    Dim strPIDMath As String = ""
                                    If Me.IsWindows = False Then
                                        Dim strCmd As String = "ps -ef|awk '{if($8==""/bin/sh"" && $9==""" & Replace(Me.Parent.startWebLogicPath, "\", "\\") & """) print $2}'"
                                        LOG.StepName = "CmdShell"
                                        'Console.WriteLine(strCmd)
                                        LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                                        If LOG.Ret <> "OK" Then
                                            LOG.AddStepNameInf(strCmd)
                                            Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.StepLogInf)
                                        Else
                                            For i = 0 To Me.mPigCmdApp.StandardOutputArray.Length - 1
                                                If strPIDMath <> "" Then
                                                    strPIDMath &= "||"
                                                End If
                                                strPIDMath &= "$3==""" & Me.mPigCmdApp.StandardOutputArray(i).ToString & """"
                                            Next
                                        End If
                                        If strPIDMath <> "" Then
                                            strCmd = "ps -ef|awk '{if(" & strPIDMath & ") print $0}'|grep "" \-Dweblogic.Name=" & Replace(Me.ServerName, "-", "\-"） & " ""|awk '{print $2}'"
                                            LOG.StepName = "CmdShell"
                                            'Console.WriteLine(strCmd)
                                            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
                                            'Console.WriteLine(Me.mPigCmdApp.StandardOutput)
                                            If LOG.Ret <> "OK" Then
                                                LOG.AddStepNameInf(strCmd)
                                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.StepLogInf)
                                                bolIsGetListenPortProcID = True
                                            Else
                                                Dim intPID As Integer = Me.mPigFunc.GECInt(Me.mPigCmdApp.StandardOutput)
                                                If intPID > 0 Then
                                                    Dim oPigProc As New PigProc(intPID)
                                                    If UCase(oPigProc.ProcessName) = "JAVA" Then
                                                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Running
                                                        Me.JavaPID = oPigProc.ProcessID
                                                        Me.JavaStartTime = oPigProc.StartTime
                                                        Me.JavaCpuTime = oPigProc.UserProcessorTime
                                                        Me.JavaMemoryUse = CDec(oPigProc.MemoryUse) / 1024 / 1024
                                                    Else
                                                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.ListenPortByOther
                                                    End If
                                                Else
                                                    bolIsGetListenPortProcID = True
                                                End If
                                            End If
                                        End If
                                    Else
                                        bolIsGetListenPortProcID = True
                                    End If
                                    If bolIsGetListenPortProcID = True Then
                                        Dim intPID As Integer
                                        LOG.StepName = "GetListenPortProcID.ListenPort"
                                        LOG.Ret = Me.mPigSysCmd.GetListenPortProcID(Me.ListenPort, intPID)
                                        If LOG.Ret <> "OK" Then Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
                                        If intPID >= 0 Then
                                            LOG.StepName = "GetPigProc.ListenPort"
                                            Dim oPigProc As PigProc = Me.mPigProcApp.GetPigProc(intPID)
                                            If Me.mPigProcApp.LastErr <> "" Then
                                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.mPigProcApp.LastErr)
                                            Else
                                                If UCase(oPigProc.ProcessName) = "JAVA" Then
                                                    Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Running
                                                    Me.JavaPID = oPigProc.ProcessID
                                                    Me.JavaStartTime = oPigProc.StartTime
                                                    Me.JavaCpuTime = oPigProc.UserProcessorTime
                                                    Me.JavaMemoryUse = CDec(oPigProc.MemoryUse) / 1024 / 1024
                                                Else
                                                    Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.ListenPortByOther
                                                End If
                                            End If
                                        Else
                                            Select Case Me.RunStatus
                                                Case WebLogicDomain.EnmDomainRunStatus.Starting, WebLogicDomain.EnmDomainRunStatus.StartPartReady
                                                    If Math.Abs(DateDiff(DateInterval.Second, Me.mStartDomainBeginTime, Now)) > Me.Parent.fParent.StartOrStopTimeout Then
                                                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.StartFail
                                                        Me.StartDomainRes = "Start server timeout."
                                                    End If
                                                Case WebLogicDomain.EnmDomainRunStatus.Stopping
                                                    If Math.Abs(DateDiff(DateInterval.Second, Me.mStopDomainBeginTime, Now)) > Me.Parent.fParent.StartOrStopTimeout Then
                                                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.StopFail
                                                        Me.StartDomainRes = "Stop server timeout."
                                                    End If
                                                Case Else
                                                    Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Stopped
                                            End Select
                                        End If
                                    End If
                            End Select
                    End Select
            End Select
            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.Running, WebLogicDomain.EnmDomainRunStatus.StartPartReady
                Case Else
                    Me.JavaPID = -1
                    Me.JavaStartTime = TEMP_DATE
                    Me.JavaCpuTime = TimeSpan.Zero
                    Me.JavaMemoryUse = 0
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Sub mPigCmdApp_AsyncRet_CmdShell_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles mPigCmdApp.AsyncRet_CmdShell_FullString
        With AsyncRet
            Select Case .AsyncThreadID
                Case Me.mStopDomainThreadID
                    Me.mStopDomainThreadID = -1
                    If .AsyncRet = "OK" And StandardError = "" Then
                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Stopped
                        Me.StopDomainRes = StandardOutput
                        RaiseEvent StopDomainSucc(StandardOutput)
                    Else
                        Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.StopFail
                        Me.StopDomainRes = .AsyncRet & Me.OsCrLf & StandardOutput & Me.OsCrLf & StandardError & Me.OsCrLf
                        RaiseEvent StopDomainFail(StandardError)
                    End If
            End Select
        End With
    End Sub

    ''' <summary>
    ''' Save as non interactive startup in production mode|生产模式下保存为不交互启动
    ''' </summary>
    ''' <param name="UserName"></param>
    ''' <param name="Password"></param>
    ''' <returns></returns>
    Public Function SaveSecurityBoot(UserName As String, Password As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "SaveSecurityBoot"
        Try
            LOG.StepName = "RefDeployStatus"
            LOG.Ret = Me.RefDeployStatus()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "RefConf"
            LOG.Ret = Me.Parent.RefConf()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.DeployStatus <> WebLogicDomain.EnmDomainDeployStatus.CreateProdModeNotSecurity Then Throw New Exception("The current status is " & Me.DeployStatus.ToString & ", SaveSecurityBoot is not allowed.")
            If Me.mIsFolderExists(Me.SecurityDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.SecurityDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.SecurityDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "OpenTextFile"
            Dim tsMain As TextStream = Me.mFS.OpenTextFile(Me.SecurityBootPath, PigFileSystem.IOMode.ForWriting, True)
            If Me.mFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.SecurityBootPath)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "WriteLine"
            With tsMain
                .WriteLine("username=" & UserName)
                .WriteLine("password=" & Password)
                .Close()
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RefAll() As String
        Dim strRet As String
        Try
            Me.Parent.RefConf()
            strRet = Me.RefDeployStatus()
            If strRet <> "OK" Then Throw New Exception(strRet)
            strRet = Me.RefRunStatus()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("RefAll", ex)
        End Try
    End Function

    Private ReadOnly Property mIsDeployReady As Boolean
        Get
            Select Case Me.DeployStatus
                Case WebLogicDomain.EnmDomainDeployStatus.CreateDevMode, WebLogicDomain.EnmDomainDeployStatus.CreateProdMode
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsRunBusy As Boolean
        Get
            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.ExecutingWLST, WebLogicDomain.EnmDomainRunStatus.Starting, WebLogicDomain.EnmDomainRunStatus.Stopping
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public Function StartServer() As String
        Dim LOG As New StruStepLog : LOG.SubName = "StartDomain"
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot start the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot start the domain.")

            If Me.mIsFileExists(Me.startManagedWebLogicPath) = False Then
                LOG.AddStepNameInf(Me.startManagedWebLogicPath)
                Throw New Exception("File not found.")
            End If

            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.Running
                    Throw New Exception("Server instance is already running")
                Case WebLogicDomain.EnmDomainRunStatus.StartPartReady
                    Throw New Exception("Server instance is successfully started part")
            End Select

            If Me.mIsFolderExists(Me.LogDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            If Me.mIsFileExists(Me.ConsolePath) = True Then
                Dim strConsolePath As String = Me.ConsolePath & "." & Me.mPigFunc.GetFmtDateTime(Now, "yyyyMMddHHmmss")
                LOG.StepName = "MoveFile"
                LOG.Ret = Me.mPigFunc.MoveFile(Me.ConsolePath, strConsolePath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.ConsolePath)
                    LOG.AddStepNameInf(strConsolePath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "AsyncCmdShell"
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "call " & Me.startManagedWebLogicPath & " " & Me.ServerName & " " & Me.Parent.RootUrl & " > " & Me.ConsolePath
            Else
                strCmd = "nohup " & Me.startManagedWebLogicPath & " " & Me.ServerName & " " & Me.Parent.RootUrl & " > " & Me.ConsolePath & " &"
            End If
            Me.Parent.fParent.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Starting
            Me.mStartDomainBeginTime = Now
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mStartDomainThreadID)
            Me.StartDomainRes = ""
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function StopServer() As String
        Dim LOG As New StruStepLog : LOG.SubName = "StopServer"
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot stop the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot stop the domain.")

            If Me.mIsFileExists(Me.stopManagedWebLogicPath) = False Then
                LOG.AddStepNameInf(Me.stopManagedWebLogicPath)
                Throw New Exception("File not found.")
            End If

            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.Running, WebLogicDomain.EnmDomainRunStatus.StartPartReady
                Case Else
                    Throw New Exception("Server instance is not running or successfully started part")
            End Select

            If Me.mIsFolderExists(Me.LogDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "call " & Me.stopManagedWebLogicPath & " " & Me.ServerName
            Else
                strCmd = Me.stopManagedWebLogicPath & " " & Me.ServerName
            End If
            Me.Parent.fParent.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Stopping
            Me.mStopDomainBeginTime = Now
            If Me.IsWindows = True Then
                LOG.StepName = "AsyncCmdShell"
                LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mStopDomainThreadID)
            Else
                Dim oPigNohup As New PigNohup(strCmd)
                LOG.StepName = "Run"
                LOG.Ret = oPigNohup.Run()
            End If
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.StopDomainRes = ""
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function HardStopServer() As String
        Dim LOG As New StruStepLog : LOG.SubName = "HardStopServer"
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot stop the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot stop the domain.")


            Select Case Me.RunStatus
                Case WebLogicDomain.EnmDomainRunStatus.Running, WebLogicDomain.EnmDomainRunStatus.StartPartReady
                Case Else
                    Throw New Exception("Domain instance is not running or successfully started part")
            End Select

            If Me.mIsFolderExists(Me.LogDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Stopping
            Dim oPigProcs As PigProcs = Me.mPigCmdApp.GetSubProcs(Me.JavaPID)
            If oPigProcs IsNot Nothing Then
                For Each oPigProc As PigProc In oPigProcs
                    Me.mKillProc(oPigProc.ProcessID)
                    Me.mPigFunc.Delay(200)
                Next
            End If
            LOG.StepName = "mKillProc"
            LOG.Ret = Me.mKillProc(Me.JavaPID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.mPigFunc.Delay(500)
            Me.StopDomainRes = ""
            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.Stopped
            Return "OK"
        Catch ex As Exception
            Me.RunStatus = WebLogicDomain.EnmDomainRunStatus.StartFail
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mKillProc(PID As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mKillProc"
        Try
            Dim strCmd As String = ""
            If Me.IsWindows = True Then
                strCmd = "taskkill /pid " & PID.ToString & " /f"
            Else
                strCmd = "kill -9 " & PID.ToString
            End If
            Dim intOutPID As Integer = -1
            LOG.StepName = "AsyncCmdShell"
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, intOutPID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
