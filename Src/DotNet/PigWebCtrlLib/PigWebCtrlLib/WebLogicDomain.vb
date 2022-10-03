'**********************************
'* Name: WebLogicDomain
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic domain
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.23
'* Create Time: 31/1/2022
'*1.1  5/2/2022   Add CheckDomain 
'*1.2  5/3/2022   Modify New
'*1.3  23/5/2022  Add CreateDomain 
'*1.4  26/5/2022  Add EnmDomainDeployStatus,EnmDomainRunStatus,EnmDomainCtrlStatus
'*1.5  27/5/2022  Add CreateDomain 
'*1.6  31/5/2022  Modify CreateDomain
'*1.7  1/6/2022  Add StartDomain,StopDomain, modify LogDirPath
'*1.8  2/6/2022  Modify StopDomain,StartDomain
'*1.9  4/6/2022  Rename CreateDomainRes to CallWlstRes, CallWlstPyPath to CallWlstPyPath, modify CreateDomain,RefRunStatus, add mCallWlstPyOpenTextFile,mWlstPublicCheck,RefAll
'*1.10 5/6/2022  Add CallWlstSucc,CallWlstFail, modify mPigCmdApp_AsyncRet_CmdShell_FullString,StartDomain
'*1.11 12/6/2022  Rename RefAll to RefAll, modify mWlstCallMain
'*1.12 13/6/2022  Add JavaPID
'*1.13 14/6/2022  Add JavaMemoryUse,JavaStartTime
'*1.14 14/6/2022  Modify AdminPort
'*1.15 15/6/2022  Add ConnectionFilterImpl, modify RefConf,RefRunStatus
'*1.15 7/7/2022   Add AdminServerLogPath,DomainLogPath,AdminServerLogPath,AccessLogPath
'*1.16 16/7/2022  Modify mWlstCallMain,CreateDomain,EnmWlstCallCmd, add UpdateCheck
'*1.17 17/7/2022  Modify AdminPort,IsAdminPortEnable
'*1.18 21/7/2022  Modify AdminPort,IsAdminPortEnable
'*1.19  26/7/2022 Modify Imports
'*1.20  29/7/2022 Modify Imports
'*1.21  1/8/2022  Add HardStopDomain
'*1.22  2/8/2022  Modify mWlstCallMain,CreateDomain
'*1.23  13/8/2022 Modify mWlstCallMain,CreateDomain, and add AsyncCreateDomain,SetT3Deny
'*1.25  28/9/2022 Add ConsoleLogTime,AccessLogTime,GetConsoleLogHasRunMode
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
Imports PigObjFsLib


Public Class WebLogicDomain
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.25.10"

    Private WithEvents mPigCmdApp As New PigCmdApp
    Private mPigSysCmd As New PigSysCmd
    Private mFS As New FileSystemObject
    Private mPigFunc As New PigFunc
    Private mPigProcApp As New PigProcApp

    Private mUpdateCheck As New UpdateCheck
    Public ReadOnly Property LastUpdateTime() As DateTime
        Get
            Return mUpdateCheck.LastUpdateTime
        End Get
    End Property

    Public ReadOnly Property IsUpdate(PropertyName As String) As Boolean
        Get
            Return mUpdateCheck.IsUpdated(PropertyName)
        End Get
    End Property

    Public ReadOnly Property HasUpdated() As Boolean
        Get
            Return mUpdateCheck.HasUpdated
        End Get
    End Property

    Public Sub UpdateCheckClear()
        mUpdateCheck.Clear()
    End Sub

    Public Event CallWlstSucc(WlstCallCmd As EnmWlstCallCmd, ResInf As String)

    Public Event CallWlstFail(WlstCallCmd As EnmWlstCallCmd, ErrInf As String)

    Public Event StopDomainSucc(ResInf As String)

    Public Event StopDomainFail(ErrInf As String)

    ''' <summary>
    ''' WLST调用命令|WLST call command
    ''' </summary>
    Public Enum EnmWlstCallCmd
        ''' <summary>
        ''' 未知|unknown
        ''' </summary>
        Unknow = 0
        ''' <summary>
        ''' 创建域|Create Domain
        ''' </summary>
        CreateDomain = 1
        ''' <summary>
        ''' 连接控制台|Connect console
        ''' </summary>
        Connect = 2
        ''' <summary>
        ''' 修改配置|Modify configuration
        ''' </summary>
        Edit = 3
        ''' <summary>
        ''' 设置管理端口|Set management port
        ''' </summary>
        SetAdminPort = 5
        ''' <summary>
        ''' 设置T3拒绝|Set T3 deny
        ''' </summary>
        SetT3Deny = 6
    End Enum

    ''' <summary>
    ''' 域部署状态|Domain deployment status
    ''' </summary>
    Public Enum EnmDomainDeployStatus
        ''' <summary>
        ''' 部署为生产模式，但未配置SecurityBoot，这样启动时需要交互输入密码|The deployment is in production mode, but securityboot is not configured, so the password needs to be entered interactively during startup
        ''' </summary>
        CreateProdModeNotSecurity = -2
        ''' <summary>
        ''' 创建域不完整|Incomplete domain creation
        ''' </summary>
        IncompleteCreation = -1
        ''' <summary>
        ''' 未创建域|Domain not created
        ''' </summary>
        NotCreate = 0
        ''' <summary>
        ''' 域创建为开发模式|Create domain as development mode
        ''' </summary>
        CreateDevMode = 2
        ''' <summary>
        ''' 域创建为生产模式并设置了SecurityBoot，这样启动时不需要交互。|The domain is created in production mode and securityboot is set, so no interaction is required during startup.
        ''' </summary>
        CreateProdMode = 3
    End Enum



    ''' <summary>
    ''' 域运行状态|Domain running status
    ''' </summary>
    Public Enum EnmDomainRunStatus
        ''' <summary>
        ''' 侦听端口被其他进程占用|Listening port is occupied by other processes
        ''' </summary>
        ListenPortByOther = -5
        ''' <summary>
        ''' 部署未就绪|Deployment not ready
        ''' </summary>
        DeployNoReady = -4
        ''' <summary>
        ''' 停止域失败|Failed to stop domain
        ''' </summary>
        StopFail = -3
        ''' <summary>
        ''' 启动域失败|Failed to start domain
        ''' </summary>
        StartFail = -2
        ''' <summary>
        ''' 执行WLST失败|WLST execution failed
        ''' </summary>
        ExecWLSTFail = -1
        ''' <summary>
        ''' 已停止|Stopped
        ''' </summary>
        Stopped = 0
        ''' <summary>
        ''' 正在启动域|Starting domain
        ''' </summary>
        Starting = 1
        ''' <summary>
        ''' 正在停止域|Stopping domain
        ''' </summary>
        Stopping = 2
        ''' <summary>
        ''' 运行中|Starting domain
        ''' </summary>
        Running = 3
        ''' <summary>
        ''' 正在执行WLST|Executing WLST
        ''' </summary>
        ExecutingWLST = 4
        ''' <summary>
        ''' 执行WLST成功|WLST executed successfully
        ''' </summary>
        ExecWLSTOK = 5
    End Enum


    Public ReadOnly Property HomeDirPath As String
    Public Property AdminUserName As String
    Public Property AdminUserPassword As String
    Friend Property CallWlstPyPath As String

    Private Property mCallWlstThreadID As Integer
    Private Property mCallWlstBeginTime
    Private Property mStartDomainThreadID As Integer
    Private Property mStartDomainBeginTime

    Private Property mStopDomainThreadID As Integer
    Private Property mStopDomainBeginTime
    'Private Property mGetConsoleLogHasRunModeThreadID As Integer

    Private Property mRefRunStatusThreadID As Integer

    Private Property mConfFileUpdtime As Date

    Private mstrStartDomainRes As String


    Public Property StartDomainRes As String
        Get
            Return mstrStartDomainRes
        End Get
        Friend Set(value As String)
            mstrStartDomainRes = value
        End Set
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

    Private mtsJavaCpuTime As TimeSpan = TimeSpan.Zero

    Public Property JavaCpuTime As TimeSpan
        Get
            Return mtsJavaCpuTime
        End Get
        Friend Set(value As TimeSpan)
            mtsJavaCpuTime = value
        End Set
    End Property

    Private mstrStopDomainRes As String

    Public Property StopDomainRes As String
        Get
            Return mstrStopDomainRes
        End Get
        Friend Set(value As String)
            mstrStopDomainRes = value
        End Set
    End Property

    Private mstrCallWlstRes As String

    Public Property CallWlstRes As String
        Get
            Return mstrCallWlstRes
        End Get
        Friend Set(value As String)
            mstrCallWlstRes = value
        End Set
    End Property


    Private mintDeployStatus As EnmDomainDeployStatus
    Public Property DeployStatus As EnmDomainDeployStatus
        Get
            Return mintDeployStatus
        End Get
        Friend Set(value As EnmDomainDeployStatus)
            mintDeployStatus = value
        End Set
    End Property

    Private mintRunStatus As EnmDomainRunStatus
    Public Property RunStatus As EnmDomainRunStatus
        Get
            Return mintRunStatus
        End Get
        Friend Set(value As EnmDomainRunStatus)
            mintRunStatus = value
        End Set
    End Property

    Private mstrDomainVersion As String
    Public Property DomainVersion As String
        Get
            Return mstrDomainVersion
        End Get
        Friend Set(value As String)
            mstrDomainVersion = value
        End Set
    End Property

    Private mbolIsProdMode As Boolean
    Public Property IsProdMode As Boolean
        Get
            Return mbolIsProdMode
        End Get
        Friend Set(value As Boolean)
            mbolIsProdMode = value
        End Set
    End Property

    Public Property IsIIopEnable As Boolean

    Private mbolHasAdminServer As Boolean
    Public Property HasAdminServer As Boolean
        Get
            Return mbolHasAdminServer
        End Get
        Friend Set(value As Boolean)
            mbolHasAdminServer = value
        End Set
    End Property
    Private mIsAdminPortEnable As Boolean
    Public Property IsAdminPortEnable As Boolean
        Get
            Return mIsAdminPortEnable
        End Get
        Set(value As Boolean)
            mIsAdminPortEnable = value
        End Set
    End Property
    'Private mdteLastOperationTime As DateTime
    'Public Property LastOperationTime As DateTime
    '    Get
    '        Return mdteLastOperationTime
    '    End Get
    '    Friend Set(value As DateTime)
    '        mdteLastOperationTime = value
    '    End Set
    'End Property


    Public Property ListenPort As Integer

    Private mAdminPort As Integer
    Public Property AdminPort As Integer
        Get
            Return mAdminPort
        End Get
        Set(value As Integer)
            Try
                Select Case value
                    Case 1 To 65535
                        If value = Me.ListenPort Then Throw New Exception("The administrator port cannot be the same as the listening port")
                        mAdminPort = value
                    Case 0
                        mAdminPort = 0
                    Case Else
                        Throw New Exception("Invalid port")
                End Select
            Catch ex As Exception
                mAdminPort = 0
                Me.SetSubErrInf("AdminPort.Set", ex)
            End Try
        End Set
    End Property
    Private mstrConnectionFilterImpl As String
    Public Property ConnectionFilterImpl As String
        Get
            Return mstrConnectionFilterImpl
        End Get
        Friend Set(value As String)
            mstrConnectionFilterImpl = value
        End Set
    End Property

    Private mstrDomainName As String
    Public Property DomainName As String
        Get
            Return mstrDomainName
        End Get
        Friend Set(value As String)
            mstrDomainName = value
        End Set
    End Property

    Friend Property fParent As WebLogicApp

    Public Sub New(HomeDirPath As String, Parent As WebLogicApp)
        MyBase.New(CLS_VERSION)
        Me.HomeDirPath = HomeDirPath
        Me.fParent = Parent
    End Sub

    Public ReadOnly Property ConfPath() As String
        Get
            Return Me.HomeDirPath & Me.OsPathSep & "config" & Me.OsPathSep & "config.xml"
        End Get
    End Property

    Public ReadOnly Property startWebLogicPath() As String
        Get
            startWebLogicPath = Me.HomeDirPath & Me.OsPathSep & "bin" & Me.OsPathSep & "startWebLogic."
            If Me.IsWindows = True Then
                startWebLogicPath &= "cmd"
            Else
                startWebLogicPath &= "sh"
            End If

        End Get
    End Property

    Public ReadOnly Property stopWebLogicPath() As String
        Get
            stopWebLogicPath = Me.HomeDirPath & Me.OsPathSep & "bin" & Me.OsPathSep & "stopWebLogic."
            If Me.IsWindows = True Then
                stopWebLogicPath &= "cmd"
            Else
                stopWebLogicPath &= "sh"
            End If

        End Get
    End Property

    Public ReadOnly Property LogDirPath() As String
        Get
            Return Me.HomeDirPath & Me.OsPathSep & "servers" & Me.OsPathSep & "AdminServer" & Me.OsPathSep & "logs"
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
                Return Date.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property ConsoleLogTime() As Date
        Get
            Try
                Me.mPigFunc.GetFileUpdateTime(Me.ConsolePath, ConsoleLogTime)
            Catch ex As Exception
                Me.SetSubErrInf("ConsoleLogTime", ex)
                Return Date.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property DomainLogPath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "Domain" & Me.ListenPort.ToString & ".log"
        End Get
    End Property

    Public ReadOnly Property AdminServerLogPath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "AdminServer.log"
        End Get
    End Property

    Public ReadOnly Property AccessLogPath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "access.log"
        End Get
    End Property

    Public ReadOnly Property SecurityDirPath() As String
        Get
            Return Me.HomeDirPath & Me.OsPathSep & "servers" & Me.OsPathSep & "AdminServer" & Me.OsPathSep & "security"
        End Get
    End Property

    Public ReadOnly Property SecurityBootPath() As String
        Get
            Return Me.SecurityDirPath & Me.OsPathSep & "boot.properties"
        End Get
    End Property

    Public Function SaveSecurityBoot(UserName As String, Password As String) As String
        Dim LOG As New PigStepLog("SaveSecurityBoot")
        Try
            LOG.StepName = "RefDeployStatus"
            LOG.Ret = Me.RefDeployStatus()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "RefConf"
            LOG.Ret = Me.RefConf()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.DeployStatus <> EnmDomainDeployStatus.CreateProdModeNotSecurity Then Throw New Exception("The current status is " & Me.DeployStatus.ToString & ", SaveSecurityBoot is not allowed.")
            If Me.mIsFolderExists(Me.SecurityDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.SecurityDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.SecurityDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "OpenTextFile"
            Dim tsMain As TextStream = Me.mFS.OpenTextFile(Me.SecurityBootPath, FileSystemObject.IOMode.ForWriting, True)
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

    Private Function mIsFolderExists(FolderPath As String) As Boolean
        Try
            Return Me.mPigFunc.IsFolderExists(FolderPath)
        Catch ex As Exception
            Me.SetSubErrInf("mIsFolderExists", ex)
            Return Nothing
        End Try
    End Function

    Private Function mIsFileExists(FilePath As String) As Boolean
        Try
            Return Me.mPigFunc.IsFileExists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("mIsFileExists", ex)
            Return Nothing
        End Try
    End Function

    Public Function RefAll() As String
        Dim strRet As String
        Try
            Me.RefConf()
            strRet = Me.RefDeployStatus()
            If strRet <> "OK" Then Throw New Exception(strRet)
            strRet = Me.RefRunStatus()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("RefAll", ex)
        End Try
    End Function

    Private ReadOnly Property mIsRunBusy As Boolean
        Get
            Select Case Me.RunStatus
                Case EnmDomainRunStatus.ExecutingWLST, EnmDomainRunStatus.Starting, EnmDomainRunStatus.Stopping
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Private ReadOnly Property mIsDeployReady As Boolean
        Get
            Select Case Me.DeployStatus
                Case EnmDomainDeployStatus.CreateDevMode, EnmDomainDeployStatus.CreateProdMode
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    Public Function StartDomain() As String
        Dim LOG As New PigStepLog("StartDomain")
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot start the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot start the domain.")

            If Me.mIsFileExists(Me.startWebLogicPath) = False Then
                LOG.AddStepNameInf(Me.startWebLogicPath)
                Throw New Exception("File not found.")
            End If

            If Me.RunStatus = EnmDomainRunStatus.Running Then Throw New Exception("Domain instance is already running")

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
                strCmd = "call " & Me.startWebLogicPath & " > " & Me.ConsolePath
            Else
                strCmd = "nohup " & Me.startWebLogicPath & " > " & Me.ConsolePath & " &"
            End If
            Me.fParent.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = EnmDomainRunStatus.Starting
            Me.mStartDomainBeginTime = Now
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mStartDomainThreadID)
            Me.StartDomainRes = ""
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mKillProc(PID As Integer) As String
        Dim LOG As New PigStepLog("mKillProc")
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

    Public Function HardStopDomain() As String
        Dim LOG As New PigStepLog("HardStopDomain")
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot stop the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot stop the domain.")

            If Me.mIsFileExists(Me.stopWebLogicPath) = False Then
                LOG.AddStepNameInf(Me.stopWebLogicPath)
                Throw New Exception("File not found.")
            End If

            If Me.RunStatus <> EnmDomainRunStatus.Running Then Throw New Exception("Domain instance is not running")

            If Me.mIsFolderExists(Me.LogDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Me.RunStatus = EnmDomainRunStatus.Stopping
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
            Me.RunStatus = EnmDomainRunStatus.Stopped
            Return "OK"
        Catch ex As Exception
            Me.RunStatus = EnmDomainRunStatus.StartFail
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function StopDomain() As String
        Dim LOG As New PigStepLog("StopDomain")
        Try
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check Status"
            If Me.mIsDeployReady = False Then Throw New Exception("The current deployment state(" & Me.DeployStatus.ToString & ") cannot stop the domain.")
            If Me.mIsRunBusy = True Then Throw New Exception("The current run state(" & Me.RunStatus.ToString & ") cannot stop the domain.")

            If Me.mIsFileExists(Me.stopWebLogicPath) = False Then
                LOG.AddStepNameInf(Me.stopWebLogicPath)
                Throw New Exception("File not found.")
            End If

            If Me.RunStatus <> EnmDomainRunStatus.Running Then Throw New Exception("Domain instance is not running")

            If Me.mIsFolderExists(Me.LogDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "AsyncCmdShell"
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "call " & Me.stopWebLogicPath
            Else
                strCmd = "nohup " & Me.stopWebLogicPath
            End If
            Me.fParent.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = EnmDomainRunStatus.Stopping
            Me.mStopDomainBeginTime = Now
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mStopDomainThreadID)
            Me.StopDomainRes = ""
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private mbolIsCreateDomainOK As Boolean
    Public Property IsCreateDomainOK As Boolean
        Get
            Return mbolIsCreateDomainOK
        End Get
        Friend Set(value As Boolean)
            mbolIsCreateDomainOK = value
        End Set
    End Property



    Private mbolIsConnectOK As Boolean
    Public Property IsConnectOK As Boolean
        Get
            Return mbolIsConnectOK
        End Get
        Friend Set(value As Boolean)
            mbolIsConnectOK = value
        End Set
    End Property

    Private mintWlstCallCmd As EnmWlstCallCmd
    Public Property WlstCallCmd As EnmWlstCallCmd
        Get
            Return mintWlstCallCmd
        End Get
        Friend Set(value As EnmWlstCallCmd)
            mintWlstCallCmd = value
        End Set
    End Property

    Private Function mWlstCallMain(WlstCallCmd As EnmWlstCallCmd, Optional ListenPort As Integer = 0, Optional AdminPort As Integer = 0, Optional IsDisableIIOP As Boolean = True, Optional IsAsync As Boolean = False, Optional ByRef CmdRes As String = "", Optional IsEnableAdminPort As Boolean = False) As String
        Dim LOG As New PigStepLog("mWlstCallMain")
        Try
            LOG.StepName = "Check CallWlst"
            If Me.mCallWlstThreadID > 0 Then Throw New Exception("Busy")
            LOG.StepName = "Check AdminUser"
            If Me.AdminUserName = "" Or Me.AdminUserPassword = "" Then Throw New Exception("Administrator user or password not set.")
            LOG.StepName = "Check Wls.Jar"
            If Me.mIsFileExists(Me.fParent.WlsJarPath) = False Then Throw New Exception(Me.fParent.WlsJarPath & " not found.")
            LOG.StepName = "RefAll"
            LOG.Ret = Me.RefAll()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.WlstCallCmd = WlstCallCmd
            LOG.StepName = "Check WlstCallCmd"
            Select Case WlstCallCmd
                Case EnmWlstCallCmd.Connect, EnmWlstCallCmd.Edit, EnmWlstCallCmd.SetT3Deny, EnmWlstCallCmd.SetAdminPort
                    If Me.RunStatus <> EnmDomainRunStatus.Running Then Throw New Exception("The current run state[" & Me.RunStatus.ToString & "] cannot do " & WlstCallCmd.ToString)
                Case EnmWlstCallCmd.CreateDomain
                    If Me.mIsDeployReady = True Then Throw New Exception("The current deployment state[" & Me.DeployStatus.ToString & "] cannot create a domain.")
                Case Else
                    Me.WlstCallCmd = EnmWlstCallCmd.Unknow
                    Throw New Exception("Invalid WlstCallCmd " & WlstCallCmd.ToString)
            End Select
            LOG.StepName = "WlstCallCmd is " & Me.WlstCallCmd.ToString
            With Me
                Dim tsMain As TextStream = Nothing
                LOG.StepName = "Set CallWlstPyPath"
                .CallWlstPyPath = Me.fParent.WorkTmpDirPath & Me.OsPathSep & Me.mPigFunc.GetPKeyValue(Me.DomainName, False) ' & ".py"
                LOG.StepName = "OpenTextFile"
                tsMain = Me.mFS.OpenTextFile(.CallWlstPyPath, FileSystemObject.IOMode.ForWriting, True)
                If Me.mFS.LastErr <> "" Then
                    LOG.AddStepNameInf(.CallWlstPyPath)
                    Throw New Exception(Me.mFS.LastErr)
                ElseIf tsMain Is Nothing Then
                    LOG.AddStepNameInf(.CallWlstPyPath)
                    Throw New Exception("tsMain Is Nothing")
                End If
                Dim strLine As String
                With tsMain
                    Select Case Me.WlstCallCmd
                        Case EnmWlstCallCmd.SetAdminPort
                            If Me.IsAdminPortEnable = True Then Throw New Exception("Can not Enable AdminPort.")
                            .WriteLine("connect('" & Me.AdminUserName & "','" & Me.AdminUserPassword & "','t3://localhost:" & Me.ListenPort & "')")
                            .WriteLine("edit()")
                            .WriteLine("startEdit()")
                            .WriteLine("cd('/')")
                            If IsEnableAdminPort = True Then
                                .WriteLine("cmo.setAdministrationPort(" & AdminPort & ")")
                                .WriteLine("cmo.setAdministrationPortEnabled(true)")
                            Else
                                .WriteLine("cmo.setAdministrationPortEnabled(false)")
                            End If
                            .WriteLine("save()")
                            .WriteLine("activate()")
                        Case EnmWlstCallCmd.SetT3Deny
                            If Me.IsAdminPortEnable = True Then Throw New Exception("Can not Enable AdminPort.")
                            .WriteLine("connect('" & Me.AdminUserName & "','" & Me.AdminUserPassword & "','t3://localhost:" & Me.ListenPort & "')")
                            .WriteLine("edit()")
                            .WriteLine("startEdit()")
                            .WriteLine("cd('/SecurityConfiguration/" + Me.DomainName + "')")
                            .WriteLine("cmo.setConnectionLoggerEnabled(true)")
                            .WriteLine("cmo.setConnectionFilter('weblogic.security.net.ConnectionFilterImpl')")
                            .WriteLine("set ('ConnectionFilterRules',jarray.array([String('127.0.0.1 127.0.0.1 " & Me.ListenPort & " allow t3 t3s'),String('0.0.0.0/0 * * deny t3 t3s')], String))")
                            .WriteLine("save()")
                            .WriteLine("activate()")
                        Case EnmWlstCallCmd.CreateDomain
                            If ListenPort <= 0 Then Throw New Exception("Invalid ListenPort is " & ListenPort.ToString)
                            Me.ListenPort = ListenPort
                            LOG.StepName = "WriteLine"
                            strLine = "readTemplate('" & Me.fParent.WlsJarPath & "')"
                            If Me.IsWindows = True Then
                                strLine = Replace(strLine, Me.OsPathSep, "//")
                            End If
                            .WriteLine(strLine)
                            .WriteLine("cd('Servers/AdminServer')")
                            .WriteLine("set('ListenAddress','')")
                            .WriteLine("set('ListenPort', " & Me.ListenPort & ")")
                            If IsDisableIIOP = True Then
                                .WriteLine("set('iiop','false')")
                            End If
                            .WriteLine("cd('/')")
                            If AdminPort > 0 Then
                                .WriteLine("cmo.setAdministrationPort(" & AdminPort.ToString & ")")
                                .WriteLine("cmo.setAdministrationPortEnabled(true)")
                            End If
                            .WriteLine("cd('Security/base_domain/User/weblogic')")
                            .WriteLine("set('Name','" & Me.AdminUserName & "')")
                            .WriteLine("cmo.setPassword('" & Me.AdminUserPassword & "')")
                            .WriteLine("setOption('OverwriteDomain','true')")
                            .WriteLine("setOption('ServerStartMode','prod')")
                            strLine = "writeDomain('" & Me.HomeDirPath & "')"
                            If Me.IsWindows = True Then
                                strLine = Replace(strLine, Me.OsPathSep, "//")
                            End If
                            .WriteLine(strLine)
                            .WriteLine("closeTemplate()")
                        Case EnmWlstCallCmd.Connect
                            Dim intPort As Integer
                            If Me.IsAdminPortEnable = True Then
                                intPort = Me.AdminPort
                            Else
                                intPort = Me.ListenPort
                            End If
                            .WriteLine("connect('" & Me.AdminUserName & "','" & Me.AdminUserPassword & "','t3://localhost:" & intPort.ToString & "')")
                    End Select
                    .WriteLine("exit()")
                    LOG.StepName = "Close"
                    .Close()
                End With
            End With
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "call " & Me.fParent.WlstPath & " " & Me.CallWlstPyPath
            Else
                strCmd = Me.fParent.WlstPath & " " & Me.CallWlstPyPath
            End If
            Me.fParent.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = EnmDomainRunStatus.ExecutingWLST
            Me.mCallWlstBeginTime = Now
            Me.CallWlstRes = ""
            If IsAsync = True Then
                LOG.StepName = "AsyncCmdShell"
                LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mCallWlstThreadID)
                If LOG.Ret <> "OK" Then
                    If Me.IsDebug = True Then LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
            Else
                LOG.StepName = "CmdShell"
                LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd)
                CmdRes = Me.mPigCmdApp.StandardOutput & Me.OsCrLf & Me.mPigCmdApp.StandardError
                If LOG.Ret <> "OK" Then
                    If Me.IsDebug = True Then LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                Me.RunStatus = EnmDomainRunStatus.ExecWLSTOK
            End If
            Return "OK"
        Catch ex As Exception
            Me.RunStatus = EnmDomainRunStatus.ExecWLSTFail
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Sub ClearAsyncRes()
        With Me
            .CallWlstRes = ""
            .mCallWlstThreadID = -1
            .StartDomainRes = ""
            .mStartDomainThreadID = -1
            .StopDomainRes = ""
            .mStopDomainThreadID = -1
            .mRefRunStatusThreadID = -1
        End With
    End Sub

    ''' <summary>
    ''' 创建域|Create domain
    ''' </summary>
    ''' <param name="ListenPort">Listening port</param>
    ''' <returns></returns>
    Public Function CreateDomain(ByRef CmdRes As String, ListenPort As Integer) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort,,,, CmdRes)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function AsyncCreateDomain(ListenPort As Integer) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort,,, True)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 创建域|Create domain
    ''' </summary>
    ''' <param name="ListenPort">侦听端口|Listening port</param>
    ''' <param name="AdminPort">管理端口|Administrator port</param>
    ''' <returns></returns>
    Public Function CreateDomain(ByRef CmdRes As String, ListenPort As Integer, AdminPort As Integer) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort, AdminPort,,, CmdRes)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function AsyncCreateDomain(ListenPort As Integer, AdminPort As Integer) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort, AdminPort,, True)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 创建域|Create domain
    ''' </summary>
    ''' <param name="ListenPort">侦听端口|Listening port</param>
    ''' <param name="AdminPort">管理端口|Administrator port</param>
    ''' <param name="IsDisableIIOP">是禁用IIOP|Is Disable IIOP</param>
    ''' <returns></returns>
    Public Function CreateDomain(ByRef CmdRes As String, ListenPort As Integer, AdminPort As Integer, IsDisableIIOP As Boolean) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort, AdminPort, IsDisableIIOP,, CmdRes)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function AsyncCreateDomain(ListenPort As Integer, AdminPort As Integer, IsDisableIIOP As Boolean) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.CreateDomain, ListenPort, AdminPort, IsDisableIIOP, True)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function Connect() As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.Connect)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function


    Public Function RefConf() As String
        Dim LOG As New PigStepLog("RefConf")
        Try
            If Me.mIsFileExists(Me.ConfPath) = False Then Throw New Exception(Me.ConfPath & " not found.")
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(Me.ConfPath)
            If oPigFile.UpdateTime <> Me.mConfFileUpdtime Then
                Dim oPigXml As New PigXml(False)
                With oPigXml
                    LOG.StepName = "InitXmlDocument"
                    LOG.Ret = .InitXmlDocument(Me.ConfPath)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(Me.ConfPath)
                        Throw New Exception(LOG.Ret)
                    End If
                End With
                With Me
                    .DomainName = oPigXml.GetXmlDocText("domain.name")
                    .DomainVersion = oPigXml.GetXmlDocText("domain.domain-version")
                    .IsProdMode = oPigXml.XmlDocGetBool("domain.production-mode-enabled")
                    If oPigXml.GetXmlDocText("domain.server.name") = "AdminServer" Then
                        .HasAdminServer = True
                    Else
                        .HasAdminServer = False
                    End If
                    .ListenPort = oPigXml.XmlDocGetInt("domain.server.listen-port")
                    .IsIIopEnable = oPigXml.XmlDocGetBoolEmpTrue("domain.server.iiop-enabled")
                    .IsAdminPortEnable = oPigXml.XmlDocGetBool("domain.administration-port-enabled")
                    .AdminPort = oPigXml.XmlDocGetInt("domain.administration-port")
                    If oPigXml.XmlDocGetStr("domain.security-configuration.connection-filter") = "weblogic.security.net.ConnectionFilterImpl" Then
                        Dim intSkipTimes As Integer = 0
                        Me.ConnectionFilterImpl = ""
                        Do While True
                            Dim strLine As String = oPigXml.GetXmlDocText("domain.security-configuration.connection-filter-rule", intSkipTimes)
                            If strLine = "" Then Exit Do
                            Me.ConnectionFilterImpl &= strLine & Me.OsCrLf
                            intSkipTimes += 1
                        Loop
                    End If
                    .mConfFileUpdtime = oPigFile.UpdateTime
                    Me.mUpdateCheck.Clear()
                End With
            End If
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function RefDeployStatus() As String
        Dim LOG As New PigStepLog("RefDeployStatus")
        Try
            If Me.mIsFolderExists(Me.HomeDirPath) = False Then
                Me.DeployStatus = EnmDomainDeployStatus.NotCreate
            ElseIf Me.mIsFileExists(Me.startWebLogicPath) = False Or Me.mIsFileExists(Me.stopWebLogicPath) = False Or Me.mIsFileExists(Me.ConfPath) = False Then
                Me.DeployStatus = EnmDomainDeployStatus.IncompleteCreation
            ElseIf Me.IsProdMode = False Then
                Me.DeployStatus = EnmDomainDeployStatus.CreateDevMode
            ElseIf Me.mIsFileExists(Me.SecurityBootPath) = False Then
                Me.DeployStatus = EnmDomainDeployStatus.CreateProdModeNotSecurity
            Else
                Me.DeployStatus = EnmDomainDeployStatus.CreateProdMode
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RefRunStatus() As String
        Dim LOG As New PigStepLog("RefRunStatus")
        Try
            Select Case Me.RunStatus
                Case EnmDomainRunStatus.ExecutingWLST
                    If Math.Abs(DateDiff(DateInterval.Second, Me.mCallWlstBeginTime, Now)) > Me.fParent.CallWlstTimeout Then
                        Me.RunStatus = EnmDomainRunStatus.ExecWLSTFail
                        Me.CallWlstRes = "WLST execution timeout."
                    End If
                Case Else
                    Select Case Me.DeployStatus
                        Case EnmDomainDeployStatus.CreateProdModeNotSecurity, EnmDomainDeployStatus.IncompleteCreation, EnmDomainDeployStatus.NotCreate
                            Me.RunStatus = EnmDomainRunStatus.DeployNoReady
                        Case Else
                            Select Case Me.RunStatus
                                Case EnmDomainRunStatus.Stopping
                                    If Math.Abs(DateDiff(DateInterval.Second, Me.mStopDomainBeginTime, Now)) > Me.fParent.StartOrStopTimeout Then
                                        Me.RunStatus = EnmDomainRunStatus.StopFail
                                        Me.StopDomainRes = "Stop domain timeout."
                                    End If
                                Case Else
                                    Dim intPID As Integer
                                    LOG.StepName = "GetListenPortProcID"
                                    LOG.Ret = Me.mPigSysCmd.GetListenPortProcID(Me.ListenPort, intPID)
                                    If LOG.Ret <> "OK" Then Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.Ret)
                                    If intPID >= 0 Then
                                        LOG.StepName = "GetPigProc"
                                        Dim oPigProc As PigProc = Me.mPigProcApp.GetPigProc(intPID)
                                        If Me.mPigProcApp.LastErr <> "" Then
                                            Me.PrintDebugLog(LOG.SubName, LOG.StepName, Me.mPigProcApp.LastErr)
                                        Else
                                            If UCase(oPigProc.ProcessName) = "JAVA" Then
                                                Me.RunStatus = EnmDomainRunStatus.Running
                                                Me.JavaPID = oPigProc.ProcessID
                                                Me.JavaStartTime = oPigProc.StartTime
                                                Me.JavaCpuTime = oPigProc.UserProcessorTime
                                                Me.JavaMemoryUse = CDec(oPigProc.MemoryUse) / 1024 / 1024
                                            Else
                                                Me.RunStatus = EnmDomainRunStatus.ListenPortByOther
                                            End If
                                        End If
                                    ElseIf Me.RunStatus = EnmDomainRunStatus.Starting Then
                                        If Math.Abs(DateDiff(DateInterval.Second, Me.mStartDomainBeginTime, Now)) > Me.fParent.StartOrStopTimeout Then
                                            Me.RunStatus = EnmDomainRunStatus.StartFail
                                            Me.StartDomainRes = "Start domain timeout."
                                        End If
                                    Else
                                        Me.RunStatus = EnmDomainRunStatus.Stopped
                                    End If
                            End Select
                    End Select
            End Select
            If Me.RunStatus <> EnmDomainRunStatus.Running Then
                Me.JavaPID = -1
                Me.JavaStartTime = TEMP_DATE
                Me.JavaCpuTime = TimeSpan.Zero
                Me.JavaMemoryUse = 0
            End If
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
                        Me.RunStatus = EnmDomainRunStatus.Stopped
                        Me.StopDomainRes = StandardOutput
                        RaiseEvent StopDomainSucc(StandardOutput)
                    Else
                        Me.RunStatus = EnmDomainRunStatus.StopFail
                        Me.StopDomainRes = .AsyncRet & Me.OsCrLf & StandardOutput & Me.OsCrLf & StandardError & Me.OsCrLf
                        RaiseEvent StopDomainFail(StandardError)
                    End If
                Case Me.mRefRunStatusThreadID
                    Me.mRefRunStatusThreadID = -1
                Case Me.mCallWlstThreadID
                    Me.mCallWlstThreadID = -1
                    If Me.mIsFileExists(Me.CallWlstPyPath) = True Then
                        Me.mPigFunc.DeleteFile(Me.CallWlstPyPath)
                    End If
                    If .AsyncRet = "OK" And StandardError = "" Then
                        Me.RunStatus = EnmDomainRunStatus.ExecWLSTOK
                        Me.CallWlstRes = StandardOutput
                        Select Case Me.WlstCallCmd
                            Case EnmWlstCallCmd.Connect
                                Me.IsConnectOK = True
                            Case EnmWlstCallCmd.CreateDomain
                                Me.IsCreateDomainOK = True
                        End Select
                        RaiseEvent CallWlstSucc(Me.WlstCallCmd, StandardOutput)
                    Else
                        Me.RunStatus = EnmDomainRunStatus.ExecWLSTFail
                        Me.CallWlstRes = StandardError
                        Select Case Me.WlstCallCmd
                            Case EnmWlstCallCmd.Connect
                                Me.IsConnectOK = False
                            Case EnmWlstCallCmd.CreateDomain
                                Me.IsCreateDomainOK = False
                        End Select
                        RaiseEvent CallWlstFail(Me.WlstCallCmd, StandardError)
                    End If
                    'Case Me.mStartDomainThreadID
                    '    Me.mStartDomainThreadID = -1
                    '    Me.PrintDebugLog("mPigCmdApp_AsyncRet_CmdShell_FullString", "mStartDomainThreadID")
                    '    If .AsyncRet = "OK" And StandardError = "" Then
                    '        Me.RunStatus = EnmDomainRunStatus.Running
                    '        Me.StartDomainRes = StandardOutput
                    '        RaiseEvent StartDomainSucc(StandardOutput)
                    '    Else
                    '        Me.StartDomainRes = .AsyncRet & Me.OsCrLf & StandardOutput & Me.OsCrLf & StandardError & Me.OsCrLf
                    '        Me.RunStatus = EnmDomainRunStatus.StartFail
                    '        RaiseEvent StartDomainFail(StandardError)
                    '    End If
            End Select

        End With
    End Sub

    Public Function SetT3Deny(ByRef CmdRes As String) As String
        Try
            Return Me.mWlstCallMain(EnmWlstCallCmd.SetT3Deny, ,,,, CmdRes)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function SetSetAdminPort(ByRef CmdRes As String, IsEnableAdminPort As Boolean, Optional AdminPort As Integer = 0) As String
        Try
            If IsEnableAdminPort = True Then
                If AdminPort <= 0 Then Throw New Exception("Invalid administrator port.")
                If AdminPort = Me.ListenPort Then Throw New Exception("The administrator port cannot be the same as the listening port.")
            End If
            Return Me.mWlstCallMain(EnmWlstCallCmd.SetAdminPort, , AdminPort,,, CmdRes, IsEnableAdminPort)
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function GetConsoleLogHasRunMode(ByRef IsHasRunMode As Boolean) As String
        Dim LOG As New PigStepLog("IsConsoleLogHasRunMode")
        Const RUN_MODE As String = "<The server started in RUNNING mode.>"
        Dim strCmd As String = ""
        Try
            If Me.IsWindows = True Then
                strCmd = "type """ & Me.ConsolePath & """|findstr """ & RUN_MODE & """"
            Else
                strCmd = "cat " & Me.ConsolePath & "|grep """ & RUN_MODE & """"
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.mPigFunc.Delay(500)
            If InStr(Me.mPigCmdApp.StandardOutput, RUN_MODE) > 0 Then
                IsHasRunMode = True
            Else
                IsHasRunMode = False
            End If
            Return "OK"
        Catch ex As Exception
            IsHasRunMode = False
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
