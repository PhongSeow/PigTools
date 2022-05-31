'**********************************
'* Name: WebLogicDomain
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic domain
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6
'* Create Time: 31/1/2022
'*1.1  5/2/2022   Add CheckDomain 
'*1.2  5/3/2022   Modify New
'*1.3  23/5/2022  Add CreateDomain 
'*1.4  26/5/2022  Add EnmDomainDeployStatus,EnmDomainRunStatus,EnmDomainCtrlStatus
'*1.5  27/5/2022  Add CreateDomain 
'*1.6  31/5/2022  Modify CreateDomain 
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
Imports PigObjFsLib

Public Class WebLogicDomain
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.6.18"

    Friend WithEvents mPigCmdApp As New PigCmdApp
    Private mFS As New FileSystemObject
    Private mPigFunc As New PigFunc

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
        ''' 创建域失败|Failed to create domain
        ''' </summary>
        CreateDomainFail = -1
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
        ''' 正在创建域|Creating domain
        ''' </summary>
        CreatingDomain = 4
        ''' <summary>
        ''' 创建域成功|Domain created successfully
        ''' </summary>
        CreateDomainOK = 5
    End Enum

    ''' <summary>
    ''' 控制状态|Control state
    ''' </summary>
    Public Enum EnmDomainCtrlStatus
        ''' <summary>
        ''' 就绪|ready
        ''' </summary>
        Ready = 0
        ''' <summary>
        ''' 启动域|Startup domain
        ''' </summary>
        StartDomain = 1
        ''' <summary>
        ''' 停止域|Stop domain
        ''' </summary>
        StopDomain = 2
        ''' <summary>
        ''' 重启域|Restart domain
        ''' </summary>
        RestartDomain = 3
        ''' <summary>
        ''' 创建域|Create domain
        ''' </summary>
        CreateDomain = 4
        ''' <summary>
        ''' 保存生产模式免输入密启动文件|Save the production mode password free startup file
        ''' </summary>
        SaveSecurityBoot = 5
    End Enum

    Public ReadOnly Property HomeDirPath As String
    Public Property AdminUserName As String
    Public Property AdminUserPassword As String
    Friend Property CreateDomainPyPath As String

    Private Property mCreateDomainThreadID As Integer

    Private mstrCreateDomainRes As String
    Public Property CreateDomainRes As String
        Get
            Return mstrCreateDomainRes
        End Get
        Friend Set(value As String)
            mstrCreateDomainRes = value
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

    Private mintCtrlStatus As EnmDomainCtrlStatus
    Public Property CtrlStatus As EnmDomainCtrlStatus
        Get
            Return mintCtrlStatus
        End Get
        Friend Set(value As EnmDomainCtrlStatus)
            mintCtrlStatus = value
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

    Private mbolIsIIopEnable As Boolean
    Public Property IsIIopEnable As Boolean
        Get
            Return mbolIsIIopEnable
        End Get
        Friend Set(value As Boolean)
            mbolIsIIopEnable = value
        End Set
    End Property

    Private mbolHasAdminServer As Boolean
    Public Property HasAdminServer As Boolean
        Get
            Return mbolHasAdminServer
        End Get
        Friend Set(value As Boolean)
            mbolHasAdminServer = value
        End Set
    End Property

    Private mbolIsAdminPortEnable As Boolean
    Public Property IsAdminPortEnable As Boolean
        Get
            Return mbolIsAdminPortEnable
        End Get
        Friend Set(value As Boolean)
            mbolIsAdminPortEnable = value
        End Set
    End Property

    Private mdteLastOperationTime As DateTime
    Public Property LastOperationTime As DateTime
        Get
            Return mdteLastOperationTime
        End Get
        Friend Set(value As DateTime)
            mdteLastOperationTime = value
        End Set
    End Property


    Private mintListenPort As String
    Public Property ListenPort As Integer
        Get
            Return mintListenPort
        End Get
        Friend Set(value As Integer)
            mintListenPort = value
        End Set
    End Property

    Private mintAdminPort As String
    Public Property AdminPort As Integer
        Get
            Return mintAdminPort
        End Get
        Friend Set(value As Integer)
            mintAdminPort = value
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
            Return Me.HomeDirPath & Me.OsPathSep & "AdminServer" & Me.OsPathSep & "logs"
        End Get
    End Property

    Public ReadOnly Property ConsolePath() As String
        Get
            Return Me.LogDirPath & Me.OsPathSep & "AdminServer" & Me.OsPathSep & "Console.log"
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

    Public Function CheckDomain() As String
        Dim LOG As New PigStepLog("CheckDomain")
        Try
            If Me.mIsFolderExists(Me.HomeDirPath) = False Then Throw New Exception("HomeDirPath " & Me.HomeDirPath & " not found.")
            If Me.mIsFileExists(Me.startWebLogicPath) = False Then Throw New Exception("startWebLogicPath  " & Me.startWebLogicPath & " not found.")
            If Me.mIsFileExists(Me.stopWebLogicPath) = False Then Throw New Exception("stopWebLogicPath  " & Me.stopWebLogicPath & " not found.")
            If Me.mIsFileExists(Me.ConfPath) = False Then Throw New Exception("ConfPath  " & Me.ConfPath & " not found.")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function CreateDomain(ListenPort As Integer, AdminUserName As String, AdminUserPassword As String) As String
        Dim LOG As New PigStepLog("CreateDomain")
        Try
            LOG.StepName = "RefDeployStatus"
            LOG.Ret = Me.RefDeployStatus()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check DeployStatus"
            If Me.DeployStatus <> EnmDomainDeployStatus.NotCreate Then Throw New Exception("A domain can be created only when DeployStatus is NotCreate")
            LOG.StepName = "RefRunStatus"
            LOG.Ret = Me.RefRunStatus()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Check RunStatus"
            If Me.RunStatus <> EnmDomainRunStatus.DeployNoReady Then Throw New Exception("A domain can be created only when RunStatus is DeployNoReady")
            LOG.StepName = "Check Wls.Jar"
            If Me.mIsFileExists(Me.fParent.WlsJarPath) = False Then Throw New Exception(Me.fParent.WlsJarPath & " not found.")
            With Me
                .ListenPort = ListenPort
                .AdminUserName = AdminUserName
                .AdminUserPassword = AdminUserPassword
            End With
            With Me
                .CreateDomainPyPath = Me.fParent.WorkTmpDirPath & Me.OsPathSep & Me.mPigFunc.GetPKeyValue(.DomainName, False) & ".py"
                LOG.StepName = "OpenTextFile"
                Dim tsMain As TextStream = Me.mFS.OpenTextFile(.CreateDomainPyPath, FileSystemObject.IOMode.ForWriting, True)
                If tsMain.LastErr <> "" Or Me.mFS.LastErr <> "" Then
                    LOG.AddStepNameInf(.CreateDomainPyPath)
                    Throw New Exception(tsMain.LastErr & "." & Me.mFS.LastErr)
                End If
                Dim strLine As String
                With tsMain
                    LOG.StepName = "WriteLine"
                    strLine = "readTemplate('" & Me.fParent.WlsJarPath & "')"
                    If Me.IsWindows = True Then
                        strLine = Replace(strLine, Me.OsPathSep, "//")
                    End If
                    .WriteLine(strLine)
                    .WriteLine("cd('Servers/AdminServer')")
                    .WriteLine("set('ListenAddress','')")
                    .WriteLine("set('ListenPort', " & Me.ListenPort & ")")
                    .WriteLine("cd('/')")
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
                    .WriteLine("exit()")
                    LOG.StepName = "Close"
                    .Close()
                End With
            End With
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "call " & Me.fParent.WlstPath & " " & Me.CreateDomainPyPath
            Else
                strCmd = Me.fParent.WlstPath & " " & Me.CreateDomainPyPath
            End If
            Me.PrintDebugLog(LOG.SubName, LOG.StepName, strCmd)
            Me.RunStatus = EnmDomainRunStatus.CreatingDomain
            Me.CreateDomainRes = ""
            LOG.StepName = "AsyncCmdShell"
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, Me.mCreateDomainThreadID)
            If LOG.Ret <> "OK" Then
                If Me.IsDebug = True Then LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function



    Public Function RefConf() As String
        Dim LOG As New PigStepLog("RefConf")
        Try
            If Me.mIsFileExists(Me.ConfPath) = False Then
                Throw New Exception(Me.ConfPath & " not found.")
            End If
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
            End With
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
                Case EnmDomainRunStatus.CreatingDomain
                Case Else
                    Select Case Me.DeployStatus
                        Case EnmDomainDeployStatus.CreateProdModeNotSecurity, EnmDomainDeployStatus.IncompleteCreation, EnmDomainDeployStatus.NotCreate
                            Me.RunStatus = EnmDomainRunStatus.DeployNoReady
                        Case Else

                    End Select
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Sub mPigCmdApp_AsyncRet_CmdShell_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles mPigCmdApp.AsyncRet_CmdShell_FullString
        With AsyncRet
            If .AsyncThreadID = Me.mCreateDomainThreadID Then
                Me.CreateDomainRes = .AsyncRet & Me.OsCrLf & StandardOutput & Me.OsCrLf & StandardError & Me.OsCrLf
                If .AsyncRet = "OK" And StandardError = "" Then
                    Me.RunStatus = EnmDomainRunStatus.CreateDomainOK
                Else
                    Me.RunStatus = EnmDomainRunStatus.CreateDomainFail
                End If
                If Me.mIsFileExists(Me.CreateDomainPyPath) = True Then
                    Me.mPigFunc.DeleteFile(Me.CreateDomainPyPath)
                End If
            End If

        End With
    End Sub
End Class
