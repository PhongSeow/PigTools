'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Application of dealing with Weblogic
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.12
'* Create Time: 31/1/2022
'* 1.1  5/2/2022   Add GetJavaVersion 
'* 1.2  6/3/2022   Add WlstPath 
'* 1.3  26/5/2022  Add SaveSecurityBoot ,modify New
'* 1.4  27/5/2022  Modify SaveSecurityBoot
'* 1.5  1/6/2022  Add IsWindows
'* 1.6  5/6/2022  Add StartOrStopTimeout
'* 1.7  26/7/2022 Modify Imports
'* 1.8  29/7/2022 Modify Imports
'* 1.9  7/9/2022  Add RunOpatch
'* 1.10 28/9/2022 Modify StartOrStopTimeout
'* 1.11 24/6/2023 Change the reference to PigObjFsLib to PigToolsLiteLib
'* 1.12 7/12/2023 Add RunOpatch
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib

''' <summary>
''' WebLogic Processing Class|WebLogic处理类
''' </summary>
Public Class WebLogicApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.12.10"
    Public ReadOnly Property HomeDirPath As String
    Public ReadOnly Property WorkTmpDirPath As String
    Public ReadOnly Property CallWlstTimeout As Integer = 300
    Public ReadOnly Property StartOrStopTimeout As Integer = 300

    Public Property WebLogicDomains As WebLogicDomains

    Private WithEvents mPigCmdApp As New PigCmdApp
    Private mFS As New PigFileSystem
    Private mPigFunc As New PigFunc

    Private mGetJavaVersionThreadID As Integer

    Public Sub New(HomeDirPath As String, WorkTmpDirPath As String)
        MyBase.New(CLS_VERSION)
        Try
            Me.WebLogicDomains = New WebLogicDomains
            Me.WebLogicDomains.fParent = Me
            Me.HomeDirPath = HomeDirPath
            Me.WorkTmpDirPath = WorkTmpDirPath
            Me.CallWlstTimeout = CallWlstTimeout
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Get java version information|获取java版本信息
    ''' </summary>
    ''' <returns></returns>
    Public Function GetJavaVersion() As String
        Try
            Dim strRet As String = Me.mPigCmdApp.AsyncCmdShell("java -version", Me.mGetJavaVersionThreadID)
            Return strRet
        Catch ex As Exception
            Return Me.GetSubErrInf("GetJavaVersion", ex)
        End Try
    End Function


    Private mstrJavaVersion As String
    Public Property JavaVersion As String
        Get
            Return mstrJavaVersion
        End Get
        Friend Set(value As String)
            mstrJavaVersion = value
        End Set
    End Property

    ''' <summary>
    ''' 操作超时时间，单位为秒|Operation timeout in seconds
    ''' </summary>
    Private mintOperationTimeout As Integer = 60
    Public Property OperationTimeout As Decimal
        Get
            Return mintOperationTimeout
        End Get
        Friend Set(value As Decimal)
            mintOperationTimeout = value
        End Set
    End Property

    Public ReadOnly Property WlsJarPath() As String
        Get
            WlsJarPath = Me.HomeDirPath & Me.OsPathSep & "wlserver" & Me.OsPathSep & "common" & Me.OsPathSep & "templates" & Me.OsPathSep & "wls" & Me.OsPathSep & "wls.jar"
        End Get
    End Property

    Public ReadOnly Property WlstPath() As String
        Get
            WlstPath = Me.HomeDirPath & Me.OsPathSep & "oracle_common" & Me.OsPathSep & "common" & Me.OsPathSep & "bin" & Me.OsPathSep & "wlst."
            If Me.IsWindows = True Then
                WlstPath &= "cmd"
            Else
                WlstPath &= "sh"
            End If
        End Get
    End Property


    Private Sub mPigCmdApp_AsyncRet_CmdShell_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles mPigCmdApp.AsyncRet_CmdShell_FullString
        If AsyncRet.AsyncThreadID = Me.mGetJavaVersionThreadID Then
            Me.JavaVersion = StandardError
        End If
    End Sub

    Public Overloads ReadOnly Property IsWindows() As String
        Get
            Return MyBase.IsWindows
        End Get
    End Property

    Public Overloads Sub SetDebug(DebugFilePath)
        MyBase.SetDebug(DebugFilePath)
    End Sub

    Public Overloads Sub PrintDebugLog(SubName As String, StepName As String, LogInf As String)
        MyBase.PrintDebugLog(SubName, StepName, LogInf)
    End Sub

    ''' <summary>
    ''' Run patch script|运行补丁脚本
    ''' </summary>
    ''' <param name="Cmd">script|脚本</param>
    ''' <param name="ResInf">Return Results|返回结果</param>
    ''' <returns></returns>
    Public Function RunOpatch(Cmd As String, ByRef ResInf As String) As String
        Dim LOG As New PigStepLog("RunOpatch")
        Try
            Dim strCmd As String = Me.HomeDirPath & Me.OsPathSep & "OPatch" & Me.OsPathSep & "opatch " & Cmd
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            ResInf = Me.mPigCmdApp.StandardOutput & Me.OsCrLf & Me.mPigCmdApp.StandardError
            Return "OK"
        Catch ex As Exception
            ResInf = ""
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

    ''' <summary>
    ''' Run patch script|运行补丁脚本
    ''' </summary>
    ''' <param name="SudoUser">sudo user|sudo 用户</param>
    ''' <param name="Cmd">script|脚本</param>
    ''' <param name="ResInf">Return Results|返回结果</param>
    ''' <returns></returns>
    Public Function RunOpatch(SudoUser As String, Cmd As String, ByRef ResInf As String) As String
        Dim LOG As New PigStepLog("RunOpatch")
        Try
            If Me.IsWindows = True Then Throw New Exception("Can only be executed on the Linux platform")
            Dim strCmd As String = Me.HomeDirPath & Me.OsPathSep & "OPatch" & Me.OsPathSep & "opatch " & Cmd
            LOG.StepName = "PigSudo.Run"
            Dim oPigSudo As New PigSudo(strCmd, SudoUser)
            LOG.Ret = oPigSudo.Run()
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            ResInf = oPigSudo.StandardOutput & Me.OsCrLf & oPigSudo.StandardError
            oPigSudo = Nothing
            Return "OK"
        Catch ex As Exception
            ResInf = ""
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

End Class
