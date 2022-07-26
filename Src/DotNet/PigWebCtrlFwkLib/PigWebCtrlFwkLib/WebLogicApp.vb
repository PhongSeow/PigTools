'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Application of dealing with Weblogic
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.7
'* Create Time: 31/1/2022
'*1.1  5/2/2022   Add GetJavaVersion 
'*1.2  6/3/2022   Add WlstPath 
'*1.3  26/5/2022  Add SaveSecurityBoot ,modify New
'*1.4  27/5/2022  Modify SaveSecurityBoot
'*1.5  1/6/2022  Add IsWindows
'*1.6  5/6/2022  Add StartOrStopTimeout
'*1.7  26/7/2022 Modify Imports
'************************************
#If NETFRAMEWORK Then
Imports PigCmdFwkLib
Imports PigToolsWinLib
#Else
Imports PigCmdLib
Imports PigToolsLiteLib
#End If
Imports PigObjFsLib

Public Class WebLogicApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.7.2"
    Public ReadOnly Property HomeDirPath As String
    Public ReadOnly Property WorkTmpDirPath As String
    Public ReadOnly Property CallWlstTimeout As Integer = 300
    Public ReadOnly Property StartOrStopTimeout As Integer = 60

    Public Property WebLogicDomains As WebLogicDomains

    Private WithEvents mPigCmdApp As New PigCmdApp
    Private mFS As New FileSystemObject
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

End Class
