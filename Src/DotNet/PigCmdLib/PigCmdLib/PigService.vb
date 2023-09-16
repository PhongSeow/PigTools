'**********************************
'* Name: PigService
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Service processing class|服务处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 2/9/2023
'* 1.1  3/9/2023   Add Refresh
'* 1.2  11/9/2023  Add StartService,StopService,Delete
'* 1.3  16/9/2023  Modify Delete,New,Refresh
'**********************************
Imports Microsoft.VisualBasic.Logging
Imports PigToolsLiteLib
''' <summary>
''' Service processing class|服务处理类
''' </summary>
Public Class PigService
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.2.26"

    Public Sub New(ServiceName As String)
        MyBase.New(CLS_VERSION)
        Me.ServiceName = ServiceName
    End Sub

    Public Enum EnmServiceState
        Unknow = 0
        Running = 1
        Stopped = 2
        Starting = 3
        Stopping = 4
    End Enum

    Public Enum EnmStartMode
        Unknow = 0
        Manual = 1
        Automatic = 2
        Disabled = 3
    End Enum

    Private ReadOnly Property mPigSysCmd As New PigSysCmd
    Private ReadOnly Property mPigCmdApp As New PigCmdApp
    Private ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property ServiceName As String

    Private mDescription As String
    Public Property Description As String
        Get
            Return mDescription
        End Get
        Friend Set(value As String)
            mDescription = value
        End Set
    End Property

    Private mPathName As String
    Public Property PathName As String
        Get
            Return mPathName
        End Get
        Friend Set(value As String)
            mPathName = value
        End Set
    End Property

    Private mStartUser As String
    Public Property StartUser As String
        Get
            Return mStartUser
        End Get
        Friend Set(value As String)
            mStartUser = value
        End Set
    End Property

    Private mDisplayName As String
    Public Property DisplayName As String
        Get
            Return mDisplayName
        End Get
        Friend Set(value As String)
            mDisplayName = value
        End Set
    End Property

    Private mProcessId As Integer
    Public Property ProcessId As Integer
        Get
            Return mProcessId
        End Get
        Friend Set(value As Integer)
            mProcessId = value
        End Set
    End Property

    Private mServiceState As EnmServiceState
    Public Property ServiceState As EnmServiceState
        Get
            Return mServiceState
        End Get
        Friend Set(value As EnmServiceState)
            mServiceState = value
        End Set
    End Property

    Private mStartMode As EnmStartMode
    Public Property StartMode As EnmStartMode
        Get
            Return mStartMode
        End Get
        Friend Set(value As EnmStartMode)
            mStartMode = value
        End Set
    End Property

    Public Function Refresh() As String
        Dim LOG As New PigStepLog("Refresh")
        Dim strWmicCmd As String = "service " & Me.ServiceName & " get Description,DisplayName,PathName,ProcessId,State,PathName,StartMode,StartName,Name"
        Try
            If Me.IsWindows = False Then Throw New Exception("Only supports Windows now.")
            Dim pxMain As PigXml = Nothing
            LOG.StepName = "GetWmicSimpleXml"
            LOG.Ret = Me.mPigSysCmd.GetWmicSimpleXml(strWmicCmd, pxMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If pxMain.XmlDocGetStr("WmicXml.Row1.Name") = "" Then Throw New Exception("Unable to obtain information for service " & Me.ServiceName)
            With Me
                .Description = pxMain.XmlDocGetStr("WmicXml.Row1.Description")
                .DisplayName = pxMain.XmlDocGetStr("WmicXml.Row1.DisplayName")
                .PathName = pxMain.XmlDocGetStr("WmicXml.Row1.PathName")
                .ProcessId = pxMain.XmlDocGetInt("WmicXml.Row1.ProcessId")
                Select Case pxMain.XmlDocGetStr("WmicXml.Row1.State")
                    Case "Running"
                        .ServiceState = EnmServiceState.Running
                    Case "Stopped"
                        .ServiceState = EnmServiceState.Stopped
                    Case "Starting"
                        .ServiceState = EnmServiceState.Starting
                    Case "Stopping"
                        .ServiceState = EnmServiceState.Stopping
                    Case Else
                        .ServiceState = EnmServiceState.Unknow
                End Select
                Select Case pxMain.XmlDocGetStr("WmicXml.Row1.StartMode")
                    Case "Auto"
                        .StartMode = EnmStartMode.Automatic
                    Case "Manual"
                        .StartMode = EnmStartMode.Manual
                    Case "Disabled"
                        .StartMode = EnmStartMode.Disabled
                    Case Else
                        .StartMode = EnmStartMode.Unknow
                End Select
                .StartUser = pxMain.XmlDocGetStr("WmicXml.Row1.StartName")
            End With
            pxMain = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strWmicCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Create(DisplayName As String, PathName As String, StartMode As EnmStartMode) As String
        Return Me.mCreate(PathName, DisplayName, StartMode)
    End Function

    Public Function Create(DisplayName As String, PathName As String, StartMode As EnmStartMode, StartUser As String, StartUserPwd As String) As String
        Return Me.mCreate(PathName, DisplayName, StartMode, StartUser, StartUserPwd)
    End Function

    Private Function mCreate(PathName As String, DisplayName As String, StartMode As EnmStartMode, Optional StartUser As String = "", Optional StartUserPwd As String = "") As String
        Dim LOG As New PigStepLog("mCreate")
        Dim strCmd As String = ""
        Try
            If Me.IsWindows = False Then Throw New Exception("Only supports Windows now.")
            If PathName = "" Then Throw New Exception("PathName not defined")
            strCmd = "sc create " & Me.ServiceName
            strCmd &= " start= "
            Select Case StartMode
                Case EnmStartMode.Automatic
                    strCmd &= "auto"
                Case EnmStartMode.Disabled
                    strCmd &= "disabled"
                Case EnmStartMode.Manual
                    strCmd &= "demand"
                Case Else
                    Throw New Exception("Invalid StartMode:" & StartMode.ToString)
            End Select
            strCmd &= " binPath= """ & PathName & """"
            strCmd &= " DisplayName= """ & DisplayName & """"
            If StartUser <> "" Then
                If InStr(StartUser, "\") = 0 Then StartUser = ".\" & StartUser
                strCmd &= " obj= " & StartUser
                strCmd &= " password= " & StartUserPwd
            End If
            LOG.StepName = ""
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(Me.mPigCmdApp.StandardOutput)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            Dim strErr As String = Me.mGetScErrInf(Me.mPigCmdApp.StandardOutput)
            If strErr <> "" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Delete() As String
        Dim LOG As New PigStepLog("Delete")
        Dim strCmd As String = "sc delete " & Me.ServiceName
        Try
            If Me.IsWindows = False Then Throw New Exception("Only supports Windows now.")
            LOG.StepName = ""
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(Me.mPigCmdApp.StandardOutput)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            Dim strErr As String = Me.mGetScErrInf(Me.mPigCmdApp.StandardOutput)
            If strErr <> "" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetScErrInf(StandardOutput As String) As String
        Try
            Dim strData As String = StandardOutput
            If InStr(StandardOutput, "WIN32_EXIT_CODE") > 0 And InStr(StandardOutput, "SERVICE_EXIT_CODE") > 0 Then
                mGetScErrInf = ""
            ElseIf InStr(strData, ":") > 0 Then
                Dim intLines As Integer = 0
                Do While True
                    Dim strLine As String = Me.mPigFunc.GetStr(strData, "", Me.OsCrLf)
                    If strLine = "" Then Exit Do
                    intLines += 1
                Loop
                If intLines > 1 Then Throw New Exception(StandardOutput)
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Function StartService() As String
        Dim LOG As New PigStepLog("StartService")
        Dim strCmd As String = "sc start " & Me.ServiceName
        Try
            If Me.IsWindows = False Then Throw New Exception("Only supports Windows now.")
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim strErr As String = Me.mGetScErrInf(Me.mPigCmdApp.StandardOutput)
            If strErr <> "" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function StopService() As String
        Dim LOG As New PigStepLog("StopService")
        Dim strCmd As String = "sc stop " & Me.ServiceName
        Try
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim strErr As String = Me.mGetScErrInf(Me.mPigCmdApp.StandardOutput)
            If strErr <> "" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
