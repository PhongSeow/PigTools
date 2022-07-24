'**********************************
'* Name: 我的主机|MyHost
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机信息处理|Host information processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 8/10/2021
'* 1.1    12/10/2021   Add New
'* 1.2    18/10/2021   Modify New
'* 1.3    23/7/2022   Modify New
'* 1.5    24/7/2022   Rename MyHost,
'**********************************
Imports PigToolsLiteLib
Imports PigCmdLib

Public Class MyHost
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.5.2"

    Public ReadOnly Property HostID As String

    ''' <summary>
    ''' 操作系统名称|Operating system name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property OSCaption As String
    Public ReadOnly Property UUID As String

    ''' <summary>
    ''' 主机名|host name
    ''' </summary>
    ''' <returns></returns>
    Public Property HostName As String
    ''' <summary>
    ''' 主IP，自动获取|Primary IP, automatic acquisition
    ''' </summary>
    ''' <returns></returns>
    Public Property HostIp As String
    ''' <summary>
    ''' IP地址列表|IP address list
    ''' </summary>
    ''' <returns></returns>
    Public Property HostIpList As String

    Private Property mPigFunc As New PigFunc
    Private Property mPigSysCmd As New PigSysCmd


    Public Sub New()
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            LOG.StepName = "GetOSCaption"
            LOG.Ret = Me.mPigSysCmd.GetOSCaption(Me.OSCaption)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "GetUUID"
            LOG.Ret = Me.mPigSysCmd.GetUUID(Me.UUID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "New PigMD5"
            Dim oPigMD5 As New PigMD5("<" & Me.OSCaption & "><" & Me.UUID & ">", PigMD5.enmTextType.UTF8)
            Me.HostID = oPigMD5.PigMD5
            oPigMD5 = Nothing
            With Me
                .HostName = Me.mPigFunc.GetHostName
                .HostIp = Me.mPigFunc.GetHostIp()
                .HostIpList = Me.mPigFunc.GetHostIpList
            End With
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

End Class
