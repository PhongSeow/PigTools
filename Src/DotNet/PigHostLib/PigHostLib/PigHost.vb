'**********************************
'* Name: 豚豚主机|PigHost
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机信息处理|Host information processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 8/10/2021
'* 1.1    12/10/2021   Add New
'* 1.2    18/10/2021   Modify New
'**********************************
Imports PigToolsLiteLib

Public Class PigHost
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"

    Public ReadOnly Property HostID As String
    Public Property HostName As String
    Public Property HostIp As String
    Public Property HostIpList As String

    Private moPigFunc As New PigFunc

    Public Sub New(HostID As String)
        MyBase.New(CLS_VERSION)
        Me.HostID = HostID
    End Sub

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Try
            With Me
                .HostName = moPigFunc.GetHostName
                .HostIp = moPigFunc.GetHostIp
                .HostIpList = moPigFunc.GetHostIpList
            End With
            Me.HostID = moPigFunc.GetPKeyValue(Me.HostName, True)
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

End Class
