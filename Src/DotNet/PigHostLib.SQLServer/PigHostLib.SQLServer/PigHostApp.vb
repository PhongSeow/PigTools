'**********************************
'* Name: 豚豚主机应用|PigHostApp
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机处理|Host processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 8/11/2021
'* 1.1    12/11/2021   Add MyHost
'* 1.2    2/1/2022   Add MyHosts
'**********************************
Public Class PigHostApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.2"

    Public ReadOnly Property MyHost As PigHost

    Private moPigHosts As PigHosts
    Public Property PigHosts As PigHosts
        Get
            Return moPigHosts
        End Get
        Friend Set(value As PigHosts)
            moPigHosts = value
        End Set
    End Property



    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.MyHost = New PigHost
    End Sub

End Class
