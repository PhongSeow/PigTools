'**********************************
'* Name: 豚豚主机用户|PigHostUser
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机用户信息处理|Host user information processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 30/10/2021
'* 1.1    2/4/2022   Add 
'**********************************
Public Class PigHostUser
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.1.2"

	Public ReadOnly Property HostUserID As String
	Public ReadOnly Property Parent As PigHost
	Private mstrUserName As String
	Public Property UserName As String
		Get
			Return mstrUserName
		End Get
		Friend Set(value As String)
			mstrUserName = value
		End Set
	End Property

	Public Sub New(UserName As String, Parent As PigHost)
		MyBase.New(CLS_VERSION)

	End Sub


End Class
