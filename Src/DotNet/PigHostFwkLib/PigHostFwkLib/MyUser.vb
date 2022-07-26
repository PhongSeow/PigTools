'**********************************
'* Name: 我的用户|MyUser
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 我的用户信息|My user information
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 30/10/2021
'* 1.1    2/4/2022   Add 
'* 1.2    24/7/2022   Change name and initialization
'* 1.6    26/7/2022   Modify Imports
'**********************************
#If NETFRAMEWORK Then
Imports PigToolsWinLib
Imports PigCmdFwkLib
#Else
Imports PigToolsLiteLib
Imports PigCmdLib
#End If

Public Class MyUser
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.3.2"

    Public ReadOnly Property UserName As String
    Private Property mPigFunc As New PigFunc

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.UserName = Me.mPigFunc.GetUserName
    End Sub


End Class
