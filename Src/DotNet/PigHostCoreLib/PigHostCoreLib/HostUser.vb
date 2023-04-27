'********************************************************************
'* Copyright 2023 Seow Phong
'*
'* Licensed under the Apache License, Version 2.0 (the "License");
'* you may Not use this file except in compliance with the License.
'* You may obtain a copy of the License at
'*
'*     http://www.apache.org/licenses/LICENSE-2.0
'*
'* Unless required by applicable law Or agreed to in writing, software
'* distributed under the License Is distributed on an "AS IS" BASIS,
'* WITHOUT WARRANTIES Or CONDITIONS OF ANY KIND, either express Or implied.
'* See the License for the specific language governing permissions And
'* limitations under the License.
'********************************************************************
'* Name: 主机用户|HostUser
'* Author: Seow Phong
'* Describe: 我的用户信息|My user information
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 30/10/2021
'* 1.1    2/4/2022   Add 
'* 1.2    24/7/2022   Change name and initialization
'* 1.3    29/7/2022   Modify Imports
'* 1.4    30/7/2022   Modify Imports
'* 1.5    4/10/2022   Modify Imports
'**********************************
Imports PigToolsLiteLib
Imports PigCmdLib

Friend Class HostUser
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.3.2"

    Public ReadOnly Property UserName As String
    Private Property mPigFunc As New PigFunc

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.UserName = Me.mPigFunc.GetUserName
    End Sub


End Class
