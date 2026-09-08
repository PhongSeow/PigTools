'**********************************
'* Name: PigJava
'* Author: Seow Phong
'* License: Copyright (c) 2025 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: JAVA information class|JAVA信息类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 19/4/2025
'**********************************
Friend Class PigJava
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "2"

    Public ReadOnly Property PID As Integer

    Public Sub New(PID As Integer)
        MyBase.New(CLS_VERSION)
        Me.PID = PID
    End Sub


End Class
