'**********************************
'* Name: 豚豚磁盘|PigDisk
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机信息集合处理|Host information collection processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 3/12/2021
'* 1.1    8/1/2022   Modif New 
'**********************************
Friend Class PigDisk
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"

    Public ReadOnly Property Parent As PigHost
    Public DiskName As String

    Public Sub New(DiskName As String, Parent As PigHost)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
        Me.DiskName = DiskName
    End Sub

End Class
