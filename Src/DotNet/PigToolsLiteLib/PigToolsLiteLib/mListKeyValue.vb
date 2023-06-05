'**********************************
'* Name: 列表键值|mListKeyValue
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://en.seowphong.com
'* Version: 1.0
'* Create Time: 24/9/2022
'**********************************
Friend Class mListKeyValue
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.6"
    Public Sub New(KeyName As String, ValueBytes As Byte())
        MyBase.New(CLS_VERSION)
        Me.KeyName = KeyName
        Me.CreateTime = Now
        Me.KeyValue = ValueBytes
    End Sub

    Public ReadOnly Property KeyName As String
    Private mKeyValue(0) As Byte
    Public Property KeyValue As Byte()
        Get
            Return mKeyValue
        End Get
        Friend Set(value As Byte())
            ReDim mKeyValue(value.Length - 1)
            value.CopyTo(mKeyValue, 0)
        End Set
    End Property

    Public Property CreateTime As Date

End Class
