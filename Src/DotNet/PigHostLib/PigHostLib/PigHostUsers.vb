'**********************************
'* Name: PigHostUsers
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigHostUser 的集合类|Collection class of PigHostUser
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 8/11/2021
'* 1.1    2/12/2021   Modify Add 
'* 1.2    10/1/2022   Modify Add 
'************************************
Imports PigToolsLiteLib

Friend Class PigHostUsers
    Inherits PigBaseMini
    Implements IEnumerable(Of PigHostUser)
    Private Const CLS_VERSION As String = "1.2.6"
    Friend Property Parent As PigHost
    Private ReadOnly moList As New List(Of PigHostUser)

    Public Sub New(Parent As PigHost)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of PigHostUser) Implements IEnumerable(Of PigHostUser).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigHostUser
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(UserName As String) As PigHostUser
        Get
            Try
                Item = Nothing
                For Each oPigHostUser As PigHostUser In moList
                    If oPigHostUser.UserName = UserName Then
                        Item = oPigHostUser
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.UserName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(UserName) As Boolean
        Try
            IsItemExists = False
            For Each oPigHostUser As PigHostUser In moList
                If oPigHostUser.UserName = UserName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigHostUser) As String
        Try
            If Me.IsItemExists(NewItem.UserName) = True Then Throw New Exception(NewItem.UserName & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function




    Public Function Add(UserName As String) As PigHostUser
        Dim LOG As New PigStepLog("Remove.UserName")
        Try
            LOG.StepName = "New PigHostUser"
            Add = New PigHostUser(UserName, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(UserName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(UserName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(UserName As String) As String
        Dim LOG As New PigStepLog("Remove.UserName")
        Try
            LOG.StepName = "For Each"
            For Each oPigHostUser As PigHostUser In moList
                If oPigHostUser.UserName = UserName Then
                    LOG.AddStepNameInf(UserName)
                    moList.Remove(oPigHostUser)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Remove(Index As Integer) As String
        Dim LOG As New PigStepLog("Remove.Index")
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddOrGet(UserName As String) As PigHostUser
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(UserName) = True Then
                AddOrGet = Me.Item(UserName)
            Else
                AddOrGet = Me.Add(UserName)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function

End Class
