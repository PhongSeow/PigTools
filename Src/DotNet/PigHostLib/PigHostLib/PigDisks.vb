'**********************************
'* Name: PigDisks
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigDisk 的集合类|Collection class of PigDisk
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 3/12/2021
'* 1.1    10/12/2020   Modify Add 
'************************************
Imports PigToolsLiteLib

Friend Class PigDisks
    Inherits PigBaseMini
    Implements IEnumerable(Of PigDisk)
    Private Const CLS_VERSION As String = "1.1.6"
    Friend ReadOnly Property Parent As PigHost
    Private ReadOnly moList As New List(Of PigDisk)

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
    Public Function GetEnumerator() As IEnumerator(Of PigDisk) Implements IEnumerable(Of PigDisk).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigDisk
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(DiskName As String) As PigDisk
        Get
            Try
                Item = Nothing
                For Each oPigDisk As PigDisk In moList
                    If oPigDisk.DiskName = DiskName Then
                        Item = oPigDisk
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.DiskName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(DiskName) As Boolean
        Try
            IsItemExists = False
            For Each oPigDisk As PigDisk In moList
                If oPigDisk.DiskName = DiskName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigDisk) As String
        Try
            If Me.IsItemExists(NewItem.DiskName) = True Then Throw New Exception(NewItem.DiskName & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function





    Public Function Add(DiskName As String, Parent As PigHost) As PigDisk
        Dim LOG As New PigStepLog("Remove.DiskName")
        Try
            LOG.StepName = "New PigDisk"
            Add = New PigDisk(DiskName, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(DiskName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(DiskName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(DiskName As String) As String
        Dim LOG As New PigStepLog("Remove.DiskName")
        Try
            LOG.StepName = "For Each"
            For Each oPigDisk As PigDisk In moList
                If oPigDisk.DiskName = DiskName Then
                    LOG.AddStepNameInf(DiskName)
                    moList.Remove(oPigDisk)
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

    Public Function AddOrGet(DiskName As String) As PigDisk
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(DiskName) = True Then
                AddOrGet = Me.Item(DiskName)
            Else
                AddOrGet = Me.Add(DiskName, Me.Parent)
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
