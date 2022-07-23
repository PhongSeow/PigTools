'**********************************
'* Name: PigHosts
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigHost 的集合类|Collection class of PigHost
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 30/10/2021
'* 1.1    3/11/2021   Modify Add 
'* 1.2    3/1/2022   Modify Item 
'************************************
Imports PigToolsLiteLib

Public Class PigHosts
    Inherits PigBaseMini
    Implements IEnumerable(Of PigHost)
    Private Const CLS_VERSION As String = "1.2.5"
    Friend Property Parent As PigHostApp
    Private ReadOnly moList As New List(Of PigHost)

    Public Sub New(Parent As PigHostApp)
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
    Public Function GetEnumerator() As IEnumerator(Of PigHost) Implements IEnumerable(Of PigHost).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigHost
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(HostID As String) As PigHost
        Get
            Try
                Item = Nothing
                For Each oPigHost As PigHost In moList
                    If oPigHost.HostID = HostID Then
                        Item = oPigHost
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.HostID", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(HostID) As Boolean
        Try
            IsItemExists = False
            For Each oPigHost As PigHost In moList
                If oPigHost.HostID = HostID Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigHost) As String
        Try
            If Me.IsItemExists(NewItem.HostID) = True Then Throw New Exception(NewItem.HostID & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function



    Public Function Add(HostID As String) As PigHost
        Dim LOG As New PigStepLog("Remove.HostID")
        Try
            LOG.StepName = "New PigHost"
            Add = New PigHost(HostID)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(HostID)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(HostID)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(HostID As String) As String
        Dim LOG As New PigStepLog("Remove.HostID")
        Try
            LOG.StepName = "For Each"
            For Each oPigHost As PigHost In moList
                If oPigHost.HostID = HostID Then
                    LOG.AddStepNameInf(HostID)
                    moList.Remove(oPigHost)
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

    Public Function AddOrGet(HostID As String) As PigHost
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(HostID) = True Then
                AddOrGet = Me.Item(HostID)
            Else
                AddOrGet = Me.Add(HostID)
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
