'**********************************
'* Name: PigHosts
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigHost 的集合类|Collection class of PigHost
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 15/10/2022
'**********************************
Imports PigToolsLiteLib
Public Class PigHosts
    Inherits PigBaseLocal
    Implements IEnumerable(Of PigHost)
    Private Const CLS_VERSION As String = "1.0.0"
    Private ReadOnly moList As New List(Of PigHost)
    Public Sub New()
        MyBase.New(CLS_VERSION)
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
                Me.SetSubErrInf("Item.Key", ex)
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
    Private Sub mAdd(NewItem As PigHost)
        Try
            If Me.IsItemExists(NewItem.HostID) = True Then Throw New Exception(NewItem.HostID & "Already exists")
            moList.Add(NewItem)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mAdd", ex)
        End Try
    End Sub
    Public Sub Add(NewItem As PigHost)
        Me.mAdd(NewItem)
    End Sub
    Public Function AddOrGet(HostID As String) As PigHost
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(HostID) = True Then
                Return Me.Item(HostID)
            Else
                Return Me.Add(HostID)
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function
    Public Function Add(HostID As String) As PigHost
        Dim LOG As New PigStepLog("Add")
        Try
            LOG.StepName = "New PigHost"
            Dim oPigHost As New PigHost(HostID)
            If oPigHost.LastErr <> "" Then Throw New Exception(oPigHost.LastErr)
            LOG.StepName = "mAdd"
            Me.mAdd(oPigHost)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Add = oPigHost
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function
    Private Sub Remove(HostID)
        Dim strStepName As String = ""
        Try
            strStepName = "For Each"
            For Each oPigHost As PigHost In moList
                If oPigHost.HostID = HostID Then
                    strStepName = "Remove " & HostID
                    moList.Remove(oPigHost)
                    Exit For
                End If
            Next
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Key", strStepName, ex)
        End Try
    End Sub
    Public Sub Remove(Index As Integer)
        Dim strStepName As String = ""
        Try
            strStepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Index", strStepName, ex)
        End Try
    End Sub
    Public Sub Clear()
        Try
            moList.Clear()
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Clear", ex)
        End Try
    End Sub
End Class