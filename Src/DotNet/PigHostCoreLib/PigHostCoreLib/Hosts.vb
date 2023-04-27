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
'* Name: PigHosts
'* Author: Seow Phong
'* Describe: PigHost 的集合类|Collection class of PigHost
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 15/10/2022
'* 1.1    18/4/2023   Modify Add
'**********************************
Imports PigToolsLiteLib
Public Class Hosts
    Inherits PigBaseLocal
    Implements IEnumerable(Of Host)
    Private Const CLS_VERSION As String = "1.1.2"
    Private ReadOnly moList As New List(Of Host)
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
    Public Function GetEnumerator() As IEnumerator(Of Host) Implements IEnumerable(Of Host).GetEnumerator
        Return moList.GetEnumerator()
    End Function
    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function
    Public ReadOnly Property Item(Index As Integer) As Host
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property Item(HostID As String) As Host
        Get
            Try
                Item = Nothing
                For Each oPigHost As Host In moList
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
            For Each oPigHost As Host In moList
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
    Private Sub mAdd(NewItem As Host)
        Try
            If Me.IsItemExists(NewItem.HostID) = True Then Throw New Exception(NewItem.HostID & "Already exists")
            moList.Add(NewItem)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mAdd", ex)
        End Try
    End Sub
    Public Sub Add(NewItem As Host, Parent As PigHostApp)
        Me.mAdd(NewItem)
    End Sub
    Public Function AddOrGet(HostID As String, Parent As PigHostApp) As Host
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(HostID) = True Then
                Return Me.Item(HostID)
            Else
                Return Me.Add(HostID, Parent)
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function
    Public Function Add(HostID As String, Parent As PigHostApp) As Host
        Dim LOG As New PigStepLog("Add")
        Try
            LOG.StepName = "New PigHost"
            Dim oPigHost As New Host(HostID, Parent)
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
            For Each oPigHost As Host In moList
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