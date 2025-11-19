'**********************************
'* Name: WeblogicServers
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: WeblogicServer 的集合类|Collection class of WeblogicServer
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 30/9/2024
'************************************
Imports PigToolsLiteLib
Public Class WeblogicServers
    Inherits PigBaseLocal
    Implements IEnumerable(Of WeblogicServer)
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "8"
    Private ReadOnly moList As New List(Of WeblogicServer)

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
    Public Function GetEnumerator() As IEnumerator(Of WeblogicServer) Implements IEnumerable(Of WeblogicServer).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As WeblogicServer
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(ServerName As String) As WeblogicServer
        Get
            Try
                Item = Nothing
                For Each oWeblogicServer As WeblogicServer In moList
                    If oWeblogicServer.ServerName = ServerName Then
                        Item = oWeblogicServer
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.ConfName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(ServerName) As Boolean
        Try
            IsItemExists = False
            For Each oWeblogicServer As WeblogicServer In moList
                If oWeblogicServer.ServerName = ServerName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As WeblogicServer) As String
        Try
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function


    Friend Function Add(ServerName As String, ListenPort As Integer, Parent As WebLogicDomain) As WeblogicServer
        Dim LOG As New StruStepLog : LOG.SubName = "Add"
        Try
            LOG.StepName = "New WeblogicServer"
            Add = New WeblogicServer(ServerName, ListenPort, Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(ServerName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(ServerName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function




    Friend Function Remove(ServerName As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "Remove"
        Try
            LOG.StepName = "For Each"
            For Each oWeblogicServer As WeblogicServer In moList
                If oWeblogicServer.ServerName = ServerName Then
                    LOG.AddStepNameInf(ServerName)
                    moList.Remove(oWeblogicServer)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function Remove(Index As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "Remove"
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function AddOrGet(ServerName As String, ListenPort As Integer, Parent As WebLogicDomain) As WeblogicServer
        Dim LOG As New StruStepLog : LOG.SubName = "AddOrGet"
        Try
            If Me.IsItemExists(ServerName) = True Then
                AddOrGet = Me.Item(ServerName)
            Else
                AddOrGet = Me.Add(ServerName, ListenPort, Parent)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Friend Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function

End Class

