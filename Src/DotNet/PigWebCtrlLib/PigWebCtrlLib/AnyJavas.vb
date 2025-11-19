'**********************************
'* Name: AnyJavas
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: AnyJava 的集合类|Collection class of AnyJava
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 19/4/2025
'************************************
Friend Class AnyJavas
    Inherits PigBaseLocal
    Implements IEnumerable(Of AnyJava)
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "2"
    Private ReadOnly moList As New List(Of AnyJava)

    Friend fParent As AnyJavaApp

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
    Public Function GetEnumerator() As IEnumerator(Of AnyJava) Implements IEnumerable(Of AnyJava).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As AnyJava
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(AppName As String) As AnyJava
        Get
            Try
                Item = Nothing
                For Each oAnyJava As AnyJava In moList
                    If oAnyJava.AppName = AppName Then
                        Item = oAnyJava
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.ConfName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(AppName) As Boolean
        Try
            IsItemExists = False
            For Each oAnyJava As AnyJava In moList
                If oAnyJava.AppName = AppName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As AnyJava) As String
        Try
            If Me.IsItemExists(NewItem.AppName) = True Then Throw New Exception(NewItem.AppName & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function


    Public Function Add(AppName As String) As AnyJava
        Dim LOG As New StruStepLog : LOG.SubName = "Add"
        Try
            LOG.StepName = "New AnyJava"
            Add = New AnyJava(AppName, Me.fParent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(AppName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(AppName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function




    Public Function Remove(AppName As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "Remove"
        Try
            LOG.StepName = "For Each"
            For Each oAnyJava As AnyJava In moList
                If oAnyJava.AppName = AppName Then
                    LOG.AddStepNameInf(AppName)
                    moList.Remove(oAnyJava)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Remove(Index As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "Remove"
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddOrGet(AppName As String) As AnyJava
        Dim LOG As New StruStepLog : LOG.SubName = "AddOrGet"
        Try
            If Me.IsItemExists(AppName) = True Then
                AddOrGet = Me.Item(AppName)
            Else
                AddOrGet = Me.Add(AppName)
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
