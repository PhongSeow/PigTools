'**********************************
'* Name: mListKeyValues
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: mListKeyValue 的集合类|Collection class of mListKeyValue
'* Home Url: https://en.seowphong.com
'* Version: 1.0
'* Create Time: 24/9/2022
'************************************
Friend Class mListKeyValues
    Inherits PigBaseMini
    Implements IEnumerable(Of mListKeyValue)
    Private Const CLS_VERSION As String = "1.0.2"
    Private ReadOnly moList As New List(Of mListKeyValue)

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of mListKeyValue) Implements IEnumerable(Of mListKeyValue).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As mListKeyValue
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(KeyName As String) As mListKeyValue
        Get
            Try
                Item = Nothing
                For Each omListKeyValue As mListKeyValue In moList
                    If omListKeyValue.KeyName = KeyName Then
                        Item = omListKeyValue
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.KeyName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(KeyName) As Boolean
        Try
            IsItemExists = False
            For Each omListKeyValue As mListKeyValue In moList
                If omListKeyValue.KeyName = KeyName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As mListKeyValue) As String
        Dim LOG As New PigStepLog("mAdd")
        Try
            If Me.IsItemExists(NewItem.KeyName) = True Then
                LOG.StepName = "Check IsItemExists"
                Throw New Exception(NewItem.KeyName & " already exists.")
            End If
            LOG.StepName = "List.Add"
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function Add(KeyName As String, ValueBytes As Byte()) As mListKeyValue
        Dim LOG As New PigStepLog("Add")
        Try
            LOG.StepName = "New mListKeyValue"
            Add = New mListKeyValue(KeyName, ValueBytes)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(KeyName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(KeyName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(KeyName As String) As String
        Dim LOG As New PigStepLog("Remove.KeyName")
        Try
            LOG.StepName = "For Each To Remove"
            For Each omListKeyValue As mListKeyValue In moList
                If omListKeyValue.KeyName = KeyName Then
                    LOG.AddStepNameInf(KeyName)
                    moList.Remove(omListKeyValue)
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

    Public Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function
End Class
