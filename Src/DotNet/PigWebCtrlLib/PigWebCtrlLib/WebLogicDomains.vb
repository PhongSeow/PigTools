'**********************************
'* Name: WebLogicDomains
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: WebLogicDomain 的集合类|Collection class of WebLogicDomain
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 12/2/2022
'* 1.1    22/12/2022   Modify Add 
'* 1.2    26/5/2022   Modify Add 
'************************************
Imports PigToolsLiteLib
Public Class WebLogicDomains
    Inherits PigBaseMini
    Implements IEnumerable(Of WebLogicDomain)
    Private Const CLS_VERSION As String = "1.8.2"
    Private ReadOnly moList As New List(Of WebLogicDomain)

    Friend fParent As WebLogicApp

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
    Public Function GetEnumerator() As IEnumerator(Of WebLogicDomain) Implements IEnumerable(Of WebLogicDomain).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As WebLogicDomain
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(HomeDirPath As String) As WebLogicDomain
        Get
            Try
                Item = Nothing
                For Each oWebLogicDomain As WebLogicDomain In moList
                    If oWebLogicDomain.HomeDirPath = HomeDirPath Then
                        Item = oWebLogicDomain
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.ConfName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(HomeDirPath) As Boolean
        Try
            IsItemExists = False
            For Each oWebLogicDomain As WebLogicDomain In moList
                If oWebLogicDomain.HomeDirPath = HomeDirPath Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As WebLogicDomain) As String
        Try
            If Me.IsItemExists(NewItem.HomeDirPath) = True Then Throw New Exception(NewItem.HomeDirPath & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function



    Public Function Add(HomeDirPath As String) As WebLogicDomain
        Dim LOG As New PigStepLog("Remove.ConfName.ConfValue")
        Try
            LOG.StepName = "New WebLogicDomain"
            Add = New WebLogicDomain(HomeDirPath, Me.fParent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(HomeDirPath)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(HomeDirPath)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function




    Public Function Remove(HomeDirPath As String) As String
        Dim LOG As New PigStepLog("Remove.ConfName")
        Try
            LOG.StepName = "For Each"
            For Each oWebLogicDomain As WebLogicDomain In moList
                If oWebLogicDomain.HomeDirPath = HomeDirPath Then
                    LOG.AddStepNameInf(HomeDirPath)
                    moList.Remove(oWebLogicDomain)
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

    Public Function AddOrGet(HomeDirPath As String) As WebLogicDomain
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(HomeDirPath) = True Then
                AddOrGet = Me.Item(HomeDirPath)
            Else
                AddOrGet = Me.Add(HomeDirPath)
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

