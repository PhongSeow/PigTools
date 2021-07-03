'**********************************
'* Name: NoScanDirItem
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Collection class of NoScanDirItem
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.2
'* Create Time: 22/6/2021
'* 1.0.2  23/6/2021   Modify Add
'************************************
Public Class NoScanDirItems
    Inherits PigBaseMini
    Implements IEnumerable(Of NoScanDirItem)
    Private Const CLS_VERSION As String = "1.0.2"

    Private moList As New List(Of NoScanDirItem)

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
    Public Function GetEnumerator() As IEnumerator(Of NoScanDirItem) Implements IEnumerable(Of NoScanDirItem).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As NoScanDirItem
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(DirPath As String) As NoScanDirItem
        Get
            Try
                Item = Nothing
                For Each oNoScanDirItem As NoScanDirItem In moList
                    If oNoScanDirItem.DirPath = DirPath Then
                        Item = oNoScanDirItem
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.Name", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub Add(NewItem As NoScanDirItem)
        Try
            moList.Add(NewItem)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.NewItem", ex)
        End Try
    End Sub

    Public Function Add(DirPath As String) As NoScanDirItem
        Dim strStepName As String = ""
        Try
            strStepName = "New NoScanDirItem"
            Dim oNoScanDirItem As New NoScanDirItem(DirPath)
            If oNoScanDirItem.LastErr <> "" Then Throw New Exception(oNoScanDirItem.LastErr)
            strStepName = "Add"
            moList.Add(oNoScanDirItem)
            Add = oNoScanDirItem
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Add.Text", strStepName, ex)
            Return Nothing
        End Try
    End Function


    Public Sub Remove(DirPath As String)
        Dim strStepName As String = ""
        Try
            strStepName = "For Each"
            For Each oNoScanDirItem As NoScanDirItem In moList
                If oNoScanDirItem.DirPath = DirPath Then
                    strStepName = "Remove"
                    moList.Remove(oNoScanDirItem)
                    Exit For
                End If
            Next
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Name", strStepName, ex)
        End Try
    End Sub

    Public Sub Remove(Index As Integer)
        Try
            moList.RemoveAt(Index)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Remove.Name", ex)
        End Try
    End Sub
End Class
