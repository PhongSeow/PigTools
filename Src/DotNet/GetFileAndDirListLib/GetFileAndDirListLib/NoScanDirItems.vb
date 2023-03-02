'**********************************
'* Name: NoScanDirItem
'* Author: Seow Phong
'* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Collection class of NoScanDirItem
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 22/6/2021
'* 1.0.2  23/6/2021   Modify Add
'* 1.1  1/3/2023   Rewrite code
'************************************
Imports PigToolsLiteLib
Public Class NoScanDirItems
	Inherits PigBaseLocal
	Implements IEnumerable(Of NoScanDirItem)
	Private Const CLS_VERSION As String = "1.1.2"
	Private ReadOnly moList As New List(Of NoScanDirItem)
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
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(DirPath) As Boolean
		Try
			IsItemExists = False
			For Each oNoScanDirItem As NoScanDirItem In moList
				If oNoScanDirItem.DirPath = DirPath Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As NoScanDirItem)
		Try
			If Me.IsItemExists(NewItem.DirPath) = True Then Throw New Exception(NewItem.DirPath & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As NoScanDirItem)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(DirPath As String) As NoScanDirItem
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(DirPath) = True Then
				Return Me.Item(DirPath)
			Else
				Return Me.Add(DirPath)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(DirPath As String) As NoScanDirItem
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New NoScanDirItem"
			Dim oNoScanDirItem As New NoScanDirItem(DirPath)
			If oNoScanDirItem.LastErr <> "" Then Throw New Exception(oNoScanDirItem.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oNoScanDirItem)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oNoScanDirItem
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(DirPath)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oNoScanDirItem As NoScanDirItem In moList
				If oNoScanDirItem.DirPath = DirPath Then
					strStepName = "Remove " & DirPath
					moList.Remove(oNoScanDirItem)
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