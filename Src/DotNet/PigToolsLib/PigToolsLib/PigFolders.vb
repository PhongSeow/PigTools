'**********************************
'* Name: PigFolders
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigFolder 的集合类|Collection class of PigFolder
'* Home Url: https://en.seowphong.com
'* Version: 1.0
'* Create Time: 11/6/2023
'************************************
Public Class PigFolders
	Inherits PigBaseMini
	Implements IEnumerable(Of PigFolder)
	Private Const CLS_VERSION As String = "1.0.0"
	Private ReadOnly moList As New List(Of PigFolder)
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
	Public Function GetEnumerator() As IEnumerator(Of PigFolder) Implements IEnumerable(Of PigFolder).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As PigFolder
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FolderPath As String) As PigFolder
		Get
			Try
				Item = Nothing
				For Each oPigFolder As PigFolder In moList
					If oPigFolder.FolderPath = FolderPath Then
						Item = oPigFolder
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FolderPath) As Boolean
		Try
			IsItemExists = False
			For Each oPigFolder As PigFolder In moList
				If oPigFolder.FolderPath = FolderPath Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As PigFolder)
		Try
			If Me.IsItemExists(NewItem.FolderPath) = True Then Throw New Exception(NewItem.FolderPath & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As PigFolder)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FolderPath As String) As PigFolder
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FolderPath) = True Then
				Return Me.Item(FolderPath)
			Else
				Return Me.Add(FolderPath)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FolderPath As String) As PigFolder
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New PigFolder"
			Dim oPigFolder As New PigFolder(FolderPath)
			If oPigFolder.LastErr <> "" Then Throw New Exception(oPigFolder.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oPigFolder)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oPigFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FolderPath)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oPigFolder As PigFolder In moList
				If oPigFolder.FolderPath = FolderPath Then
					strStepName = "Remove " & FolderPath
					moList.Remove(oPigFolder)
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