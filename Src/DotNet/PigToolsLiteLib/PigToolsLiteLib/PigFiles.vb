'**********************************
'* Name: PigFiles
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigFile 的集合类|Collection class of PigFile
'* Home Url: https://en.seowphong.com
'* Version: 1.1
'* Create Time: 11/6/2023
'* 1.1  27/7/2024   Modify PigStepLog to StruStepLog
'************************************
Public Class PigFiles
	Inherits PigBaseMini
	Implements IEnumerable(Of PigFile)
	Private Const CLS_VERSION As String = "1" & "." & "1" & "." & "2"
	Private ReadOnly moList As New List(Of PigFile)
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
	Public Function GetEnumerator() As IEnumerator(Of PigFile) Implements IEnumerable(Of PigFile).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As PigFile
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FilePath As String) As PigFile
		Get
			Try
				Item = Nothing
				For Each oPigFile As PigFile In moList
					If oPigFile.FilePath = FilePath Then
						Item = oPigFile
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FilePath) As Boolean
		Try
			IsItemExists = False
			For Each oPigFile As PigFile In moList
				If oPigFile.FilePath = FilePath Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As PigFile)
		Try
			If Me.IsItemExists(NewItem.FilePath) = True Then Throw New Exception(NewItem.FilePath & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As PigFile)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FilePath As String) As PigFile
		Dim LOG As New StruStepLog : LOG.SubName = "AddOrGet"
		Try
			If Me.IsItemExists(FilePath) = True Then
				Return Me.Item(FilePath)
			Else
				Return Me.Add(FilePath)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FilePath As String) As PigFile
		Dim LOG As New StruStepLog : LOG.SubName = "Add"
		Try
			LOG.StepName = "New PigFile"
			Dim oPigFile As New PigFile(FilePath)
			If oPigFile.LastErr <> "" Then Throw New Exception(oPigFile.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oPigFile)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oPigFile
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FilePath)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oPigFile As PigFile In moList
				If oPigFile.FilePath = FilePath Then
					strStepName = "Remove " & FilePath
					moList.Remove(oPigFile)
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