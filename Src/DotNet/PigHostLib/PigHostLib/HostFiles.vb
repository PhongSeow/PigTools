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
'* Name: DBFiles
'* Author: Seow Phong
'* Describe: DBFile 的集合类|Collection class of DBFile
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 7/3/2023
'**********************************
Imports PigToolsLiteLib
Public Class HostFiles
	Inherits PigBaseLocal
	Implements IEnumerable(Of HostFile)
	Private Const CLS_VERSION As String = "1.0.0"
	Private ReadOnly moList As New List(Of HostFile)
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
	Public Function GetEnumerator() As IEnumerator(Of HostFile) Implements IEnumerable(Of HostFile).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As HostFile
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FileID As String) As HostFile
		Get
			Try
				Item = Nothing
				For Each oDBFile As HostFile In moList
					If oDBFile.FileID = FileID Then
						Item = oDBFile
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FileID) As Boolean
		Try
			IsItemExists = False
			For Each oDBFile As HostFile In moList
				If oDBFile.FileID = FileID Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As HostFile)
		Try
			If Me.IsItemExists(NewItem.FileID) = True Then Throw New Exception(NewItem.FileID & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As HostFile)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FileID As String) As HostFile
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FileID) = True Then
				Return Me.Item(FileID)
			Else
				Return Me.Add(FileID)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FileID As String) As HostFile
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New DBFile"
			Dim oDBFile As New HostFile(FileID)
			If oDBFile.LastErr <> "" Then Throw New Exception(oDBFile.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oDBFile)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oDBFile
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FileID)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oDBFile As HostFile In moList
				If oDBFile.FileID = FileID Then
					strStepName = "Remove " & FileID
					moList.Remove(oDBFile)
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
