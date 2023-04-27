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
'* Name: DBFileSegments
'* Author: Seow Phong
'* Describe: DBFileSegment 的集合类|Collection class of DBFileSegment
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 7/3/2023
'**********************************
Imports PigToolsLiteLib
Friend Class HostFileSegments
	Inherits PigBaseLocal
	Implements IEnumerable(Of HostFileSegment)
	Private Const CLS_VERSION As String = "1.0.0"
	Private ReadOnly moList As New List(Of HostFileSegment)
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
	Public Function GetEnumerator() As IEnumerator(Of HostFileSegment) Implements IEnumerable(Of HostFileSegment).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As HostFileSegment
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FileSegID As String) As HostFileSegment
		Get
			Try
				Item = Nothing
				For Each oDBFileSegment As HostFileSegment In moList
					If oDBFileSegment.FileSegID = FileSegID Then
						Item = oDBFileSegment
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FileSegID) As Boolean
		Try
			IsItemExists = False
			For Each oDBFileSegment As HostFileSegment In moList
				If oDBFileSegment.FileSegID = FileSegID Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As HostFileSegment)
		Try
			If Me.IsItemExists(NewItem.FileSegID) = True Then Throw New Exception(NewItem.FileSegID & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As HostFileSegment)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FileSegID As String) As HostFileSegment
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FileSegID) = True Then
				Return Me.Item(FileSegID)
			Else
				Return Me.Add(FileSegID)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FileSegID As String) As HostFileSegment
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New DBFileSegment"
			Dim oDBFileSegment As New HostFileSegment(FileSegID)
			If oDBFileSegment.LastErr <> "" Then Throw New Exception(oDBFileSegment.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oDBFileSegment)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oDBFileSegment
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FileSegID)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oDBFileSegment As HostFileSegment In moList
				If oDBFileSegment.FileSegID = FileSegID Then
					strStepName = "Remove " & FileSegID
					moList.Remove(oDBFileSegment)
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
