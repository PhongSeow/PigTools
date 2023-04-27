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
'* Name: DBFileContents
'* Author: Seow Phong
'* Describe: DBFileContent 的集合类|Collection class of DBFileContent
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 7/3/2023
'**********************************
Imports PigToolsLiteLib
Friend Class HostFileContents
	Inherits PigBaseLocal
	Implements IEnumerable(Of HostFileContent)
	Private Const CLS_VERSION As String = "1.0.0"
	Private ReadOnly moList As New List(Of HostFileContent)
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
	Public Function GetEnumerator() As IEnumerator(Of HostFileContent) Implements IEnumerable(Of HostFileContent).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As HostFileContent
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FileContID As String) As HostFileContent
		Get
			Try
				Item = Nothing
				For Each oDBFileContent As HostFileContent In moList
					If oDBFileContent.FileContID = FileContID Then
						Item = oDBFileContent
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FileContID) As Boolean
		Try
			IsItemExists = False
			For Each oDBFileContent As HostFileContent In moList
				If oDBFileContent.FileContID = FileContID Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As HostFileContent)
		Try
			If Me.IsItemExists(NewItem.FileContID) = True Then Throw New Exception(NewItem.FileContID & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As HostFileContent)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FileContID As String) As HostFileContent
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FileContID) = True Then
				Return Me.Item(FileContID)
			Else
				Return Me.Add(FileContID)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FileContID As String) As HostFileContent
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New DBFileContent"
			Dim oDBFileContent As New HostFileContent(FileContID)
			If oDBFileContent.LastErr <> "" Then Throw New Exception(oDBFileContent.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oDBFileContent)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oDBFileContent
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FileContID)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oDBFileContent As HostFileContent In moList
				If oDBFileContent.FileContID = FileContID Then
					strStepName = "Remove " & FileContID
					moList.Remove(oDBFileContent)
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