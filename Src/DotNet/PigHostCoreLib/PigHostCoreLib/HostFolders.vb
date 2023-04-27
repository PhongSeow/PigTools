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
'* Name: HostFolders
'* Author: Seow Phong
'* Describe: HostFolder 的集合类|Collection class of HostFolder
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 7/3/2023
'* 1.1	8/3/2023	Modify New,Add
'* 1.2	13/3/2023	Modify Add
'* 1.3	4/4/2023	Modify Add
'* 1.5	18/4/2023	Modify Add,AddOrGet
'**********************************
Imports PigToolsLiteLib
Public Class HostFolders
	Inherits PigBaseLocal
	Implements IEnumerable(Of HostFolder)
	Private Const CLS_VERSION As String = "1.5.2"
	Private ReadOnly moList As New List(Of HostFolder)
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
	Public Function GetEnumerator() As IEnumerator(Of HostFolder) Implements IEnumerable(Of HostFolder).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As HostFolder
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FolderID As String) As HostFolder
		Get
			Try
				Item = Nothing
				For Each oHostFolder As HostFolder In moList
					If oHostFolder.FolderID = FolderID Then
						Item = oHostFolder
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FolderID) As Boolean
		Try
			IsItemExists = False
			For Each oHostFolder As HostFolder In moList
				If oHostFolder.FolderID = FolderID Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As HostFolder)
		Try
			If Me.IsItemExists(NewItem.FolderID) = True Then Throw New Exception(NewItem.FolderID & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As HostFolder)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FolderID As String, Parent As Host) As HostFolder
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FolderID) = True Then
				Return Me.Item(FolderID)
			Else
				Return Me.Add(FolderID, Parent)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(HostID As String, FolderPath As String, Parent As Host) As HostFolder
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New HostFolder"
			Dim oHostFolder As New HostFolder(HostID, FolderPath, Parent)
			If oHostFolder.LastErr <> "" Then Throw New Exception(oHostFolder.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oHostFolder)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oHostFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function

	Public Function Add(FolderID As String, Parent As Host) As HostFolder
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New HostFolder"
			Dim oHostFolder As New HostFolder(FolderID, Parent)
			If oHostFolder.LastErr <> "" Then Throw New Exception(oHostFolder.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oHostFolder)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oHostFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function

	Private Sub Remove(FolderID)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oHostFolder As HostFolder In moList
				If oHostFolder.FolderID = FolderID Then
					strStepName = "Remove " & FolderID
					moList.Remove(oHostFolder)
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
