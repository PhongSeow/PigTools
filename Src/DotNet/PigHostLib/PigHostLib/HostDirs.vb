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
'* Name: DBDirs
'* Author: Seow Phong
'* Describe: DBDir 的集合类|Collection class of DBDir
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 11/4/2023
'**********************************
Imports PigObjFsLib
Imports PigToolsLiteLib
Public Class HostDirs
	Inherits PigBaseLocal
	Implements IEnumerable(Of HostDir)
	Private Const CLS_VERSION As String = "1.0.0"
	Private ReadOnly moList As New List(Of HostDir)
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
	Public Function GetEnumerator() As IEnumerator(Of HostDir) Implements IEnumerable(Of HostDir).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As HostDir
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(DirID As String) As HostDir
		Get
			Try
				Item = Nothing
				For Each oDFDir As HostDir In moList
					If oDFDir.DirID = DirID Then
						Item = oDFDir
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(DirID) As Boolean
		Try
			IsItemExists = False
			For Each oDFDir As HostDir In moList
				If oDFDir.DirID = DirID Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As HostDir)
		Try
			If Me.IsItemExists(NewItem.DirID) = True Then Throw New Exception(NewItem.DirID & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Function AddOrGet(DirID As String, Parent As HostFolder) As HostDir
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(DirID) = True Then
				Return Me.Item(DirID)
			Else
				Return Me.Add(DirID, Parent)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(DirID As String, Parent As HostFolder) As HostDir
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New DFDir"
			Dim oDFDir As New HostDir(DirID, Parent)
			If oDFDir.LastErr <> "" Then Throw New Exception(oDFDir.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oDFDir)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oDFDir
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(DirID)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oDFDir As HostDir In moList
				If oDFDir.DirID = DirID Then
					strStepName = "Remove " & DirID
					moList.Remove(oDFDir)
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