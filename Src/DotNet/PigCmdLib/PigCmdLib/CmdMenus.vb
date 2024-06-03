'**********************************
'* Name: CmdMenus
'* Author: Seow Phong
'* License: Copyright (c) 2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: CmdMenus的集合类|Collection class of CmdMenu
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 5/3/2024
'**********************************
Imports PigToolsLiteLib
Friend Class CmdMenus
	Inherits PigBaseLocal
	Implements IEnumerable(Of CmdMenu)
	Private Const CLS_VERSION As String = "1.0.2"
	Private ReadOnly moList As New List(Of CmdMenu)
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
	Public Function GetEnumerator() As IEnumerator(Of CmdMenu) Implements IEnumerable(Of CmdMenu).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As CmdMenu
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(MenuKey As String) As CmdMenu
		Get
			Try
				Item = Nothing
				For Each oCmdMenu As CmdMenu In moList
					If oCmdMenu.MenuKey = MenuKey Then
						Item = oCmdMenu
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(MenuKey) As Boolean
		Try
			IsItemExists = False
			For Each oCmdMenu As CmdMenu In moList
				If oCmdMenu.MenuKey = MenuKey Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As CmdMenu)
		Try
			If Me.IsItemExists(NewItem.MenuKey) = True Then Throw New Exception(NewItem.MenuKey & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Function Add(Parent As PigCmdMenu, MenuType As CmdMenu.EnmMenuItemType, MenuKey As String, MenuText As String) As CmdMenu
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New CmdMenu"
			Dim oCmdMenu As New CmdMenu(Parent, MenuType, MenuKey, MenuText)
			If oCmdMenu.LastErr <> "" Then Throw New Exception(oCmdMenu.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oCmdMenu)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oCmdMenu
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Sub Remove(MenuKey As String)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oCmdMenu As CmdMenu In moList
				If oCmdMenu.MenuKey = MenuKey Then
					strStepName = "Remove " & MenuKey
					moList.Remove(oCmdMenu)
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