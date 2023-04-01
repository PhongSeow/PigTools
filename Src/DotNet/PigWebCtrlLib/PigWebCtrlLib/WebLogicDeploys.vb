'**********************************
'* Name: WebLogicDeploys
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: WebLogicDomain 的集合类|Collection class of WebLogicDomain
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 28/2/2023
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
Public Class WebLogicDeploys
	Inherits PigBaseLocal
	Implements IEnumerable(Of WebLogicDeploy)
	Private Const CLS_VERSION As String = "1.0.2"
	Private ReadOnly moList As New List(Of WebLogicDeploy)
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
	Public Function GetEnumerator() As IEnumerator(Of WebLogicDeploy) Implements IEnumerable(Of WebLogicDeploy).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As WebLogicDeploy
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(DeployName As String) As WebLogicDeploy
		Get
			Try
				Item = Nothing
				For Each oWebLogicDeploy As WebLogicDeploy In moList
					If oWebLogicDeploy.DeployName = DeployName Then
						Item = oWebLogicDeploy
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(DeployName) As Boolean
		Try
			IsItemExists = False
			For Each oWebLogicDeploy As WebLogicDeploy In moList
				If oWebLogicDeploy.DeployName = DeployName Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As WebLogicDeploy)
		Try
			If Me.IsItemExists(NewItem.DeployName) = True Then Throw New Exception(NewItem.DeployName & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As WebLogicDeploy)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(DeployName As String, Parent As WebLogicDomain) As WebLogicDeploy
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(DeployName) = True Then
				Return Me.Item(DeployName)
			Else
				Return Me.Add(DeployName, Parent)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(DeployName As String, Parent As WebLogicDomain) As WebLogicDeploy
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New WebLogicDeploy"
			Dim oWebLogicDeploy As New WebLogicDeploy(DeployName, Parent)
			If oWebLogicDeploy.LastErr <> "" Then Throw New Exception(oWebLogicDeploy.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oWebLogicDeploy)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oWebLogicDeploy
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(DeployName)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oWebLogicDeploy As WebLogicDeploy In moList
				If oWebLogicDeploy.DeployName = DeployName Then
					strStepName = "Remove " & DeployName
					moList.Remove(oWebLogicDeploy)
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