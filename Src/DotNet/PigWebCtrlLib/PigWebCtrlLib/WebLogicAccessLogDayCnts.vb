'**********************************
'* Name: WebLogicAccessLogDayCnts
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: WebLogicAccessLogDayCnt 的集合类|Collection class of WebLogicAccessLogDayCnt
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 9/11/2023
'* 1.1  28/7/2024   Modify PigStepLog to StruStepLog
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
Friend Class WebLogicAccessLogDayCnts
	Inherits PigBaseLocal
	Implements IEnumerable(Of WebLogicAccessLogDayCnt)
	Private Const CLS_VERSION As String = "1" & "." & "1" & "." & "2"
	Private ReadOnly moList As New List(Of WebLogicAccessLogDayCnt)
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
	Public Function GetEnumerator() As IEnumerator(Of WebLogicAccessLogDayCnt) Implements IEnumerable(Of WebLogicAccessLogDayCnt).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As WebLogicAccessLogDayCnt
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(DayNo As String) As WebLogicAccessLogDayCnt
		Get
			Try
				Item = Nothing
				For Each oWebLogicAccessLogDayCnt As WebLogicAccessLogDayCnt In moList
					If oWebLogicAccessLogDayCnt.DayNo = DayNo Then
						Item = oWebLogicAccessLogDayCnt
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(DayNo) As Boolean
		Try
			IsItemExists = False
			For Each oWebLogicAccessLogDayCnt As WebLogicAccessLogDayCnt In moList
				If oWebLogicAccessLogDayCnt.DayNo = DayNo Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As WebLogicAccessLogDayCnt)
		Try
			If Me.IsItemExists(NewItem.DayNo) = True Then Throw New Exception(NewItem.DayNo & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As WebLogicAccessLogDayCnt)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(DayNo As String, DayType As WebLogicAccessLogDayCnt.EnmDayType) As WebLogicAccessLogDayCnt
		Dim LOG As New StruStepLog : LOG.SubName = "AddOrGet"
		Try
			If Me.IsItemExists(DayNo) = True Then
				Return Me.Item(DayNo)
			Else
				Return Me.Add(DayNo, DayType)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(DayNo As String, DayType As WebLogicAccessLogDayCnt.EnmDayType) As WebLogicAccessLogDayCnt
		Dim LOG As New StruStepLog : LOG.SubName = "Add"
		Try
			LOG.StepName = "New WebLogicAccessLogDayCnt"
			Dim oWebLogicAccessLogDayCnt As New WebLogicAccessLogDayCnt(DayNo, DayType)
			If oWebLogicAccessLogDayCnt.LastErr <> "" Then Throw New Exception(oWebLogicAccessLogDayCnt.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oWebLogicAccessLogDayCnt)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oWebLogicAccessLogDayCnt
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(DayNo)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oWebLogicAccessLogDayCnt As WebLogicAccessLogDayCnt In moList
				If oWebLogicAccessLogDayCnt.DayNo = DayNo Then
					strStepName = "Remove " & DayNo
					moList.Remove(oWebLogicAccessLogDayCnt)
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