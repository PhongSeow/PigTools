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
'* Name: DBFileContent
'* Author: Seow Phong
'* Describe: 数据库文件内容类|Database file content class
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 7/3/2023
'* 1.1	12/10/2022	Modify Date Initial Time
'**********************************
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigToolsLiteLib
Friend Class HostFileContent
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.1.0"
	Public Sub New(FileContID As String)
		MyBase.New(CLS_VERSION)
		Me.FileContID = FileContID
	End Sub

	Public ReadOnly Property FileContID As String
	Private mUpdateCheck As New UpdateCheck
	Public ReadOnly Property LastUpdateTime() As Date
		Get
			Return mUpdateCheck.LastUpdateTime
		End Get
	End Property
	Public ReadOnly Property IsUpdate(PropertyName As String) As Boolean
		Get
			Return mUpdateCheck.IsUpdated(PropertyName)
		End Get
	End Property
	Public ReadOnly Property HasUpdated() As Boolean
		Get
			Return mUpdateCheck.HasUpdated
		End Get
	End Property
	Public Sub UpdateCheckClear()
		mUpdateCheck.Clear()
	End Sub
	Private mSize As Long
	Public Property Size() As Long
		Get
			Return mSize
		End Get
		Friend Set(value As Long)
			If value <> mSize Then
				Me.mUpdateCheck.Add("Size")
				mSize = value
			End If
		End Set
	End Property
	Private mContStatus As Integer
	Public Property ContStatus() As Integer
		Get
			Return mContStatus
		End Get
		Friend Set(value As Integer)
			If value <> mContStatus Then
				Me.mUpdateCheck.Add("ContStatus")
				mContStatus = value
			End If
		End Set
	End Property
	Private mSegContType As Integer
	Public Property SegContType() As Integer
		Get
			Return mSegContType
		End Get
		Friend Set(value As Integer)
			If value <> mSegContType Then
				Me.mUpdateCheck.Add("SegContType")
				mSegContType = value
			End If
		End Set
	End Property
	Private mSegSize As Integer
	Public Property SegSize() As Integer
		Get
			Return mSegSize
		End Get
		Friend Set(value As Integer)
			If value <> mSegSize Then
				Me.mUpdateCheck.Add("SegSize")
				mSegSize = value
			End If
		End Set
	End Property
	Private mCreateTime As DateTime = #1/1/1753#
	Public Property CreateTime() As DateTime
		Get
			Return mCreateTime
		End Get
		Friend Set(value As DateTime)
			If value <> mCreateTime Then
				Me.mUpdateCheck.Add("CreateTime")
				mCreateTime = value
			End If
		End Set
	End Property
	Private mUpdateTime As DateTime = #1/1/1753#
	Public Property UpdateTime() As DateTime
		Get
			Return mUpdateTime
		End Get
		Friend Set(value As DateTime)
			If value <> mUpdateTime Then
				Me.mUpdateCheck.Add("UpdateTime")
				mUpdateTime = value
			End If
		End Set
	End Property


	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("Size") = True Then
						If Me.Size <> .Item("Size").LngValue Then
							Me.Size = .Item("Size").LngValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("ContStatus") = True Then
						If Me.ContStatus <> .Item("ContStatus").IntValue Then
							Me.ContStatus = .Item("ContStatus").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("SegContType") = True Then
						If Me.SegContType <> .Item("SegContType").IntValue Then
							Me.SegContType = .Item("SegContType").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("SegSize") = True Then
						If Me.SegSize <> .Item("SegSize").IntValue Then
							Me.SegSize = .Item("SegSize").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("CreateTime") = True Then
						If Me.CreateTime <> .Item("CreateTime").DateValue Then
							Me.CreateTime = .Item("CreateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("UpdateTime") = True Then
						If Me.UpdateTime <> .Item("UpdateTime").DateValue Then
							Me.UpdateTime = .Item("UpdateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					Me.mUpdateCheck.Clear()
				End With
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("fFillByRs", ex)
		End Try
	End Function


	Friend Function fFillByXmlRs(ByRef InXmlRs As XmlRS, RSNo As Integer, RowNo As Integer, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If RowNo <= InXmlRs.TotalRows(RSNo) Then
				With InXmlRs
					If .IsColExists(RSNo, "Size") = True Then
						If Me.Size <> .LongValue(RSNo, RowNo, "Size") Then
							Me.Size = .LongValue(RSNo, RowNo, "Size")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "ContStatus") = True Then
						If Me.ContStatus <> .IntValue(RSNo, RowNo, "ContStatus") Then
							Me.ContStatus = .IntValue(RSNo, RowNo, "ContStatus")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "SegContType") = True Then
						If Me.SegContType <> .IntValue(RSNo, RowNo, "SegContType") Then
							Me.SegContType = .IntValue(RSNo, RowNo, "SegContType")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "SegSize") = True Then
						If Me.SegSize <> .IntValue(RSNo, RowNo, "SegSize") Then
							Me.SegSize = .IntValue(RSNo, RowNo, "SegSize")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "CreateTime") = True Then
						If Me.CreateTime <> .DateValue(RSNo, RowNo, "CreateTime") Then
							Me.CreateTime = .DateValue(RSNo, RowNo, "CreateTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "UpdateTime") = True Then
						If Me.UpdateTime <> .DateValue(RSNo, RowNo, "UpdateTime") Then
							Me.UpdateTime = .DateValue(RSNo, RowNo, "UpdateTime")
							UpdateCnt += 1
						End If
					End If
					Me.mUpdateCheck.Clear()
				End With
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("fFillByXmlRs", ex)
		End Try
	End Function


	Friend ReadOnly Property ValueMD5(Optional TextType As PigMD5.enmTextType = PigMD5.enmTextType.UTF8) As String
		Get
			Try
				Dim strText As String = ""
				With Me
					strText &= "<" & CStr(.Size) & ">"
					strText &= "<" & CStr(.ContStatus) & ">"
					strText &= "<" & CStr(.SegContType) & ">"
					strText &= "<" & CStr(.SegSize) & ">"
					strText &= "<" & Format(.CreateTime, "yyyy-MM-dd HH:mm:ss.fff") & ">"
					strText &= "<" & Format(.UpdateTime, "yyyy-MM-dd HH:mm:ss.fff") & ">"
				End With
				Dim oPigMD5 As New PigMD5(strText, TextType)
				ValueMD5 = oPigMD5.MD5
				oPigMD5 = Nothing
			Catch ex As Exception
				Me.SetSubErrInf("ValueMD5", ex)
				Return ""
			End Try
		End Get
	End Property

End Class
