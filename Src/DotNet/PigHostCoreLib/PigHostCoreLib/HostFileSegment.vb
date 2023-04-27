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
'* Name: DBFileSegment
'* Author: Seow Phong
'* Describe: 数据库文件段内容类|Database file segment content class
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 7/3/2023
'* 1.1	12/10/2022	Modify Date Initial Time
'**********************************
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigToolsLiteLib

Friend Class HostFileSegment
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.1.0"
	Public Sub New(FileSegID As String)
		MyBase.New(CLS_VERSION)
		Me.FileSegID = FileSegID
	End Sub

	Public ReadOnly Property FileSegID As String
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
	Private mFileContID As String
	Public Property FileContID() As String
		Get
			Return mFileContID
		End Get
		Friend Set(value As String)
			If value <> mFileContID Then
				Me.mUpdateCheck.Add("FileContID")
				mFileContID = value
			End If
		End Set
	End Property
	Private mSegCont As String
	Public Property SegCont() As String
		Get
			Return mSegCont
		End Get
		Friend Set(value As String)
			If value <> mSegCont Then
				Me.mUpdateCheck.Add("SegCont")
				mSegCont = value
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


	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("FileContID") = True Then
						If Me.FileContID <> .Item("FileContID").StrValue Then
							Me.FileContID = .Item("FileContID").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("SegCont") = True Then
						If Me.SegCont <> .Item("SegCont").StrValue Then
							Me.SegCont = .Item("SegCont").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("CreateTime") = True Then
						If Me.CreateTime <> .Item("CreateTime").DateValue Then
							Me.CreateTime = .Item("CreateTime").DateValue
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
					If .IsColExists(RSNo, "FileContID") = True Then
						If Me.FileContID <> .StrValue(RSNo, RowNo, "FileContID") Then
							Me.FileContID = .StrValue(RSNo, RowNo, "FileContID")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "SegCont") = True Then
						If Me.SegCont <> .StrValue(RSNo, RowNo, "SegCont") Then
							Me.SegCont = .StrValue(RSNo, RowNo, "SegCont")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "CreateTime") = True Then
						If Me.CreateTime <> .DateValue(RSNo, RowNo, "CreateTime") Then
							Me.CreateTime = .DateValue(RSNo, RowNo, "CreateTime")
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
					strText &= "<" & .FileContID & ">"
					strText &= "<" & .SegCont & ">"
					strText &= "<" & Format(.CreateTime, "yyyy-MM-dd HH:mm:ss.fff") & ">"
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
