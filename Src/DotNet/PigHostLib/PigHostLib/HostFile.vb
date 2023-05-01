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
'* Name: DBFile
'* Author: Seow Phong
'* Describe: 数据库文件类|Database file class
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6
'* Create Time: 7/3/2023
'* 1.1	8/3/2023	Modify New
'* 1.2	8/3/2023	Modify fFillByXmlRs,fFillByRs
'* 1.3	8/3/2023	Modify fFillByXmlRs,fFillByRs
'* 1.5	12/10/2022	Modify Date Initial Time
'* 1.6	24/3/2023	Change EnmContStatus to IsCheck
'**********************************
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigToolsLiteLib
Public Class HostFile
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.6.3"


	Friend ReadOnly Property HostFileSegments As New HostFileSegments
	Public Sub New(FileID As String)
		MyBase.New(CLS_VERSION)
		Me.FileID = FileID
	End Sub

	Public ReadOnly Property FileID As String
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
	Private mFileName As String = ""
	Public Property FileName() As String
		Get
			Return mFileName
		End Get
		Friend Set(value As String)
			If value <> mFileName Then
				Me.mUpdateCheck.Add("FileName")
				mFileName = value
			End If
		End Set
	End Property
	Private mFileContID As String = ""
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
	Private mFileSize As Long
	Public Property FileSize() As Long
		Get
			Return mFileSize
		End Get
		Friend Set(value As Long)
			If value <> mFileSize Then
				Me.mUpdateCheck.Add("FileSize")
				mFileSize = value
			End If
		End Set
	End Property
	Private mFastPigMD5 As String = ""
	Public Property FastPigMD5() As String
		Get
			Return mFastPigMD5
		End Get
		Friend Set(value As String)
			If value <> mFastPigMD5 Then
				Me.mUpdateCheck.Add("FastPigMD5")
				mFastPigMD5 = value
			End If
		End Set
	End Property
	Private mFileUpdateTime As DateTime = #1/1/1753#
	Public Property FileUpdateTime() As DateTime
		Get
			Return mFileUpdateTime
		End Get
		Friend Set(value As DateTime)
			If value <> mFileUpdateTime Then
				Me.mUpdateCheck.Add("FileUpdateTime")
				mFileUpdateTime = value
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
	Private mIsCheck As Boolean
	Public Property IsCheck() As Boolean
		Get
			Return mIsCheck
		End Get
		Friend Set(value As Boolean)
			If value <> mIsCheck Then
				Me.mUpdateCheck.Add("IsCheck")
				mIsCheck = value
			End If
		End Set
	End Property

	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("FileName") = True Then
						If Me.FileName <> .Item("FileName").StrValue Then
							Me.FileName = .Item("FileName").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FileContID") = True Then
						If Me.FileContID <> .Item("FileContID").StrValue Then
							Me.FileContID = .Item("FileContID").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FileSize") = True Then
						If Me.FileSize <> .Item("FileSize").LngValue Then
							Me.FileSize = .Item("FileSize").LngValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FastPigMD5") = True Then
						If Me.FastPigMD5 <> .Item("FastPigMD5").StrValue Then
							Me.FastPigMD5 = .Item("FastPigMD5").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FileUpdateTime") = True Then
						If Me.FileUpdateTime <> .Item("FileUpdateTime").DateValue Then
							Me.FileUpdateTime = .Item("FileUpdateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("CreateTime") = True Then
						If Me.CreateTime <> .Item("CreateTime").DateValue Then
							Me.CreateTime = .Item("CreateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("IsCheck") = True Then
						If Me.IsCheck <> .Item("IsCheck").BooleanValue Then
							Me.IsCheck = .Item("IsCheck").BooleanValue
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
					If .IsColExists(RSNo, "FileName") = True Then
						If Me.FileName <> .StrValue(RSNo, RowNo, "FileName") Then
							Me.FileName = .StrValue(RSNo, RowNo, "FileName")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FileContID") = True Then
						If Me.FileContID <> .StrValue(RSNo, RowNo, "FileContID") Then
							Me.FileContID = .StrValue(RSNo, RowNo, "FileContID")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FileSize") = True Then
						If Me.FileSize <> .LongValue(RSNo, RowNo, "FileSize") Then
							Me.FileSize = .LongValue(RSNo, RowNo, "FileSize")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FastPigMD5") = True Then
						If Me.FastPigMD5 <> .StrValue(RSNo, RowNo, "FastPigMD5") Then
							Me.FastPigMD5 = .StrValue(RSNo, RowNo, "FastPigMD5")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FileUpdateTime") = True Then
						If Me.FileUpdateTime <> .DateValue(RSNo, RowNo, "FileUpdateTime") Then
							Me.FileUpdateTime = .DateValue(RSNo, RowNo, "FileUpdateTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "CreateTime") = True Then
						If Me.CreateTime <> .DateValue(RSNo, RowNo, "CreateTime") Then
							Me.CreateTime = .DateValue(RSNo, RowNo, "CreateTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "IsCheck") = True Then
						If Me.IsCheck <> .BooleanValue(RSNo, RowNo, "IsCheck") Then
							Me.IsCheck = .BooleanValue(RSNo, RowNo, "IsCheck")
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

End Class
