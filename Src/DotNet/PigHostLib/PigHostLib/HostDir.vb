﻿'********************************************************************
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
'* Name: HostDir
'* Author: Seow Phong
'* Describe: 主机目录类|Host directory class
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 11/4/2023
'* 1.1	12/10/2022	Modify Date Initial Time
'* 1.2	13/10/2022	Add EnmMateStatus
'* 1.3	23/10/2022	Add HostFiles,fMergeHostFileInf, modify RefByHostFiles
'* 1.5	24/10/2022	Remove EnmMateStatus, add IsScan, modify RefByHostFiles
'********************************************************************

Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If


Public Class HostDir
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.5.12"

	Friend ReadOnly Property fParent As HostFolder
	Friend ReadOnly Property fPigFunc As New PigFunc
	Public ReadOnly Property HostFiles As New HostFiles


	Public ReadOnly Property FullDirPath As String
		Get
			Try
				FullDirPath = Me.fParent.FolderPath
				Select Case Right(FullDirPath, 1)
					Case "/", "\"
						FullDirPath = Left(FullDirPath, Len(FullDirPath) - 1)
				End Select
				If Left(Me.DirPath, 1) <> "." Then Throw New Exception("DirPath not relative path.")
				FullDirPath &= Mid(Me.DirPath, 2)
			Catch ex As Exception
				Me.SetSubErrInf("FullDirPath", ex)
				Return ""
			End Try
		End Get
	End Property

	Public Sub New(DirID As String, Parent As HostFolder)
		MyBase.New(CLS_VERSION)
		Me.fParent = Parent
		Me.DirID = DirID
	End Sub

	Public ReadOnly Property DirID As String
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
	Private mFolderID As String = ""
	Public Property FolderID() As String
		Get
			Return mFolderID
		End Get
		Friend Set(value As String)
			If value <> mFolderID Then
				Me.mUpdateCheck.Add("FolderID")
				mFolderID = value
			End If
		End Set
	End Property
	Private mDirPath As String = ""
	Public Property DirPath() As String
		Get
			Return mDirPath
		End Get
		Friend Set(value As String)
			If value <> mDirPath Then
				Me.mUpdateCheck.Add("DirPath")
				mDirPath = value
			End If
		End Set
	End Property
	Private mDirSize As Decimal
	Public Property DirSize() As Decimal
		Get
			Return mDirSize
		End Get
		Friend Set(value As Decimal)
			If value <> mDirSize Then
				Me.mUpdateCheck.Add("DirSize")
				mDirSize = value
			End If
		End Set
	End Property
	Private mDirFiles As Integer
	Public Property DirFiles() As Integer
		Get
			Return mDirFiles
		End Get
		Friend Set(value As Integer)
			If value <> mDirFiles Then
				Me.mUpdateCheck.Add("DirFiles")
				mDirFiles = value
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
	Private mDirUpdateTime As DateTime = #1/1/1753#
	Public Property DirUpdateTime() As DateTime
		Get
			Return mDirUpdateTime
		End Get
		Friend Set(value As DateTime)
			If value <> mDirUpdateTime Then
				Me.mUpdateCheck.Add("DirUpdateTime")
				mDirUpdateTime = value
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
	Private mIsScan As Boolean
	Public Property IsScan() As Boolean
		Get
			Return mIsScan
		End Get
		Friend Set(value As Boolean)
			If value <> mIsScan Then
				Me.mUpdateCheck.Add("IsScan")
				mIsScan = value
			End If
		End Set
	End Property


	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("FolderID") = True Then
						If Me.FolderID <> .Item("FolderID").StrValue Then
							Me.FolderID = .Item("FolderID").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("DirPath") = True Then
						If Me.DirPath <> .Item("DirPath").StrValue Then
							Me.DirPath = .Item("DirPath").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("DirSize") = True Then
						If Me.DirSize <> .Item("DirSize").DecValue Then
							Me.DirSize = .Item("DirSize").DecValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("DirFiles") = True Then
						If Me.DirFiles <> .Item("DirFiles").IntValue Then
							Me.DirFiles = .Item("DirFiles").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FastPigMD5") = True Then
						If Me.FastPigMD5 <> .Item("FastPigMD5").StrValue Then
							Me.FastPigMD5 = .Item("FastPigMD5").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("DirUpdateTime") = True Then
						If Me.DirUpdateTime <> .Item("DirUpdateTime").DateValue Then
							Me.DirUpdateTime = .Item("DirUpdateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("CreateTime") = True Then
						If Me.CreateTime <> .Item("CreateTime").DateValue Then
							Me.CreateTime = .Item("CreateTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("IsScan") = True Then
						If Me.IsScan <> .Item("IsScan").BooleanValue Then
							Me.IsScan = .Item("IsScan").BooleanValue
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
					If .IsColExists(RSNo, "FolderID") = True Then
						If Me.FolderID <> .StrValue(RSNo, RowNo, "FolderID") Then
							Me.FolderID = .StrValue(RSNo, RowNo, "FolderID")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "DirPath") = True Then
						If Me.DirPath <> .StrValue(RSNo, RowNo, "DirPath") Then
							Me.DirPath = .StrValue(RSNo, RowNo, "DirPath")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "DirSize") = True Then
						If Me.DirSize <> .DecValue(RSNo, RowNo, "DirSize") Then
							Me.DirSize = .DecValue(RSNo, RowNo, "DirSize")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "DirFiles") = True Then
						If Me.DirFiles <> .IntValue(RSNo, RowNo, "DirFiles") Then
							Me.DirFiles = .IntValue(RSNo, RowNo, "DirFiles")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FastPigMD5") = True Then
						If Me.FastPigMD5 <> .StrValue(RSNo, RowNo, "FastPigMD5") Then
							Me.FastPigMD5 = .StrValue(RSNo, RowNo, "FastPigMD5")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "DirUpdateTime") = True Then
						If Me.DirUpdateTime <> .DateValue(RSNo, RowNo, "DirUpdateTime") Then
							Me.DirUpdateTime = .DateValue(RSNo, RowNo, "DirUpdateTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "CreateTime") = True Then
						If Me.CreateTime <> .DateValue(RSNo, RowNo, "CreateTime") Then
							Me.CreateTime = .DateValue(RSNo, RowNo, "CreateTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "IsScan") = True Then
						If Me.IsScan <> .BooleanValue(RSNo, RowNo, "IsScan") Then
							Me.IsScan = .BooleanValue(RSNo, RowNo, "IsScan")
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

	Friend Function RefByHostFiles() As String
		Dim LOG As New PigStepLog("RefByHostFiles")
		Try
			With Me
				.DirSize = 0
				.FastPigMD5 = ""
				.DirFiles = Me.HostFiles.Count
			End With
			Dim lngCnt As Long = 0, strFastPigMD5 As String = ""
			For Each oDBFile As HostFile In Me.HostFiles
				Me.DirSize += oDBFile.FileSize
				strFastPigMD5 &= oDBFile.FastPigMD5
				If lngCnt Mod 1000 = 0 Then
					Me.fPigFunc.GetTextPigMD5(strFastPigMD5, PigMD5.enmTextType.UTF8, Me.FastPigMD5)
				End If
				lngCnt += 1
			Next
			Me.fPigFunc.GetTextPigMD5(strFastPigMD5, PigMD5.enmTextType.UTF8, Me.FastPigMD5)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Friend Function fMergeHostFileInf() As String
		Return Me.fParent.fParent.fParent.fMergeHostFileInf(Me)
	End Function

End Class