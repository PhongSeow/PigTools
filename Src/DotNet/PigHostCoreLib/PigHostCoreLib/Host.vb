'********************************************************************
'* Copyright 2021-2023 Seow Phong
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
'* Name: 豚豚主机|PigHost
'* Author: Seow Phong
'* Describe: 主机信息处理|Host information processing
'* Home Url: https://en.seowphong.com
'* Version: 1.15
'* Create Time: 8/10/2021
'* 1.1    12/10/2021   Add New
'* 1.2    18/10/2021   Modify New
'* 1.3    23/7/2022   Modify New
'* 1.5    24/7/2022   Rename MyHost,
'* 1.6    26/7/2022   Modify Imports
'* 1.7    29/7/2022   Modify Imports
'* 1.8    4/10/2022   Rename PigHost, add mNew, modify New
'* 1.9    15/10/2022  Rename PigHost, add mNew, modify New
'* 1.10   17/4/2023  Rename PigHost, add mNew, modify New
'* 1.11   18/4/2023  Regenerate partial code automatically
'* 1.12   21/4/2023  Modify RefHostFolders
'* 1.13   21/4/2023  Modify New
'* 1.15   29/4/2023  Modify RefHostFolders
'********************************************************************
Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigCmdLib

Public Class Host
    Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.15.6"

	''' <summary>
	''' 主机名|host name
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property HostID As String
    ''' <summary>
    ''' 操作系统名称|Operating system name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property OSCaption As String
    Public ReadOnly Property UUID As String

    Private Property mPigFunc As New PigFunc
    Private Property mPigSysCmd As New PigSysCmd
    Friend ReadOnly Property fParent As PigHostApp

    Public ReadOnly Property HostFolders As New HostFolders

    Public Sub New(HostID As String, Parent As PigHostApp)
        MyBase.New(CLS_VERSION)
		Me.HostID = HostID
		Me.fParent = Parent
	End Sub


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
	Private mHostName As String = ""
	Public Property HostName() As String
		Get
			Return mHostName
		End Get
		Friend Set(value As String)
			If value <> mHostName Then
				Me.mUpdateCheck.Add("HostName")
				mHostName = value
			End If
		End Set
	End Property
	Private mHostMainIp As String = ""
	Public Property HostMainIp() As String
		Get
			Return mHostMainIp
		End Get
		Friend Set(value As String)
			If value <> mHostMainIp Then
				Me.mUpdateCheck.Add("HostMainIp")
				mHostMainIp = value
			End If
		End Set
	End Property
	Private mIsUse As Boolean
	Public Property IsUse() As Boolean
		Get
			Return mIsUse
		End Get
		Friend Set(value As Boolean)
			If value <> mIsUse Then
				Me.mUpdateCheck.Add("IsUse")
				mIsUse = value
			End If
		End Set
	End Property
	Private mHostDesc As String = ""
	Public Property HostDesc() As String
		Get
			Return mHostDesc
		End Get
		Friend Set(value As String)
			If value <> mHostDesc Then
				Me.mUpdateCheck.Add("HostDesc")
				mHostDesc = value
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
	Private mScanStatus As Integer
	Public Property ScanStatus() As Integer
		Get
			Return mScanStatus
		End Get
		Friend Set(value As Integer)
			If value <> mScanStatus Then
				Me.mUpdateCheck.Add("ScanStatus")
				mScanStatus = value
			End If
		End Set
	End Property
	Private mStaticInf As String = ""
	Public Property StaticInf() As String
		Get
			Return mStaticInf
		End Get
		Friend Set(value As String)
			If value <> mStaticInf Then
				Me.mUpdateCheck.Add("StaticInf")
				mStaticInf = value
			End If
		End Set
	End Property
	Private mActiveInf As String = ""
	Public Property ActiveInf() As String
		Get
			Return mActiveInf
		End Get
		Friend Set(value As String)
			If value <> mActiveInf Then
				Me.mUpdateCheck.Add("ActiveInf")
				mActiveInf = value
			End If
		End Set
	End Property
	Private mScanBeginTime As DateTime = #1/1/1753#
	Public Property ScanBeginTime() As DateTime
		Get
			Return mScanBeginTime
		End Get
		Friend Set(value As DateTime)
			If value <> mScanBeginTime Then
				Me.mUpdateCheck.Add("ScanBeginTime")
				mScanBeginTime = value
			End If
		End Set
	End Property
	Private mScanEndTime As DateTime = #1/1/1753#
	Public Property ScanEndTime() As DateTime
		Get
			Return mScanEndTime
		End Get
		Friend Set(value As DateTime)
			If value <> mScanEndTime Then
				Me.mUpdateCheck.Add("ScanEndTime")
				mScanEndTime = value
			End If
		End Set
	End Property
	Private mActiveTime As DateTime = #1/1/1753#
	Public Property ActiveTime() As DateTime
		Get
			Return mActiveTime
		End Get
		Friend Set(value As DateTime)
			If value <> mActiveTime Then
				Me.mUpdateCheck.Add("ActiveTime")
				mActiveTime = value
			End If
		End Set
	End Property


	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("HostName") = True Then
						If Me.HostName <> .Item("HostName").StrValue Then
							Me.HostName = .Item("HostName").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("HostMainIp") = True Then
						If Me.HostMainIp <> .Item("HostMainIp").StrValue Then
							Me.HostMainIp = .Item("HostMainIp").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("IsUse") = True Then
						If Me.IsUse <> .Item("IsUse").BooleanValue Then
							Me.IsUse = .Item("IsUse").BooleanValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("HostDesc") = True Then
						If Me.HostDesc <> .Item("HostDesc").StrValue Then
							Me.HostDesc = .Item("HostDesc").StrValue
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
					If .IsItemExists("ScanStatus") = True Then
						If Me.ScanStatus <> .Item("ScanStatus").IntValue Then
							Me.ScanStatus = .Item("ScanStatus").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("StaticInf") = True Then
						If Me.StaticInf <> .Item("StaticInf").StrValue Then
							Me.StaticInf = .Item("StaticInf").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("ActiveInf") = True Then
						If Me.ActiveInf <> .Item("ActiveInf").StrValue Then
							Me.ActiveInf = .Item("ActiveInf").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("ScanBeginTime") = True Then
						If Me.ScanBeginTime <> .Item("ScanBeginTime").DateValue Then
							Me.ScanBeginTime = .Item("ScanBeginTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("ScanEndTime") = True Then
						If Me.ScanEndTime <> .Item("ScanEndTime").DateValue Then
							Me.ScanEndTime = .Item("ScanEndTime").DateValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("ActiveTime") = True Then
						If Me.ActiveTime <> .Item("ActiveTime").DateValue Then
							Me.ActiveTime = .Item("ActiveTime").DateValue
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
					If .IsColExists(RSNo, "HostName") = True Then
						If Me.HostName <> .StrValue(RSNo, RowNo, "HostName") Then
							Me.HostName = .StrValue(RSNo, RowNo, "HostName")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "HostMainIp") = True Then
						If Me.HostMainIp <> .StrValue(RSNo, RowNo, "HostMainIp") Then
							Me.HostMainIp = .StrValue(RSNo, RowNo, "HostMainIp")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "IsUse") = True Then
						If Me.IsUse <> .BooleanValue(RSNo, RowNo, "IsUse") Then
							Me.IsUse = .BooleanValue(RSNo, RowNo, "IsUse")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "HostDesc") = True Then
						If Me.HostDesc <> .StrValue(RSNo, RowNo, "HostDesc") Then
							Me.HostDesc = .StrValue(RSNo, RowNo, "HostDesc")
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
					If .IsColExists(RSNo, "ScanStatus") = True Then
						If Me.ScanStatus <> .IntValue(RSNo, RowNo, "ScanStatus") Then
							Me.ScanStatus = .IntValue(RSNo, RowNo, "ScanStatus")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "StaticInf") = True Then
						If Me.StaticInf <> .StrValue(RSNo, RowNo, "StaticInf") Then
							Me.StaticInf = .StrValue(RSNo, RowNo, "StaticInf")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "ActiveInf") = True Then
						If Me.ActiveInf <> .StrValue(RSNo, RowNo, "ActiveInf") Then
							Me.ActiveInf = .StrValue(RSNo, RowNo, "ActiveInf")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "ScanBeginTime") = True Then
						If Me.ScanBeginTime <> .DateValue(RSNo, RowNo, "ScanBeginTime") Then
							Me.ScanBeginTime = .DateValue(RSNo, RowNo, "ScanBeginTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "ScanEndTime") = True Then
						If Me.ScanEndTime <> .DateValue(RSNo, RowNo, "ScanEndTime") Then
							Me.ScanEndTime = .DateValue(RSNo, RowNo, "ScanEndTime")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "ActiveTime") = True Then
						If Me.ActiveTime <> .DateValue(RSNo, RowNo, "ActiveTime") Then
							Me.ActiveTime = .DateValue(RSNo, RowNo, "ActiveTime")
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
	Public Function RefHostFolders(Optional IsDirtyRead As Boolean = True, Optional IsShowIsUseOnly As Boolean = False) As String
		Try
			Return Me.fParent.fRefHostFolders(Me, IsDirtyRead, IsShowIsUseOnly)
		Catch ex As Exception
			Return Me.GetSubErrInf("RefHostFolders", ex)
		End Try
	End Function

End Class
