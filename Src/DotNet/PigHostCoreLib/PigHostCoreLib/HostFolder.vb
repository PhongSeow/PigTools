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
'* Name: HostFolder
'* Author: Seow Phong
'* Describe: 主机文件夹类|Host folder class
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.16
'* Create Time: 7/3/2023
'* 1.1	10/3/2023   Add fFillByRs
'* 1.2	13/3/2023   Modify New
'* 1.3	14/3/2023   Add ActiveInf,StaticInf
'* 1.5	31/3/2023   Modify fFillByXmlRs,fFillByRs
'* 1.6	3/4/2023    Add IsScanTimeout,BeginScan
'* 1.7	4/4/2023    Modify New,mGetFolderID,fFillByRs,fFillByXmlRs, add EnmFolderType,FolderType
'* 1.8	6/4/2023    Modify fGetFileAndDirListApp_ScanOK, add PigObjFsLib
'* 1.9	8/4/2023    Modify BeginScan
'* 1.10	10/4/2023   Modify fFillByRs,fFillByXmlRs,BeginScan,fGetFileAndDirListApp_ScanOK
'* 1.11	11/4/2023   Modify fGetFileAndDirListApp_ScanOK, add HostDirs,mGetDBDirID
'* 1.12	12/10/2022	Modify Date Initial Time
'* 1.13	12/10/2022	Modify New
'* 1.15	18/10/2022	Modify New
'* 1.16	23/10/2022	Modify fGetFileAndDirListApp_ScanOK,fLoadHostFiles
'**********************************
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigToolsLiteLib
Imports GetFileAndDirListLib
Imports PigObjFsLib

Public Class HostFolder
	Inherits PigBaseLocal
	Private Const CLS_VERSION As String = "1.16.10"

	Public ReadOnly Property HostDirs As New HostDirs
	Friend ReadOnly Property fParent As Host
	Friend ReadOnly Property fPigFunc As New PigFunc
	Private ReadOnly Property mFS As New FileSystemObject

	Friend WithEvents fGetFileAndDirListApp As GetFileAndDirListApp
	Public Enum EnmScanStatus
		Ready = 0
		Scanning = 1
		ScanComplete = 2
		ScanError = -1
		ScanTimeout = -2
	End Enum

	Public Enum EnmFolderType
		Unknow = 0
		WinFolder = 1
		LinuxFolder = 2
		ZipFile = 3
		WarFile = 4
		JarFile = 5
	End Enum

	Public Sub New(FolderID As String, Parent As Host)
		MyBase.New(CLS_VERSION)
		Me.FolderID = FolderID
		Me.HostID = Parent.HostID
		Me.fParent = Parent
	End Sub

	Public Sub New(FolderPath As String, HostID As String, Parent As Host)
		MyBase.New(CLS_VERSION)
		Dim strRet As String = ""
		Try
			strRet = Me.mGetFolderID(HostID, FolderPath, Me.FolderID)
			If strRet <> "OK" Then Throw New Exception(strRet)
			Me.HostID = HostID
			Me.FolderPath = FolderPath
			Dim o As Host = Me.fParent
			Me.fParent = Parent
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub


	Private Function mGetFolderID(HostID As String, FolderPath As String, ByRef OutFolderID As String) As String
		Try
			Dim strData As String = "<" & HostID & "><" & FolderPath & ">"
			Dim strRet As String = fPigFunc.GetTextPigMD5(strData, PigMD5.enmTextType.UTF8, OutFolderID)
			If strRet <> "OK" Then Throw New Exception((strRet))
			If OutFolderID = "" Then Throw New Exception("Unable to get FolderID")
			Return "OK"
		Catch ex As Exception
			OutFolderID = ""
			Return ""
		End Try
	End Function

	Private Function mGetFolderName(AbsoluteFolderPath As String) As String
		Try
			Select Case Right(AbsoluteFolderPath, 1)
				Case "/", "\"
					AbsoluteFolderPath = Left(AbsoluteFolderPath, Len(AbsoluteFolderPath) - 1)
			End Select
			mGetFolderName = Me.fPigFunc.GetFilePart(AbsoluteFolderPath, PigFunc.EnmFilePart.FileTitle)
		Catch ex As Exception
			Return ""
		End Try
	End Function


	Private mScanStatus As EnmScanStatus = EnmScanStatus.Ready
	Public Property ScanStatus() As EnmScanStatus
		Get
			Return mScanStatus
		End Get
		Friend Set(value As EnmScanStatus)
			If value <> mScanStatus Then
				Me.mUpdateCheck.Add("ScanStatus")
				mScanStatus = value
			End If
		End Set
	End Property

	Public ReadOnly Property IsScanTimeout() As Boolean
		Get
			Try
				If DateDiff(DateInterval.Minute, Me.ScanBeginTime, Now) > 10 Then
					Return True
				Else
					Return False
				End If
			Catch ex As Exception
				Return False
			End Try
		End Get
	End Property

	Private mFolderType As EnmFolderType = EnmFolderType.Unknow
	Public Property FolderType() As EnmFolderType
		Get
			Return mFolderType
		End Get
		Friend Set(value As EnmFolderType)
			If value <> mFolderType Then
				Me.mUpdateCheck.Add("FolderType")
				mFolderType = value
			End If
		End Set
	End Property

	Private mActiveInfXml As PigXml
	Public Property ActiveInf() As String
		Get
			If mActiveInfXml Is Nothing Then Me.mRefActiveInfXml()
			Return mActiveInfXml.MainXmlStr
		End Get
		Friend Set(value As String)
			Try
				If mActiveInfXml Is Nothing Then
					mActiveInfXml = New PigXml(value)
					Me.mUpdateCheck.Add("ActiveInf")
				ElseIf value <> mActiveInfXml.MainXmlStr Then
					mActiveInfXml = New PigXml(value)
					Me.mUpdateCheck.Add("ActiveInf")
				End If
			Catch ex As Exception
				Me.SetSubErrInf("ActiveInf.Set", ex)
			End Try
		End Set
	End Property

	Private Function mRefActiveInfXml() As String
		Try
			Dim strXml As String = ""
			If mActiveInfXml IsNot Nothing Then strXml = mActiveInfXml.MainXmlStr
			mActiveInfXml = New PigXml(False)
			With mActiveInfXml
				.AddEle("ErrInf", Me.mActiveInf_ErrInf)
				If .MainXmlStr <> strXml Then
					Me.mUpdateCheck.Add("ActiveInf")
				End If
			End With
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("mRefActiveInfXml", ex)
		End Try
	End Function

	Private mActiveInf_ErrInf As String
	Public Property ActiveInf_ErrInf() As String
		Get
			Return mActiveInf_ErrInf
		End Get
		Friend Set(value As String)
			If value <> mActiveInf_ErrInf Then
				mActiveInf_ErrInf = value
				Me.mRefActiveInfXml()
			End If
		End Set
	End Property

	Private mFolderPath As String = ""
	Public Property FolderPath() As String
		Get
			Return mFolderPath
		End Get
		Friend Set(value As String)
			Select Case Right(value, 1)
				Case "\", "/"
					value = Left(value, Len(value) - 1)
			End Select
			If value <> mFolderPath Then
				Me.FolderName = Me.mGetFolderName(value)
				Me.mUpdateCheck.Add("FolderPath")
				mFolderPath = value
			End If
		End Set
	End Property


	'以下可以自动生成
	Public ReadOnly Property FolderID As String
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
	Private mHostID As String
	Public Property HostID() As String
		Get
			Return mHostID
		End Get
		Friend Set(value As String)
			If value <> mHostID Then
				Me.mUpdateCheck.Add("HostID")
				mHostID = value
			End If
		End Set
	End Property
	Private mFolderName As String = ""
	Public Property FolderName() As String
		Get
			Return mFolderName
		End Get
		Friend Set(value As String)
			If value <> mFolderName Then
				Me.mUpdateCheck.Add("FolderName")
				mFolderName = value
			End If
		End Set
	End Property
	Private mFolderDesc As String = ""
	Public Property FolderDesc() As String
		Get
			Return mFolderDesc
		End Get
		Friend Set(value As String)
			If value <> mFolderDesc Then
				Me.mUpdateCheck.Add("FolderDesc")
				mFolderDesc = value
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
	Private mStaticInf As String
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



	Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
		Try
			If InRs.EOF = False Then
				With InRs.Fields
					If .IsItemExists("HostID") = True Then
						If Me.HostID <> .Item("HostID").StrValue Then
							Me.HostID = .Item("HostID").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FolderName") = True Then
						If Me.FolderName <> .Item("FolderName").StrValue Then
							Me.FolderName = .Item("FolderName").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FolderPath") = True Then
						If Me.FolderPath <> .Item("FolderPath").StrValue Then
							Me.FolderPath = .Item("FolderPath").StrValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FolderType") = True Then
						If Me.FolderType <> .Item("FolderType").IntValue Then
							Me.FolderType = .Item("FolderType").IntValue
							UpdateCnt += 1
						End If
					End If
					If .IsItemExists("FolderDesc") = True Then
						If Me.FolderDesc <> .Item("FolderDesc").StrValue Then
							Me.FolderDesc = .Item("FolderDesc").StrValue
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
					If .IsItemExists("IsUse") = True Then
						If Me.IsUse <> .Item("IsUse").BooleanValue Then
							Me.IsUse = .Item("IsUse").BooleanValue
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
					If .IsColExists(RSNo, "HostID") = True Then
						If Me.HostID <> .StrValue(RSNo, RowNo, "HostID") Then
							Me.HostID = .StrValue(RSNo, RowNo, "HostID")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FolderName") = True Then
						If Me.FolderName <> .StrValue(RSNo, RowNo, "FolderName") Then
							Me.FolderName = .StrValue(RSNo, RowNo, "FolderName")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FolderPath") = True Then
						If Me.FolderPath <> .StrValue(RSNo, RowNo, "FolderPath") Then
							Me.FolderPath = .StrValue(RSNo, RowNo, "FolderPath")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FolderType") = True Then
						If Me.FolderType <> .IntValue(RSNo, RowNo, "FolderType") Then
							Me.FolderType = .IntValue(RSNo, RowNo, "FolderType")
							UpdateCnt += 1
						End If
					End If
					If .IsColExists(RSNo, "FolderDesc") = True Then
						If Me.FolderDesc <> .StrValue(RSNo, RowNo, "FolderDesc") Then
							Me.FolderDesc = .StrValue(RSNo, RowNo, "FolderDesc")
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
					If .IsColExists(RSNo, "IsUse") = True Then
						If Me.IsUse <> .BooleanValue(RSNo, RowNo, "IsUse") Then
							Me.IsUse = .BooleanValue(RSNo, RowNo, "IsUse")
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

	Public Function BeginScan() As String
		Dim LOG As New PigStepLog("fBeginScanDBFolder")
		Try
			LOG.StepName = "fRefHostFolder"
			LOG.Ret = Me.fParent.fParent.fRefHostFolder(Me.fParent, Me)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Select Case Me.ScanStatus
				Case HostFolder.EnmScanStatus.Scanning
					If Me.IsScanTimeout = True Then
						Me.ScanStatus = HostFolder.EnmScanStatus.ScanTimeout
						LOG.StepName = "fUpdHostFolder(ScanTimeout)"
						LOG.Ret = Me.fParent.fParent.fUpdHostFolder(Me)
						If LOG.Ret <> "OK" Then
							Throw New Exception(LOG.Ret)
						Else
							Throw New Exception("Scan Timeout.")
						End If
					End If
				Case Else
					With Me
						LOG.StepName = "Before scanning"
						If .FolderType = EnmFolderType.Unknow Then
							.FolderType = .fParent.fParent.fAutoGetFolderType(.FolderPath)
						End If
						Select Case .FolderType
							Case EnmFolderType.WinFolder, EnmFolderType.LinuxFolder
							Case Else
								LOG.AddStepNameInf(.FolderPath)
								Throw New Exception("Unsupported FolderType is " & .FolderType.ToString)
						End Select
						If Me.fPigFunc.IsFolderExists(.FolderPath) = False Then
							LOG.AddStepNameInf(.FolderPath)
							Throw New Exception("Folder not found.")
						End If
						.ScanStatus = HostFolder.EnmScanStatus.Scanning
						.ScanBeginTime = Now
						If .ActiveInf_ErrInf <> "" Then .ActiveInf_ErrInf = ""
						LOG.StepName = "New GetFileAndDirListApp"
						fGetFileAndDirListApp = New GetFileAndDirListApp(.FolderPath, False, Me.fParent.fParent.LogFilePath)
						If fGetFileAndDirListApp.LastErr <> "" Then
							LOG.AddStepNameInf(.FolderPath)
							Throw New Exception(fGetFileAndDirListApp.LastErr)
						End If
						LOG.StepName = "GetFileAndDirListApp.Start"
						Me.fGetFileAndDirListApp.Start()
						If Me.fGetFileAndDirListApp.LastErr <> "" Then Throw New Exception(Me.fGetFileAndDirListApp.LastErr)
						LOG.StepName = "fUpdHostFolder(Scanning)"
						LOG.Ret = Me.fParent.fParent.fUpdHostFolder(Me)
						If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
					End With
			End Select
			Return "OK"
		Catch ex As Exception
			Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
			With Me
				.ActiveInf_ErrInf = strErr
				.ScanStatus = EnmScanStatus.ScanError
				.fParent.fParent.fUpdHostFolder(Me)
			End With
			Return strErr
		End Try
	End Function

	Private Function mGetDBDirID(DirPath As String) As String
		Try
			If Me.FolderType = EnmFolderType.WinFolder Then
				DirPath = UCase(DirPath)
			End If
			DirPath = Me.FolderID & ">" & DirPath
			mGetDBDirID = ""
			Dim strRet As String = Me.fPigFunc.GetTextPigMD5(DirPath, PigMD5.enmTextType.UTF8, mGetDBDirID)
			If strRet <> "OK" Then Throw New Exception(strRet)
		Catch ex As Exception
			Me.SetSubErrInf("mGetDBDirID", ex)
			Return ""
		End Try
	End Function


	Private Function mGetDBFileID(FilePath As String) As String
		Try
			If Me.FolderType = EnmFolderType.WinFolder Then
				FilePath = UCase(FilePath)
			End If
			FilePath = Me.FolderID & ">" & FilePath
			mGetDBFileID = ""
			Dim strRet As String = Me.fPigFunc.GetTextPigMD5(FilePath, PigMD5.enmTextType.UTF8, mGetDBFileID)
			If strRet <> "OK" Then Throw New Exception(strRet)
		Catch ex As Exception
			Me.SetSubErrInf("mGetDBFileID", ex)
			Return ""
		End Try
	End Function

	Private Sub fGetFileAndDirListApp_ScanOK() Handles fGetFileAndDirListApp.ScanOK
		Dim LOG As New PigStepLog("fGetFileAndDirListApp_ScanOK")
		Try
			LOG.StepName = "fLoadHostDirs"
			LOG.Ret = Me.fLoadHostDirs()
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			LOG.StepName = "fLoadHostFiles"
			LOG.Ret = Me.fLoadHostFiles()
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			For Each oHostDir As HostDir In Me.HostDirs
				LOG.StepName = "RefByHostFiles"
				LOG.Ret = oHostDir.RefByHostFiles()
				If LOG.Ret <> "OK" Then
					Me.fParent.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
				End If
			Next
			LOG.StepName = "fMergeHostDirInf"
			LOG.Ret = Me.fParent.fParent.fMergeHostDirInf(Me)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			LOG.StepName = "New HostFolder"
			Dim oHostFolder As New HostFolder(Me.FolderID, Me.fParent)
			LOG.StepName = "HostFolder.RefHostFolders"
			LOG.Ret = oHostFolder.fRefHostDirs(False, True)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			For Each oHostDir As HostDir In oHostFolder.HostDirs
				LOG.StepName = "Get HostDir"
				oHostDir = Me.HostDirs.Item(oHostDir.DirID)
				LOG.StepName = "fMergeHostFileInf"
				LOG.Ret = oHostDir.fMergeHostFileInf()
				If LOG.Ret <> "OK" Then
					Me.fParent.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
				End If
			Next
			oHostFolder = Nothing
			With Me
				.ScanEndTime = Now
				.ScanStatus = EnmScanStatus.ScanComplete
				LOG.StepName = "fUpdHostFolder(ScanComplete)"
				LOG.Ret = Me.fParent.fParent.fUpdHostFolder(Me)
				If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			End With
		Catch ex As Exception
			Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
			With Me
				.ScanEndTime = Now
				.ScanStatus = EnmScanStatus.ScanError
				.ActiveInf_ErrInf = strErr
				Me.fParent.fParent.fUpdHostFolder(Me)
			End With
			Me.fParent.fParent.PrintDebugLog(Me.MyClassName, strErr)
		End Try
	End Sub

	Private Sub fGetFileAndDirListApp_ScanFail(ErrInf As String) Handles fGetFileAndDirListApp.ScanFail
		Dim LOG As New PigStepLog("fGetFileAndDirListApp_ScanFail")
		Try
			With Me
				.ScanEndTime = Now
				.ScanStatus = EnmScanStatus.ScanError
				.ActiveInf_ErrInf = ErrInf
				LOG.StepName = "fUpdHostFolder(ScanError)"
				LOG.Ret = Me.fParent.fParent.fUpdHostFolder(Me)
				If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			End With
		Catch ex As Exception
			Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
			With Me
				.ScanEndTime = Now
				.ScanStatus = EnmScanStatus.ScanError
				.ActiveInf_ErrInf = strErr
				Me.fParent.fParent.fUpdHostFolder(Me)
			End With
			Me.fParent.fParent.PrintDebugLog("DBFolder", strErr)
		End Try
	End Sub

	Friend Function fLoadHostFiles() As String
		Dim LOG As New PigStepLog("fLoadHostFiles")
		Try
			LOG.StepName = "OpenTextFile_FileListPath"
			Dim tsRead As TextStream = Me.mFS.OpenTextFile(Me.fGetFileAndDirListApp.FileListPath, FileSystemObject.IOMode.ForReading)
			If Me.mFS.LastErr <> "" Then
				LOG.AddStepNameInf(Me.fGetFileAndDirListApp.FileListPath)
				Throw New Exception(Me.mFS.LastErr)
			End If
			Dim lngLineNo As Long = 1
			Dim oHostDir As HostDir = Nothing
			Do While Not tsRead.AtEndOfStream
				Dim strLine As String = tsRead.ReadLine
				Dim strFilePath As String = Me.fPigFunc.GetStr(strLine, "", vbTab)
				Dim strFileSize As String = Me.fPigFunc.GetStr(strLine, "", vbTab)
				Dim strFileUpdateTime As String = Me.fPigFunc.GetStr(strLine, "", vbTab)
				Dim strFastPigMD5 As String = strLine
				If Left(strFilePath, 1) <> "." Then
					LOG.StepName = strFilePath
					LOG.AddStepNameInf("LineNo:" & lngLineNo)
					Throw New Exception("Not a relative path")
				End If
				Dim strDirPath As String = Me.fPigFunc.GetPathPart(strFilePath, PigFunc.EnmFathPart.ParentPath)
				If strDirPath = "." Then strDirPath = Left(strFilePath, 2)
				Dim strDirID As String = Me.mGetDBDirID(strDirPath)
				Dim strFileName As String = Me.fPigFunc.GetPathPart(strFilePath, PigFunc.EnmFathPart.FileOrDirTitle)
				If oHostDir IsNot Nothing Then If oHostDir.DirID <> strDirID Then oHostDir = Nothing
				If oHostDir Is Nothing Then If Me.HostDirs.IsItemExists(strDirID) = True Then oHostDir = Me.HostDirs.Item(strDirID)
				If oHostDir Is Nothing Then
					LOG.StepName = strFilePath
					LOG.AddStepNameInf("Unable to find DBDir object")
					LOG.AddStepNameInf("LineNo is " & lngLineNo)
					Me.fParent.fParent.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
				Else
					Dim strFileID As String = Me.mGetDBFileID(strFilePath)
					With oHostDir.HostFiles.AddOrGet(strFileID)
						.FileName = strFileName
						.FileSize = Me.fPigFunc.GECLng(strFileSize)
						.FileUpdateTime = Me.fPigFunc.SQLCDate(strFileUpdateTime)
						.FastPigMD5 = strFastPigMD5
					End With
				End If
				lngLineNo += 1
			Loop
			tsRead.Close()
			tsRead = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Friend Function fLoadHostDirs() As String
		Dim LOG As New PigStepLog("fLoadHostDirs")
		Try
			LOG.StepName = "OpenTextFile_DirListPath"
			Dim tsRead As TextStream = Me.mFS.OpenTextFile(Me.fGetFileAndDirListApp.DirListPath, FileSystemObject.IOMode.ForReading)
			If Me.mFS.LastErr <> "" Then
				LOG.AddStepNameInf(Me.fGetFileAndDirListApp.DirListPath)
				Throw New Exception(Me.mFS.LastErr)
			End If
			Dim lngLineNo As Long = 1
			LOG.StepName = "HostDirs.Clear"
			Me.HostDirs.Clear()
			If Me.HostDirs.LastErr <> "" Then Throw New Exception(Me.HostDirs.LastErr)
			Do While Not tsRead.AtEndOfStream
				Dim strLine As String = tsRead.ReadLine
				Dim strDirPath As String = Me.fPigFunc.GetStr(strLine, "", vbTab)
				Dim strDirUpdateTime As String = strLine
				If Left(strDirPath, 1) <> "." Then
					LOG.StepName = strDirPath
					LOG.AddStepNameInf("LineNo:" & lngLineNo)
					Throw New Exception("Not a relative path")
				End If
				Dim strDirID As String = Me.mGetDBDirID(strDirPath)
				LOG.StepName = "HostDirs.AddOrGet(" & lngLineNo & ")"
				With Me.HostDirs.AddOrGet(strDirID, Me)
					.DirPath = strDirPath
					.DirUpdateTime = Me.fPigFunc.SQLCDate(strDirUpdateTime)
					.HostFiles.Clear()
				End With
				lngLineNo += 1
			Loop
			tsRead.Close()
			tsRead = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Friend Function fRefHostDirs(Optional IsDirtyRead As Boolean = True, Optional IsScanOnly As Boolean = True) As String
		Return Me.fParent.fParent.fRefHostDirs(Me, IsDirtyRead, IsScanOnly)
	End Function

End Class