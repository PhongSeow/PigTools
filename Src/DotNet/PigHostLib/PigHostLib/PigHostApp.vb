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
'* Name: 豚豚主机应用|PigHostApp
'* Author: Seow Phong
'* Describe: 主机应用|Host Applications
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.21
'* Create Time: 1/10/2022
'* 1.1  1/10/2022   Add RefDBConn
'* 1.2  15/10/2022  Add mCreateTable_HostInf, modify RefDBConn
'* 1.3	8/3/2023    Add mCreateTable_HFContSegInf,mCreateTable_DFPathInf, modify RefDBConn
'* 1.4	13/3/2023   Add mExecCmdSQLSrvText
'* 1.5	13/3/2023   Modify mNew
'* 1.6	15/3/2023   Modify New
'* 1.7	30/3/2023   Modify RefHostFolders,mAddNewHostFolder
'* 1.8	31/3/2023   Modify RefHostFolders
'* 1.9	3/4/2023    Modify RefHostFolders, add fRefHostFolder
'* 1.10	4/4/2023    Modify fBeginScanHostFolder
'* 1.11	5/4/2023    Add GetFolderType
'* 1.12	6/4/2023    Add GetFolderType, modify fPrintErrLogInf,fMergeDFPathInf
'* 1.13	8/4/2023    Modify RefDBConn
'* 1.15	10/4/2023   Modify New,AutoGetFolderType,fUpdHFFileInf,fUpdHostFolder and add fScanHFFileInf,fAddHFFileInf,mCreateTable_HFFolderInf
'* 1.16	11/4/2023   Add fCreateTable_TmpFileList,fInsTmpFileList,mCreateTable_HFDirInf, modify RefDBConn
'* 1.17	17/4/2023   Add mCreateTable_HostConfInf, modify RefDBConn,RefHostFolder
'* 1.18	18/4/2023   Add mCreateTable_HostConfInf, modify RefDBConn,RefHostFolder,fRefHostFolder,AddNewHostFolder,RefHostFolder
'* 1.19	23/4/2023   Modify fMergeHostDirInf
'* 1.20	24/4/2023   Modify fMergeHostFileInf
'* 1.21	29/4/2023   Modify AddNewHostFolder,mAddNewHostFolder
'* 1.22	30/4/2023   Modify AddNewHostFolder,mAddNewHostFolder
'**********************************
Imports System.Data
Imports System.IO
Imports System.IO.Compression
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
Imports PigSQLSrvLib
#Else
Imports Microsoft.Data.SqlClient
Imports PigSQLSrvCoreLib
#End If
Imports PigObjFsLib
Imports PigToolsLiteLib
Imports PigCmdLib


Public Class PigHostApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.22.28"

    Public ReadOnly Property BaseDirPath As String
    Private ReadOnly Property mMyHostID As String
    Private Property mConnSQLSrv As ConnSQLSrv
    Private Property mPigFunc As New PigFunc
    Private Property mPigSysCmd As New PigSysCmd
    Private Property mLogFilePath As String
    Private Property mLastExecCmdSQLSrvTextTime As DateTime
    Public ReadOnly Property Hosts As New Hosts
    Private Property mSQLSrvTools As SQLSrvTools



    Public ReadOnly Property TempPath As String
        Get
            Return Me.BaseDirPath & Me.OsPathSep & "Temp"
        End Get
    End Property

    Public ReadOnly Property ConfPath As String
        Get
            Return Me.BaseDirPath & Me.OsPathSep & "Conf"
        End Get
    End Property

    Public ReadOnly Property LogPath As String
        Get
            Return Me.BaseDirPath & Me.OsPathSep & "Log"
        End Get
    End Property

    Public ReadOnly Property MyHost(Optional IsRefresh As Boolean = False) As Host
        Get
            Try
                Dim strRet As String = ""
                If Me.Hosts.IsItemExists(Me.mMyHostID) = False Then
                    strRet = Me.RefMyHost()
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return Me.Hosts.Item(Me.mMyHostID)
            Catch ex As Exception
                Me.SetSubErrInf("MyHost.Get", ex)
                Return Nothing
            End Try
        End Get
    End Property


    Public Sub New(ConnSQLSrv As ConnSQLSrv, Optional BaseDirPath As String = "")
        MyBase.New(CLS_VERSION)
        Try
            Dim strRet As String = ""
            If BaseDirPath = "" Then
                Me.BaseDirPath = Me.mPigFunc.GetPathPart(Me.AppPath, PigFunc.EnmFilePart.Path)
            Else
                Me.BaseDirPath = BaseDirPath
            End If
            strRet = Me.mGetMyHostID(Me.mMyHostID)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.mRefLogFilePath()
            Me.mNew(ConnSQLSrv)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub


    Private Sub mNew(ConnSQLSrv As ConnSQLSrv)
        Dim LOG As New PigStepLog("mNew")
        Try
            Dim strError As String = ""
            LOG.StepName = "CreateFolder"
            If Me.mPigFunc.IsFolderExists(Me.TempPath) = False Then
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.TempPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.TempPath)
                    strError &= LOG.StepLogInf & ";"
                End If
            End If
            If Me.mPigFunc.IsFolderExists(Me.LogPath) = False Then
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.LogPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.LogPath)
                    strError &= LOG.StepLogInf & ";"
                End If
            End If
            If Me.mPigFunc.IsFolderExists(Me.ConfPath) = False Then
                LOG.Ret = Me.mPigFunc.CreateFolder(Me.ConfPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.ConfPath)
                    strError &= LOG.StepLogInf & ";"
                End If
            End If
            If strError <> "" Then Throw New Exception(strError)
            LOG.StepName = "Set ConnSQLSrv"
            Me.mConnSQLSrv = ConnSQLSrv
            LOG.StepName = "RefDBConn"
            LOG.Ret = Me.RefDBConn(True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.ClearErr()
        Catch ex As Exception
            Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Me.fPrintErrLogInf(strErr)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    '''' <summary>
    '''' Refresh database connection|刷新数据库连接
    '''' </summary>
    '''' <returns></returns>
    Public Function RefDBConn(Optional IsCreateDBObj As Boolean = False) As String
        Dim LOG As New PigStepLog("RefDBConn")
        Try
            If Me.mConnSQLSrv.IsDBConnReady = False Then
                LOG.StepName = "OpenOrKeepActive"
                Me.mConnSQLSrv.ClearErr()
                Me.mConnSQLSrv.OpenOrKeepActive()
                If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
            End If
            If IsCreateDBObj = True Then
                Me.mSQLSrvTools = New SQLSrvTools(Me.mConnSQLSrv)
                'LOG.StepName = "IsDBObjExists(_ptHFConfInf)"
                'If oSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFConfInf") = False Then
                '    LOG.StepName = "mCreateTable_DFConfInf"
                '    LOG.Ret = mCreateTable_DFConfInf()
                '    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                'End If
                LOG.StepName = "IsDBObjExists(_ptHFContInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFContInf") = False Then
                    LOG.StepName = "mCreateTable_HFContInf"
                    LOG.Ret = Me.mCreateTable_HFContInf
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "IsDBObjExists(_ptHFContSegInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFContSegInf") = False Then
                    LOG.StepName = "mCreateTable_HFContSegInf"
                    LOG.Ret = Me.mCreateTable_HFContSegInf()
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "IsDBObjExists(_ptHostInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHostInf") = False Then
                    LOG.StepName = "mCreateTable_HostInf"
                    LOG.Ret = Me.mCreateTable_HostInf
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "IsDBObjExists(_ptHFFolderInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFFolderInf") = False Then
                    LOG.StepName = "mCreateTable_HFFolderInf"
                    LOG.Ret = Me.mCreateTable_HFFolderInf()
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "IsDBObjExists(_ptHFDirInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFDirInf") = False Then
                    LOG.StepName = "mCreateTable_HFDirInf"
                    LOG.Ret = Me.mCreateTable_HFDirInf
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "IsDBObjExists(_ptHFFileInf)"
                If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptHFFileInf") = False Then
                    LOG.StepName = "mCreateTable_HFFileInf"
                    LOG.Ret = Me.mCreateTable_HFFileInf
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                End If
            End If
            Me.mIsDBReady = True
            Return "OK"
        Catch ex As Exception
            Me.mIsDBReady = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HostInf() As String
        Const SUB_NAME As String = "mCreateTable_HostInf"
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strTabName As String = ""
            Dim strSQL As String = ""
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE dbo._ptHostInf(")
                .AddMultiLineText(strSQL, "HostID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",HostName varchar(256) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",HostMainIp varchar(30) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",IsUse bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",HostDesc varchar(1024) NOT NULL DEFAULT ('')", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime  DEFAULT (getdate())", 1)
                .AddMultiLineText(strSQL, ",UpdateTime datetime NOT NULL DEFAULT (getdate())", 1)
                .AddMultiLineText(strSQL, ",ScanStatus int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",StaticInf varchar(8000) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",ActiveInf varchar(max) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",ScanBeginTime datetime ", 1)
                .AddMultiLineText(strSQL, ",ScanEndTime datetime ", 1)
                .AddMultiLineText(strSQL, ",ActiveTime datetime ", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHostInf ADD CONSTRAINT PK_ptHostInf PRIMARY KEY CLUSTERED(HostID)")
                .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptHostInf ON _ptHostInf(HostMainIp,HostName)")
            End With
            strStepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                strStepName = "ExecuteNonQuery"
                strRet = .ExecuteNonQuery()
                If strRet <> "OK" Then
                    Me.PrintDebugLog(SUB_NAME, strStepName, .DebugStr)
                    Throw New Exception(strRet)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
        End Try
    End Function

    Private mIsDBReady As Boolean
    Public ReadOnly Property IsDBReady As Boolean
        Get
            Return Me.mIsDBReady
        End Get
    End Property

    Private Function mGetMyHostID(ByRef MyHostID As String) As String
        Dim LOG As New PigStepLog("mGetMyHostID")
        Try
            Dim strUUID As String = ""
            LOG.Ret = Me.mPigSysCmd.GetUUID(strUUID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "GetTextPigMD5"
            LOG.Ret = Me.mPigFunc.GetTextPigMD5(strUUID, PigMD5.enmTextType.UTF8, MyHostID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            MyHostID = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fUpdHostFolder(ByRef InObj As HostFolder) As String
        Dim LOG As New PigStepLog("fUpdHostFolder")
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Generate SQL"
            Dim strUpdCols As String = ""
            If InObj.IsUpdate("HostID") = True Then strUpdCols &= ",HostID=@HostID"
            If InObj.IsUpdate("FolderName") = True Then strUpdCols &= ",FolderName=@FolderName"
            If InObj.IsUpdate("FolderPath") = True Then strUpdCols &= ",FolderPath=@FolderPath"
            If InObj.IsUpdate("FolderType") = True Then strUpdCols &= ",FolderType=@FolderType"
            If InObj.IsUpdate("FolderDesc") = True Then strUpdCols &= ",FolderDesc=@FolderDesc"
            If InObj.IsUpdate("CreateTime") = True Then strUpdCols &= ",CreateTime=@CreateTime"
            If InObj.IsUpdate("UpdateTime") = True Then strUpdCols &= ",UpdateTime=@UpdateTime"
            If InObj.IsUpdate("IsUse") = True Then strUpdCols &= ",IsUse=@IsUse"
            If InObj.IsUpdate("ScanStatus") = True Then strUpdCols &= ",ScanStatus=@ScanStatus"
            If InObj.IsUpdate("StaticInf") = True Then strUpdCols &= ",StaticInf=@StaticInf"
            If InObj.IsUpdate("ActiveInf") = True Then strUpdCols &= ",ActiveInf=@ActiveInf"
            If InObj.IsUpdate("ScanBeginTime") = True Then strUpdCols &= ",ScanBeginTime=@ScanBeginTime"
            If InObj.IsUpdate("ScanEndTime") = True Then strUpdCols &= ",ScanEndTime=@ScanEndTime"
            If InObj.IsUpdate("ActiveTime") = True Then strUpdCols &= ",ActiveTime=@ActiveTime"
            If strUpdCols = "" Then Throw New Exception("There is nothing to update")
            strUpdCols = Mid(strUpdCols, 2)
            strSQL = "UPDATE dbo._ptHFFolderInf SET " & strUpdCols & " WHERE FolderID=@FolderID"
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InObj.FolderID
                If InObj.IsUpdate("HostID") = True Then
                    .AddPara("@HostID", Data.SqlDbType.VarChar, 64)
                    .ParaValue("@HostID") = InObj.HostID
                End If
                If InObj.IsUpdate("FolderName") = True Then
                    .AddPara("@FolderName", Data.SqlDbType.VarChar, 256)
                    .ParaValue("@FolderName") = InObj.FolderName
                End If
                If InObj.IsUpdate("FolderPath") = True Then
                    .AddPara("@FolderPath", Data.SqlDbType.VarChar, 2048)
                    .ParaValue("@FolderPath") = InObj.FolderPath
                End If
                If InObj.IsUpdate("FolderType") = True Then
                    .AddPara("@FolderType", Data.SqlDbType.Int)
                    .ParaValue("@FolderType") = InObj.FolderType
                End If
                If InObj.IsUpdate("FolderDesc") = True Then
                    .AddPara("@FolderDesc", Data.SqlDbType.VarChar, 512)
                    .ParaValue("@FolderDesc") = InObj.FolderDesc
                End If
                If InObj.IsUpdate("CreateTime") = True Then
                    .AddPara("@CreateTime", Data.SqlDbType.DateTime)
                    .ParaValue("@CreateTime") = InObj.CreateTime
                End If
                If InObj.IsUpdate("UpdateTime") = True Then
                    .AddPara("@UpdateTime", Data.SqlDbType.DateTime)
                    .ParaValue("@UpdateTime") = InObj.UpdateTime
                End If
                If InObj.IsUpdate("IsUse") = True Then
                    .AddPara("@IsUse", Data.SqlDbType.Bit)
                    .ParaValue("@IsUse") = InObj.IsUse
                End If
                If InObj.IsUpdate("ScanStatus") = True Then
                    .AddPara("@ScanStatus", Data.SqlDbType.Int)
                    .ParaValue("@ScanStatus") = InObj.ScanStatus
                End If
                If InObj.IsUpdate("StaticInf") = True Then
                    .AddPara("@StaticInf", Data.SqlDbType.VarChar, 8000)
                    .ParaValue("@StaticInf") = InObj.StaticInf
                End If
                If InObj.IsUpdate("ActiveInf") = True Then
                    .AddPara("@ActiveInf", Data.SqlDbType.VarChar, 8000)
                    .ParaValue("@ActiveInf") = InObj.ActiveInf
                End If
                If InObj.IsUpdate("ScanBeginTime") = True Then
                    .AddPara("@ScanBeginTime", Data.SqlDbType.DateTime)
                    .ParaValue("@ScanBeginTime") = InObj.ScanBeginTime
                End If
                If InObj.IsUpdate("ScanEndTime") = True Then
                    .AddPara("@ScanEndTime", Data.SqlDbType.DateTime)
                    .ParaValue("@ScanEndTime") = InObj.ScanEndTime
                End If
                If InObj.IsUpdate("ActiveTime") = True Then
                    .AddPara("@ActiveTime", Data.SqlDbType.DateTime)
                    .ParaValue("@ActiveTime") = InObj.ActiveTime
                End If
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RefHostFolder(ByRef InHost As Host, ByRef InHostFolder As HostFolder, Optional IsDirtyRead As Boolean = True) As String
        Dim LOG As New PigStepLog("RefHostFolder")
        Dim strSQL As String = "SELECT * FROM dbo._ptHFFolderInf"
        If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
        strSQL &= " WHERE FolderID=@FolderID"
        Try
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InHostFolder.FolderID
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim strFolderID As String = rsMain.Fields.Item("FolderID").StrValue
            LOG.StepName = "HostFolders.AddOrGet"
            InHostFolder = InHost.HostFolders.AddOrGet(strFolderID, InHost)
            If InHost.HostFolders.LastErr <> "" Then
                LOG.AddStepNameInf(strFolderID)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "fFillByRs"
            LOG.Ret = InHostFolder.fFillByRs(rsMain)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strFolderID)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 自动识别文件夹或打包文件类型|Automatically recognize folder or packaged file types
    ''' </summary>
    ''' <param name="FolderOrFilePath">文件夹或文件路径|Folder or file path</param>
    ''' <returns></returns>
    Friend Function fAutoGetFolderType(FolderOrFilePath As String) As HostFolder.EnmFolderType
        Dim LOG As New PigStepLog("fAutoGetFolderType")
        Try
            If Me.mPigFunc.IsFolderExists(FolderOrFilePath) = True Then
                If Me.IsWindows Then
                    Return HostFolder.EnmFolderType.WinFolder
                Else
                    Return HostFolder.EnmFolderType.LinuxFolder
                End If
            ElseIf Me.mPigFunc.IsFolderExists(FolderOrFilePath) = True Then
                Dim strExtName As String = Me.mPigFunc.GetPathPart(FolderOrFilePath, PigFunc.EnmFathPart.ExtName)
                Select Case LCase(strExtName)
                    Case "jar"
                        Return HostFolder.EnmFolderType.JarFile
                    Case "zip"
                        Return HostFolder.EnmFolderType.ZipFile
                    Case "war"
                        Return HostFolder.EnmFolderType.WarFile
                    Case Else
                        Return HostFolder.EnmFolderType.Unknow
                End Select
            Else
                Return HostFolder.EnmFolderType.Unknow
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return HostFolder.EnmFolderType.Unknow
        End Try
    End Function
    Friend ReadOnly Property LogFilePath As String
        Get
            Return Me.mLogFilePath
        End Get
    End Property

    Private Sub mRefLogFilePath()
        Try
            Me.mLogFilePath = Me.LogPath & Me.OsPathSep & Me.MyClassName & "_" & Format(Now, "yyyyMMdd") & ".log"
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("RefLogFilePath", ex)
        End Try
    End Sub


    Private Function mExecCmdSQLSrvText(InCmdSQLSrvText As CmdSQLSrvText, ByRef RsOut As Recordset, Optional IsTxRes As Boolean = True, Optional IsRefLastExec As Boolean = True, Optional IsAllowEOF As Boolean = False, Optional IsNextRecordset As Boolean = False) As String
        Dim LOG As New PigStepLog("mExecCmdSQLSrvText")
        Try
            If Me.mConnSQLSrv.IsDBConnReady = False Then
                LOG.StepName = "OpenOrKeepActive"
                Me.mConnSQLSrv.OpenOrKeepActive()
                If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
            End If
            LOG.StepName = "ActiveConnection"
            InCmdSQLSrvText.ActiveConnection = Me.mConnSQLSrv.Connection
            LOG.StepName = "Execute"
            Me.PrintDebugLog(LOG.SubName, LOG.StepName, InCmdSQLSrvText.DebugStr)
            Dim rsMain As Recordset = InCmdSQLSrvText.Execute()
            If InCmdSQLSrvText.LastErr <> "" Then
                LOG.AddStepNameInf(InCmdSQLSrvText.SQLText)
                LOG.AddStepNameInf(InCmdSQLSrvText.DebugStr)
                Throw New Exception(InCmdSQLSrvText.LastErr)
            End If
            If rsMain.LastErr <> "" Then
                LOG.AddStepNameInf(InCmdSQLSrvText.SQLText)
                Throw New Exception(rsMain.LastErr)
            End If
            If rsMain.EOF = True And IsAllowEOF = False Then
                Throw New Exception("No results returned")
            ElseIf IsTxRes = True Then
                If rsMain.Fields.Item(0).StrValue <> "OK" Then
                    Throw New Exception(rsMain.Fields.Item(0).StrValue)
                ElseIf IsNextRecordset = True Then
                    LOG.StepName = "NextRecordset"
                    RsOut = rsMain.NextRecordset()
                    If rsMain.LastErr <> "" Then Throw New Exception(rsMain.LastErr)
                    If RsOut Is Nothing Then Throw New Exception("RsOut Is Nothing")
                End If
            Else
                RsOut = rsMain
            End If
            rsMain = Nothing
            If IsRefLastExec = True Then
                Me.mLastExecCmdSQLSrvTextTime = Now
            End If
            Return "OK"
        Catch ex As Exception
            RsOut = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fSetDelHostDirInf(ByRef InHostFolder As HostFolder) As String
        Dim LOG As New PigStepLog("fSetDelHostDirInf")
        Dim strSQL As String = ""
        Try
            Dim oCmdSQLSrvText As CmdSQLSrvText
            LOG.StepName = "Generate SQL"
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf")
                .AddMultiLineText(strSQL, "SET IsDel=1")
                .AddMultiLineText(strSQL, "WHERE DirID=@DirID AND IsDel=0", 1)
            End With
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
                For Each oHostDir As HostDir In InHostFolder.HostDirs
                    If oHostDir.IsDel = True Then
                        .ParaValue("@DirID") = oHostDir.DirID
                        Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                        LOG.StepName = "ExecuteNonQuery"
                        LOG.Ret = .ExecuteNonQuery()
                        If LOG.Ret <> "OK" Then
                            Me.fPrintErrLogInf(LOG.StepLogInf)
                        End If
                    End If
                Next
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fMergeHostDirInf(ByRef InHostDir As HostDir) As String
        Dim LOG As New PigStepLog("fMergeHostDirInf")
        Dim strSQL As String = ""
        Try
            Dim oCmdSQLSrvText As CmdSQLSrvText
            LOG.StepName = "Generate SQL"
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "DECLARE @Rows int")
                .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT TOP 1 1 FROM dbo._ptHFDirInf WHERE DirID=@DirID)")
                .AddMultiLineText(strSQL, "BEGIN")
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptHFDirInf(DirID,FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,IsDel,IsScan,LastScanTime,MaxFileUpdateTime)", 1)
                .AddMultiLineText(strSQL, "VALUES(@DirID,@FolderID,@DirPath,@DirSize,@DirFiles,@FastPigMD5,@DirUpdateTime,0,0,GETDATE(),@MaxFileUpdateTime)", 1)
                .AddMultiLineText(strSQL, "SET @Rows=@@ROWCOUNT", 1)
                .AddMultiLineText(strSQL, "END")
                .AddMultiLineText(strSQL, "ELSE")
                .AddMultiLineText(strSQL, "BEGIN")
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf", 1)
                .AddMultiLineText(strSQL, "SET DirSize=@DirSize,DirFiles=@DirFiles,FastPigMD5=@FastPigMD5,DirUpdateTime=@DirUpdateTime,MaxFileUpdateTime=@MaxFileUpdateTime", 1)
                .AddMultiLineText(strSQL, "WHERE DirID=@DirID", 1)
                .AddMultiLineText(strSQL, "AND (DirSize!=@DirSize OR DirFiles!=@DirFiles OR FastPigMD5!=@FastPigMD5 OR DirUpdateTime!=@DirUpdateTime OR MaxFileUpdateTime!=@MaxFileUpdateTime)", 2)
                .AddMultiLineText(strSQL, "SET @Rows=@@ROWCOUNT", 1)
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf SET LastScanTime=GETDATE() WHERE DirID=@DirID", 1)
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf SET DirPath=@DirPath WHERE DirID=@DirID", 1)
                .AddMultiLineText(strSQL, "END")
                .AddMultiLineText(strSQL, "SELECT @Rows Rows")
            End With
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .AddPara("@DirPath", Data.SqlDbType.VarChar, 4096)
                .AddPara("@DirSize", Data.SqlDbType.Money)
                .AddPara("@DirFiles", Data.SqlDbType.Int)
                .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
                .AddPara("@DirUpdateTime", Data.SqlDbType.DateTime)
                .AddPara("@MaxFileUpdateTime", Data.SqlDbType.DateTime)
                .ParaValue("@DirID") = InHostDir.DirID
                .ParaValue("@FolderID") = InHostDir.fParent.FolderID
                .ParaValue("@DirPath") = InHostDir.DirPath
                .ParaValue("@DirSize") = InHostDir.DirSize
                .ParaValue("@DirFiles") = InHostDir.DirFiles
                .ParaValue("@FastPigMD5") = InHostDir.FastPigMD5
                .ParaValue("@DirUpdateTime") = InHostDir.DirUpdateTime
                .ParaValue("@MaxFileUpdateTime") = InHostDir.MaxFileUpdateTime
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                Dim rsMain As Recordset = Nothing
                LOG.StepName = "mExecCmdSQLSrvSp"
                LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If rsMain.Fields.Item("Rows").IntValue > 0 Then
                    InHostDir.IsScan = True
                Else
                    InHostDir.IsScan = False
                End If
            End With
            Return "OK"
        Catch ex As Exception
            InHostDir.IsScan = False
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fMergeHostDirInf(ByRef InHostFolder As HostFolder) As String
        Dim LOG As New PigStepLog("fMergeHostDirInf")
        Dim strTmpTabName As String = ""
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Generate SQL"
            strTmpTabName = "_ptmpDirList_" & InHostFolder.FolderID
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM sysobjects WHERE name='" & strTmpTabName & "' AND xtype='U') DROP TABLE dbo." & strTmpTabName)
                .AddMultiLineText(strSQL, "CREATE TABLE dbo." & strTmpTabName & "(")
                .AddMultiLineText(strSQL, "DirID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirPath varchar(4096) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirSize money NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirFiles int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirUpdateTime datetime NOT NULL", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_" & strTmpTabName & " PRIMARY KEY CLUSTERED(DirPath)", 1)
                .AddMultiLineText(strSQL, ")")
            End With
            LOG.StepName = "ExecuteNonQuery"
            Dim oCmdSQLSrvText As CmdSQLSrvText
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            oCmdSQLSrvText = Nothing
            strSQL = "INSERT INTO dbo." & strTmpTabName & "(DirID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime)VALUES(@DirID,@DirPath,@DirSize,@DirFiles,@FastPigMD5,@DirUpdateTime)"
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
                .AddPara("@DirPath", Data.SqlDbType.VarChar, 4096)
                .AddPara("@DirSize", Data.SqlDbType.Money)
                .AddPara("@DirFiles", Data.SqlDbType.Int)
                .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
                .AddPara("@DirUpdateTime", Data.SqlDbType.DateTime)
                .ActiveConnection = Me.mConnSQLSrv.Connection
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
            End With
            For Each oHostDir As HostDir In InHostFolder.HostDirs
                With oCmdSQLSrvText
                    .ParaValue("@DirID") = oHostDir.DirID
                    .ParaValue("@DirPath") = oHostDir.DirPath
                    .ParaValue("@DirSize") = oHostDir.DirSize
                    .ParaValue("@DirFiles") = oHostDir.DirFiles
                    .ParaValue("@FastPigMD5") = oHostDir.FastPigMD5
                    .ParaValue("@DirUpdateTime") = oHostDir.DirUpdateTime
                    LOG.Ret = .ExecuteNonQuery
                    If LOG.Ret <> "OK" Then Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End With
            Next
            oCmdSQLSrvText = Nothing
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptHFDirInf(DirID,FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,IsDel,IsScan)")
                .AddMultiLineText(strSQL, "SELECT DirID,@FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,0,1")
                .AddMultiLineText(strSQL, "FROM " & strTmpTabName & " t", 1)
                .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM dbo._ptHFDirInf p WHERE t.DirID=p.DirID AND p.FolderID=@FolderID)")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf")
                .AddMultiLineText(strSQL, "SET IsDel=1")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf p")
                .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM " & strTmpTabName & " t WHERE p.DirID=t.DirID)")
                .AddMultiLineText(strSQL, "AND p.IsDel=0 AND p.FolderID=@FolderID", 1)
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf")
                .AddMultiLineText(strSQL, "SET DirPath=t.DirPath")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
                .AddMultiLineText(strSQL, "WHERE p.DirPath!=t.DirPath")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf")
                .AddMultiLineText(strSQL, "SET IsDel=0")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
                .AddMultiLineText(strSQL, "WHERE p.IsDel=1")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf")
                .AddMultiLineText(strSQL, "SET DirUpdateTime=t.DirUpdateTime,DirSize=t.DirSize,DirFiles=t.DirFiles,FastPigMD5=t.FastPigMD5,IsScan=1")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
                .AddMultiLineText(strSQL, "WHERE (p.DirUpdateTime!=t.DirUpdateTime OR p.DirSize!=t.DirSize OR p.DirFiles!=t.DirFiles OR p.FastPigMD5!=t.FastPigMD5)")
                .AddMultiLineText(strSQL, "AND p.IsDel=0", 1)
                '---------
                If InHostFolder.fParent.fParent.IsDebug = False Then
                    .AddMultiLineText(strSQL, "DROP TABLE dbo." & strTmpTabName)
                End If
            End With
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InHostFolder.FolderID
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.Ret = .ExecuteNonQuery
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fRefHostFolder(ByRef InHost As Host, ByRef InHostFolder As HostFolder, Optional IsDirtyRead As Boolean = True) As String
        Dim LOG As New PigStepLog("fRefHostFolder")
        Dim strSQL As String = "SELECT * FROM dbo._ptHFFolderInf"
        If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
        strSQL &= " WHERE FolderID=@FolderID"
        Try
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InHostFolder.FolderID
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim strFolderID As String = rsMain.Fields.Item("FolderID").StrValue
            LOG.StepName = "HostFolders.AddOrGet"
            InHostFolder = InHost.HostFolders.AddOrGet(strFolderID, InHost)
            If InHost.HostFolders.LastErr <> "" Then
                LOG.AddStepNameInf(strFolderID)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "fFillByRs"
            LOG.Ret = InHostFolder.fFillByRs(rsMain)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strFolderID)
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 新建数据库文件夹|New Database Folder
    ''' </summary>
    ''' <param name="InHost">主机对象|host object</param>
    ''' <param name="FolderPath">文件夹绝对路径|Folder absolute path</param>
    ''' <param name="IsLocalPath">是否本地路径|Is Local path</param>
    ''' <returns></returns>
    Public Function AddNewHostFolder(InHost As Host, FolderPath As String, Optional IsLocalPath As Boolean = False, Optional TimeoutMinutes As Integer = 10) As String
        Dim LOG As New PigStepLog("AddNewHostFolder")
        Try
            If IsLocalPath = True Then
                LOG.StepName = "Check FolderPath"
                If Me.mPigFunc.IsFolderExists(FolderPath) = False Then Throw New Exception("Invalid folder")
            End If
            LOG.StepName = "New HostFolder"
            Dim oHostFolder As New HostFolder(FolderPath, InHost.HostID, InHost)
            If oHostFolder.LastErr <> "" Then Throw New Exception(oHostFolder.LastErr)
            oHostFolder.StaticInf_TimeoutMinutes = TimeoutMinutes
            LOG.StepName = "mAddNewHostFolder"
            LOG.Ret = Me.mAddNewHostFolder(InHost, oHostFolder)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(FolderPath)
            Dim strRet As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Me.fPrintErrLogInf(strRet)
            Return strRet
        End Try
    End Function

    ''' <summary>
    ''' 新建数据库文件夹|Add new database folder
    ''' </summary>
    ''' <returns></returns>
    Private Function mAddNewHostFolder(ByRef InHost As Host, InObj As HostFolder) As String
        Dim LOG As New PigStepLog("mAddNewHostFolder")
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Check entries"
            If InObj.FolderName = "" Then Throw New Exception("Unable to get folder name")
            If InObj.FolderPath = "" Then Throw New Exception("No folder absolute path specified")
            If InObj.FolderID = "" Then Throw New Exception("Unable to get folder ID")
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "DECLARE @ExistsFolderPath varchar(2048)")
                .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM dbo._ptHFFolderInf WHERE FolderID=@FolderID)")
                .AddMultiLineText(strSQL, "SELECT 'Folder ID already exists'", 1)
                .AddMultiLineText(strSQL, "ELSE IF EXISTS(SELECT 1 FROM dbo._ptHFFolderInf WHERE HostID=@HostID AND FolderPath=@FolderPath)")
                .AddMultiLineText(strSQL, "SELECT 'Host ID and Folder Path already exists'", 1)
                .AddMultiLineText(strSQL, "ELSE IF EXISTS(SELECT TOP 1 1 FROM _ptHFFolderInf p WHERE LEFT(@FolderPath,LEN(p.FolderPath))=p.FolderPath AND p.IsUse=1 AND HostID=@HostID)")
                .AddMultiLineText(strSQL, "SELECT TOP 1 'There is a defined top-level folder named '+FolderPath FROM _ptHFFolderInf p WHERE LEFT(@FolderPath,LEN(p.FolderPath))=p.FolderPath AND p.IsUse=1 AND HostID=@HostID", 1)
                .AddMultiLineText(strSQL, "ELSE")
                .AddMultiLineText(strSQL, "BEGIN")
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptHFFolderInf(FolderID,HostID,FolderName,FolderPath,FolderType,FolderDesc,IsUse,ScanStatus,StaticInf)", 1)
                .AddMultiLineText(strSQL, "VALUES(@FolderID,@HostID,@FolderName,@FolderPath,@FolderType,@FolderDesc,@IsUse,@ScanStatus,@StaticInf)", 1)
                .AddMultiLineText(strSQL, "SELECT 'OK'", 1)
                .AddMultiLineText(strSQL, "END")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InObj.FolderID
                .AddPara("@HostID", Data.SqlDbType.VarChar, 64)
                .ParaValue("@HostID") = InObj.HostID
                .AddPara("@FolderName", Data.SqlDbType.VarChar, 256)
                .ParaValue("@FolderName") = InObj.FolderName
                .AddPara("@FolderPath", Data.SqlDbType.VarChar, 2048)
                .ParaValue("@FolderPath") = InObj.FolderPath
                .AddPara("@FolderType", Data.SqlDbType.Int)
                .ParaValue("@FolderType") = InObj.FolderType
                .AddPara("@FolderDesc", Data.SqlDbType.VarChar, 512)
                .ParaValue("@FolderDesc") = InObj.FolderDesc
                .AddPara("@IsUse", Data.SqlDbType.Bit)
                .ParaValue("@IsUse") = InObj.IsUse
                .AddPara("@ScanStatus", Data.SqlDbType.Int)
                .ParaValue("@ScanStatus") = InObj.ScanStatus
                .AddPara("@StaticInf", Data.SqlDbType.VarChar, 8000)
                .ParaValue("@StaticInf") = InObj.StaticInf
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "HostFolders.AddOrGet"
            If InHost.HostFolders.LastErr <> "" Then InHost.HostFolders.ClearErr()
            InHost.HostFolders.AddOrGet(InObj.FolderID, InHost)
            If InHost.HostFolders.LastErr <> "" Then Throw New Exception(InHost.HostFolders.LastErr)
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fPrintErrLogInf(ErrInf As String) As String
        Try
            Me.mPigFunc.ASyncOptLogInf(ErrInf, Me.mLogFilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fPrintErrLogInf", ex)
        End Try
    End Function

    Friend Function fRefHostFolders(ByRef InHost As Host, Optional IsDirtyRead As Boolean = True, Optional IsShowIsUseOnly As Boolean = False) As String
        Dim LOG As New PigStepLog("fRefHostFolders")
        Dim strSQL As String = "SELECT * FROM dbo._ptHFFolderInf"
        If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
        strSQL &= " WHERE HostID=@HostID"
        If IsShowIsUseOnly = True Then strSQL &= " AND IsUse=1"
        strSQL &= " ORDER BY FolderName"
        Try
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@HostID", SqlDbType.VarChar, 32)
                .ParaValue("@HostID") = InHost.HostID
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "HostFolders.Clear"
            InHost.HostFolders.Clear()
            Do While Not rsMain.EOF
                Dim oHostFolder As HostFolder
                Dim strFolderID As String = rsMain.Fields.Item("FolderID").StrValue
                LOG.StepName = "HostFolders.AddOrGet"
                oHostFolder = InHost.HostFolders.AddOrGet(strFolderID, InHost)
                If InHost.HostFolders.LastErr <> "" Then
                    LOG.AddStepNameInf(strFolderID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                LOG.StepName = "fFillByRs"
                LOG.Ret = oHostFolder.fFillByRs(rsMain)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strFolderID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                rsMain.MoveNext()
            Loop
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Friend Function fRefHostDirs(ByRef InHostFolder As HostFolder, Optional IsDirtyRead As Boolean = True) As String
        Dim LOG As New PigStepLog("fRefHostDirs")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                strSQL = ""
                Select Case InHostFolder.StaticInf_ScanLevel
                    Case HostFolder.EnmScanLevel.Complete
                        .AddMultiLineText(strSQL, "SELECT *")
                    Case HostFolder.EnmScanLevel.Standard
                        .AddMultiLineText(strSQL, "SELECT DirID,DirSize,DirFiles,FastPigMD5,DirUpdateTime,MaxFileUpdateTime")
                    Case HostFolder.EnmScanLevel.Fast
                        .AddMultiLineText(strSQL, "SELECT DirID,DirSize,DirFiles,DirUpdateTime,MaxFileUpdateTime")
                    Case HostFolder.EnmScanLevel.VeryFast
                        .AddMultiLineText(strSQL, "SELECT DirID,DirUpdateTime")
                    Case Else
                        .AddMultiLineText(strSQL, "SELECT DirID,DirUpdateTime")
                End Select
                If IsDirtyRead = True Then
                    .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf WITH (NOLOCK)")
                Else
                    .AddMultiLineText(strSQL, "FROM dbo._ptHFDirInf")
                End If
                .AddMultiLineText(strSQL, "WHERE FolderID=@FolderID AND IsDel=0")
                .AddMultiLineText(strSQL, "ORDER BY  DirID")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FolderID", SqlDbType.VarChar, 32)
                .ParaValue("@FolderID") = InHostFolder.FolderID
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "HostDirs.Clear"
            InHostFolder.HostDirs.Clear()
            Do While Not rsMain.EOF
                Dim oHostDir As HostDir
                Dim strDirID As String = rsMain.Fields.Item("DirID").StrValue
                LOG.StepName = "HostDirs.AddOrGet"
                oHostDir = InHostFolder.HostDirs.AddOrGet(strDirID, InHostFolder)
                If InHostFolder.HostDirs.LastErr <> "" Then
                    LOG.AddStepNameInf(strDirID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                LOG.StepName = "fFillByRs"
                LOG.Ret = oHostDir.fFillByRs(rsMain)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strDirID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                rsMain.MoveNext()
            Loop
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            For Each oHostDir As HostDir In InHostFolder.HostDirs
                oHostDir.IsDel = True
            Next
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RefHosts(Optional IsMyHostOnly As Boolean = False, Optional IsDirtyRead As Boolean = True) As String
        Dim LOG As New PigStepLog("RefHosts")
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Generate SQL"
            strSQL = "SELECT * FROM dbo._ptHostInf"
            If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
            If IsMyHostOnly = True Then strSQL &= " WHERE HostID=@HostID"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@HostID", SqlDbType.VarChar, 32)
                .ParaValue("@HostID") = Me.mMyHostID
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Hosts.Clear"
            Me.Hosts.Clear()
            Do While Not rsMain.EOF
                Dim oHost As Host
                Dim strHostID As String = rsMain.Fields.Item("HostID").StrValue
                LOG.StepName = "Hosts.AddOrGet"
                oHost = Me.Hosts.AddOrGet(strHostID, Me)
                If Me.Hosts.LastErr <> "" Then
                    LOG.AddStepNameInf(strHostID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                LOG.StepName = "fFillByRs"
                LOG.Ret = oHost.fFillByRs(rsMain)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strHostID)
                    Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End If
                rsMain.MoveNext()
            Loop
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HostConfInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HostConfInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHostConfInf(")
                .AddMultiLineText(strSQL, "ConfID varchar(64) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",ConfValue varchar(8000) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",CodeDesc varchar(512) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_ptHostConfInf PRIMARY KEY CLUSTERED(ConfID)", 1)
                .AddMultiLineText(strSQL, ")")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HFContInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HFContInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHFContInf(")
                .AddMultiLineText(strSQL, "FileContID varchar(32) NOT NULL", 1) '文件的FullPigMD5
                .AddMultiLineText(strSQL, ",FileSize bigint NOT NULL", 1)
                .AddMultiLineText(strSQL, ",ContStatus int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",SegContType int", 1)
                .AddMultiLineText(strSQL, ",SegSize int", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_ptHFContInf PRIMARY KEY CLUSTERED(FileContID)", 1)
                .AddMultiLineText(strSQL, ")")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HFContSegInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HFContSegInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHFContSegInf(")
                .AddMultiLineText(strSQL, "FileContID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",SegNo int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",SegCont varchar(max) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_HFContSegInf PRIMARY KEY CLUSTERED(FileContID,SegNo)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFContSegInf WITH CHECK ADD CONSTRAINT FK_HFContSegInf_ptHFContInf FOREIGN KEY(FileContID) REFERENCES _ptHFContInf(FileContID)")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFContSegInf CHECK CONSTRAINT FK_HFContSegInf_ptHFContInf")
                '.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFContSegInf WITH CHECK ADD CONSTRAINT FK_HFContSegInf_ptHFSegInf FOREIGN KEY(ContSegID) REFERENCES _ptHFSegInf(ContSegID)")
                '.AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFContSegInf CHECK CONSTRAINT FK_HFContSegInf_ptHFSegInf")
                '.AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFContSegInf_FileContID ON _ptHFContSegInf(FileContID)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HFFolderInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HFFolderInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHFFolderInf(")
                .AddMultiLineText(strSQL, "FolderID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",HostID varchar(32)", 1)
                .AddMultiLineText(strSQL, ",FolderName varchar(256) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FolderPath varchar(2048) NOT NULL", 1)  '绝对路径
                .AddMultiLineText(strSQL, ",FolderType int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FolderDesc varchar(512) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
                .AddMultiLineText(strSQL, ",IsUse bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",ScanStatus int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",StaticInf varchar(8000) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",ActiveInf varchar(max) NOT NULL DEFAULT('')", 1)
                .AddMultiLineText(strSQL, ",ScanBeginTime datetime ", 1)
                .AddMultiLineText(strSQL, ",ScanEndTime datetime ", 1)
                .AddMultiLineText(strSQL, ",ActiveTime datetime ", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_ptHFFolderInf PRIMARY KEY CLUSTERED(FolderID)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFolderInf WITH CHECK ADD CONSTRAINT FK_ptHFFolderInf_ptHostInf FOREIGN KEY(HostID) REFERENCES _ptHostInf(HostID)")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFolderInf CHECK CONSTRAINT FK_ptHFFolderInf_ptHostInf")
                .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptHFFolderInf_HostID_FolderPath ON _ptHFFolderInf(FolderPath,HostID)")
                .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFFolderInf_FolderName ON _ptHFFolderInf(FolderName)")
                .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFFolderInf_HostID ON _ptHFFolderInf(HostID)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mCreateTable_HFDirInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HFDirInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHFDirInf(")
                .AddMultiLineText(strSQL, "DirID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FolderID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirPath varchar(4096) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirSize money NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirFiles int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",DirUpdateTime datetime NOT NULL", 1)
                .AddMultiLineText(strSQL, ",MaxFileUpdateTime datetime NOT NULL", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, ",IsDel bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",IsScan bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",LastScanTime datetime", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK__ptHFDirInf PRIMARY KEY CLUSTERED(DirID)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFDirInf WITH CHECK ADD CONSTRAINT FK_ptHFDirInf_ptHFFolderInf FOREIGN KEY(FolderID) REFERENCES _ptHFFolderInf(FolderID)")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFDirInf CHECK CONSTRAINT FK_ptHFDirInf_ptHFFolderInf")
                .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFDirInf_FolderID ON _ptHFDirInf(FolderID)")
                .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptHFDirInf_FilePath_FolderID ON _ptHFDirInf(DirPath,FolderID)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function
    Private Function mCreateTable_HFFileInf() As String
        Dim LOG As New PigStepLog("mCreateTable_HFFileInf")
        Dim strSQL As String = ""
        Try
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptHFFileInf(")
                .AddMultiLineText(strSQL, "FileID varchar(36) NOT NULL DEFAULT(newid())", 1)
                .AddMultiLineText(strSQL, ",DirID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileName varchar(256) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileContID varchar(32)", 1)
                .AddMultiLineText(strSQL, ",FileSize int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileUpdateTime datetime NOT NULL", 1)
                .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
                .AddMultiLineText(strSQL, ",IsDel bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",IsCheck bit NOT NULL", 1)
                .AddMultiLineText(strSQL, ",LastCheckTime datetime", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_ptHFFileInf PRIMARY KEY CLUSTERED(FileID)", 1)
                .AddMultiLineText(strSQL, ")")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFileInf WITH CHECK ADD CONSTRAINT FK_ptHFFileInf_ptHFDirInf FOREIGN KEY(DirID) REFERENCES _ptHFDirInf(DirID)")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFileInf CHECK CONSTRAINT FK_ptHFFileInf_ptHFDirInf")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFileInf WITH CHECK ADD CONSTRAINT FK_ptHFFileInf_ptHFContInf FOREIGN KEY(FileContID) REFERENCES _ptHFContInf(FileContID)")
                .AddMultiLineText(strSQL, "ALTER TABLE dbo._ptHFFileInf CHECK CONSTRAINT FK_ptHFFileInf_ptHFContInf")
                .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFFileInf_FileContID ON _ptHFFileInf(FileContID)")
                .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptHFFileInf_FileName ON _ptHFFileInf(FileName)")
                .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptHFFileInf_DirID_FileName ON _ptHFFileInf(DirID,FileName)")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.StepName = "ExecuteNonQuery"
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then
                    Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RefMyHost() As String
        Dim LOG As New PigStepLog("RefMyHost")
        Try
            LOG.StepName = "RefHosts"
            LOG.Ret = Me.RefHosts(True, False)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.Hosts.IsItemExists(Me.mMyHostID) = False Then
                LOG.StepName = "New Host"
                Dim oHost As New Host(Me.mMyHostID, Me)
                With oHost
                    .HostName = Me.mPigFunc.GetHostName
                    .HostMainIp = Me.mPigFunc.GetHostIp()
                    .IsUse = True
                End With
                LOG.StepName = "AddNewHost"
                LOG.Ret = Me.AddNewHost(oHost)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
        Return Me.AddNewHost(Me.MyHost)
    End Function

    Public Function AddNewHost(InObj As Host) As String
        Dim LOG As New PigStepLog("AddNewHost")
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Check entries"
            If InObj.HostName = "" Then Throw New Exception("No host name provided.")
            If InObj.HostMainIp = "" Then Throw New Exception("No host main ip provided.")
            With Me.mPigFunc
                .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM dbo._ptHostInf WHERE HostID=@HostID)")
                .AddMultiLineText(strSQL, "SELECT 'HostID already exists'", 1)
                .AddMultiLineText(strSQL, "ELSE IF EXISTS(SELECT 1 FROM dbo._ptHostInf WHERE HostName=@HostName AND HostMainIp=@HostMainIp)")
                .AddMultiLineText(strSQL, "SELECT 'HostName and HostMainIp already exists'", 1)
                .AddMultiLineText(strSQL, "ELSE")
                .AddMultiLineText(strSQL, "BEGIN")
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptHostInf(HostID,HostName,HostMainIp,HostDesc,IsUse,ScanStatus)", 1)
                .AddMultiLineText(strSQL, "VALUES(@HostID,@HostName,@HostMainIp,@HostDesc,@IsUse,@ScanStatus)", 1)
                .AddMultiLineText(strSQL, "SELECT 'OK'", 1)
                .AddMultiLineText(strSQL, "END")
            End With
            LOG.StepName = "New CmdSQLSrvText"
            Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@HostID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@HostID") = InObj.HostID
                .AddPara("@HostName", Data.SqlDbType.VarChar, 256)
                .ParaValue("@HostName") = InObj.HostName
                .AddPara("@HostMainIp", Data.SqlDbType.VarChar, 30)
                .ParaValue("@HostMainIp") = InObj.HostMainIp
                .AddPara("@IsUse", Data.SqlDbType.Bit)
                .ParaValue("@IsUse") = InObj.IsUse
                .AddPara("@HostDesc", Data.SqlDbType.VarChar, 1024)
                .ParaValue("@HostDesc") = InObj.HostDesc
                .AddPara("@ScanStatus", Data.SqlDbType.Int)
                .ParaValue("@ScanStatus") = InObj.ScanStatus
            End With
            Dim rsMain As Recordset = Nothing
            LOG.StepName = "mExecCmdSQLSrvSp"
            LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Hosts.AddOrGet"
            If Me.Hosts.LastErr <> "" Then Me.Hosts.ClearErr()
            Me.Hosts.AddOrGet(InObj.HostID, Me)
            If Me.Hosts.LastErr <> "" Then Throw New Exception(Me.Hosts.LastErr)
            LOG.StepName = "Close"
            If rsMain IsNot Nothing Then rsMain.Close()
            rsMain = Nothing
            oCmdSQLSrvText = Nothing
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Friend Function fMergeHostFileInf(ByRef InHostDir As HostDir) As String
        Dim LOG As New PigStepLog("fMergeHostFileInf")
        Dim strTmpTabName As String = ""
        Dim strSQL As String = ""
        Try
            LOG.StepName = "Generate SQL"
            strTmpTabName = "_ptmpFileList_" & InHostDir.DirID
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM sysobjects WHERE name='" & strTmpTabName & "' AND xtype='U') DROP TABLE dbo." & strTmpTabName)
                .AddMultiLineText(strSQL, "CREATE TABLE dbo." & strTmpTabName & "(")
                .AddMultiLineText(strSQL, "FileID varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileName varchar(256) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileSize int NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
                .AddMultiLineText(strSQL, ",FileUpdateTime datetime NOT NULL", 1)
                .AddMultiLineText(strSQL, "CONSTRAINT PK_" & strTmpTabName & " PRIMARY KEY CLUSTERED(FileID)", 1)
                .AddMultiLineText(strSQL, ")")
            End With
            LOG.StepName = "ExecuteNonQuery"
            Dim oCmdSQLSrvText As CmdSQLSrvText
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .ActiveConnection = Me.mConnSQLSrv.Connection
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                LOG.Ret = .ExecuteNonQuery()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            oCmdSQLSrvText = Nothing
            strSQL = "INSERT INTO dbo." & strTmpTabName & "(FileID,FileName,FileSize,FastPigMD5,FileUpdateTime)VALUES(@FileID,@FileName,@FileSize,@FastPigMD5,@FileUpdateTime)"
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@FileID", Data.SqlDbType.VarChar, 32)
                .AddPara("@FileName", Data.SqlDbType.VarChar, 256)
                .AddPara("@FileSize", Data.SqlDbType.Int)
                .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
                .AddPara("@FileUpdateTime", Data.SqlDbType.DateTime)
                .ActiveConnection = Me.mConnSQLSrv.Connection
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
            End With
            For Each InHostFile As HostFile In InHostDir.HostFiles
                With oCmdSQLSrvText
                    .ParaValue("@FileID") = InHostFile.FileID
                    .ParaValue("@FileName") = InHostFile.FileName
                    .ParaValue("@FileSize") = InHostFile.FileSize
                    .ParaValue("@FastPigMD5") = InHostFile.FastPigMD5
                    .ParaValue("@FileUpdateTime") = InHostFile.FileUpdateTime
                    LOG.Ret = .ExecuteNonQuery
                    If LOG.Ret <> "OK" Then Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
                End With
            Next
            oCmdSQLSrvText = Nothing
            With Me.mPigFunc
                strSQL = ""
                .AddMultiLineText(strSQL, "INSERT INTO dbo._ptHFFileInf(FileID,DirID,FileName,FileSize,FastPigMD5,FileUpdateTime,IsDel,IsCheck)")
                .AddMultiLineText(strSQL, "SELECT FileID,@DirID,FileName,FileSize,FastPigMD5,FileUpdateTime,0,0")
                .AddMultiLineText(strSQL, "FROM " & strTmpTabName & " t", 1)
                .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM dbo._ptHFFileInf p WHERE t.FileID=p.FileID AND p.DirID=@DirID)")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFFileInf")
                .AddMultiLineText(strSQL, "SET IsDel=1")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFFileInf p")
                .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM " & strTmpTabName & " t WHERE p.FileID=t.FileID)")
                .AddMultiLineText(strSQL, "AND p.IsDel=0 AND p.DirID=@DirID", 1)
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFFileInf")
                .AddMultiLineText(strSQL, "SET FileName=t.FileName")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFFileInf p JOIN " & strTmpTabName & " t ON p.FileID=t.FileID")
                .AddMultiLineText(strSQL, "WHERE p.FileName!=t.FileName")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFFileInf")
                .AddMultiLineText(strSQL, "SET IsDel=0")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFFileInf p JOIN " & strTmpTabName & " t ON p.FileID=t.FileID")
                .AddMultiLineText(strSQL, "WHERE p.IsDel=1")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFFileInf")
                .AddMultiLineText(strSQL, "SET IsCheck=0")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFFileInf p JOIN " & strTmpTabName & " t ON p.FileID=t.FileID")
                .AddMultiLineText(strSQL, "WHERE p.FileContID IS NULL AND IsCheck=1")
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFFileInf")
                .AddMultiLineText(strSQL, "SET FileSize=t.FileSize,FileUpdateTime=t.FileUpdateTime,FastPigMD5=t.FastPigMD5,IsCheck=1")
                .AddMultiLineText(strSQL, "FROM dbo._ptHFFileInf p JOIN " & strTmpTabName & " t ON p.FileID=t.FileID")
                .AddMultiLineText(strSQL, "WHERE (p.FileSize!=t.FileSize OR p.FileUpdateTime!=t.FileUpdateTime OR p.FastPigMD5!=t.FastPigMD5)")
                .AddMultiLineText(strSQL, "AND p.FileContID IS NOT NULL AND p.IsDel=0", 1)
                '---------
                .AddMultiLineText(strSQL, "UPDATE dbo._ptHFDirInf SET IsScan=0,LastScanTime=GETDATE() WHERE DirID=@DirID")
                '---------
                If InHostDir.fParent.fParent.fParent.IsDebug = False Then
                    .AddMultiLineText(strSQL, "DROP TABLE dbo." & strTmpTabName)
                End If
            End With
            LOG.StepName = "ExecuteNonQuery"
            oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
            With oCmdSQLSrvText
                .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
                .ParaValue("@DirID") = InHostDir.DirID
                Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
                .ActiveConnection = Me.mConnSQLSrv.Connection
                LOG.Ret = .ExecuteNonQuery
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Return "OK"
        Catch ex As Exception
            Me.fPrintErrLogInf(strSQL)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Overloads Function SetDebug() As String
        Try
            MyBase.SetDebug(Me.LogFilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetDebug", ex)
        End Try
    End Function


End Class
