'**********************************
'* Name: DBFileApp
'* Author: Seow Phong
'* Describe: Database file application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.16
'* Create Time: 5/3/2023
'* 1.1	6/3/2023    Add mCreateTable_DFConfInf
'* 1.2	7/3/2023    Modify mCreateTable_DFConfInf
'* 1.3	8/3/2023    Add mCreateTable_DFContSegInf,mCreateTable_DFPathInf, modify RefDBConn
'* 1.4	13/3/2023   Add mExecCmdSQLSrvText
'* 1.5	13/3/2023   Modify mNew
'* 1.6	15/3/2023   Modify New
'* 1.7	30/3/2023   Modify RefDBFolders,mAddNewDBFolder
'* 1.8	31/3/2023   Modify RefDBFolders
'* 1.9	3/4/2023    Modify RefDBFolders, add fRefDBFolder
'* 1.10	4/4/2023    Modify fBeginScanDBFolder
'* 1.11	5/4/2023    Add GetFolderType
'* 1.12	6/4/2023    Add GetFolderType, modify mPrintErrLogInf,fMergeDFPathInf
'* 1.13	8/4/2023    Modify RefDBConn
'* 1.15	10/4/2023   Modify New,AutoGetFolderType,fUpdDFFileInf,fUpdDBFolder and add fScanDFFileInf,fAddDFFileInf,mCreateTable_DFFolderInf
'* 1.16	11/4/2023   Add fCreateTable_TmpFileList,fInsTmpFileList,mCreateTable_DFDirInf, modify RefDBConn
'**********************************
Imports System.Data
Imports System.IO
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
Imports System.IO.Compression
Imports PigObjFsLib
#Else
Imports Microsoft.Data.SqlClient
#End If
Imports PigToolsLiteLib



''' <summary>
''' Application of file processing through database|通过数据库处理文件的应用
''' </summary>
Friend Class fDBFileApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.16.2"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    'Public Sub New(ConnSQLSrv As ConnSQLSrv, Optional LogDirPath As String = "")
    '    MyBase.New(CLS_VERSION)
    '    If LogDirPath <> "" Then
    '        Me.mLogDirPath = LogDirPath
    '    Else
    '        Dim strPath As String = Me.mPigFunc.GetPathPart(Me.AppPath, PigFunc.EnmFathPart.ParentPath) & Me.OsPathSep & "Log"
    '        Me.mLogDirPath = strPath
    '    End If
    '    If Me.mPigFunc.IsFolderExists(Me.mLogDirPath) = False Then
    '        Me.mPigFunc.CreateFolder(Me.mLogDirPath)
    '    End If
    '    Me.mRefLogFilePath()
    '    Me.mNew(ConnSQLSrv)
    'End Sub

    'Public ReadOnly Property DBFolders As New DBFolders
    'Private Property mConnSQLSrv As ConnSQLSrv
    'Private ReadOnly Property mPigFunc As New PigFunc
    'Private Property mSQLSrvTools As SQLSrvTools

    'Private ReadOnly Property mLogDirPath As String
    'Private Property mLogFilePath As String
    'Private ReadOnly Property mFS As New FileSystemObject
    'Private Property mLastExecCmdSQLSrvTextTime As DateTime


    'Friend ReadOnly Property LogFilePath As String
    '    Get
    '        Return Me.mLogDirPath
    '    End Get
    'End Property

    'Private Sub mNew(ConnSQLSrv As ConnSQLSrv)
    '    Dim LOG As New PigStepLog("mNew")
    '    Try
    '        LOG.StepName = "Set ConnSQLSrv"
    '        Me.mConnSQLSrv = ConnSQLSrv
    '        LOG.StepName = "RefDBConn"
    '        LOG.Ret = Me.RefDBConn(True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Me.mPrintErrLogInf(strErr)
    '        Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Sub

    'Private mIsDBReady As Boolean
    'Public ReadOnly Property IsDBReady As Boolean
    '    Get
    '        Return Me.mIsDBReady
    '    End Get
    'End Property

    '''' <summary>
    '''' Refresh database connection|刷新数据库连接
    '''' </summary>
    '''' <returns></returns>
    'Public Function RefDBConn(Optional IsCreateDBObj As Boolean = False) As String
    '    Dim LOG As New PigStepLog("RefDBConn")
    '    Try
    '        If Me.mConnSQLSrv.IsDBConnReady = False Then
    '            LOG.StepName = "OpenOrKeepActive"
    '            Me.mConnSQLSrv.ClearErr()
    '            Me.mConnSQLSrv.OpenOrKeepActive()
    '            If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
    '        End If
    '        If IsCreateDBObj = True Then
    '            Me.mSQLSrvTools = New SQLSrvTools(Me.mConnSQLSrv)
    '            LOG.StepName = "IsDBObjExists(_ptDFConfInf)"
    '            'If oSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFConfInf") = False Then
    '            '    LOG.StepName = "mCreateTable_DFConfInf"
    '            '    LOG.Ret = mCreateTable_DFConfInf()
    '            '    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            'End If
    '            'LOG.StepName = "IsDBObjExists(_ptDFPathInf)"
    '            LOG.StepName = "IsDBObjExists(_ptDFContInf)"
    '            If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFContInf") = False Then
    '                LOG.StepName = "mCreateTable_DFContInf"
    '                LOG.Ret = Me.mCreateTable_DFContInf
    '                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            End If
    '            LOG.StepName = "IsDBObjExists(_ptDFContSegInf)"
    '            If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFContSegInf") = False Then
    '                LOG.StepName = "mCreateTable_DFContSegInf"
    '                LOG.Ret = Me.mCreateTable_DFContSegInf()
    '                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            End If
    '            LOG.StepName = "IsDBObjExists(_ptDFFolderInf)"
    '            If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFFolderInf") = False Then
    '                LOG.StepName = "mCreateTable_DFFolderInf"
    '                LOG.Ret = Me.mCreateTable_DFFolderInf()
    '                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            End If
    '            LOG.StepName = "IsDBObjExists(_ptDFDirInf)"
    '            If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFDirInf") = False Then
    '                LOG.StepName = "mCreateTable_DFDirInf"
    '                LOG.Ret = Me.mCreateTable_DFDirInf
    '                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            End If
    '            LOG.StepName = "IsDBObjExists(_ptDFFileInf)"
    '            If Me.mSQLSrvTools.IsDBObjExists(SQLSrvTools.EnmDBObjType.UserTable, "_ptDFFileInf") = False Then
    '                LOG.StepName = "mCreateTable_DFFileInf"
    '                LOG.Ret = Me.mCreateTable_DFFileInf
    '                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            End If
    '        End If
    '        Me.mIsDBReady = True
    '        Return "OK"
    '    Catch ex As Exception
    '        Me.mIsDBReady = False
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function


    'Private Function mCreateTable_DFConfInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFConfInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFConfInf(")
    '            .AddMultiLineText(strSQL, "ConfID varchar(64) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",ConfValue varchar(8000) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",CodeDesc varchar(512) NOT NULL DEFAULT('')", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_ptDFConfInf PRIMARY KEY CLUSTERED(ConfID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mCreateTable_DFFolderInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFFolderInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFFolderInf(")
    '            .AddMultiLineText(strSQL, "FolderID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",HostID varchar(64)", 1)
    '            .AddMultiLineText(strSQL, ",FolderName varchar(256) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FolderPath varchar(2048) NOT NULL", 1)  '绝对路径
    '            .AddMultiLineText(strSQL, ",FolderType int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FolderDesc varchar(512) NOT NULL DEFAULT('')", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
    '            .AddMultiLineText(strSQL, ",IsUse bit NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",ScanStatus int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",StaticInf varchar(8000) NOT NULL DEFAULT('')", 1)
    '            .AddMultiLineText(strSQL, ",ActiveInf varchar(max) NOT NULL DEFAULT('')", 1)
    '            .AddMultiLineText(strSQL, ",ScanBeginTime datetime ", 1)
    '            .AddMultiLineText(strSQL, ",ScanEndTime datetime ", 1)
    '            .AddMultiLineText(strSQL, ",ActiveTime datetime ", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_ptDFFolderInf PRIMARY KEY CLUSTERED(FolderID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '            .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptDFFolderInf_HostID_FolderPath ON _ptDFFolderInf(FolderPath,HostID)")
    '            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFFolderInf_FolderName ON _ptDFFolderInf(FolderName)")
    '            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFFolderInf_HostID ON _ptDFFolderInf(HostID)")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fInsTmpFileList(FileListPath As String, FolderID As String) As String
    '    Dim LOG As New PigStepLog("fInsTmpFileList")
    '    Dim strTmpTabName As String = "_ptmpFileList" & FolderID
    '    Dim strSQL As String = "INSERT INTO dbo." & strTmpTabName & "(FilePath,FileSize,FileUpdateTime,FastPigMD5)VALUES(@FilePath,@FileSize,@FileUpdateTime,@FastPigMD5)"
    '    Try
    '        LOG.StepName = "OpenTextFile_FileListPath"
    '        Dim tsRead As TextStream = Me.mFS.OpenTextFile(FileListPath, FileSystemObject.IOMode.ForReading)
    '        If Me.mFS.LastErr <> "" Then
    '            LOG.AddStepNameInf(FileListPath)
    '            Throw New Exception(Me.mFS.LastErr)
    '        End If
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FilePath", Data.SqlDbType.VarChar, 4096)
    '            .AddPara("@FileSize", Data.SqlDbType.BigInt)
    '            .AddPara("@FileUpdateTime", Data.SqlDbType.DateTime)
    '            .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '        End With
    '        Dim lngLineNo As Long = 1
    '        Do While Not tsRead.AtEndOfStream
    '            Dim strLine As String = tsRead.ReadLine
    '            Dim strFilePath As String = Me.mPigFunc.GetStr(strLine, "", vbTab)
    '            Dim strFileSize As String = Me.mPigFunc.GetStr(strLine, "", vbTab)
    '            Dim strFileUpdateTime As String = Me.mPigFunc.GetStr(strLine, "", vbTab)
    '            Dim strFastPigMD5 As String = strLine
    '            If Left(strFilePath, 1) <> "." Then
    '                LOG.StepName = strFilePath
    '                LOG.AddStepNameInf("LineNo:" & lngLineNo)
    '                Throw New Exception("Not a relative path")
    '            End If
    '            'Dim strFileID As String = ""
    '            'LOG.StepName = "Check data or mGetFileID"
    '            'If IsNumeric(strFileSize) = False Then
    '            '    LOG.Ret = "FileSize is " & strFileSize & ", not a numeric"
    '            'ElseIf IsDate(strFileUpdateTime) = False Then
    '            '    LOG.Ret = "FileUpdateTime is " & strFileUpdateTime & ", not a datetime"
    '            'ElseIf Len(strFastPigMD5) <> 32 Then
    '            '    LOG.Ret = "Invalid FastPigMD5(" & strFastPigMD5 & ")"
    '            'Else
    '            '    LOG.Ret = Me.mGetFileID(strFilePath, strFileID)
    '            'End If
    '            'If LOG.Ret <> "OK" Then
    '            '    Me.fParent.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            'Else
    '            '    LOG.StepName = "fScanDFFileInf"
    '            '    LOG.Ret = Me.fParent.fInsTmpFileList(Me.FolderID, strFilePath, CLng(strFileSize), CDate(strFileUpdateTime), strFastPigMD5)
    '            '    If LOG.Ret <> "OK" Then Me.fParent.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            'End If
    '            lngLineNo += 1
    '        Loop
    '        tsRead.Close()
    '        tsRead = Nothing

    '        'With oCmdSQLSrvText
    '        '    .AddPara("@FilePath", Data.SqlDbType.VarChar, 4096)
    '        '    .ParaValue("@FilePath") = FilePath
    '        '    .AddPara("@FileSize", Data.SqlDbType.BigInt)
    '        '    .ParaValue("@FileSize") = FileSize
    '        '    .AddPara("@FileUpdateTime", Data.SqlDbType.DateTime)
    '        '    .ParaValue("@FileUpdateTime") = FileUpdateTime
    '        '    .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '        '    .ParaValue("@FastPigMD5") = FastPigMD5
    '        '    .ActiveConnection = Me.mConnSQLSrv.Connection
    '        '    LOG.StepName = "ExecuteNonQuery"
    '        '    LOG.Ret = .ExecuteNonQuery()
    '        '    If LOG.Ret <> "OK" Then
    '        '        Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '        '        Throw New Exception(LOG.Ret)
    '        '    End If
    '        'End With

    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fMergeDBDirInf(ByRef InDBFolder As DBFolder) As String
    '    Dim LOG As New PigStepLog("fMergeDBDirInf")
    '    Dim strTmpTabName As String = ""
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Generate SQL"
    '        strTmpTabName = "_ptmpDirList_" & InDBFolder.FolderID
    '        With Me.mPigFunc
    '            strSQL = ""
    '            .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM sysobjects WHERE name='" & strTmpTabName & "' AND xtype='U') DROP TABLE dbo." & strTmpTabName)
    '            .AddMultiLineText(strSQL, "CREATE TABLE dbo." & strTmpTabName & "(")
    '            .AddMultiLineText(strSQL, "DirID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirPath varchar(4096) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirSize money NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirFiles int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirUpdateTime datetime NOT NULL", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_" & strTmpTabName & " PRIMARY KEY CLUSTERED(DirPath)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '        End With
    '        LOG.StepName = "ExecuteNonQuery"
    '        Dim oCmdSQLSrvText As CmdSQLSrvText
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        End With
    '        oCmdSQLSrvText = Nothing
    '        strSQL = "INSERT INTO dbo." & strTmpTabName & "(DirID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime)VALUES(@DirID,@DirPath,@DirSize,@DirFiles,@FastPigMD5,@DirUpdateTime)"
    '        LOG.StepName = "ExecuteNonQuery"
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
    '            .AddPara("@DirPath", Data.SqlDbType.VarChar, 4096)
    '            .AddPara("@DirSize", Data.SqlDbType.Money)
    '            .AddPara("@DirFiles", Data.SqlDbType.Int)
    '            .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '            .AddPara("@DirUpdateTime", Data.SqlDbType.DateTime)
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '        End With
    '        For Each oDBDir As DBDir In InDBFolder.DBDirs
    '            With oCmdSQLSrvText
    '                .ParaValue("@DirID") = oDBDir.DirID
    '                .ParaValue("@DirPath") = oDBDir.DirPath
    '                .ParaValue("@DirSize") = oDBDir.DirSize
    '                .ParaValue("@DirFiles") = oDBDir.DirFiles
    '                .ParaValue("@FastPigMD5") = oDBDir.FastPigMD5
    '                .ParaValue("@DirUpdateTime") = oDBDir.DirUpdateTime
    '                LOG.Ret = .ExecuteNonQuery
    '                If LOG.Ret <> "OK" Then Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            End With
    '        Next
    '        oCmdSQLSrvText = Nothing
    '        With Me.mPigFunc
    '            strSQL = ""
    '            .AddMultiLineText(strSQL, "INSERT INTO dbo._ptDFDirInf(DirID,FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,IsDel,MateStatus)")
    '            .AddMultiLineText(strSQL, "SELECT DirID,@FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime," & DBDir.EnmMateStatus.Different)
    '            .AddMultiLineText(strSQL, "FROM " & strTmpTabName & " t", 1)
    '            .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM dbo._ptDFDirInf p WHERE t.DirID=p.DirID AND p.FolderID=@FolderID)")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET IsDel=1")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p")
    '            .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM " & strTmpTabName & " t WHERE p.DirID=t.DirID AND p.FolderID=@FolderID)")
    '            .AddMultiLineText(strSQL, "AND p.IsDel=0", 1)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirPath=t.DirPath")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.DirPath!=t.DirPath")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirSize=t.DirSize,DirFiles=t.DirFiles,MateStatus=" & DBDir.EnmMateStatus.Different)
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE (p.DirSize!=t.DirSize OR p.DirFiles!=t.DirFiles) AND p.MateStatus!=" & DBDir.EnmMateStatus.Different)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET FastPigMD5=t.FastPigMD5,MateStatus=" & DBDir.EnmMateStatus.BasicallySame)
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.FastPigMD5!=t.FastPigMD5 AND p.MateStatus=" & DBDir.EnmMateStatus.Same)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirUpdateTime=t.DirUpdateTime")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.DirUpdateTime!=t.DirUpdateTime")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET IsDel=0")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.IsDel=1")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET MateStatus=" & DBDir.EnmMateStatus.Different)
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.MateStatus NOT IN(" & DBDir.EnmMateStatus.Same & "," & DBDir.EnmMateStatus.Different & "," & DBDir.EnmMateStatus.BasicallySame & ")")
    '            '---------
    '            '.AddMultiLineText(strSQL, "DROP TABLE dbo." & strTmpTabName)
    '        End With
    '        LOG.StepName = "ExecuteNonQuery"
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InDBFolder.FolderID
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.Ret = .ExecuteNonQuery
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Me.PrintDebugLog(Me.MyClassName, strErr)
    '        Return strErr
    '    End Try
    'End Function

    'Friend Function fMergeDBFileInf(ByRef InDBDir As DBDir, Optional IsIncBasicallySame As Boolean = False) As String
    '    Dim LOG As New PigStepLog("fMergeDBFileInf")
    '    Dim strTmpTabName As String = ""
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Generate SQL"
    '        strTmpTabName = "_ptmpFileList_" & InDBDir.FolderID
    '        With Me.mPigFunc
    '            strSQL = ""
    '            .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM sysobjects WHERE name='" & strTmpTabName & "' AND xtype='U') DROP TABLE dbo." & strTmpTabName)
    '            .AddMultiLineText(strSQL, "CREATE TABLE dbo." & strTmpTabName & "(")
    '            .AddMultiLineText(strSQL, "FileID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileName varchar(256) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileSize int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileUpdateTime datetime NOT NULL", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_" & strTmpTabName & " PRIMARY KEY CLUSTERED(FileID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '        End With
    '        LOG.StepName = "ExecuteNonQuery"
    '        Dim oCmdSQLSrvText As CmdSQLSrvText
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        End With
    '        oCmdSQLSrvText = Nothing
    '        strSQL = "INSERT INTO dbo." & strTmpTabName & "(FileID,FileName,FileSize,FastPigMD5,FileUpdateTime)VALUES(@FileID,@FileName,@FileSize,@FastPigMD5,@FileUpdateTime)"
    '        LOG.StepName = "ExecuteNonQuery"
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
    '            .AddPara("@DirPath", Data.SqlDbType.VarChar, 4096)
    '            .AddPara("@DirSize", Data.SqlDbType.Money)
    '            .AddPara("@DirFiles", Data.SqlDbType.Int)
    '            .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '            .AddPara("@DirUpdateTime", Data.SqlDbType.DateTime)
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '        End With
    '        For Each oDBFile As DBFile In InDBDir.DBFiles
    '            With oCmdSQLSrvText
    '                .ParaValue("@DirID") = oDBDir.DirID
    '                .ParaValue("@DirPath") = oDBDir.DirPath
    '                .ParaValue("@DirSize") = oDBDir.DirSize
    '                .ParaValue("@DirFiles") = oDBDir.DirFiles
    '                .ParaValue("@FastPigMD5") = oDBDir.FastPigMD5
    '                .ParaValue("@DirUpdateTime") = oDBDir.DirUpdateTime
    '                LOG.Ret = .ExecuteNonQuery
    '                If LOG.Ret <> "OK" Then Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            End With
    '        Next
    '        oCmdSQLSrvText = Nothing
    '        With Me.mPigFunc
    '            strSQL = ""
    '            .AddMultiLineText(strSQL, "INSERT INTO dbo._ptDFDirInf(DirID,FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,IsDel)")
    '            .AddMultiLineText(strSQL, "SELECT DirID,@FolderID,DirPath,DirSize,DirFiles,FastPigMD5,DirUpdateTime,0")
    '            .AddMultiLineText(strSQL, "FROM " & strTmpTabName & " t", 1)
    '            .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM dbo._ptDFDirInf p WHERE t.DirID=p.DirID AND p.FolderID=@FolderID)")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET IsDel=1")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p")
    '            .AddMultiLineText(strSQL, "WHERE NOT EXISTS(SELECT 1 FROM " & strTmpTabName & " t WHERE p.DirID=t.DirID AND p.FolderID=@FolderID)")
    '            .AddMultiLineText(strSQL, "AND p.IsDel=0", 1)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirPath=t.DirPath")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.DirPath!=t.DirPath")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirSize=t.DirSize,DirFiles=t.DirFiles,MateStatus=" & DBDir.EnmMateStatus.Different)
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE (p.DirSize!=t.DirSize OR p.DirFiles!=t.DirFiles) AND p.MateStatus!=" & DBDir.EnmMateStatus.Different)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET FastPigMD5=t.FastPigMD5,MateStatus=" & DBDir.EnmMateStatus.BasicallySame)
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.FastPigMD5!=t.FastPigMD5 AND p.MateStatus!=" & DBDir.EnmMateStatus.BasicallySame)
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET DirUpdateTime=t.DirUpdateTime")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.DirUpdateTime!=t.DirUpdateTime")
    '            '---------
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFDirInf")
    '            .AddMultiLineText(strSQL, "SET IsDel=0")
    '            .AddMultiLineText(strSQL, "FROM dbo._ptDFDirInf p JOIN " & strTmpTabName & " t ON p.DirID=t.DirID")
    '            .AddMultiLineText(strSQL, "WHERE p.IsDel=1")
    '            '---------
    '            '.AddMultiLineText(strSQL, "DROP TABLE dbo." & strTmpTabName)
    '        End With
    '        LOG.StepName = "ExecuteNonQuery"
    '        oCmdSQLSrvText = New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InDBDir.FolderID
    '            Me.PrintDebugLog(LOG.StepLogInf, .DebugStr)
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.Ret = .ExecuteNonQuery
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Me.PrintDebugLog(Me.MyClassName, strErr)
    '        Return strErr
    '    End Try
    'End Function

    'Friend Function fCreateTable_TmpFileList(FolderID As String) As String
    '    Dim LOG As New PigStepLog("fCreateTable_TmpFileList")
    '    Dim strSQL As String = ""
    '    Try
    '        Dim strTmpTabName As String = "_ptmpFileList" & FolderID
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM sysobjects WHERE name='" & strTmpTabName & "' AND xtype='U') DROP TABLE dbo." & strTmpTabName)
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo." & strTmpTabName)
    '            .AddMultiLineText(strSQL, "(")
    '            .AddMultiLineText(strSQL, "DirID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, "FilePath varchar(4096) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileSize bigint NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileUpdateTime datetime NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_" & strTmpTabName & " PRIMARY KEY CLUSTERED(FilePath)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mCreateTable_DFDirInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFDirInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFDirInf(")
    '            .AddMultiLineText(strSQL, "DirID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FolderID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirPath varchar(4096) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirSize money NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirFiles int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",DirUpdateTime datetime NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, ",IsDel bit NOT NULL DEFAULT(0)", 1)
    '            .AddMultiLineText(strSQL, ",MateStatus int NOT NULL DEFAULT(0)", 1)
    '            .AddMultiLineText(strSQL, ",LastScanTime datetime", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK__ptDFDirInf PRIMARY KEY CLUSTERED(DirID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFDirInf WITH CHECK ADD CONSTRAINT FK_ptDFDirInf_ptDFFolderInf FOREIGN KEY(FolderID) REFERENCES _ptDFFolderInf(FolderID)")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFDirInf CHECK CONSTRAINT FK_ptDFDirInf_ptDFFolderInf")
    '            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFDirInf_FolderID ON _ptDFDirInf(FolderID)")
    '            .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptDFDirInf_FilePath_FolderID ON _ptDFDirInf(DirPath,FolderID)")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mCreateTable_DFFileInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFFileInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFFileInf(")
    '            .AddMultiLineText(strSQL, "FileID varchar(36) NOT NULL DEFAULT(newid())", 1)
    '            .AddMultiLineText(strSQL, ",DirID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileName varchar(256) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileContID varchar(32)", 1)
    '            .AddMultiLineText(strSQL, ",FileSize int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FastPigMD5 varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",FileUpdateTime datetime NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, ",IsDel bit NOT NULL DEFAULT(0)", 1)
    '            .AddMultiLineText(strSQL, ",ContStatus int NOT NULL DEFAULT(0)", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_ptDFFileInf PRIMARY KEY CLUSTERED(FileID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFFileInf WITH CHECK ADD CONSTRAINT FK_ptDFFileInf_ptDFDirInf FOREIGN KEY(DirID) REFERENCES _ptDFDirInf(DirID)")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFFileInf CHECK CONSTRAINT FK_ptDFFileInf_ptDFDirInf")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFFileInf WITH CHECK ADD CONSTRAINT FK_ptDFFileInf_ptDFContInf FOREIGN KEY(FileContID) REFERENCES _ptDFContInf(FileContID)")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFFileInf CHECK CONSTRAINT FK_ptDFFileInf_ptDFContInf")
    '            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFFileInf_FileContID ON _ptDFFileInf(FileContID)")
    '            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFFileInf_FileName ON _ptDFFileInf(FileName)")
    '            .AddMultiLineText(strSQL, "CREATE UNIQUE INDEX UI_ptDFFileInf_DirID_FileName ON _ptDFFileInf(DirID,FileName)")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mCreateTable_DFContInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFContInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFContInf(")
    '            .AddMultiLineText(strSQL, "FileContID varchar(32) NOT NULL", 1) '文件的FullPigMD5
    '            .AddMultiLineText(strSQL, ",FileSize bigint NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",ContStatus int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",SegContType int", 1)
    '            .AddMultiLineText(strSQL, ",SegSize int", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, ",UpdateTime datetime NULL", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_ptDFContInf PRIMARY KEY CLUSTERED(FileContID)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mCreateTable_DFContSegInf() As String
    '    Dim LOG As New PigStepLog("mCreateTable_DFContSegInf")
    '    Dim strSQL As String = ""
    '    Try
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFContSegInf(")
    '            .AddMultiLineText(strSQL, "FileContID varchar(32) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",SegNo int NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",SegCont varchar(max) NOT NULL", 1)
    '            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    '            .AddMultiLineText(strSQL, "CONSTRAINT PK_DFContSegInf PRIMARY KEY CLUSTERED(FileContID,SegNo)", 1)
    '            .AddMultiLineText(strSQL, ")")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFContSegInf WITH CHECK ADD CONSTRAINT FK_DFContSegInf_ptDFContInf FOREIGN KEY(FileContID) REFERENCES _ptDFContInf(FileContID)")
    '            .AddMultiLineText(strSQL, "ALTER TABLE _ptDFContSegInf CHECK CONSTRAINT FK_DFContSegInf_ptDFContInf")
    '            '.AddMultiLineText(strSQL, "ALTER TABLE _ptDFContSegInf WITH CHECK ADD CONSTRAINT FK_DFContSegInf_ptDFSegInf FOREIGN KEY(ContSegID) REFERENCES _ptDFSegInf(ContSegID)")
    '            '.AddMultiLineText(strSQL, "ALTER TABLE _ptDFContSegInf CHECK CONSTRAINT FK_DFContSegInf_ptDFSegInf")
    '            '.AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFContSegInf_FileContID ON _ptDFContSegInf(FileContID)")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .ActiveConnection = Me.mConnSQLSrv.Connection
    '            LOG.StepName = "ExecuteNonQuery"
    '            LOG.Ret = .ExecuteNonQuery()
    '            If LOG.Ret <> "OK" Then
    '                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End With
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    ''Private Function mCreateTable_DFSegInf() As String
    ''    Dim LOG As New PigStepLog("mCreateTable_DFContInf")
    ''    Dim strSQL As String = ""
    ''    Try
    ''        With Me.mPigFunc
    ''            .AddMultiLineText(strSQL, "CREATE TABLE	dbo._ptDFSegInf(")
    ''            .AddMultiLineText(strSQL, "ContSegID varchar(32) NOT NULL", 1)  '段内容的PigMD5
    ''            .AddMultiLineText(strSQL, ",SegCont varchar(max) NOT NULL", 1)
    ''            .AddMultiLineText(strSQL, ",CreateTime datetime NOT NULL DEFAULT(GetDate())", 1)
    ''            .AddMultiLineText(strSQL, "CONSTRAINT PK_ptDFSegInf PRIMARY KEY CLUSTERED(ContSegID)", 1)
    ''            .AddMultiLineText(strSQL, ")")
    ''            .AddMultiLineText(strSQL, "CREATE INDEX IDX_ptDFSegInf_FileContID ON _ptDFSegInf(ContSegID)")
    ''        End With
    ''        LOG.StepName = "New CmdSQLSrvText"
    ''        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    ''        With oCmdSQLSrvText
    ''            .ActiveConnection = Me.mConnSQLSrv.Connection
    ''            LOG.StepName = "ExecuteNonQuery"
    ''            LOG.Ret = .ExecuteNonQuery()
    ''            If LOG.Ret <> "OK" Then
    ''                Me.PrintDebugLog(LOG.SubName, LOG.StepLogInf)
    ''                Throw New Exception(LOG.Ret)
    ''            End If
    ''        End With
    ''        Return "OK"
    ''    Catch ex As Exception
    ''        LOG.AddStepNameInf(strSQL)
    ''        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    ''    End Try
    ''End Function

    '''' <summary>
    '''' 新建数据库文件夹|New Database Folder
    '''' </summary>
    '''' <param name="HostID">主机标识|Host ID</param>
    '''' <param name="FolderPath">文件夹绝对路径|Folder absolute path</param>
    '''' <param name="IsLocalPath">是否本地路径|Is Local path</param>
    '''' <returns></returns>
    'Public Function AddNewDBFolder(HostID As String, FolderPath As String, Optional IsLocalPath As Boolean = False) As String
    '    Dim LOG As New PigStepLog("AddNewDBFolder")
    '    Try
    '        If IsLocalPath = True Then
    '            LOG.StepName = "Check FolderPath"
    '            If Me.mPigFunc.IsFolderExists(FolderPath) = False Then Throw New Exception("Invalid folder")
    '        End If
    '        LOG.StepName = "New DBFolder"
    '        Dim oDBFolder As New DBFolder(HostID, FolderPath, Me)
    '        If oDBFolder.LastErr <> "" Then Throw New Exception(oDBFolder.LastErr)
    '        LOG.StepName = "mAddNewDBFolder"
    '        LOG.Ret = Me.mAddNewDBFolder(oDBFolder)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(FolderPath)
    '        Dim strRet As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Me.mPrintErrLogInf(strRet)
    '        Return strRet
    '    End Try
    'End Function

    'Friend Function fRefDBFolder(ByRef InDBFolder As DBFolder, Optional IsDirtyRead As Boolean = True) As String
    '    Dim LOG As New PigStepLog("RefDBFolders")
    '    Dim strSQL As String = "SELECT * FROM dbo._ptDFFolderInf"
    '    If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
    '    strSQL &= " WHERE FolderID=@FolderID"
    '    Try
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InDBFolder.FolderID
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        Dim strFolderID As String = rsMain.Fields.Item("FolderID").StrValue
    '        LOG.StepName = "DBFolders.AddOrGet"
    '        InDBFolder = Me.DBFolders.AddOrGet(strFolderID, Me)
    '        If Me.DBFolders.LastErr <> "" Then
    '            LOG.AddStepNameInf(strFolderID)
    '            Throw New Exception(LOG.Ret)
    '        End If
    '        LOG.StepName = "fFillByRs"
    '        LOG.Ret = InDBFolder.fFillByRs(rsMain)
    '        If LOG.Ret <> "OK" Then
    '            LOG.AddStepNameInf(strFolderID)
    '            Throw New Exception(LOG.Ret)
    '        End If
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fRefDBDirs(ByRef InDBFolder As DBFolder, Optional IsDirtyRead As Boolean = True) As String
    '    Dim LOG As New PigStepLog("fRefDBDir")
    '    Dim strSQL As String = "SELECT * FROM dbo._ptDFDirInf"
    '    If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
    '    strSQL &= " WHERE FolderID=@FolderID"
    '    strSQL &= " ORDER BY DirID"
    '    Try
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InDBFolder.FolderID
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        LOG.StepName = "DBDirs.Clear"
    '        InDBFolder.DBDirs.Clear()
    '        Do While Not rsMain.EOF
    '            Dim strDirID As String = rsMain.Fields.Item("DirID").StrValue
    '            LOG.StepName = "DBFolders.AddOrGet"
    '            Dim oDBDir As DBDir = InDBFolder.DBDirs.AddOrGet(strDirID, InDBFolder)
    '            If InDBFolder.DBDirs.LastErr <> "" Then
    '                LOG.AddStepNameInf(strDirID)
    '                Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            End If
    '            LOG.StepName = "fFillByRs"
    '            LOG.Ret = oDBDir.fFillByRs(rsMain)
    '            If LOG.Ret <> "OK" Then
    '                LOG.AddStepNameInf(strDirID)
    '                Me.PrintDebugLog(Me.MyClassName, LOG.StepLogInf)
    '            End If
    '            rsMain.MoveNext()
    '        Loop
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fRefDBDir(ByRef InDBFolder As DBFolder, ByRef InObj As DBDir, Optional IsDirtyRead As Boolean = True) As String
    '    Dim LOG As New PigStepLog("fRefDBDir")
    '    Dim strSQL As String = "SELECT * FROM dbo._ptDFDirInf"
    '    If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
    '    strSQL &= " WHERE DirID=@DirID"
    '    Try
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@DirID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@DirID") = InObj.DirID
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        Dim strDirID As String = rsMain.Fields.Item("DirID").StrValue
    '        LOG.StepName = "DBFolders.AddOrGet"
    '        InObj = InDBFolder.DBDirs.AddOrGet(strDirID, InDBFolder)
    '        If InDBFolder.DBDirs.LastErr <> "" Then
    '            LOG.AddStepNameInf(strDirID)
    '            Throw New Exception(LOG.Ret)
    '        End If
    '        LOG.StepName = "fFillByRs"
    '        LOG.Ret = InDBFolder.fFillByRs(rsMain)
    '        If LOG.Ret <> "OK" Then
    '            LOG.AddStepNameInf(strDirID)
    '            Throw New Exception(LOG.Ret)
    '        End If
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fUpdDBFolder(ByRef InObj As DBFolder) As String
    '    Dim LOG As New PigStepLog("fUpdDBFolder")
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Generate SQL"
    '        Dim strUpdCols As String = ""
    '        If InObj.IsUpdate("HostID") = True Then strUpdCols &= ",HostID=@HostID"
    '        If InObj.IsUpdate("FolderName") = True Then strUpdCols &= ",FolderName=@FolderName"
    '        If InObj.IsUpdate("FolderPath") = True Then strUpdCols &= ",FolderPath=@FolderPath"
    '        If InObj.IsUpdate("FolderType") = True Then strUpdCols &= ",FolderType=@FolderType"
    '        If InObj.IsUpdate("FolderDesc") = True Then strUpdCols &= ",FolderDesc=@FolderDesc"
    '        If InObj.IsUpdate("CreateTime") = True Then strUpdCols &= ",CreateTime=@CreateTime"
    '        If InObj.IsUpdate("UpdateTime") = True Then strUpdCols &= ",UpdateTime=@UpdateTime"
    '        If InObj.IsUpdate("IsUse") = True Then strUpdCols &= ",IsUse=@IsUse"
    '        If InObj.IsUpdate("ScanStatus") = True Then strUpdCols &= ",ScanStatus=@ScanStatus"
    '        If InObj.IsUpdate("StaticInf") = True Then strUpdCols &= ",StaticInf=@StaticInf"
    '        If InObj.IsUpdate("ActiveInf") = True Then strUpdCols &= ",ActiveInf=@ActiveInf"
    '        If InObj.IsUpdate("ScanBeginTime") = True Then strUpdCols &= ",ScanBeginTime=@ScanBeginTime"
    '        If InObj.IsUpdate("ScanEndTime") = True Then strUpdCols &= ",ScanEndTime=@ScanEndTime"
    '        If InObj.IsUpdate("ActiveTime") = True Then strUpdCols &= ",ActiveTime=@ActiveTime"
    '        If strUpdCols = "" Then Throw New Exception("There is nothing to update")
    '        strUpdCols = Mid(strUpdCols, 2)
    '        strSQL = "UPDATE dbo._ptDFFolderInf SET " & strUpdCols & " WHERE FolderID=@FolderID"
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InObj.FolderID
    '            If InObj.IsUpdate("HostID") = True Then
    '                .AddPara("@HostID", Data.SqlDbType.VarChar, 64)
    '                .ParaValue("@HostID") = InObj.HostID
    '            End If
    '            If InObj.IsUpdate("FolderName") = True Then
    '                .AddPara("@FolderName", Data.SqlDbType.VarChar, 256)
    '                .ParaValue("@FolderName") = InObj.FolderName
    '            End If
    '            If InObj.IsUpdate("FolderPath") = True Then
    '                .AddPara("@FolderPath", Data.SqlDbType.VarChar, 2048)
    '                .ParaValue("@FolderPath") = InObj.FolderPath
    '            End If
    '            If InObj.IsUpdate("FolderType") = True Then
    '                .AddPara("@FolderType", Data.SqlDbType.Int)
    '                .ParaValue("@FolderType") = InObj.FolderType
    '            End If
    '            If InObj.IsUpdate("FolderDesc") = True Then
    '                .AddPara("@FolderDesc", Data.SqlDbType.VarChar, 512)
    '                .ParaValue("@FolderDesc") = InObj.FolderDesc
    '            End If
    '            If InObj.IsUpdate("CreateTime") = True Then
    '                .AddPara("@CreateTime", Data.SqlDbType.DateTime)
    '                .ParaValue("@CreateTime") = InObj.CreateTime
    '            End If
    '            If InObj.IsUpdate("UpdateTime") = True Then
    '                .AddPara("@UpdateTime", Data.SqlDbType.DateTime)
    '                .ParaValue("@UpdateTime") = InObj.UpdateTime
    '            End If
    '            If InObj.IsUpdate("IsUse") = True Then
    '                .AddPara("@IsUse", Data.SqlDbType.Bit)
    '                .ParaValue("@IsUse") = InObj.IsUse
    '            End If
    '            If InObj.IsUpdate("ScanStatus") = True Then
    '                .AddPara("@ScanStatus", Data.SqlDbType.Int)
    '                .ParaValue("@ScanStatus") = InObj.ScanStatus
    '            End If
    '            If InObj.IsUpdate("StaticInf") = True Then
    '                .AddPara("@StaticInf", Data.SqlDbType.VarChar, 8000)
    '                .ParaValue("@StaticInf") = InObj.StaticInf
    '            End If
    '            If InObj.IsUpdate("ActiveInf") = True Then
    '                .AddPara("@ActiveInf", Data.SqlDbType.VarChar, 8000)
    '                .ParaValue("@ActiveInf") = InObj.ActiveInf
    '            End If
    '            If InObj.IsUpdate("ScanBeginTime") = True Then
    '                .AddPara("@ScanBeginTime", Data.SqlDbType.DateTime)
    '                .ParaValue("@ScanBeginTime") = InObj.ScanBeginTime
    '            End If
    '            If InObj.IsUpdate("ScanEndTime") = True Then
    '                .AddPara("@ScanEndTime", Data.SqlDbType.DateTime)
    '                .ParaValue("@ScanEndTime") = InObj.ScanEndTime
    '            End If
    '            If InObj.IsUpdate("ActiveTime") = True Then
    '                .AddPara("@ActiveTime", Data.SqlDbType.DateTime)
    '                .ParaValue("@ActiveTime") = InObj.ActiveTime
    '            End If
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fScanDFFileInf(FileID As String, FileSize As Long, FastPigMD5 As String, FilePath As String, FileUpdateTime As Date) As String
    '    Dim LOG As New PigStepLog("fScanDFFileInf")
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Generate SQL"
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "DECLARE @Rows int = 0")
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET LastScanTime=GetDate() WHERE FileID=@FileID", 1)
    '            .AddMultiLineText(strSQL, "IF @@ROWCOUNT>0")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FastPigMD5=@FastPigMD5 WHERE FileID=@FileID AND FastPigMD5!=@FastPigMD5", 1)
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FileSize=@FileSize WHERE FileID=@FileID AND FileSize!=@FileSize")
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FileUpdateTime=@FileUpdateTime WHERE FileID=@FileID AND FileUpdateTime!=@FileUpdateTime")
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "END")
    '            .AddMultiLineText(strSQL, "SELECT @Rows Rows")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FileID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FileID") = FileID
    '            .AddPara("@FileSize", Data.SqlDbType.BigInt)
    '            .ParaValue("@FileSize") = FileSize
    '            .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FastPigMD5") = FastPigMD5
    '            .AddPara("@FileUpdateTime", Data.SqlDbType.DateTime)
    '            .ParaValue("@FileUpdateTime") = FileUpdateTime
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        If rsMain.Fields.Item("Rows").LongValue = 0 Then
    '            LOG.StepName = ""
    '            LOG.Ret = Me.fAddDFFileInf(FileID, FileSize, FastPigMD5, FilePath, FileUpdateTime)
    '            If LOG.Ret <> "OK" Then
    '                If rsMain IsNot Nothing Then rsMain.Close()
    '                rsMain = Nothing
    '                oCmdSQLSrvText = Nothing
    '                Throw New Exception(LOG.Ret)
    '            End If
    '        End If
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Friend Function fAddDFFileInf(FileID As String, FileSize As Long, FastPigMD5 As String, FilePath As String, FileUpdateTime As Date) As String
    '    Dim LOG As New PigStepLog("fAddDFFileInf")
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Generate SQL"
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "DECLARE @Rows int = 0")
    '            .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM dbo._ptDFFileInf WHERE FileID=@FileID)")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "INSERT INTO dbo._ptDFFileInf(FileID,FileSize,FastPigMD5,FilePath,FileUpdateTime,LastScanTime)VALUES(@FileID,@FileSize,@FastPigMD5,@FilePath,@FileUpdateTime,GetDate())", 1)
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "END")
    '            .AddMultiLineText(strSQL, "ELSE")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FastPigMD5=@FastPigMD5 WHERE FileID=@FileID AND FastPigMD5!=@FastPigMD5", 1)
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET LastScanTime=GetDate() WHERE FileID=@FileID", 1)
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FilePath=@FilePath WHERE FileID=@FileID AND FilePath!=@FilePath", 1)
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET FileUpdateTime=@FileUpdateTime WHERE FileID=@FileID AND FileUpdateTime!=@FileUpdateTime")
    '            .AddMultiLineText(strSQL, "SET @Rows=@Rows+@@ROWCOUNT", 1)
    '            .AddMultiLineText(strSQL, "UPDATE dbo._ptDFFileInf SET LastScanTime=GetDate() WHERE FileID=@FileID", 1)
    '            .AddMultiLineText(strSQL, "END")
    '            .AddMultiLineText(strSQL, "SELECT @Rows Rows")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FileID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FileID") = FileID
    '            .AddPara("@FileSize", Data.SqlDbType.BigInt)
    '            .ParaValue("@FileSize") = FileSize
    '            .AddPara("@FastPigMD5", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FastPigMD5") = FastPigMD5
    '            .AddPara("@FilePath", Data.SqlDbType.VarChar, 4096)
    '            .ParaValue("@FilePath") = FilePath
    '            .AddPara("@FileUpdateTime", Data.SqlDbType.DateTime)
    '            .ParaValue("@FileUpdateTime") = FileUpdateTime
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function


    'Public Function RefDBFolders(Optional IsDirtyRead As Boolean = True) As String
    '    Dim LOG As New PigStepLog("RefDBFolders")
    '    Dim strSQL As String = "SELECT * FROM dbo._ptDFFolderInf"
    '    If IsDirtyRead = True Then strSQL &= " WITH (NOLOCK)"
    '    strSQL &= " ORDER BY FolderName"
    '    Try
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        LOG.StepName = "DBFolders.Clear"
    '        Me.DBFolders.Clear()
    '        Do While Not rsMain.EOF
    '            Dim oDBFolder As DBFolder
    '            Dim strFolderID As String = rsMain.Fields.Item("FolderID").StrValue
    '            LOG.StepName = "DBFolders.AddOrGet"
    '            oDBFolder = Me.DBFolders.AddOrGet(strFolderID, Me)
    '            If Me.DBFolders.LastErr <> "" Then
    '                LOG.AddStepNameInf(strFolderID)
    '                Me.mPrintErrLogInf(LOG.StepLogInf)
    '            End If
    '            LOG.StepName = "fFillByRs"
    '            LOG.Ret = oDBFolder.fFillByRs(rsMain)
    '            If LOG.Ret <> "OK" Then
    '                LOG.AddStepNameInf(strFolderID)
    '                Me.mPrintErrLogInf(LOG.StepLogInf)
    '            End If
    '            rsMain.MoveNext()
    '        Loop
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function


    'Friend Function fGetDBPathID(RelativePath As String) As String
    '    Try
    '        Dim strPigMD5 As String = ""
    '        Dim strRet As String = Me.mPigFunc.GetTextPigMD5(RelativePath, PigMD5.enmTextType.UTF8, strPigMD5)
    '        If strRet <> "OK" Then Throw New Exception(strRet)
    '        Return strPigMD5
    '    Catch ex As Exception
    '        Me.SetSubErrInf("fGetDBPathID", ex)
    '        Return ""
    '    End Try
    'End Function

    ''Friend Function fMergeDFPathInf(RelativePath As String) As String
    ''    Dim LOG As New PigStepLog("fAddOrGetPathID")
    ''    Dim strSQL As String = ""
    ''    Try
    ''        Dim strPathID As String = Me.fGetDBPathID(RelativePath)
    ''        With Me.mPigFunc
    ''            .AddMultiLineText(strSQL, "IF NOT EXISTS(SELECT 1 FROM dbo._ptDFPathInf WHERE PathID=@PathID)")
    ''            .AddMultiLineText(strSQL, "INSERT INTO dbo._ptDFPathInf(PathID,RelativePath,PathDesc)VALUES(@PathID,@RelativePath,'')", 1)
    ''        End With
    ''        LOG.StepName = "New CmdSQLSrvText"
    ''        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    ''        With oCmdSQLSrvText
    ''            .AddPara("@PathID", Data.SqlDbType.VarChar, 32)
    ''            .ParaValue("@PathID") = strPathID
    ''            .AddPara("@RelativePath", Data.SqlDbType.VarChar, 8000)
    ''            .ParaValue("@RelativePath") = RelativePath
    ''        End With
    ''        Dim rsMain As Recordset = Nothing
    ''        LOG.StepName = "mExecCmdSQLSrvSp"
    ''        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain, False,, True)
    ''        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    ''        rsMain.Close()
    ''        Return "OK"
    ''    Catch ex As Exception
    ''        LOG.AddStepNameInf(strSQL)
    ''        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    ''    End Try
    ''End Function

    '''' <summary>
    '''' 新建数据库文件夹|Add new database folder
    '''' </summary>
    '''' <returns></returns>
    'Private Function mAddNewDBFolder(InObj As DBFolder) As String
    '    Dim LOG As New PigStepLog("mAddNewDBFolder")
    '    Dim strSQL As String = ""
    '    Try
    '        LOG.StepName = "Check entries"
    '        If InObj.FolderName = "" Then Throw New Exception("Unable to get folder name")
    '        If InObj.FolderPath = "" Then Throw New Exception("No folder absolute path specified")
    '        If InObj.FolderID = "" Then Throw New Exception("Unable to get folder ID")
    '        With Me.mPigFunc
    '            .AddMultiLineText(strSQL, "IF EXISTS(SELECT 1 FROM dbo._ptDFFolderInf WHERE FolderID=@FolderID)")
    '            .AddMultiLineText(strSQL, "SELECT 'Folder ID already exists'", 1)
    '            .AddMultiLineText(strSQL, "ELSE IF EXISTS(SELECT 1 FROM dbo._ptDFFolderInf WHERE HostID=@HostID AND FolderPath=@FolderPath)")
    '            .AddMultiLineText(strSQL, "SELECT 'Host ID and Folder Path already exists'", 1)
    '            .AddMultiLineText(strSQL, "ELSE")
    '            .AddMultiLineText(strSQL, "BEGIN")
    '            .AddMultiLineText(strSQL, "INSERT INTO dbo._ptDFFolderInf(FolderID,HostID,FolderName,FolderPath,FolderType,FolderDesc,IsUse,ScanStatus)", 1)
    '            .AddMultiLineText(strSQL, "VALUES(@FolderID,@HostID,@FolderName,@FolderPath,@FolderType,@FolderDesc,@IsUse,@ScanStatus)", 1)
    '            .AddMultiLineText(strSQL, "SELECT 'OK'", 1)
    '            .AddMultiLineText(strSQL, "END")
    '        End With
    '        LOG.StepName = "New CmdSQLSrvText"
    '        Dim oCmdSQLSrvText As New CmdSQLSrvText(strSQL)
    '        With oCmdSQLSrvText
    '            .AddPara("@FolderID", Data.SqlDbType.VarChar, 32)
    '            .ParaValue("@FolderID") = InObj.FolderID
    '            .AddPara("@HostID", Data.SqlDbType.VarChar, 64)
    '            .ParaValue("@HostID") = InObj.HostID
    '            .AddPara("@FolderName", Data.SqlDbType.VarChar, 256)
    '            .ParaValue("@FolderName") = InObj.FolderName
    '            .AddPara("@FolderPath", Data.SqlDbType.VarChar, 2048)
    '            .ParaValue("@FolderPath") = InObj.FolderPath
    '            .AddPara("@FolderType", Data.SqlDbType.Int)
    '            .ParaValue("@FolderType") = InObj.FolderType
    '            .AddPara("@FolderDesc", Data.SqlDbType.VarChar, 512)
    '            .ParaValue("@FolderDesc") = InObj.FolderDesc
    '            .AddPara("@IsUse", Data.SqlDbType.Bit)
    '            .ParaValue("@IsUse") = InObj.IsUse
    '            .AddPara("@ScanStatus", Data.SqlDbType.Int)
    '            .ParaValue("@ScanStatus") = InObj.ScanStatus
    '        End With
    '        Dim rsMain As Recordset = Nothing
    '        LOG.StepName = "mExecCmdSQLSrvSp"
    '        LOG.Ret = Me.mExecCmdSQLSrvText(oCmdSQLSrvText, rsMain)
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        LOG.StepName = "DBFolders.AddOrGet"
    '        If Me.DBFolders.LastErr <> "" Then Me.DBFolders.ClearErr()
    '        Me.DBFolders.AddOrGet(InObj.FolderID, Me)
    '        If Me.DBFolders.LastErr <> "" Then Throw New Exception(Me.DBFolders.LastErr)
    '        LOG.StepName = "Close"
    '        If rsMain IsNot Nothing Then rsMain.Close()
    '        rsMain = Nothing
    '        oCmdSQLSrvText = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strSQL)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mExecCmdSQLSrvText(InCmdSQLSrvText As CmdSQLSrvText, ByRef RsOut As Recordset, Optional IsTxRes As Boolean = True, Optional IsRefLastExec As Boolean = True, Optional IsAllowEOF As Boolean = False, Optional IsNextRecordset As Boolean = False) As String
    '    Dim LOG As New PigStepLog("mExecCmdSQLSrvText")
    '    Try
    '        If Me.mConnSQLSrv.IsDBConnReady = False Then
    '            LOG.StepName = "OpenOrKeepActive"
    '            Me.mConnSQLSrv.OpenOrKeepActive()
    '            If Me.mConnSQLSrv.LastErr <> "" Then Throw New Exception(Me.mConnSQLSrv.LastErr)
    '        End If
    '        LOG.StepName = "ActiveConnection"
    '        InCmdSQLSrvText.ActiveConnection = Me.mConnSQLSrv.Connection
    '        LOG.StepName = "Execute"
    '        Me.PrintDebugLog(LOG.SubName, LOG.StepName, InCmdSQLSrvText.DebugStr)
    '        Dim rsMain As Recordset = InCmdSQLSrvText.Execute()
    '        If InCmdSQLSrvText.LastErr <> "" Then
    '            LOG.AddStepNameInf(InCmdSQLSrvText.SQLText)
    '            LOG.AddStepNameInf(InCmdSQLSrvText.DebugStr)
    '            Throw New Exception(InCmdSQLSrvText.LastErr)
    '        End If
    '        If rsMain.LastErr <> "" Then
    '            LOG.AddStepNameInf(InCmdSQLSrvText.SQLText)
    '            Throw New Exception(rsMain.LastErr)
    '        End If
    '        If rsMain.EOF = True And IsAllowEOF = False Then
    '            Throw New Exception("No results returned")
    '        ElseIf IsTxRes = True Then
    '            If rsMain.Fields.Item(0).StrValue <> "OK" Then
    '                Throw New Exception(rsMain.Fields.Item(0).StrValue)
    '            ElseIf IsNextRecordset = True Then
    '                LOG.StepName = "NextRecordset"
    '                RsOut = rsMain.NextRecordset()
    '                If rsMain.LastErr <> "" Then Throw New Exception(rsMain.LastErr)
    '                If RsOut Is Nothing Then Throw New Exception("RsOut Is Nothing")
    '            End If
    '        Else
    '            RsOut = rsMain
    '        End If
    '        rsMain = Nothing
    '        If IsRefLastExec = True Then
    '            Me.mLastExecCmdSQLSrvTextTime = Now
    '        End If
    '        Return "OK"
    '    Catch ex As Exception
    '        RsOut = Nothing
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    'Private Function mPrintErrLogInf(ErrInf As String) As String
    '    Try
    '        Me.mPigFunc.ASyncOptLogInf(ErrInf, Me.mLogFilePath)
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf("mPrintErrLogInf", ex)
    '    End Try
    'End Function

    'Private Sub mRefLogFilePath()
    '    Try
    '        Me.mLogFilePath = Me.mLogDirPath & Me.OsPathSep & Me.MyClassName & "_" & Format(Now, "yyyyMMdd") & ".log"
    '        Me.SetDebug(Me.mLogFilePath)
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        Me.SetSubErrInf("RefLogFilePath", ex)
    '    End Try
    'End Sub

    '''' <summary>
    '''' 自动识别文件夹或打包文件类型|Automatically recognize folder or packaged file types
    '''' </summary>
    '''' <param name="FolderOrFilePath">文件夹或文件路径|Folder or file path</param>
    '''' <returns></returns>
    'Public Function AutoGetFolderType(FolderOrFilePath As String) As DBFolder.EnmFolderType
    '    Dim LOG As New PigStepLog("GetFolderType")
    '    Try
    '        If Me.mPigFunc.IsFolderExists(FolderOrFilePath) = True Then
    '            If Me.IsWindows Then
    '                Return DBFolder.EnmFolderType.WinFolder
    '            Else
    '                Return DBFolder.EnmFolderType.LinuxFolder
    '            End If
    '        ElseIf Me.mPigFunc.IsFolderExists(FolderOrFilePath) = True Then
    '            Dim strExtName As String = Me.mPigFunc.GetPathPart(FolderOrFilePath, PigFunc.EnmFathPart.ExtName)
    '            Select Case LCase(strExtName)
    '                Case "jar"
    '                    Return DBFolder.EnmFolderType.JarFile
    '                Case "zip"
    '                    Return DBFolder.EnmFolderType.ZipFile
    '                Case "war"
    '                    Return DBFolder.EnmFolderType.WarFile
    '                Case Else
    '                    Return DBFolder.EnmFolderType.Unknow
    '            End Select
    '        Else
    '            Return DBFolder.EnmFolderType.Unknow
    '        End If
    '    Catch ex As Exception
    '        Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        Return DBFolder.EnmFolderType.Unknow
    '    End Try
    'End Function

    'Public Function fRefDBDir() As String
    '    Dim LOG As New PigStepLog("fRefDBDir")
    '    Try
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function


End Class
