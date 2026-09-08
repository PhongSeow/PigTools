'**********************************
'* Name: DirScanApp
'* Author: Seow Phong
'* License: Copyright (c) 2025 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Directory scanning application|目录扫描应用
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 1/7/2025
'* 1.1  5/7/2025  Add ScanDir, OptLogInf, RefLogFilePath
'* 1.1  6/7/2025  Modify ScanDir, add LoadConf
'**********************************
Imports System.IO
Imports PigToolsLiteLib
Public Class DirScanApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "2" & "." & "18"

    Public ReadOnly Property ConfFilePath As String
    Private ReadOnly Property BaseDir As String
    Private ReadOnly Property PigFunc As New PigFunc
    Private ReadOnly Property UseTime As New UseTime
    Private Shadows ReadOnly Property IsWindows As Boolean
    Private Property FS As New PigFileSystem
    Private Property LogFilePath As String
    Private Property PXConf As PigXml
    Private Property IsLoadConf As Boolean = False
    Private Property RootDir As PigFolder

    Public Enum EnmSaveFileFmt
        Bcp = 0
        Xml = 1
    End Enum

    Public Enum EnmSaveFileType
        SingleFile = 0
        Distribution = 1
    End Enum

    Private Enum EnmScanStatus
        UnChanged = 0
        Deleted = 1
        Changed = 2
        IsNew = 3
    End Enum

    Private Structure StruDir
        Public DirPath As String
        Public CreateTime As String
        Public UpdateTime As String
        Public ScanType As EnmScanStatus


        Public Function BcpLine(Key As String) As String
            Return Key & vbTab & Me.DirPath & vbTab & Me.CreateTime & vbTab & Me.UpdateTime
        End Function

        Public Function XmlLine(Key As String) As String
            Return "<Key>" & Key & "</Key><DirPath>" & Me.DirPath & "</DirPath><CreateTime>" & Me.CreateTime & "</CreateTime><UpdateTime>" & Me.UpdateTime & "</UpdateTime>"
        End Function


        Public Sub ChkOrUpdate(ByRef InDirectoryInfo As DirectoryInfo)
            With Me
                ScanType = EnmScanStatus.UnChanged
                If .CreateTime <> InDirectoryInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss") Then
                    .CreateTime = InDirectoryInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                    ScanType = EnmScanStatus.Changed
                End If
                If .UpdateTime <> InDirectoryInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") Then
                    .UpdateTime = InDirectoryInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    ScanType = EnmScanStatus.Changed
                End If
            End With
        End Sub

        Public Sub AddNew(ByRef InDirectoryInfo As DirectoryInfo)
            With Me
                .DirPath = InDirectoryInfo.Name
                .CreateTime = InDirectoryInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                .UpdateTime = InDirectoryInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                .ScanType = EnmScanStatus.IsNew
            End With
        End Sub
    End Structure

    Private Structure StruFile
        Public FilePath As String
        Public FileSize As Long
        Public CreateTime As String
        Public UpdateTime As String
        Public ScanType As EnmScanStatus

        Public Function BcpLine(Key As String) As String
            Return Key & vbTab & Me.FilePath & vbTab & Me.FileSize.ToString & vbTab & Me.CreateTime & vbTab & Me.UpdateTime
        End Function
        Public Function XmlLine(Key As String) As String
            Return "<Key>" & Key & "</Key><FilePath>" & Me.FilePath & "</FilePath><FileSize>" & Me.FileSize.ToString & "</FileSize><CreateTime>" & Me.CreateTime & "</CreateTime><UpdateTime>" & Me.UpdateTime & "</UpdateTime>"
        End Function

        Public Sub ChkOrUpdate(ByRef InFileInfo As FileInfo)
            With Me
                ScanType = EnmScanStatus.UnChanged
                If .FileSize <> InFileInfo.Length Then
                    .FileSize = InFileInfo.Length
                    ScanType = EnmScanStatus.Changed
                End If
                If .CreateTime <> InFileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss") Then
                    .CreateTime = InFileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                    ScanType = EnmScanStatus.Changed
                End If
                If .UpdateTime <> InFileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") Then
                    .UpdateTime = InFileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    ScanType = EnmScanStatus.Changed
                End If
            End With
        End Sub
        Public Sub AddNew(ByRef InFileInfo As FileInfo)
            With Me
                .FilePath = InFileInfo.Name
                .FileSize = InFileInfo.Length
                .CreateTime = InFileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                .UpdateTime = InFileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                .ScanType = EnmScanStatus.IsNew
            End With
        End Sub
    End Structure


    Private mWorkDirPath As String = ""
    Public Property WorkDirPath As String
        Get
            Return mWorkDirPath
        End Get
        Friend Set(value As String)
            mWorkDirPath = value
        End Set
    End Property

    Private mScanDirPath As String = ""
    Public Property ScanDirPath As String
        Get
            Return mScanDirPath
        End Get
        Friend Set(value As String)
            mScanDirPath = value
        End Set
    End Property

    Private mOnlyScanExtNameRegexMatch As String = ""
    Public Property OnlyScanExtNameRegexMatch As String
        Get
            Return mOnlyScanExtNameRegexMatch
        End Get
        Friend Set(value As String)
            mOnlyScanExtNameRegexMatch = value
        End Set
    End Property


    Private mNoScanExtNameRegexMatch As String = ""
    Public Property NoScanExtNameRegexMatch As String
        Get
            Return mNoScanExtNameRegexMatch
        End Get
        Friend Set(value As String)
            mNoScanExtNameRegexMatch = value
        End Set
    End Property

    Private mNoScanExtNameList As String = ""
    Public Property NoScanExtNameList As String
        Get
            Return mNoScanExtNameList
        End Get
        Friend Set(value As String)
            mNoScanExtNameList = value
        End Set
    End Property

    Private mOnlyScanExtNameList As String = ""
    Public Property OnlyScanExtNameList As String
        Get
            Return mOnlyScanExtNameList
        End Get
        Friend Set(value As String)
            mOnlyScanExtNameList = value
        End Set
    End Property


    Private mOnlyScanFileLike As String = ""
    Public Property OnlyScanFileLike As String
        Get
            Return mOnlyScanFileLike
        End Get
        Friend Set(value As String)
            mOnlyScanFileLike = value
        End Set
    End Property

    Private mNoScanPathLike As String = ""
    Public Property NoScanPathLike As String
        Get
            Return mNoScanPathLike
        End Get
        Friend Set(value As String)
            mNoScanPathLike = value
        End Set
    End Property

    Private mOnlyScanPathLike As String = ""
    Public Property OnlyScanPathLike As String
        Get
            Return mOnlyScanPathLike
        End Get
        Friend Set(value As String)
            mOnlyScanPathLike = value
        End Set
    End Property

    Private mOnlyScanPathRegexMatch As String = ""
    Public Property OnlyScanPathRegexMatch As String
        Get
            Return mOnlyScanPathRegexMatch
        End Get
        Friend Set(value As String)
            mOnlyScanPathRegexMatch = value
        End Set
    End Property

    Private mNoScanFileLike As String = ""
    Public Property NoScanFileLike As String
        Get
            Return mNoScanFileLike
        End Get
        Friend Set(value As String)
            mNoScanFileLike = value
        End Set
    End Property

    Private mNoScanPathRegexMatch As String = ""
    Public Property NoScanPathRegexMatch As String
        Get
            Return mNoScanPathRegexMatch
        End Get
        Friend Set(value As String)
            mNoScanPathRegexMatch = value
        End Set
    End Property

    Private mNoScanFileRegexMatch As String = ""
    Public Property NoScanFileRegexMatch As String
        Get
            Return mNoScanFileRegexMatch
        End Get
        Friend Set(value As String)
            mNoScanFileRegexMatch = value
        End Set
    End Property

    Private mOnlyScanFileRegexMatch As String = ""
    Public Property OnlyScanFileRegexMatch As String
        Get
            Return mOnlyScanFileRegexMatch
        End Get
        Friend Set(value As String)
            mOnlyScanFileRegexMatch = value
        End Set
    End Property

    Private mSaveFileFmt As EnmSaveFileFmt = EnmSaveFileFmt.Bcp
    Public Property SaveFileFmt As EnmSaveFileFmt
        Get
            Return mSaveFileFmt
        End Get
        Friend Set(value As EnmSaveFileFmt)
            mSaveFileFmt = value
        End Set
    End Property

    Private mSaveFileType As EnmSaveFileType = EnmSaveFileType.SingleFile
    Public Property SaveFileType As EnmSaveFileType
        Get
            Return mSaveFileType
        End Get
        Friend Set(value As EnmSaveFileType)
            mSaveFileType = value
        End Set
    End Property

    Friend ReadOnly Property NowDirPath As String
        Get
            Return Me.WorkDirPath & Me.OsPathSep & "Now"
        End Get
    End Property

    Friend ReadOnly Property HisDirPath As String
        Get
            Return Me.WorkDirPath & Me.OsPathSep & "His"
        End Get
    End Property

    Friend ReadOnly Property LogDirPath As String
        Get
            Return Me.BaseDir & Me.OsPathSep & "Log"
        End Get
    End Property

    Friend ReadOnly Property TempDirPath As String
        Get
            Return Me.BaseDir & Me.OsPathSep & "Temp"
        End Get
    End Property

    Public Sub New(ConfFilePath As String)
        MyBase.New(CLS_VERSION)
        Me.IsWindows = MyBase.IsWindows
        Me.BaseDir = Me.PigFunc.GetFilePart(Me.AppPath, PigFunc.EnmFilePart.Path)
        Me.ConfFilePath = ConfFilePath
        Me.RefLogFilePath()
    End Sub


    Friend ReadOnly Property MainDirListPath As String
        Get
            Return Me.NowDirPath & Me.OsPathSep & "MainDirList.dat"
        End Get
    End Property

    Friend ReadOnly Property MainFileListPath As String
        Get
            Return Me.NowDirPath & Me.OsPathSep & "MainFileList.dat"
        End Get
    End Property

    Friend ReadOnly Property MainScanPath As String
        Get
            Return Me.NowDirPath & Me.OsPathSep & "Main.sca"
        End Get
    End Property

    Public Function SaveConf() As String
        Dim LOG As New PigStepLog("LoadConf")
        Try
            Dim strBakFilePath As String = Me.ConfFilePath & "." & Now.ToString("yyyyMMddHHmmss")
            LOG.StepName = "CopyFile"
            LOG.Ret = Me.FS.CopyFile(Me.ConfFilePath, strBakFilePath, True)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.ConfFilePath)
                LOG.AddStepNameInf(strBakFilePath)
                Me.OptLogInf(LOG.StepLogInf)
            End If
            Me.PXConf = New PigXml(True)
            With Me.PXConf
                .AddEleLeftSign("Root")
                .AddEle("ScanDirPath", Me.ScanDirPath, 1, True)
                .AddEle("WorkDirPath", Me.WorkDirPath, 1, True)
                .AddEle("SaveFileFmt", CStr(Me.SaveFileFmt), 1)
                .AddEle("SaveFileType", CStr(Me.SaveFileType), 1)
                .AddEle("NoScanExtNameList", Me.NoScanExtNameList, 1, True)
                .AddEle("NoScanExtNameRegexMatch", Me.NoScanExtNameRegexMatch, 1, True)
                .AddEle("NoScanPathLike", Me.NoScanPathLike, 1, True)
                .AddEle("NoScanPathRegexMatch", Me.NoScanPathRegexMatch, 1, True)
                .AddEle("NoScanFileLike", Me.NoScanFileLike, 1, True)
                .AddEle("NoScanFileRegexMatch", Me.NoScanFileRegexMatch, 1, True)
                .AddEle("OnlyScanExtNameList", Me.OnlyScanExtNameList, 1, True)
                .AddEle("OnlyScanExtNameRegexMatch", Me.OnlyScanExtNameRegexMatch, 1, True)
                .AddEle("OnlyScanPathLike", Me.OnlyScanPathLike, 1, True)
                .AddEle("OnlyScanPathRegexMatch", Me.OnlyScanPathRegexMatch, 1, True)
                .AddEle("OnlyScanFileLike", Me.OnlyScanFileLike, 1, True)
                .AddEle("OnlyScanFileRegexMatch", Me.OnlyScanFileRegexMatch, 1, True)
                .AddEleRightSign("Root")
            End With
            LOG.StepName = "SaveTextToFile"
            LOG.Ret = Me.PigFunc.SaveTextToFile(Me.ConfFilePath, Me.PXConf.MainXmlStr)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(Me.ConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            Me.IsLoadConf = True
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function LoadConf() As String
        Dim LOG As New PigStepLog("LoadConf")
        Try
            If Me.PigFunc.IsFileExists(Me.ConfFilePath) = False Then
                LOG.StepName = "SaveConf"
                LOG.Ret = Me.SaveConf()
                Me.OptLogInf(LOG.StepLogInf)
                If LOG.Ret <> "Ok" Then
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "InitXmlDocument"
            Me.PXConf = New PigXml(False)
            LOG.Ret = Me.PXConf.InitXmlDocument(Me.ConfFilePath)
            If LOG.Ret <> "Ok" Then
                LOG.AddStepNameInf(Me.ConfFilePath)
                Throw New Exception(LOG.Ret)
            End If
            With Me
                .ScanDirPath = Me.PXConf.XmlDocGetStr("Root.ScanDirPath")
                .WorkDirPath = Me.PXConf.XmlDocGetStr("Root.WorkDirPath")
                .SaveFileFmt = Me.PXConf.XmlDocGetInt("Root.SaveFileFmt")
                .SaveFileType = Me.PXConf.XmlDocGetInt("Root.SaveFileType")
                .NoScanExtNameList = Me.PXConf.XmlDocGetStr("Root.NoScanExtNameList")
                .NoScanExtNameRegexMatch = Me.PXConf.XmlDocGetStr("Root.NoScanExtNameRegexMatch")
                .NoScanPathLike = Me.PXConf.XmlDocGetStr("Root.NoScanPathLike")
                .NoScanPathRegexMatch = Me.PXConf.XmlDocGetStr("Root.NoScanPathRegexMatch")
                .NoScanFileLike = Me.PXConf.XmlDocGetStr("Root.NoScanFileLike")
                .NoScanFileRegexMatch = Me.PXConf.XmlDocGetStr("Root.NoScanFileRegexMatch")
                .OnlyScanExtNameList = Me.PXConf.XmlDocGetStr("Root.OnlyScanExtNameList")
                .OnlyScanExtNameRegexMatch = Me.PXConf.XmlDocGetStr("Root.OnlyScanExtNameRegexMatch")
                .OnlyScanPathLike = Me.PXConf.XmlDocGetStr("Root.OnlyScanPathLike")
                .OnlyScanPathRegexMatch = Me.PXConf.XmlDocGetStr("Root.OnlyScanPathRegexMatch")
                .OnlyScanFileLike = Me.PXConf.XmlDocGetStr("Root.OnlyScanFileLike")
                .OnlyScanFileRegexMatch = Me.PXConf.XmlDocGetStr("Root.OnlyScanFileRegexMatch")
                .IsLoadConf = True
            End With
            Return "OK"
        Catch ex As Exception
            Me.IsLoadConf = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function RefLogFilePath() As String
        Try
            If Me.PigFunc.IsFolderExists(Me.LogDirPath) = False Then
                Me.PigFunc.CreateFolder(Me.LogDirPath)
            End If
            Me.LogFilePath = Me.LogDirPath & Me.OsPathSep & Me.AppTitle & Me.PigFunc.GetFmtDateTime(Now, "yyyyMMDD") & ".log"
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("RefLogFilePath", ex)
        End Try
    End Function

    Friend Function OptLogInf(LogInf As String) As String
        Try
            Me.PigFunc.OptLogInf(LogInf, Me.LogFilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("OptLogInf", ex)
        End Try
    End Function

    Private Function IsDoFile(InPath As String) As Boolean
        Try
            IsDoFile = True
            If Me.NoScanExtNameList <> "" Then
                Dim strNoScanExtNameList As String = ";" & Me.NoScanExtNameList & ";"
                Dim strExtName As String = Me.PigFunc.GetFilePart(InPath, PigFunc.EnmFilePart.ExtName)
                If Me.IsWindows = True Then
                    If InStr(UCase(strNoScanExtNameList), UCase(strExtName)) > 0 Then
                        Return False
                    End If
                ElseIf InStr(strNoScanExtNameList, strExtName) > 0 Then
                    Return False
                End If
            End If

            If Me.NoScanFileLike <> "" And IsDoFile = True Then
                If Me.IsWindows = True Then
                    If InStr(UCase(InPath), UCase(Me.NoScanFileLike)) > 0 Then
                        Return False
                    End If
                ElseIf InStr(InPath, Me.NoScanFileLike) > 0 Then
                    Return False
                End If
            End If

            If Me.NoScanFileRegexMatch <> "" And IsDoFile = True Then
                If Text.RegularExpressions.Regex.IsMatch(InPath, Me.NoScanFileRegexMatch) = True Then
                    Return False
                End If
            End If

            If Me.OnlyScanExtNameList <> "" And IsDoFile = True Then
                Dim strOnlyScanExtNameList As String = ";" & Me.OnlyScanExtNameList & ";"
                Dim strExtName As String = Me.PigFunc.GetFilePart(InPath, PigFunc.EnmFilePart.ExtName)
                If Me.IsWindows = True Then
                    If InStr(UCase(strOnlyScanExtNameList), UCase(strExtName)) = 0 Then
                        Return False
                    End If
                ElseIf InStr(strOnlyScanExtNameList, strExtName) = 0 Then
                    Return False
                End If
            End If
            If Me.OnlyScanFileLike <> "" And IsDoFile = True Then
                If Me.IsWindows = True Then
                    If InStr(UCase(InPath), UCase(Me.OnlyScanFileLike)) = 0 Then
                        Return False
                    End If
                ElseIf InStr(InPath, Me.OnlyScanFileLike) = 0 Then
                    Return False
                End If
            End If
            If Me.OnlyScanFileRegexMatch <> "" And IsDoFile = True Then
                If Text.RegularExpressions.Regex.IsMatch(InPath, Me.OnlyScanFileRegexMatch) = False Then
                    Return False
                End If
            End If
        Catch ex As Exception
            Me.OptLogInf(Me.GetSubErrInf("IsDoFile", ex))
            Return False
        End Try

    End Function


    Private Function IsDoPath(InPath As String) As Boolean
        Try
            IsDoPath = True
            If Me.NoScanPathLike <> "" Then
                If Me.IsWindows = True Then
                    If InStr(UCase(InPath), UCase(Me.NoScanPathLike)) > 0 Then
                        IsDoPath = False
                    End If
                ElseIf InStr(InPath, Me.NoScanPathLike) > 0 Then
                    IsDoPath = False
                End If
            End If
            If Me.NoScanPathRegexMatch <> "" And IsDoPath = True Then
                If Text.RegularExpressions.Regex.IsMatch(InPath, Me.NoScanPathRegexMatch) = True Then
                    IsDoPath = False
                End If
            End If
            If Me.OnlyScanPathLike <> "" And IsDoPath = True Then
                If Me.IsWindows = True Then
                    If InStr(UCase(InPath), UCase(Me.OnlyScanPathLike)) = 0 Then
                        IsDoPath = False
                    End If
                ElseIf InStr(InPath, Me.OnlyScanPathLike) = 0 Then
                    IsDoPath = False
                End If
            End If
            If Me.OnlyScanPathRegexMatch <> "" And IsDoPath = True Then
                If Text.RegularExpressions.Regex.IsMatch(InPath, Me.OnlyScanPathRegexMatch) = False Then
                    IsDoPath = False
                End If
            End If
        Catch ex As Exception
            Me.OptLogInf(Me.GetSubErrInf("IsDoPath", ex))
            Return False
        End Try
    End Function

    Public Function Scan()
        Dim LOG As New PigStepLog("Scan")
        Try
            If Me.IsLoadConf = False Then
                LOG.StepName = "LoadConf"
                LOG.Ret = Me.LoadConf
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            Me.UseTime.GoBegin()
            LOG.StepName = "Check parameters"
            If Me.PigFunc.IsFolderExists(Me.ScanDirPath) = False Then
                LOG.AddStepNameInf(Me.ScanDirPath)
                LOG.Ret = "Scan Directory does not exist."
                Throw New Exception(LOG.Ret)
            End If
            If Me.PigFunc.IsFolderExists(Me.WorkDirPath) = False Then
                LOG.StepName = "Create Work Directory"
                LOG.Ret = Me.PigFunc.CreateFolder(Me.WorkDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.WorkDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            LOG.StepName = "Check SaveFileType"
            Select Case Me.SaveFileType
                Case EnmSaveFileType.SingleFile
                    LOG.StepName = "RefSubPigFolders"
                    Me.RootDir = New PigFolder(Me.ScanDirPath)
                    LOG.Ret = Me.RootDir.RefSubPigFolders()
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(Me.ScanDirPath)
                        Throw New Exception(LOG.Ret)
                    End If
                    Dim aStruDir(-1) As StruDir
                    For Each oPigFolder As PigFolder In Me.RootDir.SubPigFolders
                        If Me.IsDoPath(oPigFolder.FolderPath) = True Then
                            Dim oFolderInfo As New DirectoryInfo(oPigFolder.FolderPath)
                            For Each oFileInf As FileInfo In oFolderInfo.GetFiles
                                If Me.IsDoFile(oFileInf.FullName) = True Then

                                End If
                            Next
                        End If
                    Next

                Case EnmSaveFileType.Distribution
                    LOG.Ret = "Distribution is not currently supported."
                    Throw New Exception(LOG.Ret)
                Case Else
                    LOG.Ret = "SaveFileType is not supported."
                    Throw New Exception(LOG.Ret)
            End Select
            Me.UseTime.ToEnd()
            Return "OK"
        Catch ex As Exception
            Me.UseTime.ToEnd()
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function MirrorDirPath(CurrDirPath As String) As String
        Try
            Dim strCurrDirPath As String = CurrDirPath
            If Me.IsWindows = True Then
                strCurrDirPath = Replace(strCurrDirPath, ":\", "\")
            End If
            MirrorDirPath = Me.NowDirPath & Me.OsPathSep & strCurrDirPath
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function DeletedDirPath(CurrDirPath As String) As String
        Try
            Dim strCurrDirPath As String = CurrDirPath
            If Me.IsWindows = True Then
                strCurrDirPath = Replace(strCurrDirPath, ":\", "\")
            End If
            DeletedDirPath = Me.HisDirPath & Me.OsPathSep & strCurrDirPath
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function DirListPath(CurrDirPath As String) As String
        Return Me.MirrorDirPath(CurrDirPath) & Me.OsPathSep & "DirList.txt"
    End Function

    Private Function FileListPath(CurrDirPath As String) As String
        Return Me.MirrorDirPath(CurrDirPath) & Me.OsPathSep & "FileList.txt"
    End Function

    Private Function ScanDirDistribution(ByRef InDirectoryInfo As DirectoryInfo, ByRef Ret As String) As String
        Dim LOG As New PigStepLog("ScanDirDistribution")
        Dim strInDirPath As String = InDirectoryInfo.FullName
        Try
            'Dim strDirListPath As String = Me.DirListPath(strInDirPath)
            'If strDirListPath = "" Then
            '    LOG.Ret = "Can not get DirListPath."
            '    Throw New Exception(LOG.Ret)
            'End If
            'Dim strFileListPath As String = Me.FileListPath(strInDirPath)
            'If strFileListPath = "" Then
            '    LOG.AddStepNameInf("InDirPath=" & strInDirPath)
            '    LOG.Ret = "Can not get FileListPath."
            '    Throw New Exception(LOG.Ret)
            'End If
            'Dim tsAny As TextStream, bolIsSave As Boolean = False
            ''----------------------
            'LOG.StepName = "OpenTextFile.FileListPath.ForReading"
            'LOG.AddStepNameInf(strFileListPath)
            'tsAny = Me.FS.OpenTextFile(strFileListPath, PigFileSystem.IOMode.ForReading, True)
            'If tsAny Is Nothing Then
            '    LOG.Ret = "Is Nothing"
            '    Throw New Exception(LOG.Ret)
            'ElseIf Me.FS.LastErr <> "" Then
            '    LOG.Ret = Me.FS.LastErr
            '    Throw New Exception(LOG.Ret)
            'End If
            'Dim aStruFile(-1) As StruFile
            'LOG.StepName = "ReadLine"
            'Do While Not tsAny.AtEndOfStream
            '    Dim strLine As String = tsAny.ReadLine & vbTab
            '    Dim lngCnt As Long = UBound(aStruFile) + 1
            '    ReDim Preserve aStruFile(lngCnt)
            '    With aStruFile(lngCnt)
            '        .FileTitle = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .FileSize = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .CreateTime = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .UpdateTime = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .ScanType = EnmScanStatus.Deleted
            '    End With
            'Loop
            'tsAny.Close()
            ''----------------------
            'LOG.StepName = "GetFiles"
            'For Each oFileInf As FileInfo In InDirectoryInfo.GetFiles
            '    Dim bolIsFind As Boolean = False
            '    For Each oStruFile As StruFile In aStruFile
            '        If oStruFile.FileTitle = oFileInf.Name Then
            '            oStruFile.ChkOrUpdate(oFileInf)
            '            bolIsFind = True
            '            Exit For
            '        End If
            '    Next
            '    If bolIsFind = False Then
            '        Dim lngCnt As Long = UBound(aStruFile) + 1
            '        ReDim Preserve aStruFile(lngCnt)
            '        aStruFile(lngCnt).AddNew(oFileInf)
            '    End If
            'Next
            'bolIsSave = False
            'For Each oStruFile As StruFile In aStruFile
            '    Select Case oStruFile.ScanType
            '        Case EnmScanStatus.Changed, EnmScanStatus.Deleted, EnmScanStatus.IsNew
            '            bolIsSave = True
            '            Exit For
            '    End Select
            'Next
            'If bolIsSave = True Then
            '    LOG.StepName = "OpenTextFile.FileListPath.ForWriting"
            '    LOG.AddStepNameInf(strFileListPath)
            '    tsAny = Me.FS.OpenTextFile(strFileListPath, PigFileSystem.IOMode.ForWriting, True)
            '    If tsAny Is Nothing Then
            '        LOG.Ret = "Is Nothing"
            '        Throw New Exception(LOG.Ret)
            '    ElseIf Me.FS.LastErr <> "" Then
            '        LOG.Ret = Me.FS.LastErr
            '        Throw New Exception(LOG.Ret)
            '    End If
            '    For Each oStruFile As StruFile In aStruFile
            '        If oStruFile.ScanType = EnmScanStatus.Deleted Then
            '            IO.Directory.Move(Me.MirrorDirPath(strInDirPath), Me.HisDirPath & Me.OsPathSep & oStruFile.FileTitle)) 
            '        Else
            '            tsAny.WriteLine(oStruFile.Line)
            '        End If
            '    Next
            '    tsAny.Close()
            'End If
            ''----------------------
            'LOG.StepName = "OpenTextFile.DirListPath"
            'LOG.AddStepNameInf(strDirListPath)
            'tsAny = Me.FS.OpenTextFile(strDirListPath, PigFileSystem.IOMode.ForReading, True)
            'If tsAny Is Nothing Then
            '    LOG.Ret = "Is Nothing"
            '    Throw New Exception(LOG.Ret)
            'ElseIf Me.FS.LastErr <> "" Then
            '    LOG.Ret = Me.FS.LastErr
            '    Throw New Exception(LOG.Ret)
            'End If
            'Dim aStruDir(-1) As StruDir
            'LOG.StepName = "ReadLine"
            'Do While Not tsAny.AtEndOfStream
            '    Dim strLine As String = tsAny.ReadLine & vbTab
            '    Dim lngCnt As Long = UBound(aStruDir) + 1
            '    ReDim Preserve aStruDir(lngCnt)
            '    With aStruDir(lngCnt)
            '        .DirName = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .CreateTime = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .UpdateTime = Me.PigFunc.GetStr(strLine, "", vbTab)
            '        .ScanType = EnmScanStatus.Deleted
            '    End With
            'Loop
            'tsAny.Close()
            ''----------------------
            'LOG.StepName = "GetDirectories"
            'For Each oInDirectoryInfo In InDirectoryInfo.GetDirectories
            '    Dim strRet As String = ""
            '    strRet = Me.ScanDir(oInDirectoryInfo, Ret)
            '    If strRet <> "OK" Then
            '        Ret &= strInDirPath & ":Err=" & strRet & Me.OsCrLf
            '    End If
            'Next



            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf("InDirPath=" & strInDirPath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
