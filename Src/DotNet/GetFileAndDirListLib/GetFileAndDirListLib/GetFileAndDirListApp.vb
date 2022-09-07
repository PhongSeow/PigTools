'**********************************
'* Name: GetFileAndDirListApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Get file and directory list application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 21/6/2021
'* 1.0.2  23/6/2021   Add LogFilePath,IsAbsolutePath,RootDirPath
'* 1.0.3  23/6/2021   Modify Start
'* 1.0.4  9/7/2021   Modify Start
'* 1.0.5  25/7/2021  Add OpenDebug, Modify LogFilePath
'* 1.0.6  27/7/2021  Remove OpenDebug
'* 1.1  28/8/2021  Change Imports PigToolsLib to PigToolsLiteLib, Modify mIsNoScanDir,mGetDirList
'* 1.2  3/9/2021  Modify Start
'************************************
Imports PigObjFsLib
Imports PigToolsLiteLib

Public Class GetFileAndDirListApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.8"
    Private moFS As FileSystemObject
    Private moPigFunc As PigFunc
    Property mstrLogFilePath As String
    Public Property LogFilePath() As String
        Get
            Return mstrLogFilePath
        End Get
        Set(value As String)
            mstrLogFilePath = value
            If Me.IsDebug = True Then
                Me.SetDebug(Me.LogFilePath)
            End If
        End Set
    End Property

    Property mbolIsAbsolutePath As Boolean
    Public Property IsAbsolutePath() As Boolean
        Get
            Return mbolIsAbsolutePath
        End Get
        Friend Set(value As Boolean)
            mbolIsAbsolutePath = value
        End Set
    End Property

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.mNew()
    End Sub

    Public Sub New(IsAbsolutePath As Boolean, Optional LogFilePath As String = "")
        MyBase.New(CLS_VERSION)
        Me.mNew(IsAbsolutePath,, LogFilePath)
    End Sub

    Public Sub New(RootDirPath As String, IsAbsolutePath As Boolean, Optional LogFilePath As String = "")
        MyBase.New(CLS_VERSION)
        Me.mNew(IsAbsolutePath, RootDirPath, LogFilePath)
    End Sub

    Public Sub New(RootDirPath As String, Optional LogFilePath As String = "")
        MyBase.New(CLS_VERSION)
        Me.mNew(, RootDirPath, LogFilePath)
    End Sub

    Private Sub mNew(Optional IsAbsolutePath As Boolean = False, Optional RootDirPath As String = "", Optional LogFilePath As String = "")
        With Me
            If LogFilePath <> "" Then
                .LogFilePath = LogFilePath
            Else
                .LogFilePath = Me.AppPath & Me.OsPathSep & Me.AppTitle & ".log"
            End If
            .IsAbsolutePath = IsAbsolutePath
            If RootDirPath = "" Then
                .RootDirPath = Me.AppPath
            Else
                .RootDirPath = RootDirPath
            End If
            .NoScanDirListPath = .RootDirPath & Me.OsPathSep & "NoScanDir.txt"
            .DirListPath = .RootDirPath & Me.OsPathSep & "DirList.txt"
            .FileListPath = .RootDirPath & Me.OsPathSep & "FileList.txt"
            .NoScanDirItems = New NoScanDirItems
        End With
    End Sub


    ''' <summary>
    ''' Directory list file path not scanned
    ''' </summary>
    Private mstrNoScanDirListPath As String
    Public Property NoScanDirListPath() As String
        Get
            Return mstrNoScanDirListPath
        End Get
        Friend Set(ByVal value As String)
            mstrNoScanDirListPath = value
        End Set
    End Property

    ''' <summary>
    ''' Directory list file path
    ''' </summary>
    Private mstrDirListPath As String
    Public Property DirListPath() As String
        Get
            Return mstrDirListPath
        End Get
        Friend Set(ByVal value As String)
            mstrDirListPath = value
        End Set
    End Property

    ''' <summary>
    ''' File list file path
    ''' </summary>
    Private mstrFileListPath As String
    Public Property FileListPath() As String
        Get
            Return mstrFileListPath
        End Get
        Friend Set(ByVal value As String)
            mstrFileListPath = value
        End Set
    End Property

    ''' <summary>
    ''' Root path
    ''' </summary>
    Private mstrRootDirPath As String
    Public Property RootDirPath() As String
        Get
            Return mstrRootDirPath
        End Get
        Friend Set(ByVal value As String)
            mstrRootDirPath = value
        End Set
    End Property

    Public Sub Start(Optional IsFullPigMD5 As Boolean = False)
        Const VSSVER_SCC_FILE As String = "vssver.scc"
        Dim LOG As New PigStepLog("Start")
        Try
            Dim strLine As String = ""
            moFS = New FileSystemObject
            moPigFunc = New PigFunc
            Dim tsDir As TextStream
            LOG.StepName = "OpenTextFile(NoScanDirListPath)"
            tsDir = moFS.OpenTextFile(Me.NoScanDirListPath, FileSystemObject.IOMode.ForReading, True)
            If moFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.NoScanDirListPath)
                Throw New Exception(moFS.LastErr)
            End If
            Me.NoScanDirItems = New NoScanDirItems
            If tsDir.AtEndOfStream = False Then
                LOG.StepName = "ReadAll for NoScanDirList"
                Dim strNoScanDirList As String = tsDir.ReadAll
                If tsDir.LastErr <> "" Then
                    LOG.AddStepNameInf(Me.NoScanDirListPath)
                    Throw New Exception(tsDir.LastErr)
                End If
                tsDir.Close()
                If Right(strNoScanDirList, Me.OsCrLf.Length) <> Me.OsCrLf Then strNoScanDirList = strNoScanDirList & Me.OsCrLf
                Do While True
                    strLine = moPigFunc.GetStr(strNoScanDirList, "", Me.OsCrLf)
                    If strLine = "" Then Exit Do
                    Me.NoScanDirItems.Add(strLine)
                Loop
            Else
                tsDir.Close()
            End If
            LOG.StepName = "OpenTextFile"
            tsDir = moFS.OpenTextFile(Me.DirListPath, FileSystemObject.IOMode.ForWriting, True)
            If moFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.DirListPath)
                Throw New Exception(moFS.LastErr)
            End If
            Dim tsFile As TextStream
            LOG.StepName = "OpenTextFile"
            tsFile = moFS.OpenTextFile(Me.FileListPath, FileSystemObject.IOMode.ForWriting, True)
            If moFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.DirListPath)
                Throw New Exception(moFS.LastErr)
            End If
            Dim oFolder As Folder
            LOG.StepName = "GetFolder"
            oFolder = moFS.GetFolder(Me.RootDirPath)
            If moFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.RootDirPath)
                Throw New Exception(moFS.LastErr)
            End If
            LOG.StepName = "mGetDirList"
            Dim strDirList As String = Me.mGetDirList(Me.RootDirPath)
            If Me.LastErr <> "" Then
                LOG.AddStepNameInf(Me.RootDirPath)
                Throw New Exception(Me.LastErr)
            End If
            Do While True
                strLine = moPigFunc.GetStr(strDirList, "", Me.OsCrLf)
                If strLine = "" Then Exit Do
                If Left(strLine, 1) = "#" Then
                    moPigFunc.OptLogInf(strLine & " is an error directory, and it will not be processed", Me.LogFilePath)
                ElseIf Me.mIsNoScanDir(strLine) = True Then
                    moPigFunc.OptLogInf(strLine & " does not need to scan directory", Me.LogFilePath)
                ElseIf moFS.FolderExists(strLine) = False Then
                    moPigFunc.OptLogInf(strLine & " no longer exists, no processing", Me.LogFilePath)
                Else
                    LOG.StepName = "Get current directory " & strLine
                    Dim strDirPath = strLine
                    Dim lngSubLen As Long
                    If Me.IsAbsolutePath = False Then
                        lngSubLen = Len(strDirPath) - Len(Me.RootDirPath) - 1
                        If lngSubLen <= 0 Then
                            strDirPath = "." & Me.OsPathSep
                        Else
                            strDirPath = "." & Me.OsPathSep & Right(strDirPath, lngSubLen)
                        End If
                    End If
                    LOG.StepName = "GetFolder"
                    oFolder = moFS.GetFolder(strLine)
                    If moFS.LastErr <> "" Then
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, strLine)
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, moFS.LastErr)
                    End If
                    LOG.StepName = "WriteLine"
                    tsDir.WriteLine(strDirPath & vbTab & Me.moPigFunc.GetFmtDateTime(oFolder.DateLastModified))
                    If tsDir.LastErr <> "" Then
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, strDirPath)
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, tsDir.LastErr)
                    End If
                    LOG.StepName = "Traverse the files in the current directory" & strLine
                    For Each oFile In oFolder.Files
                        If LCase(moPigFunc.GetFilePart(oFile.Name)) = VSSVER_SCC_FILE Then
                        ElseIf oFile.Path = Me.FileListPath Then
                        ElseIf oFile.Path = Me.DirListPath Then
                        Else
                            Dim strFilePath As String = oFile.Path
                            If Me.IsAbsolutePath = False Then
                                strFilePath = "." & Me.OsPathSep & Right(strFilePath, Len(strFilePath) - Len(Me.RootDirPath) - 1)
                            End If
                            LOG.StepName = "New PigFile"
                            Dim oPigFile As New PigFile(oFile.Path)
                            Dim strFastPigMD5 As String = "", strFullPigMD5 As String = ""
                            Dim pmFast As PigMD5 = Nothing, pmFull As PigMD5 = Nothing
                            LOG.StepName = "GetFastPigMD5"
                            LOG.Ret = oPigFile.GetFastPigMD5(pmFast)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf(oFile.Path)
                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.StepLogInf)
                            Else
                                strFastPigMD5 = pmFast.PigMD5
                            End If
                            pmFast = Nothing
                            If IsFullPigMD5 = True Then
                                LOG.StepName = "GetFastPigMD5"
                                LOG.Ret = oPigFile.GetFullPigMD5(pmFull)
                                If LOG.Ret <> "OK" Then
                                    Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.StepLogInf)
                                Else
                                    strFullPigMD5 = pmFull.PigMD5
                                End If
                                pmFull = Nothing
                            End If
                            LOG.StepName = "WriteLine"
                            strLine = strFilePath & vbTab & oFile.Size & vbTab & Me.moPigFunc.GetFmtDateTime(oFile.DateLastModified) & vbTab & strFastPigMD5
                            If IsFullPigMD5 = True Then strLine &= vbTab & strFullPigMD5
                            tsFile.WriteLine(strLine)
                            If tsFile.LastErr <> "" Then
                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, strLine)
                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, tsFile.LastErr)
                            End If
                        End If
                    Next
                End If
            Loop
            tsDir.Close()
            tsFile.Close()
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Me.PrintDebugLog(LOG.SubName, "Catch Exception", Me.LastErr)
        End Try
    End Sub

    Private Function mGetDirList(BaseDir As String) As String
        Try
            '出错的以#开头
            Dim oFolder As Folder, strRet As String = "", strOptRes As String = ""
            oFolder = moFS.GetFolder(BaseDir)
            Me.mGetDirList(oFolder, strRet, strOptRes)
            oFolder = Nothing
            Return strRet
        Catch ex As Exception
            Return Me.GetSubErrInf("mGetDirList", ex)
        End Try
    End Function

    ''' <summary>
    ''' 递归获取子目录列表|Get subdirectory list recursively
    ''' </summary>
    ''' <param name="oFolder"></param>
    ''' <param name="DirList"></param>
    ''' <param name="OptRes"></param>
    Private Sub mGetDirList(ByRef oFolder As Folder, ByRef DirList As String, ByRef OptRes As String)
        Try
            Dim oSubFolder As Folder
            If DirList = "" Then
                DirList = oFolder.Path & Me.OsCrLf
            Else
                DirList = DirList & oFolder.Path & Me.OsCrLf
            End If
            Dim lngCount As Long
#If NETCOREAPP Or NET40_OR_GREATER Then
            lngCount=oFolder.SubFolders.Count
#Else
            lngCount = oFolder.SubFolders.Length
#End If
            If lngCount = 0 Then Exit Sub
            For Each oSubFolder In oFolder.SubFolders
                OptRes = "OK"
                Me.mGetDirList(oSubFolder, DirList, OptRes）
                If OptRes <> "OK" Then
                    'Error in getting the subdirectory of AAA, the error is BBB
                    DirList = DirList & "#Error in getting the subdirectory of " & oSubFolder.Path & ", the error is " & OptRes & Me.OsCrLf
                End If
            Next
        Catch ex As Exception
            OptRes &= ex.Message.ToString
        End Try
    End Sub

    ''' <summary>
    ''' Do not scan directory list
    ''' </summary>
    Private moNoScanDirItems As NoScanDirItems
    Public Property NoScanDirItems() As NoScanDirItems
        Get
            Return moNoScanDirItems
        End Get
        Friend Set(ByVal value As NoScanDirItems)
            moNoScanDirItems = value
        End Set
    End Property

    Private Function mIsNoScanDir(ChkDirPath As String) As Boolean
        Try
            Dim bolIsNoScanDir As Boolean
            For Each oNoScanDirItem As NoScanDirItem In Me.NoScanDirItems
                bolIsNoScanDir = oNoScanDirItem.IsMeNoScan(ChkDirPath)
                If bolIsNoScanDir = True Then
                    Exit For
                End If
            Next
            Return bolIsNoScanDir
        Catch ex As Exception
            Me.SetSubErrInf("mIsNoScanDir", ex)
            Return False
        End Try
    End Function


End Class
