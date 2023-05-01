'**********************************
'* Name: GetFileAndDirListApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Get file and directory list application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 21/6/2021
'* 1.0.2  23/6/2021   Add LogFilePath,IsAbsolutePath,RootDirPath
'* 1.0.3  23/6/2021   Modify Start
'* 1.0.4  9/7/2021   Modify Start
'* 1.0.5  25/7/2021  Add OpenDebug, Modify LogFilePath
'* 1.0.6  27/7/2021  Remove OpenDebug
'* 1.1  28/8/2021  Change Imports PigToolsLib to PigToolsLiteLib, Modify mIsNoScanDir,mGetDirList
'* 1.2  3/9/2021  Modify Start
'* 1.3  27/2/2023  Hide internal properties
'* 1.5  28/2/2023  Modify Start, add WorkPath,StatusFilePath,RunStatus
'* 1.6  1/3/2023  Modify mNew,SaveStatus,RefStatus,add 
'* 1.7  2/3/2023  Modify mNew,Start,RefNoScanDir, add mStart
'* 1.8  29/4/2023  Add ScanSubFolders,mGetSubFolder,ScanSubFolders
'************************************
Imports System.Threading
Imports PigObjFsLib
Imports PigToolsLiteLib

''' <summary>
''' Get the current directory or subdirectory file or directory processing class|获取当前目录或子目录文件或目录处理类
''' </summary>
Public Class GetFileAndDirListApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.8.16"
    Private Property mFS As New FileSystemObject
    Private Property mPigFunc As New PigFunc

    Public Event ScanOK()
    Public Event ScanFail(ErrInf As String)
    Public Event FindFolderOK(oFolder As Folder, FindFolders As Long)
    Public Event FindFolderErr(oFolder As Folder, ErrInf As String)
    Public Event FindFolderNoScan(oFolder As Folder)
    Public Event FindFolderEnd(FindFolders As Long)

    Private mStatusFilePath As String
    Public Property StatusFilePath() As String
        Get
            Return mStatusFilePath
        End Get
        Friend Set(value As String)
            mStatusFilePath = value
        End Set
    End Property


    Private mRunStatus As EnmRunStatus
    Public Property RunStatus() As EnmRunStatus
        Get
            Return mRunStatus
        End Get
        Friend Set(value As EnmRunStatus)
            mRunStatus = value
        End Set
    End Property

    Private mCurrProcThreadID As String
    Public Property CurrProcThreadID() As String
        Get
            Return mCurrProcThreadID
        End Get
        Friend Set(value As String)
            mCurrProcThreadID = value
        End Set
    End Property

    Private mTimeoutTime As Date
    Public Property TimeoutTime() As Date
        Get
            Return mTimeoutTime
        End Get
        Friend Set(value As Date)
            mTimeoutTime = value
        End Set
    End Property

    Private mEndTime As Date
    Public Property EndTime() As Date
        Get
            Return mEndTime
        End Get
        Friend Set(value As Date)
            mEndTime = value
        End Set
    End Property

    Private mStartTime As Date
    Public Property StartTime() As Date
        Get
            Return mStartTime
        End Get
        Friend Set(value As Date)
            mStartTime = value
        End Set
    End Property

    Private mScanFiles As Long
    Public Property ScanFiles() As Long
        Get
            Return mScanFiles
        End Get
        Friend Set(value As Long)
            mScanFiles = value
        End Set
    End Property

    Private mScanFolders As Long
    Public Property ScanFolders() As Long
        Get
            Return mScanFolders
        End Get
        Friend Set(value As Long)
            mScanFolders = value
        End Set
    End Property

    Private mUseSeconds As Decimal
    Public Property UseSeconds() As Decimal
        Get
            Return mUseSeconds
        End Get
        Friend Set(value As Decimal)
            mUseSeconds = value
        End Set
    End Property

    Private mActiveTime As Date
    Public Property ActiveTime() As Date
        Get
            Return mActiveTime
        End Get
        Friend Set(value As Date)
            mActiveTime = value
        End Set
    End Property

    Private Sub mResetStatus()
        With Me
            .ActiveTime = Date.MinValue
            .UseSeconds = -1
            .TimeoutTime = Date.MinValue
            .EndTime = Date.MinValue
            .RunStatus = EnmRunStatus.Free
            .ScanFiles = 0
            .ScanFolders = 0
        End With
    End Sub

    Private mLogFilePath As String
    Public Property LogFilePath() As String
        Get
            Return mLogFilePath
        End Get
        Set(value As String)
            mLogFilePath = value
            If Me.IsDebug = True Then
                Me.SetDebug(Me.LogFilePath)
            End If
        End Set
    End Property

    Private mIsAbsolutePath As Boolean
    Public Property IsAbsolutePath() As Boolean
        Get
            Return mIsAbsolutePath
        End Get
        Friend Set(value As Boolean)
            mIsAbsolutePath = value
        End Set
    End Property

    Private mIsFullPigMD5 As Boolean
    Public Property IsFullPigMD5() As Boolean
        Get
            Return mIsFullPigMD5
        End Get
        Friend Set(value As Boolean)
            mIsFullPigMD5 = value
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
        Dim LOG As New PigStepLog("mNew")
        Try
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
                .mDataPath = .RootDirPath & Me.OsPathSep & "." & Replace(Me.AppTitle, "Lib", "")
                If Me.mPigFunc.IsFolderExists(.mDataPath) = False Then
                    LOG.StepName = "CreateFolder"
                    LOG.Ret = Me.mPigFunc.CreateFolder(Me.mDataPath)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(Me.mDataPath)
                        Throw New Exception(LOG.Ret)
                    End If
                End If
                .NoScanDirListPath = .mDataPath & Me.OsPathSep & "NoScanDir.txt"
                .DirListPath = .mDataPath & Me.OsPathSep & "DirList.txt"
                .FileListPath = .mDataPath & Me.OsPathSep & "FileList.txt"
                .StatusFilePath = .mDataPath & Me.OsPathSep & "Status.xml"
            End With
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Private Property mDataPath As String

    ''' <summary>
    ''' Directory list file path not scanned
    ''' </summary>
    Private mNoScanDirListPath As String
    Public Property NoScanDirListPath() As String
        Get
            Return mNoScanDirListPath
        End Get
        Friend Set(ByVal value As String)
            mNoScanDirListPath = value
        End Set
    End Property

    ''' <summary>
    ''' Directory list file path
    ''' </summary>
    Private mDirListPath As String
    Public Property DirListPath() As String
        Get
            Return mDirListPath
        End Get
        Friend Set(ByVal value As String)
            mDirListPath = value
        End Set
    End Property

    ''' <summary>
    ''' File list file path
    ''' </summary>
    Private mFileListPath As String
    Public Property FileListPath() As String
        Get
            Return mFileListPath
        End Get
        Friend Set(ByVal value As String)
            mFileListPath = value
        End Set
    End Property

    ''' <summary>
    ''' Root path
    ''' </summary>
    Private mRootDirPath As String
    Public Property RootDirPath() As String
        Get
            Return mRootDirPath
        End Get
        Friend Set(ByVal value As String)
            mRootDirPath = value
        End Set
    End Property

    Public Function ScanSubFolders() As String
        Try
            Me.RefNoScanDir()
            Dim oThread As New Thread(AddressOf mScanSubFolders)
            oThread.Start()
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ScanSubFolders", ex)
        End Try
    End Function

    Private Sub mScanSubFolders()
        Dim LOG As New PigStepLog("mScanSubFolders")
        Try
            Dim lngFindCnt As Long = 0
            Dim oRootFolder As Folder = Me.mFS.GetFolder(Me.mRootDirPath)
            Me.mGetSubFolder(oRootFolder, lngFindCnt)
            RaiseEvent FindFolderEnd(lngFindCnt)
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub



    Private Sub mGetSubFolder(ByRef ParentFolder As Folder, ByRef FindCnt As Long)
        Dim strParentFolderPath As String = ParentFolder.Path
        Try
            If Me.mIsNoScanDir(strParentFolderPath) = True Then
                RaiseEvent FindFolderNoScan(ParentFolder)
                Exit Sub
            End If
            FindCnt += 1
            RaiseEvent FindFolderOK(ParentFolder, FindCnt)
            Dim lngSubCount As Long
#If NETCOREAPP Or NET40_OR_GREATER Then
            lngSubCount = ParentFolder.SubFolders.Count
#Else
            lngSubCount = ParentFolder.SubFolders.Length
#End If
            If lngSubCount = 0 Then Exit Sub
            For Each oSubFolder In ParentFolder.SubFolders
                Me.mGetSubFolder(oSubFolder, FindCnt）
            Next
        Catch ex As Exception
            RaiseEvent FindFolderErr(ParentFolder, ex.Message.ToString)
        End Try
    End Sub


    Public Sub Start(Optional IsFullPigMD5 As Boolean = False, Optional TimeoutMinutes As Integer = 10)
        Try
            Me.RefStatus()
            Select Case Me.RunStatus
                Case EnmRunStatus.Running
                    If Math.Abs(DateDiff(DateInterval.Minute, Me.ActiveTime, Now)) < 5 Then
                        Throw New Exception("Another thread is running.")
                    End If
            End Select
            Dim struMain As mStruStart
            With struMain
                .IsFullPigMD5 = IsFullPigMD5
                .TimeoutMinutes = TimeoutMinutes
            End With
            Dim oThread As New Thread(AddressOf mStart)
            oThread.Start(struMain)
            oThread = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Start", ex)
        End Try
    End Sub

    Public Function RefNoScanDir() As String
        Dim LOG As New PigStepLog("RefNoScanDir")
        Try
            LOG.StepName = "New NoScanDirItems"
            Me.NoScanDirItems = New NoScanDirItems
            Me.NoScanDirItems.Add(Me.mDataPath)
            Dim strGitPath As String = Me.RootDirPath & Me.OsPathSep & ".git"
            Me.NoScanDirItems.Add(strGitPath)
            Dim strNoScanDirList As String = ""
            If Me.mPigFunc.IsFileExists(Me.NoScanDirListPath) = True Then
                LOG.StepName = "GetFileText"
                LOG.Ret = Me.mPigFunc.GetFileText(Me.NoScanDirListPath, strNoScanDirList)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.NoScanDirListPath)
                    Me.PrintDebugLog(LOG.SubName, LOG.StepName, LOG.StepLogInf)
                End If
                strNoScanDirList &= Me.OsCrLf
                Do While True
                    Dim strDir As String = Me.mPigFunc.GetStr(strNoScanDirList, "", Me.OsCrLf)
                    If strDir = "" Then Exit Do
                    Me.NoScanDirItems.AddOrGet(strDir)
                Loop
            Else
                Me.mPigFunc.SaveTextToFile(Me.NoScanDirListPath, "")
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Structure mStruStart
        Public IsFullPigMD5 As Boolean
        Public TimeoutMinutes As Integer
    End Structure
    Private Sub mStart(StruMain As mStruStart)
        Dim LOG As New PigStepLog("Start")
        Const VSSVER_SCC_FILE As String = "vssver.scc"
        Dim bolIsSaveStatus As Boolean = False
        Dim oUseTime As New UseTime
        Try
            Me.mResetStatus()
            Me.StartTime = Now
            Me.IsFullPigMD5 = StruMain.IsFullPigMD5
            Me.TimeoutTime = Now.AddMinutes(StruMain.TimeoutMinutes)
            oUseTime.GoBegin()
            Me.RefNoScanDir()
            LOG.StepName = "SaveStatus"
            LOG.Ret = Me.SaveStatus(EnmCtrlStatus.Start, False)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim strLine As String = ""
            Dim tsDir As TextStream
            LOG.StepName = "OpenTextFile"
            tsDir = mFS.OpenTextFile(Me.DirListPath, FileSystemObject.IOMode.ForWriting, True)
            If mFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.DirListPath)
                Throw New Exception(mFS.LastErr)
            End If
            Dim tsFile As TextStream
            LOG.StepName = "OpenTextFile"
            tsFile = mFS.OpenTextFile(Me.FileListPath, FileSystemObject.IOMode.ForWriting, True)
            If mFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.DirListPath)
                Throw New Exception(mFS.LastErr)
            End If
            Dim oFolder As Folder
            LOG.StepName = "GetFolder"
            oFolder = mFS.GetFolder(Me.RootDirPath)
            If mFS.LastErr <> "" Then
                LOG.AddStepNameInf(Me.RootDirPath)
                Throw New Exception(mFS.LastErr)
            End If
            LOG.StepName = "mGetDirList"
            Dim strDirList As String = Me.mGetDirList(Me.RootDirPath)
            If Me.LastErr <> "" Then
                LOG.AddStepNameInf(Me.RootDirPath)
                Throw New Exception(Me.LastErr)
            End If
            Dim dteLastSaveTime As Date = Date.MinValue
            Do While True
                strLine = mPigFunc.GetStr(strDirList, "", Me.OsCrLf)
                If strLine = "" Then Exit Do
                If Left(strLine, 1) = "#" Then
                    mPigFunc.OptLogInf(strLine & " is an error directory, and it will not be processed", Me.LogFilePath)
                ElseIf Me.mIsNoScanDir(strLine) = True Then
                    mPigFunc.OptLogInf(strLine & " does not need to scan directory", Me.LogFilePath)
                ElseIf mFS.FolderExists(strLine) = False Then
                    mPigFunc.OptLogInf(strLine & " no longer exists, no processing", Me.LogFilePath)
                Else
                    Me.ScanFolders += 1
                    If Math.Abs(DateDiff(DateInterval.Second, dteLastSaveTime, Now)) > 30 Then
                        Me.UseSeconds = oUseTime.Step2NowSeconds
                        Me.SaveStatus(EnmCtrlStatus.Inoperation, False)
                        dteLastSaveTime = Now
                    End If
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
                    oFolder = mFS.GetFolder(strLine)
                    If mFS.LastErr <> "" Then
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, strLine)
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, mFS.LastErr)
                    End If
                    LOG.StepName = "WriteLine"
                    tsDir.WriteLine(strDirPath & vbTab & Me.mPigFunc.GetFmtDateTime(oFolder.DateLastModified))
                    If tsDir.LastErr <> "" Then
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, strDirPath)
                        Me.PrintDebugLog(LOG.SubName, LOG.StepName, tsDir.LastErr)
                    End If
                    LOG.StepName = "Traverse the files in the current directory" & strLine
                    For Each oFile In oFolder.Files
                        If LCase(mPigFunc.GetFilePart(oFile.Name)) = VSSVER_SCC_FILE Then
                        ElseIf oFile.Path = Me.FileListPath Then
                        ElseIf oFile.Path = Me.DirListPath Then
                        Else
                            Me.ScanFiles += 1
                            If Math.Abs(DateDiff(DateInterval.Second, dteLastSaveTime, Now)) > 30 Then
                                Me.UseSeconds = oUseTime.Step2NowSeconds
                                Me.SaveStatus(EnmCtrlStatus.Inoperation, False)
                                dteLastSaveTime = Now
                            End If
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
                            If StruMain.IsFullPigMD5 = True Then
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
                            strLine = strFilePath & vbTab & oFile.Size & vbTab & Me.mPigFunc.GetFmtDateTime(oFile.DateLastModified) & vbTab & strFastPigMD5
                            If StruMain.IsFullPigMD5 = True Then strLine &= vbTab & strFullPigMD5
                            tsFile.WriteLine(strLine)
                            If tsFile.LastErr <> "" Then
                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, strLine)
                                Me.PrintDebugLog(LOG.SubName, LOG.StepName, tsFile.LastErr)
                            End If
                        End If
                    Next
                End If
                If Now > Me.TimeoutTime Then
                    Me.RunStatus = EnmRunStatus.RunTimeout
                    Throw New Exception("Scan timeout.")
                End If
            Loop
            tsDir.Close()
            tsFile.Close()
            oUseTime.ToEnd()
            Me.UseSeconds = oUseTime.AllDiffSeconds
            Me.EndTime = Now
            Me.SaveStatus(EnmCtrlStatus.ScanOK, False)
            RaiseEvent ScanOK()
            Me.ClearErr()
        Catch ex As Exception
            Dim strRet As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            oUseTime.ToEnd()
            Me.UseSeconds = oUseTime.AllDiffSeconds
            If Me.RunStatus = EnmRunStatus.RunTimeout Then
                Me.SaveStatus(EnmCtrlStatus.ScanTimeout, False)
            Else
                Me.SaveStatus(EnmCtrlStatus.ScanFail, False)
            End If
            RaiseEvent ScanFail(strRet)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Private Function mGetDirList(BaseDir As String) As String
        Try
            '出错的以#开头
            Dim oFolder As Folder, strRet As String = "", strOptRes As String = ""
            oFolder = mFS.GetFolder(BaseDir)
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

    Public Enum EnmRunStatus
        Free = 0
        Running = 1
        RunOK = 2
        RunFail = -1
        RunTimeout = -2
    End Enum

    Public Enum EnmCtrlStatus
        Start = 0
        ScanOK = 1
        Inoperation = 2
        ScanFail = 3
        ScanTimeout = 4
    End Enum



    Public Function RefStatus() As String
        Dim LOG As New PigStepLog("RefStatus")
        Try
            Dim strXml As String = ""
            LOG.StepName = "GetFileText"
            LOG.Ret = Me.mPigFunc.GetFileText(Me.StatusFilePath, strXml)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim oPigXml As New PigXml(True, False)
            With oPigXml
                LOG.StepName = "SetMainXml"
                .Clear()
                .SetMainXml(strXml)
                If .LastErr <> "" Then
                    LOG.AddStepNameInf(strXml)
                    Throw New Exception(.LastErr)
                End If
                LOG.StepName = "InitXmlDocument"
                LOG.Ret = .InitXmlDocument()
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strXml)
                    Throw New Exception(.LastErr)
                End If
                With Me
                    .RunStatus = oPigXml.XmlDocGetStr("Root.RunStatus")
                    .CurrProcThreadID = oPigXml.XmlDocGetStr("Root.ProcThreadID")
                    .IsAbsolutePath = oPigXml.XmlDocGetBool("Root.IsAbsolutePath")
                    .IsFullPigMD5 = oPigXml.XmlDocGetBool("Root.IsFullPigMD5")
                    .StartTime = oPigXml.XmlDocGetDate("Root.StartTime")
                    .TimeoutTime = oPigXml.XmlDocGetDate("Root.TimeoutTime")
                    .EndTime = oPigXml.XmlDocGetDate("Root.EndTime")
                    .UseSeconds = oPigXml.XmlDocGetDec("Root.UseSeconds")
                    .ScanFiles = oPigXml.XmlDocGetLong("Root.ScanFiles")
                    .ScanFolders = oPigXml.XmlDocGetLong("Root.ScanFolders")
                    .ActiveTime = oPigXml.XmlDocGetDate("Root.ActiveTime")
                End With
            End With
            oPigXml = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function SaveStatus(CtrlStatus As EnmCtrlStatus, Optional IsRefStatus As Boolean = True) As String
        Dim LOG As New PigStepLog("SaveStatus")
        Try
            If IsRefStatus = True Then Me.RefStatus()
            LOG.StepName = "Check CtrlStatus"
            Dim intNewRunStatus As EnmRunStatus = EnmRunStatus.Free
            Select Case CtrlStatus
                Case EnmCtrlStatus.Start
                    Select Case Me.RunStatus
                        Case EnmRunStatus.Free, EnmRunStatus.RunFail, EnmRunStatus.RunOK
                            intNewRunStatus = EnmRunStatus.Running
                        Case Else
                            Throw New Exception("The current state is " & Me.RunStatus.ToString & " and cannot be " & CtrlStatus.ToString)
                    End Select
                Case EnmCtrlStatus.ScanOK
                    intNewRunStatus = EnmRunStatus.RunOK
                Case EnmCtrlStatus.ScanFail
                    intNewRunStatus = EnmRunStatus.RunFail
                Case EnmCtrlStatus.ScanTimeout
                    intNewRunStatus = EnmRunStatus.RunTimeout
                Case EnmCtrlStatus.Inoperation
                    intNewRunStatus = Me.RunStatus
            End Select
            LOG.StepName = "Set oPigXml"
            Dim oPigXml As New PigXml(True)
            With oPigXml
                .AddEleLeftSign("Root")
                .AddEle("RunStatus", intNewRunStatus)
                .AddEle("ProcThreadID", Me.mPigFunc.GetProcThreadID)
                '                .AddEle("LogFilePath", Me.LogFilePath)
                .AddEle("IsAbsolutePath", Me.IsAbsolutePath)
                .AddEle("IsFullPigMD5", Me.IsFullPigMD5)
                .AddEle("StartTime", Me.StartTime)
                .AddEle("TimeoutTime", Me.TimeoutTime)
                .AddEle("EndTime", Me.EndTime)
                .AddEle("UseSeconds", Me.UseSeconds)
                .AddEle("ScanFiles", Me.ScanFiles)
                .AddEle("ScanFolders", Me.ScanFolders)
                .AddEle("ActiveTime", Now)
                .AddEleRightSign("Root")
                LOG.StepName = "SaveTextToFile"
                LOG.Ret = Me.mPigFunc.SaveTextToFile(Me.StatusFilePath, .MainXmlStr)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Me.StatusFilePath)
                    Throw New Exception(LOG.Ret)
                End If
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
