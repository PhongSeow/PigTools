'**********************************
'* Name: PigFolder
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 目录处理|Directory processing
'* Home Url: https://en.seowphong.com
'* Version: 1.3
'* Create Time: 4/6/2023
'* 1.1  11/5/2023   Add IsRootFolder
'* 1.2  11/6/2023   Add CreationTime,UpdateTime
'* 1.3  13/6/2023   Add RefSubPigFolders,RefPigFiles,FindSubFolders, modify mGetSubDirList
'**********************************
Imports System.IO

Public Class PigFolder
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.16"

    Public ReadOnly Property FolderPath As String
    Private Property mFolderInfo As DirectoryInfo

    'Public Event FindFolder(FolderPath As String, TotalFindFolders As Long)
    'Public Event ScanFolderErr(ErrInf As String)

    'Public Event FindFile(FilePath As String, TotalFindFiles As Long)
    'Public Event ScanFileErr(ErrInf As String)

    Public Sub New(FolderPath As String)
        MyBase.New(CLS_VERSION)
        Me.FolderPath = FolderPath
    End Sub

    '''' <summary>
    '''' 扫描目录|Scan directory
    '''' </summary>
    '''' <param name="IsSubFolders">是否包括子目录|Is include subdirectories</param>
    '''' <returns></returns>
    'Public Function ScanFolder(IsSubFolders As Boolean) As String
    '    Dim LOG As New PigStepLog("ScanFolder")
    '    Dim strFolderPath As String = ""
    '    Dim lngTotalFindFolders As Long = 0
    '    Try
    '        LOG.StepName = "mInitFolderInf"
    '        LOG.Ret = Me.mInitFolderInf
    '        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '        If IsSubFolders = False Then
    '            LOG.StepName = "For Each"
    '            For Each oDirectoryInfo As DirectoryInfo In Me.mFolderInfo.GetDirectories
    '                strFolderPath = oDirectoryInfo.FullName
    '                lngTotalFindFolders += 1
    '                RaiseEvent FindFolder(oDirectoryInfo.FullName, lngTotalFindFolders)
    '            Next
    '        Else
    '            Dim astrDir(0) As String, strRet As String = ""
    '            LOG.StepName = "mGetSubDirList"
    '            LOG.Ret = Me.mGetSubDirList(Me.mFolderInfo, astrDir, strRet)
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            If strRet <> "" Then Throw New Exception(strRet)
    '            For i = 0 To astrDir.Length - 1
    '                strFolderPath = astrDir(i)
    '                lngTotalFindFolders += 1
    '                RaiseEvent FindFolder(strFolderPath, lngTotalFindFolders)
    '            Next
    '        End If
    '        Return "OK"
    '    Catch ex As Exception
    '        LOG.AddStepNameInf(strFolderPath)
    '        LOG.AddStepNameInf(lngTotalFindFolders)
    '        Dim strErr As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '        RaiseEvent ScanFolderErr(strErr)
    '        Return strErr
    '    End Try
    'End Function

    Private Function mInitFolderInf() As String
        Try
            If mFolderInfo Is Nothing Then
                mFolderInfo = New DirectoryInfo(Me.FolderPath)
            Else
                mFolderInfo.Refresh()
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mInitFolderInf", ex)
        End Try
    End Function

    Private Function mGetSubPigFolders(ByRef InDirectoryInfo As DirectoryInfo, ByRef InPigFolders As PigFolders, ByRef Ret As String) As String
        Try
            Dim lngCount As Long
#If NET40_OR_GREATER Or NETCOREAPP Then
            lngCount = InDirectoryInfo.GetDirectories.LongCount
#Else
            lngCount = InDirectoryInfo.GetDirectories.Length
#End If

            If lngCount > 0 Then
                Dim strOsCrLf As String = Me.OsCrLf
                For Each oInDirectoryInfo In InDirectoryInfo.GetDirectories
                    Dim strRet As String = ""
                    InPigFolders.AddOrGet(InDirectoryInfo.FullName)
                    strRet = Me.mGetSubPigFolders(oInDirectoryInfo, InPigFolders, Ret)
                    If strRet <> "OK" Then Ret &= oInDirectoryInfo.FullName & ":Err=" & strRet & strOsCrLf
                Next
            Else
                InPigFolders.AddOrGet(InDirectoryInfo.FullName)
            End If
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public ReadOnly Property IsRootFolder() As Boolean
        Get
            Try
                If mFolderInfo Is Nothing Then Me.mInitFolderInf()
                If Me.mFolderInfo.Parent Is Nothing Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IsRootFolder", ex)
                Return Nothing
            End Try
        End Get
    End Property

    'Public ReadOnly Property SubFolders(IsSubFolders As Boolean) As PigFolder()
    '    Get
    '        Try
    '            If IsSubFolders = True Then
    '                Dim strRet As String = "", saDir(-1) As String
    '                Me.mGetSubDirList(Me.mFolderInfo, saDir, strRet)
    '                Dim aFolder() As PigFolder
    '                ReDim aFolder(-1)
    '                For i = 0 To saDir.LongLength - 1
    '                    Dim oDir As New DirectoryInfo(saDir(i))
    '                    ReDim Preserve aFolder(aFolder.LongLength)
    '                    Dim oFolder As New PigFolder(oDir.FullName)
    '                    aFolder(aFolder.LongLength - 1) = oFolder
    '                Next
    '                Return aFolder
    '            Else
    '                Return Me.mSubFolders
    '            End If
    '        Catch ex As Exception
    '            Me.SetSubErrInf("SubFolders", ex)
    '            Return Nothing
    '        End Try
    '    End Get
    'End Property

    'Private ReadOnly Property mSubFolders() As PigFolder()
    '    Get
    '        Try
    '            Dim aFolder() As PigFolder
    '            ReDim aFolder(-1)
    '            For Each oDirectoryInfo In Me.mFolderInfo.GetDirectories
    '                ReDim Preserve aFolder(aFolder.LongLength)
    '                Dim oFolder As New PigFolder(oDirectoryInfo.FullName)
    '                aFolder(aFolder.LongLength - 1) = oFolder
    '            Next
    '            Return aFolder
    '        Catch ex As Exception
    '            Me.SetSubErrInf("mSubFolders", ex)
    '            Return Nothing
    '        End Try
    '    End Get
    'End Property

    Public ReadOnly Property CreationTime() As Date
        Get
            Try
                If mFolderInfo Is Nothing Then Me.mInitFolderInf()
                Return Me.mFolderInfo.CreationTime
            Catch ex As Exception
                Me.SetSubErrInf("CreationTime", ex)
                Return Date.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property UpdateTime() As Date
        Get
            Try
                If mFolderInfo Is Nothing Then Me.mInitFolderInf()
                Return Me.mFolderInfo.LastWriteTime
            Catch ex As Exception
                Me.SetSubErrInf("UpdateTime", ex)
                Return Date.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property FolderName() As String
        Get
            Try
                If mFolderInfo Is Nothing Then Me.mInitFolderInf()
                Return Me.mFolderInfo.Name
            Catch ex As Exception
                Me.SetSubErrInf("FolderName", ex)
                Return ""
            End Try
        End Get
    End Property

    'Public ReadOnly Property Size() As Long
    '    Get
    '        Try
    '            If mFolderInfo Is Nothing Then Me.mInitFolderInf()
    '            Size = 0
    '            For Each oFile In Me.mFolderInfo.GetFiles
    '                Size += oFile.Length
    '            Next
    '            Dim strRet As String = "", saDir(-1) As String
    '            Me.mGetSubDirList(Me.mFolderInfo, saDir, strRet)
    '            For i = 0 To saDir.LongLength - 1
    '                Dim oDir As New DirectoryInfo(saDir(i))
    '                For Each oFile In oDir.GetFiles
    '                    Size += oFile.Length
    '                Next
    '            Next
    '            Return Size
    '        Catch ex As Exception
    '            Me.SetSubErrInf("Size", ex)
    '            Return -1
    '        End Try
    '    End Get
    'End Property

    Private mPigFiles As PigFiles
    Public ReadOnly Property PigFiles As PigFiles
        Get
            If mPigFiles Is Nothing Then Me.mPigFiles = New PigFiles
            Return mPigFiles
        End Get
    End Property

    Private mSubPigFolders As PigFolders
    Public ReadOnly Property SubPigFolders As PigFolders
        Get
            If mSubPigFolders Is Nothing Then Me.mSubPigFolders = New PigFolders
            Return mSubPigFolders
        End Get
    End Property

    ''' <summary>
    ''' 扫描子目录|Scan for subdirectories
    ''' </summary>
    ''' <param name="IsDeep">是否深入扫描子目录|Do you want to drill down to scan subdirectories</param>
    ''' <param name="OutPigFolders">输出的目录集合|Output directory collection</param>
    ''' <returns></returns>
    Public Function FindSubFolders(IsDeep As Boolean, ByRef OutPigFolders As PigFolders) As String
        Dim LOG As New PigStepLog("FindSubFolders")
        Try
            If IsDeep = True Then
                LOG.Ret = Me.RefMe()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                OutPigFolders = New PigFolders
                Dim strRet As String = ""
                LOG.StepName = "mGetSubPigFolders"
                LOG.Ret = Me.mGetSubPigFolders(Me.mFolderInfo, OutPigFolders, strRet)
                If LOG.Ret <> "OK" Then
                    If strRet <> "" Then LOG.AddStepNameInf(strRet)
                    Throw New Exception(LOG.Ret)
                ElseIf strRet <> "" Then
                    Throw New Exception(strRet)
                End If
            Else
                LOG.StepName = "RefSubPigFolders"
                LOG.Ret = Me.RefSubPigFolders()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "Set OutPigFolders"
                OutPigFolders = Me.SubPigFolders
            End If
            Return "OK"
        Catch ex As Exception
            OutPigFolders = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function RefMe() As String
        Dim strRet As String = ""
        Try
            strRet = mInitFolderInf()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.mFolderInfo.Refresh()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("RefMe", ex)
        End Try

    End Function

    Public Function RefSubPigFolders() As String
        Dim LOG As New PigStepLog("RefSubPigFolders")
        Try
            LOG.Ret = Me.RefMe()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "SubPigFolders.Clear"
            If Me.mSubPigFolders Is Nothing Then
                Me.mSubPigFolders = New PigFolders
            Else
                Me.mSubPigFolders.Clear()
            End If
            If Me.mSubPigFolders.LastErr <> "" Then Throw New Exception(Me.mSubPigFolders.LastErr)
            LOG.StepName = "For Each"
            For Each oDirectoryInfo In Me.mFolderInfo.GetDirectories
                Me.mSubPigFolders.Add(oDirectoryInfo.FullName)
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function RefPigFiles() As String
        Dim LOG As New PigStepLog("RefPigFiles")
        Try
            LOG.Ret = Me.RefMe()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "PigFiles.Clear"
            If Me.mPigFiles Is Nothing Then
                Me.mPigFiles = New PigFiles
            Else
                Me.mPigFiles.Clear()
            End If
            If Me.mPigFiles.LastErr <> "" Then Throw New Exception(Me.mPigFiles.LastErr)
            LOG.StepName = "For Each"
            For Each oFileInf In Me.mFolderInfo.GetFiles
                Me.mPigFiles.Add(oFileInf.FullName)
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(Log.SubName, Log.StepName, ex)
        End Try
    End Function

End Class
