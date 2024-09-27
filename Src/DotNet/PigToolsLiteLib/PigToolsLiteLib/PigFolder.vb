'**********************************
'* Name: PigFolder
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 目录处理|Directory processing
'* Home Url: https://en.seowphong.com
'* Version: 1.7
'* Create Time: 4/6/2023
'* 1.1  11/5/2023   Add IsRootFolder
'* 1.2  11/6/2023   Add CreationTime,UpdateTime
'* 1.3  13/6/2023   Add RefSubPigFolders,RefPigFiles,FindSubFolders, modify mGetSubDirList
'* 1.5  22/11/2023  Add FilesSize,mGetFastPigMD5,GetFastPigMD5,mGetSubPigFolders
'* 1.6  27/7/2024   Modify PigStepLog to StruStepLog
'* 1.7  26/9/2024   Modify RefPigFiles, add mRefPigFiles,DelOldFiles
'**********************************
Imports System.IO

Public Class PigFolder
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "7" & "." & "20"

    Public ReadOnly Property FolderPath As String
    Private Property mFolderInfo As DirectoryInfo
    Public Enum EnmGetFastPigMD5Type
        FileSize_Files = 0
        FileSize_Files_UpdateTime = 1
        FileFastPigMD5 = 2
        CurrDirInfo = 3
    End Enum

    Public Sub New(FolderPath As String)
        MyBase.New(CLS_VERSION)
        Me.FolderPath = FolderPath
    End Sub


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
            Dim lngCount As Long = 0
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
                    If strRet <> "OK" Then
                        Ret &= oInDirectoryInfo.FullName & ":Err=" & strRet & strOsCrLf
                    End If
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

    ''' <summary>
    ''' The file space size of the current directory, in MB|当前目录的文件空间大小
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property FilesSize() As Long
        Get
            Try
                FilesSize = 0
                For Each oPigFile As PigFile In Me.PigFiles
                    FilesSize += oPigFile.Size
                Next
            Catch ex As Exception
                Me.SetSubErrInf("FilesSize", ex)
                Return -1
            End Try
        End Get
    End Property


    Private mPigFiles As PigFiles
    Public ReadOnly Property PigFiles As PigFiles
        Get
            Try
                If mPigFiles Is Nothing Then
                    Dim strRet As String = Me.RefPigFiles()
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return mPigFiles
            Catch ex As Exception
                Me.SetSubErrInf("PigFiles", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Private mSubPigFolders As PigFolders
    Public ReadOnly Property SubPigFolders As PigFolders
        Get
            If mSubPigFolders Is Nothing Then Me.mSubPigFolders = New PigFolders
            Return mSubPigFolders
        End Get
    End Property

    Public Structure StruDelOldFileRes
        Dim ScanFiles As Long
        Dim DelOKFiles As Long
        Dim DelFailFiles As Long
        Dim ErrInf As String
    End Structure

    Private Function mDeleteFile(FilePath As String) As String
        Try
            File.Delete(FilePath)
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' Delete outdated files|删除过时的文件
    ''' </summary>
    ''' <param name="IsDeep">Does not include subdirectories|是不包含子目录</param>
    ''' <param name="KeepDays">Keep days|保留天数</param>
    ''' <param name="Res">Return information|返回信息</param>
    ''' <returns></returns>
    Public Function DelOldFiles(IsDeep As Boolean, KeepDays As Integer, ByRef Res As StruDelOldFileRes) As String
        Try
            If KeepDays < 0 Then Throw New Exception("Invalid KeepDays")
            Return Me.mDelOldFiles(IsDeep, Res, 0, KeepDays, 0)
        Catch ex As Exception
            Return Me.GetSubErrInf("DelOldFiles", ex)
        End Try
    End Function

    ''' <summary>
    ''' Delete outdated files|删除过时的文件
    ''' </summary>
    ''' <param name="IsDeep">Does not include subdirectories|是不包含子目录</param>
    ''' <param name="KeepHours">Keep hours|保留小时数</param>
    ''' <param name="Res">Return information|返回信息</param>
    ''' <returns></returns>
    Public Function DelOldFiles(IsDeep As Boolean, KeepHours As Decimal, ByRef Res As StruDelOldFileRes) As String
        Try
            If KeepHours < 0 Then Throw New Exception("Invalid KeepHours")
            Return Me.mDelOldFiles(IsDeep, Res, 0, 0, KeepHours)
        Catch ex As Exception
            Return Me.GetSubErrInf("DelOldFiles", ex)
        End Try
    End Function


    ''' <summary>
    ''' Delete outdated files|删除过时的文件
    ''' </summary>
    ''' <param name="IsDeep">Does not include subdirectories|是不包含子目录</param>
    ''' <param name="KeepDays">Keep days|保留天数</param>
    ''' <param name="MaxScanFiles">Maximum number of scanned files|最大扫描文件数</param>
    ''' <param name="Res">Return information|返回信息</param>
    ''' <returns></returns>
    ''' <summary>
    Public Function DelOldFiles(IsDeep As Boolean, KeepDays As Integer, MaxScanFiles As Long, ByRef Res As StruDelOldFileRes) As String
        Try
            If KeepDays < 0 Then Throw New Exception("Invalid KeepDays")
            Return Me.mDelOldFiles(IsDeep, Res, MaxScanFiles, KeepDays, 0)
        Catch ex As Exception
            Return Me.GetSubErrInf("DelOldFiles", ex)
        End Try
    End Function

    ''' <summary>
    ''' Delete outdated files|删除过时的文件
    ''' </summary>
    ''' <param name="IsDeep">Does not include subdirectories|是不包含子目录</param>
    ''' <param name="KeepHours">Keep hours|保留小时数</param>
    ''' <param name="MaxScanFiles">Maximum number of scanned files|最大扫描文件数</param>
    ''' <param name="Res">Return information|返回信息</param>
    ''' <returns></returns>
    ''' <summary>
    Public Function DelOldFiles(IsDeep As Boolean, KeepHours As Decimal, MaxScanFiles As Long, ByRef Res As StruDelOldFileRes) As String
        Try
            If KeepHours < 0 Then Throw New Exception("Invalid KeepDays")
            If MaxScanFiles <= 0 Then Throw New Exception("Invalid MaxScanFiles")
            Return Me.mDelOldFiles(IsDeep, Res, MaxScanFiles, 0, KeepHours)
        Catch ex As Exception
            Return Me.GetSubErrInf("DelOldFiles", ex)
        End Try
    End Function


    Private Function mDelOldFiles(IsDeep As Boolean, ByRef Res As StruDelOldFileRes, Optional MaxScanFiles As Long = 0, Optional KeepDays As Integer = 0, Optional KeepHours As Decimal = 0) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mDelOldFiles"
        Try
            Dim oPigFolders As PigFolders = Nothing
            If IsDeep = False Then
                oPigFolders = New PigFolders
            Else
                LOG.StepName = "FindSubFolders"
                LOG.Ret = Me.FindSubFolders(True, oPigFolders)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            If oPigFolders.IsItemExists(Me.FolderPath) = False Then
                LOG.StepName = "oPigFolders.Add"
                LOG.AddStepNameInf(Me.FolderPath)
                oPigFolders.Add(Me.FolderPath)
            End If
            LOG.StepName = "For Each"
            With Res
                .DelFailFiles = 0
                .DelOKFiles = 0
                .ErrInf = ""
                .ScanFiles = 0
            End With
            Dim intMaxScanFiles As Long = MaxScanFiles, bolIsMaxScan As Boolean
            If intMaxScanFiles <= 0 Then
                bolIsMaxScan = False
            Else
                bolIsMaxScan = True
            End If
            For Each oPigFolder As PigFolder In oPigFolders
                LOG.StepName = "RefPigFiles"
                If bolIsMaxScan = False Then
                    LOG.Ret = oPigFolder.RefPigFiles()
                ElseIf intMaxScanFiles < 0 Then
                    Exit For
                Else
                    LOG.Ret = oPigFolder.RefPigFiles(intMaxScanFiles)
                End If
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(oPigFolder.FolderPath)
                    Res.ErrInf &= LOG.StepLogInf & Me.OsCrLf
                Else
                    Res.ScanFiles += oPigFolder.PigFiles.Count
                    intMaxScanFiles -= oPigFolder.PigFiles.Count
                    LOG.StepName = "For Each mDeleteFile"
                    For Each oPigFile As PigFile In oPigFolder.PigFiles
                        If KeepDays > 0 Then
                            If DateDiff(DateInterval.Day, oPigFile.UpdateTime, Now) > KeepDays Then
                                LOG.Ret = mDeleteFile(oPigFile.FilePath)
                                If LOG.Ret <> "OK" Then
                                    Res.ErrInf &= "Delete File " & oPigFile.FilePath & " Err:" & LOG.Ret
                                    Res.DelFailFiles += 1
                                Else
                                    Res.DelOKFiles += 1
                                End If
                            End If
                        ElseIf KeepHours > 0 Then
                            If CDec(DateDiff(DateInterval.Minute, oPigFile.UpdateTime, Now)) / 60 > KeepHours Then
                                LOG.Ret = mDeleteFile(oPigFile.FilePath)
                                If LOG.Ret <> "OK" Then
                                    Res.ErrInf &= "Delete File " & oPigFile.FilePath & " Err:" & LOG.Ret
                                    Res.DelFailFiles += 1
                                Else
                                    Res.DelOKFiles += 1
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' 扫描子目录|Scan for subdirectories
    ''' </summary>
    ''' <param name="IsDeep">是否深入扫描子目录|Do you want to drill down to scan subdirectories</param>
    ''' <param name="OutPigFolders">输出的目录集合|Output directory collection</param>
    ''' <returns></returns>
    Public Function FindSubFolders(IsDeep As Boolean, ByRef OutPigFolders As PigFolders) As String
        Dim LOG As New StruStepLog : LOG.SubName = "FindSubFolders"
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
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("RefMe", ex)
        End Try

    End Function

    Public Function RefSubPigFolders() As String
        Dim LOG As New StruStepLog : LOG.SubName = "RefSubPigFolders"
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

    ''' <summary>
    ''' Refresh the collection of files in the current directory|刷新当前目录下的文件集合
    ''' </summary>
    ''' <returns></returns>
    Public Function RefPigFiles() As String
        Return Me.mRefPigFiles(0)
    End Function

    ''' <summary>
    ''' Refresh the collection of files in the current directory|刷新当前目录下的文件集合
    ''' </summary>
    ''' <param name="MaxScanFiles">The maximum number of scanned files|扫描最多的文件数</param>
    ''' <returns></returns>
    Public Function RefPigFiles(MaxScanFiles As Long) As String
        Return Me.mRefPigFiles(MaxScanFiles)
    End Function

    Private Function mRefPigFiles(MaxScanFiles As Long) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mRefPigFiles"
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
            Dim lngScanFiles As Long = 0
            For Each oFileInf In Me.mFolderInfo.GetFiles
                If MaxScanFiles > 0 Then
                    lngScanFiles += 1
                    If lngScanFiles > MaxScanFiles Then Exit For
                End If
                Me.mPigFiles.Add(oFileInf.FullName)
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As EnmGetFastPigMD5Type, ScanSize As Integer) As String
        Return Me.mGetFastPigMD5(FastPigMD5, GetFastPigMD5Type, ScanSize)
    End Function

    Public Function GetFastPigMD5(ByRef FastPigMD5 As String, GetFastPigMD5Type As EnmGetFastPigMD5Type, ScanSize As Integer) As String
        Try
            Dim strRet As String = ""
            Dim oPigMD5 As PigMD5 = Nothing
            strRet = Me.mGetFastPigMD5(oPigMD5, GetFastPigMD5Type, ScanSize)
            If strRet <> "OK" Then Throw New Exception(strRet)
            FastPigMD5 = oPigMD5.PigMD5
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            FastPigMD5 = ""
            Return Me.GetSubErrInf("GetFastPigMD5", ex)
        End Try
    End Function

    Public Function GetFastPigMD5(ByRef FastPigMD5 As String, GetFastPigMD5Type As EnmGetFastPigMD5Type) As String
        Try
            Dim strRet As String = ""
            Dim oPigMD5 As PigMD5 = Nothing
            strRet = Me.mGetFastPigMD5(oPigMD5, GetFastPigMD5Type)
            If strRet <> "OK" Then Throw New Exception(strRet)
            FastPigMD5 = oPigMD5.PigMD5
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            FastPigMD5 = ""
            Return Me.GetSubErrInf("GetFastPigMD5", ex)
        End Try
    End Function

    Public Function GetFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As EnmGetFastPigMD5Type) As String
        Return Me.mGetFastPigMD5(FastPigMD5, GetFastPigMD5Type)
    End Function

    Private Function mGetFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As EnmGetFastPigMD5Type, Optional ScanSize As Integer = 20480) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mGetFastPigMD5"
        Try
            LOG.StepName = "RefPigFiles"
            LOG.Ret = Me.RefPigFiles
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "RefSubPigFolders"
            LOG.Ret = Me.RefSubPigFolders()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim pbMain As New PigBytes
            Select Case GetFastPigMD5Type
                Case EnmGetFastPigMD5Type.FileSize_Files
                    LOG.Ret = pbMain.SetValue(Me.PigFiles.Count)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf("PigFiles.Count")
                        Throw New Exception(LOG.Ret)
                    End If
                    LOG.StepName = "SetValue"
                    LOG.Ret = pbMain.SetValue(Me.FilesSize)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf("FilesSize")
                        Throw New Exception(LOG.Ret)
                    End If
                Case Else
                    Select Case GetFastPigMD5Type
                        Case EnmGetFastPigMD5Type.FileSize_Files_UpdateTime
                            LOG.Ret = pbMain.SetValue(Me.PigFiles.Count)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf("PigFiles.Count")
                                Throw New Exception(LOG.Ret)
                            End If
                            LOG.StepName = "SetValue"
                            LOG.Ret = pbMain.SetValue(Me.FilesSize)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf("FilesSize")
                                Throw New Exception(LOG.Ret)
                            End If
                        Case EnmGetFastPigMD5Type.CurrDirInfo
                            LOG.Ret = pbMain.SetValue(Me.UpdateTime)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf("UpdateTime")
                                Throw New Exception(LOG.Ret)
                            End If
                    End Select
                    Dim oList As New List(Of String)
                    For Each oPigFolder As PigFolder In Me.SubPigFolders
                        oList.Add(oPigFolder.FolderPath)
                    Next
                    oList.Sort()
                    For i = 0 To oList.Count - 1
                        With Me.SubPigFolders.Item(oList.Item(i))
                            Dim ptFolder As New PigText(.FolderName, PigText.enmTextType.UTF8)
                            LOG.Ret = pbMain.SetValue(ptFolder.TextBytes)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf(.FolderPath)
                                LOG.AddStepNameInf(.FolderName)
                                Throw New Exception(LOG.Ret)
                            End If
                            LOG.Ret = pbMain.SetValue(.UpdateTime)
                            If LOG.Ret <> "OK" Then
                                LOG.AddStepNameInf(.FolderPath)
                                LOG.AddStepNameInf(.UpdateTime)
                                Throw New Exception(LOG.Ret)
                            End If
                            ptFolder = Nothing
                        End With
                    Next
                    '---------
                    oList.Clear()
                    For Each oPigFile As PigFile In Me.PigFiles
                        oList.Add(oPigFile.FilePath)
                    Next
                    oList.Sort()
                    For i = 0 To oList.Count - 1
                        With Me.mPigFiles.Item(oList.Item(i))
                            Select Case GetFastPigMD5Type
                                Case EnmGetFastPigMD5Type.CurrDirInfo
                                    Dim ptFile As New PigText(.FileTitle, PigText.enmTextType.UTF8)
                                    LOG.Ret = pbMain.SetValue(ptFile.TextBytes)
                                    If LOG.Ret <> "OK" Then
                                        LOG.AddStepNameInf(.FilePath)
                                        LOG.AddStepNameInf(.FileTitle)
                                        Throw New Exception(LOG.Ret)
                                    End If
                                    LOG.Ret = pbMain.SetValue(.Size)
                                    If LOG.Ret <> "OK" Then
                                        LOG.AddStepNameInf(.FilePath)
                                        LOG.AddStepNameInf(.Size)
                                        Throw New Exception(LOG.Ret)
                                    End If
                                    LOG.Ret = pbMain.SetValue(.UpdateTime)
                                    If LOG.Ret <> "OK" Then
                                        LOG.AddStepNameInf(.FilePath)
                                        LOG.AddStepNameInf(.UpdateTime)
                                        Throw New Exception(LOG.Ret)
                                    End If
                                    ptFile = Nothing
                                Case EnmGetFastPigMD5Type.FileSize_Files_UpdateTime
                                    LOG.Ret = pbMain.SetValue(.UpdateTime)
                                    If LOG.Ret <> "OK" Then
                                        LOG.AddStepNameInf(.FilePath)
                                        LOG.AddStepNameInf(.UpdateTime)
                                        Throw New Exception(LOG.Ret)
                                    End If
                                Case EnmGetFastPigMD5Type.FileFastPigMD5
                                    Dim oPigMD5 As PigMD5 = Nothing
                                    LOG.Ret = .GetFastPigMD5(oPigMD5, ScanSize)
                                    If LOG.Ret <> "OK" Then
                                        LOG.AddStepNameInf(.FilePath)
                                        LOG.AddStepNameInf("GetFastPigMD5")
                                        Throw New Exception(LOG.Ret)
                                    End If
                                    If oPigMD5 IsNot Nothing Then
                                        LOG.Ret = pbMain.SetValue(oPigMD5.PigMD5Bytes)
                                        If LOG.Ret <> "OK" Then
                                            LOG.AddStepNameInf(.FilePath)
                                            Throw New Exception(LOG.Ret)
                                        End If
                                    End If
                            End Select
                        End With
                    Next
            End Select
            LOG.StepName = "New PigMD5"
            FastPigMD5 = New PigMD5(pbMain.Main)
            If FastPigMD5.LastErr <> "" Then
                LOG.Ret = FastPigMD5.LastErr
                Throw New Exception(LOG.Ret)
            End If
            pbMain = Nothing
            Return "OK"
        Catch ex As Exception
            FastPigMD5 = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
