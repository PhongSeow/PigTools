﻿'**********************************
'* Name: PigFile
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: File processing,Handle file reading, writing, information, etc
'* Home Url: https://en.seowphong.com
'* Version: 1.12
'* Create Time: 4/11/2019
'* 1.0.2  2019-11-5   增加mSaveFile
'* 1.0.3  2019-11-20  增加 CopyFileTo
'* 1.0.5  2020-2-13   增加 IsPigMD5
'* 1.0.6  2020-2-15   部分功能迁到 fGEFileSystem
'* 1.0.7  2020-3-5    增加 SetData
'* 1.0.8  2020-10-27  修改 LoadFile
'* 1.0.9  1/2/2021 Err.Raise change to Throw New Exception|Err.Raise改为Throw New Exception
'* 1.0.10  26/7/2021 Modify LoadFile
'* 1.1  10/5/2022    Add PigMD5, modify LoadFile
'* 1.2  10/8/2022    Add GetFastPigMD5
'* 1.3  16/8/2022    Add SegLoadFile, modify GetFastPigMD5,mGetMyMD5
'* 1.5  12/9/2022    Modify New
'* 1.6  30/9/2022    Add GetTailText,GetTopText
'* 1.7  2/5/2023     Modify GetFastPigMD5
'* 1.8  4/5/2023     Modify GetFastPigMD5,mGetMyMD5,SegLoadFile,GetTailText, add mGetFastPigMD5
'* 1.9  10/11/2023   Add GetTextRows, modify GetTailText
'* 1.10 4/12/2023   Add GetFullText
'* 1.11  27/7/2024   Modify PigStepLog to StruStepLog
'* 1.12  27/8/2024   Modify FileShare.Read to FileShare.ReadWrite
'**********************************
Imports System.IO
'Imports Microsoft.VisualBasic.Logging
''' <summary>
''' File processing set|文件处理集
''' </summary>
Public Class PigFile
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "11" & "." & "8"
    Private mstrFilePath As String '文件路径
    Private moFileInfo As FileInfo '文件信息
    Public GbMain As PigBytes '主数据数组

    Private mintKeepFileVerCnt As Integer = 10 '保留文件的版本数
    Private mintKeepFileDays As Integer = 365 '保留文件的时间，以天为单位。

    Public Sub New(FilePath As String)
        MyBase.New(CLS_VERSION)
        mstrFilePath = FilePath
        Me.GbMain = New PigBytes
        ReDim Me.GbMain.Main(-1)
    End Sub

    ''' <summary>保留文件的时间，以天为单位。</summary>
    Public Property KeepFileDays As Integer
        Get
            KeepFileDays = mintKeepFileDays
        End Get
        Set(value As Integer)
            If value <= 0 Then value = 1
            mintKeepFileDays = value
        End Set
    End Property


    ''' <summary>保留文件的版本数</summary>
    Public Property KeepFileVerCnt As Integer
        Get
            KeepFileVerCnt = mintKeepFileVerCnt
        End Get
        Set(value As Integer)
            If value <= 0 Then value = 1
            mintKeepFileVerCnt = value
        End Set
    End Property

    ''' <summary>复制文件到</summary>
    ''' <param name="TarFile">目标文件路径</param>
    Public Overloads Function CopyFileTo(TarFile As String) As String
        Try
            System.IO.File.Copy(Me.FilePath, TarFile, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CopyFileTo", ex, False)
        End Try
    End Function

    ''' <summary>复制文件到</summary>
    ''' <param name="TarFile">目标文件路径</param>
    ''' <param name="IsOverWrite">是否强制覆盖</param>
    Public Overloads Function CopyFileTo(TarFile As String, IsOverWrite As Boolean) As String
        Try
            System.IO.File.Copy(Me.FilePath, TarFile, IsOverWrite)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CopyFileTo", ex, False)
        End Try
    End Function

    ''' <summary>设置数据</summary>
    Public Overloads Function SetData(DataBytes As Byte()) As String
        Try
            If Me.GbMain IsNot Nothing Then Me.GbMain = Nothing
            Me.GbMain = New PigBytes(DataBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetData", ex, False)
        End Try
    End Function

    ''' <summary>设置数据</summary>
    Public Overloads Function SetData(GETextData As PigText) As String
        Try
            If Not Me.GbMain Is Nothing Then Me.GbMain = Nothing
            Me.GbMain = New PigBytes(GETextData.TextBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetData", ex, False)
        End Try
    End Function

    Public Event EnvSegLoadFile(SegNo As Integer, SegBytes As Byte(), IsEnd As Boolean)

    ''' <summary>
    ''' 异步分段导入文件|Asynchronous segmented load file
    ''' </summary>
    ''' <param name="SegmentSize">一段大小|Segment size</param>
    ''' <returns></returns>
    Public Function SegLoadFile(SegmentSize As Long) As String
        Dim LOG As New StruStepLog : LOG.SubName = "SegLoadFile"
        Dim sfAny As FileStream = Nothing
        Dim brAny As BinaryReader = Nothing
        Try
            LOG.StepName = "New FileStream"
            sfAny = New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            LOG.StepName = "New BinaryReader"
            brAny = New BinaryReader(sfAny)
            LOG.StepName = "New PigBytes"
            Dim abSegFile(-1) As Byte, intRetSize As Integer = 0, intSegNo As Integer = 0, lngPos As Long = 0
            Dim lngFileSize As Long = Me.Size, intGetBytes As Integer = SegmentSize, bolIsEnd As Boolean = False
            Do While True
                If (lngPos + SegmentSize) > lngFileSize Then
                    intGetBytes = lngFileSize - lngPos
                    bolIsEnd = True
                End If
                ReDim abSegFile(intGetBytes - 1)
                LOG.StepName = "Read"
                intRetSize = brAny.Read(abSegFile, 0, intGetBytes)
                If intRetSize <= 0 Then
                    Throw New Exception("RetSize is " & intRetSize)
                ElseIf intRetSize <> intGetBytes Then
                    Throw New Exception("RetSize not equal to GetBytes")
                End If
                lngPos += intGetBytes
                RaiseEvent EnvSegLoadFile(intSegNo, abSegFile, bolIsEnd)
                If bolIsEnd = True Then Exit Do
                intSegNo += 1
            Loop
            LOG.StepName = "Close"
            brAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            If brAny IsNot Nothing Then brAny.Close()
            If sfAny IsNot Nothing Then sfAny.Close()
            LOG.AddStepNameInf(mstrFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the number of text lines in the file|获取文件的文本行数
    ''' </summary>
    ''' <returns></returns>
    Public Function GetTextRows() As Long
        Try
            Dim sfAny As FileStream = Nothing
            Dim srAny As StreamReader = Nothing
            sfAny = New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            srAny = New StreamReader(sfAny)
            GetTextRows = 0
            Do While Not srAny.EndOfStream
                srAny.ReadLine()
                GetTextRows += 1
            Loop
            srAny.Close()
            sfAny.Close()
        Catch ex As Exception
            Me.SetSubErrInf("GetTextRows", ex)
            Return -1
        End Try
    End Function


    ''' <summary>
    ''' 读取文本文件的尾部|Read the tail of a text file
    ''' </summary>
    ''' <param name="Rows"></param>
    ''' <returns></returns>
    Public Function GetTailText(Rows As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetTailText"
        Dim sfAny As FileStream = Nothing
        Dim srAny As StreamReader = Nothing
        Try
            Dim lngTotalRows As Long = Me.GetTextRows
            If lngTotalRows < 0 Then Throw New Exception("Unable to obtain the number of file lines")
            If Rows > lngTotalRows Then Rows = lngTotalRows
            Dim lngSkip As Long = lngTotalRows - Rows
            LOG.StepName = "New FileStream"
            sfAny = New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            LOG.StepName = "New StreamReader"
            srAny = New StreamReader(sfAny)
            GetTailText = ""
            Dim strCrLf As String = Me.OsCrLf
            Do While Not srAny.EndOfStream
                If lngSkip > 0 Then
                    srAny.ReadLine()
                    lngSkip -= 1
                Else
                    GetTailText &= srAny.ReadLine & strCrLf
                End If
            Loop
            srAny.Close()
            sfAny.Close()
        Catch ex As Exception
            If srAny IsNot Nothing Then srAny.Close()
            If sfAny IsNot Nothing Then sfAny.Close()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 读取文本文件的顶部|Read the top of a text file
    ''' </summary>
    ''' <param name="Rows"></param>
    ''' <returns></returns>
    Public Function GetTopText(Rows As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetTopText"
        Try
            LOG.StepName = "New FileStream"
            Dim sfAny As New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            LOG.StepName = "New StreamReader"
            Dim srAny = New StreamReader(sfAny)
            Dim strCrLf As String = Me.OsCrLf
            GetTopText = ""
            Dim i As Integer = 0
            Do While Not srAny.EndOfStream
                GetTopText &= srAny.ReadLine & strCrLf
                i += 1
                If i >= Rows Then Exit Do
            Loop
            srAny.Close()
            sfAny.Close()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function


    ''' <summary>导入数据</summary>
    Public Function LoadFile() As String
        Dim LOG As New StruStepLog : LOG.SubName = "LoadFile"
        Try
            If Me.GbMain IsNot Nothing Then Me.GbMain = Nothing
            LOG.StepName = "New FileStream"
            Dim sfAny As New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            LOG.StepName = "New BinaryReader"
            Dim brAny = New BinaryReader(sfAny)
            LOG.StepName = "New PigBytes"
            Me.GbMain = New PigBytes
            LOG.StepName = "GbMain ReadBytes"
            GbMain.Main = brAny.ReadBytes(Me.Size)
            LOG.StepName = "Close"
            brAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(mstrFilePath)
            If Me.IsDebug Or Me.IsHardDebug Then
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex, True)
            Else
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            End If
        End Try
    End Function

    ''' <summary>保存数据</summary>
    ''' <param name="IsKeepVerFile">是否保留版本文件</param>
    Public Overloads Function SaveFile(IsKeepVerFile As Boolean) As String
        SaveFile = Me.mSaveFile(IsKeepVerFile)
    End Function

    ''' <summary>保存数据</summary>
    Public Overloads Function SaveFile() As String
        SaveFile = Me.mSaveFile(True)
    End Function

    ''' <summary>保存数据</summary>
    ''' <param name="IsKeepVerFile">是否保留版本文件</param>
    Private Function mSaveFile(IsKeepVerFile As Boolean) As String
        Dim strStepName As String = ""
        Try
            strStepName = "Check the data"    '检查数据
            If Me.GbMain Is Nothing Then Throw New Exception("No data")
            If IsKeepVerFile = True Then
                strStepName = "Locate the current backup file"  '定位当前备份文件
                Dim i As Integer, strBakFile As String = "", bolIsFind As Boolean = False
                For i = 1 To Me.KeepFileVerCnt
                    strBakFile = Me.FilePath & "." & i.ToString
                    If Me.IsExists(strBakFile) = False Then
                        bolIsFind = True
                        Exit For
                    End If
                Next
                strStepName = "Locating backup file"   '定位备份文件
                If bolIsFind = False Then
                    strBakFile = Me.FilePath & "." & Me.KeepFileVerCnt.ToString
                    Dim oGEFile As New PigFile(strBakFile)
                    If oGEFile.IsExists = False Then
                        bolIsFind = True
                    Else
                        For i = 1 To Me.KeepFileVerCnt
                            strBakFile = Me.FilePath & "." & i.ToString
                            Dim oGEFile2 As New PigFile(strBakFile)
                            If oGEFile2.IsExists = False Then
                                bolIsFind = True
                                Exit For
                            End If
                            If oGEFile2.UpdateTime < oGEFile.UpdateTime Then
                                bolIsFind = True
                                Exit For
                            End If
                            oGEFile2 = Nothing
                        Next
                        If bolIsFind = False Then
                            bolIsFind = True
                            strBakFile = Me.FilePath & ".1"
                        End If
                    End If
                End If
                If bolIsFind = False Then Throw New Exception("Unable to determine backup file")
                strStepName = "Backup" & Me.FilePath & " to " & strBakFile

            End If
            strStepName = "New FileStream(" & mstrFilePath & ")"
            Dim sfAny As New FileStream(mstrFilePath, FileMode.Create, FileAccess.Write, FileShare.None, Me.GbMain.Main.Length, False)
            'strStepName = "New StreamWriter"
            'Dim swAny = New StreamWriter(sfAny)
            strStepName = "New BinaryWriter"
            Dim bwAny = New BinaryWriter(sfAny)
            strStepName = "Write Bytes"
            bwAny.Write(Me.GbMain.Main)
            strStepName = "Close"
            bwAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSaveFile", ex, False)
        End Try
    End Function


    ''' <summary>本文件MD5</summary>
    Public ReadOnly Property MD5 As String
        Get
            If mPigMD5 IsNot Nothing Then
                Return mPigMD5.MD5
            Else
                Dim strRet As String = Me.mGetMyMD5()
                If strRet = "OK" Then
                    Return mPigMD5.MD5
                Else
                    Return ""
                End If
            End If
        End Get
    End Property

    ''' <summary>本文件豚豚MD5</summary>
    Public ReadOnly Property PigMD5 As String
        Get
            If mPigMD5 IsNot Nothing Then
                Return mPigMD5.PigMD5
            Else
                Dim strRet As String = Me.mGetMyMD5()
                If strRet = "OK" Then
                    Return mPigMD5.PigMD5
                Else
                    Return ""
                End If
            End If
        End Get
    End Property

    Public Function GetFullPigMD5(ByRef FullPigMD5 As String) As String
        Try
            If Me.mPigMD5 Is Nothing Then
                Dim strRet As String = Me.mGetMyMD5()
                If strRet <> "OK" Then Throw New Exception（strRet)
            End If
            FullPigMD5 = Me.mPigMD5.PigMD5
            Return "OK"
        Catch ex As Exception
            FullPigMD5 = ""
            Return Me.GetSubErrInf("GetFullPigMD5", ex)
        End Try
    End Function

    Public Function GetFullPigMD5(ByRef FullPigMD5 As PigMD5) As String
        Try
            If Me.mPigMD5 Is Nothing Then
                Dim strRet As String = Me.mGetMyMD5()
                If strRet <> "OK" Then Throw New Exception（strRet)
            End If
            FullPigMD5 = Me.mPigMD5
            Return "OK"
        Catch ex As Exception
            FullPigMD5 = Nothing
            Return Me.GetSubErrInf("GetFullPigMD5", ex)
        End Try
    End Function

    Private mPigMD5 As PigMD5
    Private Function mGetMyMD5() As String
        Dim LOG As New StruStepLog : LOG.SubName = "mGetMyMD5"
        Const SEGMENTSIZE As Integer = 20971520
        Dim sfAny As FileStream = Nothing
        Dim brAny As BinaryReader = Nothing
        Try
            If mPigMD5 Is Nothing Then
                LOG.StepName = "New FileStream"
                sfAny = New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                LOG.StepName = "New BinaryReader"
                brAny = New BinaryReader(sfAny)
                LOG.StepName = "New PigBytes"
                Dim pbMD5 As New PigBytes
                Dim abSegFile(-1) As Byte, intRetSize As Integer = 0, intSegNo As Integer = 0, lngPos As Long = 0
                Dim lngFileSize As Long = Me.Size, intGetBytes As Integer = SEGMENTSIZE, bolIsEnd As Boolean = False
                Do While True
                    If (lngPos + SEGMENTSIZE) > lngFileSize Then
                        intGetBytes = lngFileSize - lngPos
                        bolIsEnd = True
                    End If
                    ReDim abSegFile(intGetBytes - 1)
                    LOG.StepName = "Read"
                    intRetSize = brAny.Read(abSegFile, 0, intGetBytes)
                    If intRetSize <= 0 Then
                        Throw New Exception("RetSize is " & intRetSize)
                    ElseIf intRetSize <> intGetBytes Then
                        Throw New Exception("RetSize not equal to GetBytes")
                    End If
                    lngPos += intGetBytes
                    Dim oPigMD5 As New PigMD5(abSegFile)
                    pbMD5.SetValue(oPigMD5.PigMD5Bytes)
                    oPigMD5 = Nothing
                    If bolIsEnd = True Then Exit Do
                    intSegNo += 1
                Loop
                LOG.StepName = "Close"
                brAny.Close()
                sfAny.Close()
                LOG.StepName = "New PigMD5"
                Me.mPigMD5 = New PigMD5(pbMD5.Main)
                If Me.mPigMD5.LastErr <> "" Then Throw New Exception(Me.mPigMD5.LastErr)
            End If
            Return "OK"
        Catch ex As Exception
            If brAny IsNot Nothing Then brAny.Close()
            If sfAny IsNot Nothing Then sfAny.Close()
            mPigMD5 = Nothing
            LOG.AddStepNameInf(mstrFilePath)
            If Me.IsDebug Or Me.IsHardDebug Then
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex, True)
            Else
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            End If
        End Try
    End Function



    ''' <summary>文件完整路径</summary>
    Public ReadOnly Property FilePath As String
        Get
            FilePath = mstrFilePath
        End Get
    End Property

    Private Sub mInitFileInf()
        Try
            If moFileInfo Is Nothing Then
                moFileInfo = New FileInfo(mstrFilePath)
            Else
                moFileInfo.Refresh()
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mInitFileInf", ex, False)
        End Try
    End Sub

    ''' <summary>The directory path of the file|文件的目录路径</summary>
    Public ReadOnly Property DirPath() As String
        Get
            Try
                Me.mInitFileInf()
                DirPath = moFileInfo.DirectoryName
            Catch ex As Exception
                DirPath = ""
            End Try
        End Get
    End Property

    ''' <summary>File name|文件名</summary>
    Public ReadOnly Property FileTitle() As String
        Get
            Try
                Me.mInitFileInf()
                FileTitle = moFileInfo.Name
            Catch ex As Exception
                FileTitle = ""
            End Try
        End Get
    End Property

    ''' <summary>File extension|文件的扩展名</summary>
    Public ReadOnly Property ExtName() As String
        Get
            Try
                Me.mInitFileInf()
                ExtName = moFileInfo.Extension
            Catch ex As Exception
                ExtName = ""
            End Try
        End Get
    End Property

    ''' <summary>File size|文件大小</summary>
    Public ReadOnly Property Size() As Long
        Get
            Try
                Me.mInitFileInf()
                Size = moFileInfo.Length
            Catch ex As Exception
                Size = -1
            End Try
        End Get
    End Property

    ''' <summary>File creation time|文件的创建时间</summary>
    Public ReadOnly Property CreationTime() As DateTime
        Get
            Try
                Me.mInitFileInf()
                CreationTime = moFileInfo.CreationTime
            Catch ex As Exception
                CreationTime = #1/1/1900#
            End Try
        End Get
    End Property

    ''' <summary>File update time|文件的更新时间</summary>
    Public ReadOnly Property UpdateTime() As DateTime
        Get
            Try
                Me.mInitFileInf()
                UpdateTime = moFileInfo.LastWriteTime
            Catch ex As Exception
                UpdateTime = #1/1/1900#
            End Try
        End Get
    End Property

    ''' <summary>Does the file exist|文件是否存在</summary>
    Public Overloads ReadOnly Property IsExists As Boolean
        Get
            Try
                Me.mInitFileInf()
                IsExists = moFileInfo.Exists
            Catch ex As Exception
                IsExists = False
            End Try
        End Get
    End Property

    ''' <summary>Does the file exist|文件是否存在</summary>
    ''' <param name="FilePath">specify the path to a file|指定文件路径</param>
    Public Overloads ReadOnly Property IsExists(FilePath As String) As Boolean
        Get
            Try
                Dim oFileInfo = New FileInfo(FilePath)
                IsExists = oFileInfo.Exists
                oFileInfo = Nothing
            Catch ex As Exception
                IsExists = False
            End Try
        End Get
    End Property

    Public Function GetFastPigMD5(ByRef FastPigMD5 As String, Optional ScanSize As Integer = 20480) As String
        Try
            Dim oPigMD5 As PigMD5 = Nothing
            GetFastPigMD5 = Me.mGetFastPigMD5(oPigMD5, ScanSize)
            If oPigMD5 Is Nothing Then
                FastPigMD5 = ""
            Else
                FastPigMD5 = oPigMD5.PigMD5
                oPigMD5 = Nothing
            End If
        Catch ex As Exception
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

    Public Function GetFastPigMD5(ByRef FastPigMD5 As PigMD5, Optional ScanSize As Integer = 20480) As String
        Return Me.mGetFastPigMD5(FastPigMD5, ScanSize)
    End Function


    Private Function mGetFastPigMD5(ByRef FastPigMD5 As PigMD5, Optional ScanSize As Integer = 20480) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mGetFastPigMD5"
        Dim i As Long
        Dim intSize As Long = Me.Size
        Dim sfAny As FileStream = Nothing
        Dim brAny As BinaryReader = Nothing
        Try
            FastPigMD5 = Nothing
            Dim abData(-1) As Byte, intAdd As Integer = 0
            LOG.StepName = "Check file"
            If Me.IsExists(Me.FilePath) = False Then Throw New Exception("File not found")
            If Me.Size = 0 Then
                FastPigMD5 = Nothing
            Else
                LOG.StepName = "New FileStream"
                sfAny = New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                LOG.StepName = "New BinaryReader"
                brAny = New BinaryReader(sfAny)
                If ScanSize > Me.Size Then ScanSize = Me.Size
                LOG.StepName = "Set abData"
                ReDim abData(ScanSize - 1)
                intAdd = Me.Size / ScanSize - 1
                If intAdd <= 0 Then intAdd = 1
                Dim intPos As Integer = 1
                LOG.StepName = "Read"
                Dim abOne(Me.Size - 1) As Byte, intRet As Integer
                For i = 0 To ScanSize - 1
                    intPos = i * intAdd
                    intRet = brAny.Read(abOne, intPos, 1)
                    If intRet < 0 Then
                        LOG.AddStepNameInf("i=" & i.ToString)
                        LOG.AddStepNameInf("Pos=" & intPos.ToString)
                        LOG.AddStepNameInf("Ret=" & intRet.ToString)
                        Throw New Exception("Error")
                    End If
                    abData(i) = abOne(intPos)
                Next
                LOG.StepName = "New PigMD5.3"
                FastPigMD5 = New PigMD5(abData)
                If FastPigMD5.LastErr <> "" Then
                    Throw New Exception(FastPigMD5.LastErr)
                End If
                LOG.StepName = "Close"
                brAny.Close()
                sfAny.Close()
            End If
            Return "OK"
        Catch ex As Exception
            If brAny IsNot Nothing Then brAny.Close()
            If sfAny IsNot Nothing Then sfAny.Close()
            FastPigMD5 = Nothing
            LOG.AddStepNameInf(Me.FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFullText(ByRef FileText As PigText, TextType As PigText.enmTextType) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetFullText"
        Try
            LOG.StepName = "Check GbMain.Main.Length"
            If Me.GbMain.Main.Length <> Me.Size Then
                LOG.StepName = "LoadFile"
                LOG.Ret = Me.LoadFile
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "New PigText"
            FileText = New PigText(Me.GbMain.Main, TextType)
            If FileText.LastErr <> "" Then Throw New Exception(FileText.LastErr)
            Return "OK"
        Catch ex As Exception
            FileText = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
