'**********************************
'* Name: PigFileSystem
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 用于目录及文件操作|Used for directory and file operations
'* Home Url: https://en.seowphong.com
'* Version: 1.6
'* Create Time: 10/6/2023
'* 1.1 11/6/2023   Add GetPigFile,GetPigFolder
'* 1.2 15/6/2023   Add DeleteFile,CopyFile,IsFileExists,MoveFile
'* 1.3 24/6/2023   Add IOMode,OpenTextFile
'* 1.5 5/12/2023   Add OpenTextFile
'* 1.6 19/12/2023  Modify OpenTextFile, add mOpenTextFile
'**********************************
Imports System.IO

Public Class PigFileSystem
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "6" & "." & "10"

    Public Enum IOMode
        ForAppending = 8
        ForReading = 1
        ForWriting = 2
    End Enum

    Private Property mObj As Object

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function GetPigFile(FilePath As String) As PigFile
        Try
            GetPigFile = New PigFile(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("GetPigFile", ex)
            Return Nothing
        End Try
    End Function

    Public Function GetPigFolder(FolderPath As String) As PigFolder
        Try
            GetPigFolder = New PigFolder(FolderPath)
        Catch ex As Exception
            Me.SetSubErrInf("GetPigFolder", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Delete file|删除文件
    ''' </summary>
    ''' <param name="FilePath">File absolute path|文件绝对路径</param>
    ''' <returns></returns>
    Public Function DeleteFile(FilePath As String) As String
        Try
            File.Delete(FilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFile", ex)
        End Try
    End Function


    ''' <summary>
    ''' Move file|移动文件
    ''' </summary>
    ''' <param name="SourceFile">Source file|源文件</param>
    ''' <param name="TargetFile">Target file|目标文件</param>
    ''' <param name="IsOverWrite">Overwrite or not|是否覆盖</param>
    ''' <returns></returns>
    Public Function MoveFile(SourceFile As String, TargetFile As String, Optional IsOverwrite As Boolean = True) As String
        Dim LOG As New PigStepLog("MoveFile")
        Try
            If SourceFile = TargetFile Then Throw New Exception("Cannot move file itself.")
            If Me.IsFileExists(TargetFile) = True Then
                If IsOverwrite = True Then
                    LOG.StepName = "DeleteFile"
                    LOG.Ret = Me.DeleteFile(TargetFile)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Else
                    Throw New Exception("The target file already exists.")
                End If
            End If
            LOG.StepName = "Move"
            IO.File.Move(SourceFile, TargetFile)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(SourceFile)
            LOG.AddStepNameInf(TargetFile)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Copy file|复制文件
    ''' </summary>
    ''' <param name="SourceFile">Source file|源文件</param>
    ''' <param name="TargetFile">Target file|目标文件</param>
    ''' <param name="IsOverWrite">Overwrite or not|是否覆盖</param>
    ''' <returns></returns>
    Public Function CopyFile(SourceFile As String, TargetFile As String, Optional IsOverwrite As Boolean = True) As String
        Dim LOG As New PigStepLog("CopyFile")
        Try
            If SourceFile = TargetFile Then Throw New Exception("Cannot copy file itself.")
            If IsOverwrite = False Then
                If Me.IsFileExists(TargetFile) = True Then Throw New Exception("The target file already exists.")
            End If
            LOG.StepName = "Copy"
            IO.File.Copy(SourceFile, TargetFile, IsOverwrite)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(SourceFile)
            LOG.AddStepNameInf(TargetFile)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Determine if the file exists|判断文件是否存在
    ''' </summary>
    ''' <param name="FilePath">File absolute path|文件绝对路径</param>
    ''' <returns></returns>
    Public Function IsFileExists(FilePath As String) As Boolean
        Try
            Return IO.File.Exists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("IsFileExists", ex)
            Return Nothing
        End Try
    End Function

    Public Function OpenTextFile(FilePath As String, IOMode As IOMode, Optional Create As Boolean = False) As TextStream
        Return Me.mOpenTextFile(FilePath, IOMode, PigText.enmTextType.UnknowOrBin, Create)
    End Function

    Public Function OpenTextFile(FilePath As String, IOMode As IOMode, TextType As PigText.enmTextType, Optional Create As Boolean = False) As TextStream
        Return Me.mOpenTextFile(FilePath, IOMode, TextType, Create)
    End Function


    Private Function mOpenTextFile(FilePath As String, IOMode As IOMode, TextType As PigText.enmTextType, Optional Create As Boolean = False) As TextStream
        Const SUB_NAME As String = "mOpenTextFile"
        Dim strStepName As String = ""
        Try
            Select Case TextType
                Case PigText.enmTextType.UnknowOrBin
                    mOpenTextFile = New TextStream()
                Case Else
                    mOpenTextFile = New TextStream(TextType)
            End Select
            strStepName = "Init"
            Me.PrintDebugLog(SUB_NAME, strStepName, FilePath)
            mOpenTextFile.Init(FilePath, IOMode, Create)
            If mOpenTextFile.LastErr <> "" Then Throw New Exception(mOpenTextFile.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function OpenTextFileAsc(FilePath As String, IOMode As IOMode, Optional Create As Boolean = False) As TextStreamAsc
        Const SUB_NAME As String = "OpenTextFileAsc"
        Dim strStepName As String = ""
        Try
#If NETFRAMEWORK Then
            If Me.mObj Is Nothing Then
                strStepName = "CreateObject"
                Me.mObj = CreateObject("Scripting.FileSystemObject")
            End If
            strStepName = "New TextStreamAsc"
            OpenTextFileAsc = New TextStreamAsc
            If OpenTextFileAsc.LastErr <> "" Then Throw New Exception(OpenTextFileAsc.LastErr)
            strStepName = "OpenTextFile"
            OpenTextFileAsc.Obj = Me.mObj.OpenTextFile(FilePath, IOMode, Create)
#Else
            Throw New Exception("This function can only be run under NETFRAMEWORK")
#End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return Nothing
        End Try
    End Function

End Class
