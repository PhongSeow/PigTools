'**********************************
'* Name: PigFileSystem
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 用于目录及文件操作|Used for directory and file operations
'* Home Url: https://en.seowphong.com
'* Version: 1.2
'* Create Time: 10/6/2023
'* 1.1 11/6/2023   Add GetPigFile,GetPigFolder
'* 1.2 15/6/2023   Add DeleteFile,CopyFile,IsFileExists,MoveFile
'**********************************
Imports System.IO

Public Class PigFileSystem
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.18"

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
        Dim LOG As New PigStepLog("MoveFile")
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


End Class
