'**********************************
'* Name: CmdZip
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Command line operations for commonly used compressed packages|常用压缩包的命令行操作
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 15/11/2023
'* 1.1  16/11/2023  Modify New,AddArchive, add ExtractArchive
'* 1.2  17/11/2023  Add mAddArchive, modify AddArchive
'* 1.3  20/11/2023  Add mAddArchive,mExtractArchive, modify AddArchive,ExtractArchive
'* 1.5  21/7/2024  Modify PigFunc to PigFuncLite
'**********************************
Imports PigToolsLiteLib

Public Class CmdZip
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "5" & "." & "38"

    Public Enum EnmZipType
        _7_Zip = 0
        WinRar = 1
        Tar = 2
    End Enum

    Public ReadOnly Property ZipType As EnmZipType
    Public ReadOnly Property ZipExePath As String
    Private ReadOnly Property mPigCmdApp As New PigCmdApp
    Private ReadOnly Property mPigFunc As New PigFuncLite

    Public Sub New(ZipType As EnmZipType, ZipExePath As String)
        MyBase.New(CLS_VERSION)
        Me.ZipExePath = ZipExePath
        Me.ZipType = ZipType
    End Sub

    Private Function mAddArchive(SrcDir As String, ZipFilePath As String, SuDoUser As String) As String
        Dim LOG As New PigStepLog("mAddArchive")
        Dim strCmd As String = ""
        Try
            If SuDoUser <> "" And Me.IsWindows = True Then Throw New Exception("Sudo can only run on the Linux platform.")
            If Me.mPigFunc.IsFileExists(Me.ZipExePath) = False Then
                LOG.StepName = "Check execution files"
                LOG.AddStepNameInf(Me.ZipExePath)
                LOG.Ret = "File not found."
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "Check the directory to be compressed"
            If Me.mPigFunc.IsFolderExists(SrcDir) = False Then
                LOG.AddStepNameInf(SrcDir)
                LOG.Ret = "Folder not found."
                Throw New Exception(LOG.Ret)
            ElseIf SrcDir = "/" And Me.IsWindows = False Or Me.IsWindows = True And Right(SrcDir, 2) = ":\" Then
                LOG.AddStepNameInf(SrcDir)
                LOG.Ret = "The root directory cannot be compressed"
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "Check ZipType"
            Select Case Me.ZipType
                Case EnmZipType.WinRar
                    If Me.IsWindows = False Then Throw New Exception("Only supports running on Windows.")
                Case EnmZipType.Tar
                    If Me.IsWindows = True Then Throw New Exception("Only supports running on Linux.")
            End Select
            Dim strZipFileDirPath As String = Me.mPigFunc.GetFilePart(ZipFilePath, PigFunc.EnmFilePart.Path)
            If Me.mPigFunc.IsFolderExists(strZipFileDirPath) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(strZipFileDirPath)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strZipFileDirPath)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Select Case Me.ZipType
                Case EnmZipType._7_Zip
                    If Me.IsWindows = True Then
                        strCmd = """" & Me.ZipExePath & """"
                    Else
                        strCmd = Me.ZipExePath
                    End If
                    strCmd &= " a -t7z "
                    If Me.IsWindows = True Then
                        strCmd &= """" & ZipFilePath & """ """ & SrcDir & """"
                    Else
                        strCmd &= ZipFilePath & " " & SrcDir
                    End If
                Case EnmZipType.Tar
                    Dim strDirPath As String = Me.mPigFunc.GetFilePart(SrcDir, PigFunc.EnmFilePart.Path)
                    Dim strFileTitle As String = Me.mPigFunc.GetFilePart(SrcDir, PigFunc.EnmFilePart.FileTitle)
                    strCmd = "cd " & strDirPath & " && " & Me.ZipExePath & " cvzf " & ZipFilePath & " " & strFileTitle
                Case EnmZipType.WinRar
                    Dim strDirPath As String = Me.mPigFunc.GetFilePart(SrcDir, PigFunc.EnmFilePart.Path)
                    Dim strFileTitle As String = Me.mPigFunc.GetFilePart(SrcDir, PigFunc.EnmFilePart.FileTitle)
                    Dim strZipExePath As String = """" & Me.ZipExePath & """"
                    strCmd = "cd /d """ & strDirPath & """ && call " & strZipExePath & " a """ & ZipFilePath & """ " & strFileTitle
                Case Else
                    LOG.AddStepNameInf(Me.ZipType.ToString)
                    Throw New Exception("Currently not supported")
            End Select
            If SuDoUser <> "" Then
                Dim oPigSudo As New PigSudo(strCmd, SuDoUser, False)
                LOG.StepName = "PigSudo.Run"
                LOG.Ret = oPigSudo.Run
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Else
                LOG.StepName = "CmdShell"
                LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If Me.mPigCmdApp.StandardError <> "" Then Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function



    ''' <summary>
    ''' Generate compressed packages for directories, using relative path compression|生成目录的压缩包，使用相对路径压缩
    ''' </summary>
    ''' <param name="SrcDir">Source directory path to compress|要压缩的源目录路径</param>
    ''' <param name="ZipFilePath">Generated compressed package path|生成的压缩包路径</param>
    ''' <returns></returns>
    Public Function AddArchive(SrcDir As String, ZipFilePath As String) As String
        Return Me.mAddArchive(SrcDir, ZipFilePath, "")
    End Function

    ''' <summary>
    ''' Generate compressed packages for directories, using relative path compression|生成目录的压缩包，使用相对路径压缩
    ''' </summary>
    ''' <param name="SrcDir">Source directory path to compress|要压缩的源目录路径</param>
    ''' <param name="ZipFilePath">Generated compressed package path|生成的压缩包路径</param>
    ''' <param name="SudoUser">sudo user</param>
    ''' <returns></returns>
    Public Function AddArchive(SrcDir As String, ZipFilePath As String, SudoUser As String) As String
        Return Me.mAddArchive(SrcDir, ZipFilePath, SudoUser)
    End Function

    ''' <summary>
    ''' Unzip the compressed package to the target directory, using a relative path|解压压缩包到目标目录，使用相对路径
    ''' </summary>
    ''' <param name="TargetDir">Target directory|目标目录</param>
    ''' <param name="ZipFilePath">Generated compressed package path|生成的压缩包路径</param>
    ''' <param name="SudoUser">sudo user</param>
    ''' <returns></returns>

    Public Function ExtractArchive(TargetDir As String, ZipFilePath As String, SudoUser As String) As String
        Return Me.mExtractArchive(TargetDir, ZipFilePath, SudoUser)
    End Function

    ''' <summary>
    ''' Unzip the compressed package to the target directory, using a relative path|解压压缩包到目标目录，使用相对路径
    ''' </summary>
    ''' <param name="TargetDir">Target directory|目标目录</param>
    ''' <param name="ZipFilePath">Generated compressed package path|生成的压缩包路径</param>
    ''' <returns></returns>

    Public Function ExtractArchive(TargetDir As String, ZipFilePath As String) As String
        Return Me.mExtractArchive(TargetDir, ZipFilePath, "")
    End Function


    Private Function mExtractArchive(TargetDir As String, ZipFilePath As String, SuDoUser As String) As String
        Dim LOG As New PigStepLog("mExtractArchive")
        Dim strCmd As String = ""
        Try
            If SuDoUser <> "" And Me.IsWindows = True Then Throw New Exception("Sudo can only run on the Linux platform.")
            If Me.mPigFunc.IsFileExists(Me.ZipExePath) = False Then
                LOG.StepName = "Check execution files"
                LOG.AddStepNameInf(Me.ZipExePath)
                LOG.Ret = "File not found."
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "Check ZipType"
            Select Case Me.ZipType
                Case EnmZipType.WinRar
                    If Me.IsWindows = False Then Throw New Exception("Only supports running on Windows.")
                Case EnmZipType.Tar
                    If Me.IsWindows = True Then Throw New Exception("Only supports running on Linux.")
            End Select
            If Me.mPigFunc.IsFolderExists(TargetDir) = False Then
                LOG.StepName = "CreateFolder"
                LOG.Ret = Me.mPigFunc.CreateFolder(TargetDir)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(TargetDir)
                    Throw New Exception(LOG.Ret)
                End If
            End If
            Select Case Me.ZipType
                Case EnmZipType._7_Zip
                    If Me.IsWindows = True Then
                        strCmd = "cd """ & TargetDir & """ && """ & Me.ZipExePath & """"
                    Else
                        strCmd = "cd " & TargetDir & " && " & Me.ZipExePath
                    End If
                    strCmd &= " x -t7z "
                    If Me.IsWindows = True Then
                        strCmd &= """" & ZipFilePath & """"
                    Else
                        strCmd &= ZipFilePath
                    End If
                Case EnmZipType.Tar
                    strCmd = "cd " & TargetDir & " && " & Me.ZipExePath & " xzsf " & ZipFilePath
                Case EnmZipType.WinRar
                    Dim strZipExePath As String = """" & Me.ZipExePath & """"
                    strCmd = "cd /d """ & TargetDir & """ && call " & strZipExePath & " x """ & ZipFilePath & """"
                Case Else
                    LOG.StepName = "Check ZipType"
                    LOG.AddStepNameInf(Me.ZipType.ToString)
                    Throw New Exception("Currently not supported")
            End Select
            If SuDoUser <> "" Then
                Dim oPigSudo As New PigSudo(strCmd, SuDoUser, False)
                LOG.StepName = "PigSudo.Run"
                LOG.Ret = oPigSudo.Run
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Else
                LOG.StepName = "CmdShell"
                LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If Me.mPigCmdApp.StandardError <> "" Then Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
