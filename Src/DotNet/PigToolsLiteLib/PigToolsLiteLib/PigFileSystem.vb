'**********************************
'* Name: PigFileSystem
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 用于目录及文件操作|Used for directory and file operations
'* Home Url: https://en.seowphong.com
'* Version: 1.1
'* Create Time: 10/6/2023
'* 1.1 11/6/2023   Add GetPigFile,GetPigFolder
'**********************************
Imports System.IO

Public Class PigFileSystem
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"

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

End Class
