'**********************************
'* Name: FileSystemObject
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.FileSystemObject of VB6
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 31/12/2020
'* 1.0.2 15/1/2021   Err.Raise change to Throw New Exception
'* 1.0.3 23/1/2021   pFileSystemObject rename to FileSystemObject
'* 1.0.4 26/1/2021   pIOMode rename to IOMode
'* 1.0.5 27/1/2021   Add AppPath,AppTitle,IsWindows,OsCrLf,OsPathSep
'* 1.0.6 25/7/2021   Modify OpenTextFile
'* 1.1 13/3/2021   Add Obj
'**********************************
Imports System.IO
Friend Class mFileSystemObject
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"

#If NETFRAMEWORK Then
    Friend Obj As Object
#End If

    Public Enum IOMode
        ForAppending = 8
        ForReading = 1
        ForWriting = 2
    End Enum


    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Sub MoveFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
        Try
            IO.File.Move(Source, Destination)
        Catch ex As Exception
            Me.SetSubErrInf("MoveFile", ex)
        End Try
    End Sub

    Public Sub CopyFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
        Try
            IO.File.Copy(Source, Destination, OverWriteFiles)
        Catch ex As Exception
            Me.SetSubErrInf("CopyFile", ex)
        End Try
    End Sub


    Public ReadOnly Property FileExists(FileSpec As String) As Boolean
        Get
            Try
                Return IO.File.Exists(FileSpec)
            Catch ex As Exception
                Me.SetSubErrInf("FileExists", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property FolderExists(FolderSpec As String) As Boolean
        Get
            Try
                Return Directory.Exists(FolderSpec)
            Catch ex As Exception
                Me.SetSubErrInf("FolderExists", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function OpenTextFile(FilePath As String, IOMode As IOMode, Optional Create As Boolean = False) As mTextStream
        Const SUB_NAME As String = "OpenTextFile"
        Dim strStepName As String = ""
        Try
            OpenTextFile = New mTextStream
            strStepName = "Init"
            Me.PrintDebugLog(SUB_NAME, strStepName, FilePath)
            OpenTextFile.Init(FilePath, IOMode, Create)
            If OpenTextFile.LastErr <> "" Then Throw New Exception(OpenTextFile.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function OpenTextFileAsc(FilePath As String, IOMode As IOMode, Optional Create As Boolean = False) As mTextStreamAsc
        Const SUB_NAME As String = "OpenTextFileAsc"
        Dim strStepName As String = ""
        Try
#If NETFRAMEWORK Then
            If Me.Obj Is Nothing Then
                strStepName = "CreateObject"
                Me.Obj = CreateObject("Scripting.FileSystemObject")
            End If
            strStepName = "New TextStreamAsc"
            OpenTextFileAsc = New mTextStreamAsc
            If OpenTextFileAsc.LastErr <> "" Then Throw New Exception(OpenTextFileAsc.LastErr)
            strStepName = "OpenTextFile"
            OpenTextFileAsc.Obj = Me.Obj.OpenTextFile(FilePath, IOMode, Create)
#Else
            Throw New Exception("This function can only be run under NETFRAMEWORK")
#End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Return Nothing
        End Try
    End Function

    Public Shadows ReadOnly Property AppPath() As String
        Get
            Return MyBase.AppPath
        End Get
    End Property

    Public Shadows ReadOnly Property AppTitle() As String
        Get
            Return MyBase.AppTitle
        End Get
    End Property

    Public Shadows ReadOnly Property IsWindows() As Boolean
        Get
            Return MyBase.IsWindows
        End Get
    End Property

    Public Shadows ReadOnly Property OsCrLf() As String
        Get
            Return MyBase.OsCrLf
        End Get
    End Property

    Public Shadows ReadOnly Property OsPathSep() As String
        Get
            Return MyBase.OsPathSep
        End Get
    End Property

End Class
