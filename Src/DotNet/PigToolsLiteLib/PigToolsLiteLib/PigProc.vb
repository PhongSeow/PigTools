'**********************************
'* Name: 豚豚进程|PigProc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 进程|Process
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 20/3/2022
'**********************************
Public Class PigProc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.2"

    Private moProcess As Process

    Public Sub New(ProcessID As Integer)
        MyBase.New(CLS_VERSION)
        Try
            moProcess = Process.GetProcessById(ProcessID)
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub


    Public ReadOnly Property ProcessID As Integer
        Get
            Try
                Return moProcess.Id
            Catch ex As Exception
                Me.SetSubErrInf("ProcessID", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcessName As String
        Get
            Try
                Return moProcess.ProcessName
            Catch ex As Exception
                Me.SetSubErrInf("ProcessName", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ModuleName As String
        Get
            Try
                Return moProcess.MainModule.ModuleName
            Catch ex As Exception
                Me.SetSubErrInf("ModuleName", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property FilePath As String
        Get
            Try
                Return moProcess.MainModule.FileName
            Catch ex As Exception
                Me.SetSubErrInf("FilePath", ex)
                Return ""
            End Try
        End Get
    End Property


End Class
