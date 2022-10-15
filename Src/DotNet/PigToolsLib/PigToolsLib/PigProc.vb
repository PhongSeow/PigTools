'**********************************
'* Name: 豚豚进程|PigProc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 进程|Process
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 20/3/2022
'* 1.1    2/4/2022   Modify new
'* 1.2    1/8/2022   Add Close
'* 1.3    1/8/2022   Add Kill
'* 1.5    5/9/2022   Modify StartTime
'**********************************
Public Class PigProc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.2"

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

    Public ReadOnly Property MemoryUse As Long
        Get
            Try
                Return moProcess.WorkingSet64
            Catch ex As Exception
                Me.SetSubErrInf("MemoryUse", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property StartTime As Date
        Get
            Try
                Return CDate(Format(moProcess.StartTime, "yyyy-MM-dd HH:mm:ss.fff"))
            Catch ex As Exception
                Me.SetSubErrInf("StartTime", ex)
                Return #1/1/1900#
            End Try
        End Get
    End Property

    Public ReadOnly Property TotalProcessorTime As TimeSpan
        Get
            Try
                Return moProcess.TotalProcessorTime
            Catch ex As Exception
                Me.SetSubErrInf("TotalProcessorTime", ex)
                Return TimeSpan.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property UserProcessorTime As TimeSpan
        Get
            Try
                Return moProcess.UserProcessorTime
            Catch ex As Exception
                Me.SetSubErrInf("UserProcessorTime", ex)
                Return TimeSpan.MinValue
            End Try
        End Get
    End Property

    Public Function Close() As String
        Try
            moProcess.Close()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Close", ex)
        End Try
    End Function

    Public Function Kill() As String
        Try
            moProcess.Kill()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Close", ex)
        End Try
    End Function

End Class
