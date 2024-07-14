'**********************************
'* Name: 豚豚进程|PigProc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Process class|进程类
'* Home Url: https://en.seowphong.com
'* Version: 1.7
'* Create Time: 20/3/2022
'* 1.1    2/4/2022   Modify new
'* 1.2    1/8/2022   Add Close
'* 1.3    1/8/2022   Add Kill
'* 1.5    5/9/2022   Modify StartTime
'* 1.6    27/6/2024  Add ThreadsCount,ProcFileName,ProcUserName,ProcWorkingDirectory
'* 1.7    12/7/2024  Modify ProcWorkingDirectory etc.
'**********************************
''' <summary>
''' Process class|进程类
''' </summary>
Public Class PigProc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "7" & "." & "18"

    Private moProcess As Process

    Public Sub New(ProcessID As Integer)
        MyBase.New(CLS_VERSION)
        Try
            moProcess = Process.GetProcessById(ProcessID)
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Public ReadOnly Property ThreadsCount As Integer
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.Threads IsNot Nothing Then
                    Return moProcess.Threads.Count
                Else
                    Return -1
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ThreadsCount.Get", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcFileName As String
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.StartInfo IsNot Nothing Then
                    Return moProcess.StartInfo.FileName
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ProcFileName.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcUserName As String
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.StartInfo IsNot Nothing Then
                    Return moProcess.StartInfo.UserName
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ProcUserName.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcWorkingDirectory As String
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.StartInfo IsNot Nothing Then
                    Return moProcess.StartInfo.WorkingDirectory
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ProcWorkingDirectory.Get", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcessID As Integer
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return moProcess.Id
                Else
                    Return -1
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ProcessID", ex)
                Return -1
            End Try
        End Get
    End Property

    Public ReadOnly Property ProcessName As String
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return moProcess.ProcessName
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ProcessName", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ModuleName As String
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.MainModule IsNot Nothing Then
                    Return moProcess.MainModule.ModuleName
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("ModuleName", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property FilePath As String
        Get
            Try
                If moProcess IsNot Nothing AndAlso moProcess.MainModule IsNot Nothing Then
                    Return moProcess.MainModule.FileName
                Else
                    Return ""
                End If
            Catch ex As Exception
                Me.SetSubErrInf("FilePath", ex)
                Return ""
            End Try
        End Get
    End Property

    Public ReadOnly Property MemoryUse As Long
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return moProcess.WorkingSet64
                Else
                    Return 0
                End If
            Catch ex As Exception
                Me.SetSubErrInf("MemoryUse", ex)
                Return 0
            End Try
        End Get
    End Property

    Public ReadOnly Property StartTime As Date
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return CDate(Format(moProcess.StartTime, "yyyy-MM-dd HH:mm:ss.fff"))
                Else
                    Return #1/1/2000#
                End If
            Catch ex As Exception
                Me.SetSubErrInf("StartTime", ex)
                Return #1/1/2000#
            End Try
        End Get
    End Property

    Public ReadOnly Property TotalProcessorTime As TimeSpan
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return moProcess.TotalProcessorTime
                Else
                    Return TimeSpan.MinValue
                End If
            Catch ex As Exception
                Me.SetSubErrInf("TotalProcessorTime", ex)
                Return TimeSpan.MinValue
            End Try
        End Get
    End Property

    Public ReadOnly Property UserProcessorTime As TimeSpan
        Get
            Try
                If moProcess IsNot Nothing Then
                    Return moProcess.UserProcessorTime
                Else
                    Return TimeSpan.MinValue
                End If
            Catch ex As Exception
                Me.SetSubErrInf("UserProcessorTime", ex)
                Return TimeSpan.MinValue
            End Try
        End Get
    End Property

    Public Function Close() As String
        Try
            If moProcess IsNot Nothing Then
                moProcess.Close()
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Close", ex)
        End Try
    End Function

    Public Function Kill() As String
        Try
            If moProcess IsNot Nothing Then
                moProcess.Kill()
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Close", ex)
        End Try
    End Function

End Class
