﻿'**********************************
'* Name: 豚豚进程应用|PigProcApp
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Process Processing Class|进程处理类
'* Home Url: https://en.seowphong.com
'* Version: 1.9
'* Create Time: 20/3/2022
'* 1.1    26/3/2022  Modify GetPigProc(PID), Add GetPigProcs
'* 1.2    1/8/2022   Add KillProc,KillProcs
'* 1.3    1/8/2022   Modify KillProc
'* 1.4    16/8/2022  Add IsOtherExeExists
'* 1.5    17/8/2022  Add KillOtherExe
'* 1.6    15/5/2023  Modify IsOtherExeExists
'* 1.7    24/6/2024  Add IsProcExists
'* 1.8    12/7/2024  Add GetPigProc
'* 1.9    27/7/2024   Modify PigStepLog to StruStepLog
'**********************************

Imports System.Security.Cryptography

''' <summary>
''' Process Processing Class|进程处理类
''' </summary>
Public Class PigProcApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "9" & "." & "12"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function GetPigProc(PID As Integer) As PigProc
        Dim LOG As New StruStepLog : LOG.SubName = "GetPigProc"
        Try
            Dim bolIsExists As Boolean = False
            Me.IsProcExists(PID, bolIsExists)
            If bolIsExists = True Then
                GetPigProc = New PigProc(PID)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function IsProcExists(ProcName As String, ByRef IsExists As Boolean) As String
        Try
            If Process.GetProcessesByName(ProcName).Length > 0 Then
                IsExists = True
            Else
                IsExists = False
            End If
            Return "OK"
        Catch ex As Exception
            IsExists = False
            Return ex.Message.ToString
        End Try
    End Function

    Public Function IsProcExists(PID As Integer, ByRef IsExists As Boolean) As String
        Try
            Dim oProcess As Process = Process.GetProcessById(PID)
            If oProcess Is Nothing Then
                IsExists = False
            Else
                IsExists = True
            End If
            Return "OK"
        Catch ex As Exception
            IsExists = False
            Return ex.Message.ToString
        End Try
    End Function

    Public Function GetPigProcs(ProcName As String) As PigProcs
        Dim LOG As New StruStepLog : LOG.SubName = "GetPigProcs"
        Try
            LOG.StepName = "GetProcessesByName"
            Dim abProcess As Process() = Process.GetProcessesByName(ProcName)
            LOG.StepName = "New PigProcs"
            GetPigProcs = New PigProcs
            Dim bolIsExists As Boolean = False
            For Each oProcess As Process In abProcess
                Me.IsProcExists(oProcess.Id, bolIsExists)
                If bolIsExists = True Then
                    LOG.StepName = "Add"
                    GetPigProcs.Add(oProcess.Id)
                    If GetPigProcs.LastErr <> "" Then
                        LOG.Ret = GetPigProcs.LastErr
                        Throw New Exception(LOG.Ret)
                    End If
                End If
            Next
        Catch ex As Exception
            LOG.AddStepNameInf(ProcName)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function KillProc(PID As Long) As String
        Dim LOG As New StruStepLog : LOG.SubName = "KillProc"
        Try
            LOG.StepName = "GetPigProc"
            Dim oPigProc As PigProc = Me.GetPigProc(PID)
            If oPigProc Is Nothing Then Throw New Exception("Unable to get process")
            LOG.StepName = "Kill"
            LOG.Ret = oPigProc.Kill()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            oPigProc = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function KillProcs(ProcName As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "KillProcs"
        Try
            LOG.StepName = "GetPigProcs"
            Dim oPigProcs As PigProcs = Me.GetPigProcs(ProcName)
            If oPigProcs Is Nothing Then Throw New Exception("Unable to get processes")
            Dim strErr As String = ""
            LOG.StepName = "Kill"
            For Each oPigProc As PigProc In oPigProcs
                LOG.Ret = oPigProc.Kill
                If LOG.Ret <> "OK" Then strErr &= LOG.Ret
            Next
            If strErr <> "" Then Throw New Exception(strErr)
            oPigProcs = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function IsOtherExeExists(ExeName As String) As Boolean
        Dim LOG As New StruStepLog : LOG.SubName = "IsOtherExeExists"
        Try
            LOG.StepName = "GetPigProcs"
            Dim oPigProcs As PigProcs = Me.GetPigProcs(ExeName)
            If oPigProcs Is Nothing Then
                Throw New Exception("oPigProcs Is Nothing")
            ElseIf oPigProcs.LastErr <> "" Then
                Throw New Exception(oPigProcs.LastErr)
            Else
                LOG.StepName = "For Each"
                Dim intPID As Integer = Me.fMyPID
                IsOtherExeExists = False
                For Each oPigProc As PigProc In oPigProcs
                    If oPigProc.ProcessID <> intPID Then
                        IsOtherExeExists = True
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            LOG.AddStepNameInf(ExeName)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function KillOtherExe(ExeName As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "KillOtherExe"
        Try
            LOG.StepName = "GetPigProcs"
            Dim oPigProcs As PigProcs = Me.GetPigProcs(ExeName)
            If oPigProcs Is Nothing Then
                Throw New Exception("oPigProcs Is Nothing")
            ElseIf oPigProcs.LastErr <> "" Then
                Throw New Exception(oPigProcs.LastErr)
            Else
                LOG.StepName = "For Each"
                Dim intPID As Integer = Me.fMyPID
                Dim strError As String = ""
                For Each oPigProc As PigProc In oPigProcs
                    If oPigProc.ProcessID <> intPID Then
                        LOG.StepName = "Kill"
                        LOG.Ret = oPigProc.Kill()
                        If LOG.Ret <> "OK" Then strError &= LOG.Ret & ";"
                    End If
                Next
                If strError <> "" Then Throw New Exception(strError)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(ExeName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
