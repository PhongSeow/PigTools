﻿'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2.1
'* Create Time: 15/1/2022
'* 1.1    31/1/2022   Add CallFile
'* 1.1    1/3/2022   Add CmdShell
'************************************

Imports PigCmdLib

Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public PigCmdApp As New PigCmdApp
    Public PID As Long
    Public Cmd As String
    Public Para As String
    Public Line As String
    Public PigConsole As New PigConsole
    Public Pwd As String
    Public Ret As String
    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to HideShell")
            Console.WriteLine("Press B to CallFile")
            Console.WriteLine("Press C to CmdShell")
            Console.WriteLine("Press D to GetPwdStr")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("HideShell")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Cmd=" & Me.Cmd)
                    Console.CursorVisible = True
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.CmdOrFilePath = Me.Line
                    End If
                    Me.PID = PigCmdApp.HideShell(Me.CmdOrFilePath)
                    Console.WriteLine("LastErr=" & Me.PigCmdApp.LastErr)
                    Console.WriteLine("PID=" & Me.PID)
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("CallFile")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Console.WriteLine("FilePath=" & Me.CmdOrFilePath)
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.CmdOrFilePath = Me.Line
                    End If
                    Console.WriteLine("Para=" & Me.Para)
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.Para = Me.Line
                    End If
                    Me.Ret = PigCmdApp.CallFile(Me.CmdOrFilePath, Me.Para)
                    Console.WriteLine("CallFile=" & Me.Ret)
                    If Me.Ret = "OK" Then
                        Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                        Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                        Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                    End If
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("CmdShell")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Cmd=" & Me.Cmd)
                    Console.CursorVisible = True
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.Cmd = Me.Line
                    End If
                    Me.Ret = PigCmdApp.CmdShell(Me.Cmd)
                    Console.WriteLine("CmdShell=" & Me.Ret)
                    If Me.Ret = "OK" Then
                        Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                        Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                        Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                    End If
                Case ConsoleKey.D
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPwdStr")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Enter the password and press ENTER to end")
                    Console.CursorVisible = True
                    Me.Pwd = Me.PigConsole.GetPwdStr()
                    If Me.PigConsole.LastErr <> "" Then
                        Console.WriteLine(Me.PigConsole.LastErr)
                    Else
                        Console.WriteLine("Password is ")
                        Console.WriteLine(Me.Pwd)
                    End If
            End Select
            Console.CursorVisible = False
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class