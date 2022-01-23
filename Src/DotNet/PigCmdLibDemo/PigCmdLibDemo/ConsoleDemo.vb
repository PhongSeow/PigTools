'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 15/1/2022
'* 1.1    23/1/2022   Add GetPwdStr
'************************************

Imports PigCmdLib

Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public PigCmdApp As New PigCmdApp
    Public PID As Long
    Public Line As String
    Public PigConsole As New PigConsole
    Public Pwd As String
    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to HideShell")
            Console.WriteLine("Press B to GetPwdStr")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("HideShell")
                    Console.WriteLine("*******************")
                    Console.WriteLine("CmdFilePath=" & Me.CmdOrFilePath)
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
