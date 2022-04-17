'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6.2
'* Create Time: 15/1/2022
'* 1.1    31/1/2022   Add CallFile
'* 1.2    1/3/2022   Add CmdShell
'* 1.3    19/3/2022  Modify GetLine,GetPwdStr
'* 1.4    20/3/2022  Modify GetPwdStr
'* 1.5    1/4/2022  Add PigCmdAppDemo
'* 1.6    1/4/2022  Modify PigCmdAppDemo
'************************************

Imports PigCmdLib
Imports PigToolsLiteLib


Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public PigCmdApp As New PigCmdApp
    Public PID As String
    Public Cmd As String
    Public Para As String
    Public Line As String
    Public PigConsole As New PigConsole
    Public Pwd As String
    Public Ret As String
    Public PigFunc As New PigFunc
    Public MenuKey As String
    Public MenuDefinition As String

    Public Sub PigCmdAppDemo()
        Me.MenuDefinition = "HideShell#HideShell|"
        Me.MenuDefinition &= "CallFile#CallFile|"
        Me.MenuDefinition &= "CmdShell#CmdShell|"
        Me.MenuDefinition &= "GetParentProc#GetParentProc|"
        Do While True
            Me.PigConsole.SimpleMenu("PigCmdAppDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "HideShell"
                    Console.WriteLine("*******************")
                    Console.WriteLine("HideShell")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input Command", Me.CmdOrFilePath)
                    Me.PID = PigCmdApp.HideShell(Me.CmdOrFilePath)
                    Console.WriteLine("LastErr=" & Me.PigCmdApp.LastErr)
                    Console.WriteLine("PID=" & Me.PID)
                Case "CallFile"
                    Console.WriteLine("*******************")
                    Console.WriteLine("CallFile")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input FilePath", Me.CmdOrFilePath)
                    Me.PigConsole.GetLine("Input Para", Me.Para)
                    Me.Ret = PigCmdApp.CallFile(Me.CmdOrFilePath, Me.Para)
                    Console.WriteLine("CallFile=" & Me.Ret)
                    If Me.Ret = "OK" Then
                        Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                        Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                        Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                    End If
                Case "CmdShell"
                    Console.WriteLine("*******************")
                    Console.WriteLine("CmdShell")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input Command", Me.Cmd)
                    Me.Ret = PigCmdApp.CmdShell(Me.Cmd)
                    Console.WriteLine("CmdShell=" & Me.Ret)
                    If Me.Ret = "OK" Then
                        Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                        Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                        Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                    End If
                Case "GetParentProc"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetParentProc")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Cmd=" & Me.Cmd)
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input PID", Me.PID)
                    If IsNumeric(Me.PID) = False Then
                        Console.WriteLine("PID is not Numeric")
                    Else
                        Dim oParentPigProc As PigProc = PigCmdApp.GetParentProc(CInt(Me.PID))
                        If PigCmdApp.LastErr <> "" Then
                            Console.WriteLine(PigCmdApp.LastErr)
                        Else
                            Console.WriteLine("ProcessID=" & oParentPigProc.ProcessID)
                            Console.WriteLine("ProcessName=" & oParentPigProc.ProcessName)
                            Console.WriteLine("FilePath=" & oParentPigProc.FilePath)
                            Console.WriteLine("TotalProcessorTime=" & oParentPigProc.TotalProcessorTime.ToString)
                            Console.WriteLine("UserProcessorTime=" & oParentPigProc.UserProcessorTime.ToString)
                            Console.WriteLine("MemoryUse=" & Me.PigFunc.GetHumanSize(oParentPigProc.MemoryUse))
                            Console.WriteLine("StartTime=" & oParentPigProc.StartTime.ToString)
                        End If
                    End If
            End Select
        Loop
        'Do While True
        '    Console.WriteLine("*******************")
        '    Console.WriteLine("PigCmdAppDemo")
        '    Console.WriteLine("*******************")
        '    Console.WriteLine("Press Q to Up")
        '    Console.WriteLine("Press A to HideShell")
        '    Console.WriteLine("Press B to CallFile")
        '    Console.WriteLine("Press C to CmdShell")
        '    Console.WriteLine("Press D to GetParentProc")
        '    Console.WriteLine("*******************")
        '    Console.CursorVisible = False
        '    Select Case Console.ReadKey(True).Key
        '        Case ConsoleKey.Q
        '            Exit Do
        '        Case ConsoleKey.A
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("HideShell")
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("Cmd=" & Me.Cmd)
        '            Console.CursorVisible = True
        '            Me.Line = Console.ReadLine
        '            If Me.Line <> "" Then
        '                Me.CmdOrFilePath = Me.Line
        '            End If
        '            Me.PID = PigCmdApp.HideShell(Me.CmdOrFilePath)
        '            Console.WriteLine("LastErr=" & Me.PigCmdApp.LastErr)
        '            Console.WriteLine("PID=" & Me.PID)
        '        Case ConsoleKey.B
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("CallFile")
        '            Console.WriteLine("*******************")
        '            Console.CursorVisible = True
        '            Console.WriteLine("FilePath=" & Me.CmdOrFilePath)
        '            Me.Line = Console.ReadLine
        '            If Me.Line <> "" Then
        '                Me.CmdOrFilePath = Me.Line
        '            End If
        '            Console.WriteLine("Para=" & Me.Para)
        '            Me.Line = Console.ReadLine
        '            If Me.Line <> "" Then
        '                Me.Para = Me.Line
        '            End If
        '            Me.Ret = PigCmdApp.CallFile(Me.CmdOrFilePath, Me.Para)
        '            Console.WriteLine("CallFile=" & Me.Ret)
        '            If Me.Ret = "OK" Then
        '                Console.WriteLine("PID=" & Me.PigCmdApp.PID)
        '                Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
        '                Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
        '            End If
        '        Case ConsoleKey.C
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("CmdShell")
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("Cmd=" & Me.Cmd)
        '            Console.CursorVisible = True
        '            Me.Line = Console.ReadLine
        '            If Me.Line <> "" Then
        '                Me.Cmd = Me.Line
        '            End If
        '            Me.Ret = PigCmdApp.CmdShell(Me.Cmd)
        '            Console.WriteLine("CmdShell=" & Me.Ret)
        '            If Me.Ret = "OK" Then
        '                Console.WriteLine("PID=" & Me.PigCmdApp.PID)
        '                Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
        '                Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
        '            End If
        '        Case ConsoleKey.D
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("GetParentProc")
        '            Console.WriteLine("*******************")
        '            Console.WriteLine("Cmd=" & Me.Cmd)
        '            Console.CursorVisible = True
        '            Me.PigConsole.GetLine("Input PID", Me.PID)
        '            If IsNumeric(Me.PID) = False Then
        '                Console.WriteLine("PID is not Numeric")
        '            Else
        '                Dim oParentPigProc As PigProc = PigCmdApp.GetParentProc(CInt(Me.PID))
        '                If PigCmdApp.LastErr <> "" Then
        '                    Console.WriteLine(PigCmdApp.LastErr)
        '                Else
        '                    Console.WriteLine("ProcessID=" & oParentPigProc.ProcessID)
        '                    Console.WriteLine("ProcessName=" & oParentPigProc.ProcessName)
        '                    Console.WriteLine("FilePath=" & oParentPigProc.FilePath)
        '                    Console.WriteLine("TotalProcessorTime=" & oParentPigProc.TotalProcessorTime.ToString)
        '                    Console.WriteLine("UserProcessorTime=" & oParentPigProc.UserProcessorTime.ToString)
        '                    Console.WriteLine("MemoryUse=" & Me.PigFunc.GetHumanSize(oParentPigProc.MemoryUse))
        '                    Console.WriteLine("StartTime=" & oParentPigProc.StartTime.ToString)
        '                End If
        '            End If
        '    End Select
        '    Console.CursorVisible = False
        'Loop
    End Sub

    Public Sub PigConsoleDemo()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("PigCmdAppDemo")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Up")
            Console.WriteLine("Press A to GetPwdStr")
            Console.WriteLine("Press B to GetLine")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetPwdStr")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.Pwd = Me.PigConsole.GetPwdStr("Enter the password and press ENTER to end")
                    If Me.PigConsole.LastErr <> "" Then
                        Console.WriteLine(Me.PigConsole.LastErr)
                    Else
                        Console.WriteLine("Password is ")
                        Console.WriteLine(Me.Pwd)
                    End If
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetLine")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.Ret = Me.PigConsole.GetLine("Enter the Line", Me.Line, True)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.Write("Line is :")
                        Console.WriteLine(Me.Line)
                    End If
            End Select
            Console.CursorVisible = False
        Loop
    End Sub

    Public Sub Main()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "1#PigCmdAppDemo|2#PigConsoleDemo"
            Me.PigConsole.SimpleMenu("Main menu", Me.MenuDefinition, Me.MenuKey)
            Select Case Me.MenuKey
                Case "1"
                    Me.PigCmdAppDemo()
                Case "2"
                    Me.PigConsoleDemo()
                Case ""
                    If Me.PigConsole.IsYesOrNo("Is exit no?") = True Then
                        Exit Do
                    End If
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
