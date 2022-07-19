'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.11.2
'* Create Time: 15/1/2022
'* 1.1    31/1/2022   Add CallFile
'* 1.2    1/3/2022   Add CmdShell
'* 1.3    19/3/2022  Modify GetLine,GetPwdStr
'* 1.4    20/3/2022  Modify GetPwdStr
'* 1.5    1/4/2022  Add PigCmdAppDemo
'* 1.6    1/4/2022  Modify PigCmdAppDemo
'* 1.7    29/4/2022  Modify PigCmdAppDemo
'* 1.8    19/5/2022  Modify PigCmdAppDemo
'* 1.9    2/6/2022  Modify Main, add PigSysCmdDemo
'* 1.10   7/6/2022  Modify PigSysCmdDemo, add GetOSCaption
'* 1.11   17/6/2022  Modify PigSysCmdDemo, add GetProcListenPortList
'************************************

Imports PigCmdLib
Imports PigToolsLiteLib


Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public WithEvents PigCmdApp As New PigCmdApp
    Public PigSysCmd As New PigSysCmd
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
    Public OutThreadID As Integer
    Public ListenPort As Integer
    Public OSCaption As String

    Public Sub PigCmdAppDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "HideShell#HideShell|"
            Me.MenuDefinition &= "CallFile#CallFile|"
            Me.MenuDefinition &= "CmdShell#CmdShell|"
            Me.MenuDefinition &= "GetParentProc#GetParentProc|"
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
                    If Me.PigConsole.IsYesOrNo("Is asynchronous processing") = True Then
                        Me.Ret = PigCmdApp.AsyncCallFile(Me.CmdOrFilePath, Me.Para, Me.OutThreadID)
                        Console.WriteLine(vbCrLf & "OutThreadID=" & Me.OutThreadID)
                        Console.WriteLine("Delay(1000)")
                        Me.PigFunc.Delay(1000)
                    Else
                        Me.Ret = PigCmdApp.CallFile(Me.CmdOrFilePath, Me.Para)
                        Me.Ret = PigCmdApp.CmdShell(Me.Cmd)
                        Console.WriteLine("CallFile=" & Me.Ret)
                        If Me.Ret = "OK" Then
                            Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                            Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                            Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                        End If
                    End If
                Case "CmdShell"
                    Console.WriteLine("*******************")
                    Console.WriteLine("CmdShell")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input Command", Me.Cmd)
                    If Me.PigConsole.IsYesOrNo("Is asynchronous processing") = True Then
                        Me.Ret = PigCmdApp.AsyncCmdShell(Me.Cmd, Me.OutThreadID)
                        Console.WriteLine(vbCrLf & "OutThreadID=" & Me.OutThreadID)
                        Console.WriteLine("Delay(1000)")
                        Me.PigFunc.Delay(1000)
                    Else
                        Me.Ret = PigCmdApp.CmdShell(Me.Cmd)
                        Console.WriteLine("CmdShell=" & Me.Ret)
                        If Me.Ret = "OK" Then
                            Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                            Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                            Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                        End If
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
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigSysCmdDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "GetListenPortProcID#GetListenPortProcID|"
            Me.MenuDefinition &= "GetProcListenPortList#GetProcListenPortList|"
            Me.MenuDefinition &= "GetOSCaption#GetOSCaption|"
            Me.PigConsole.SimpleMenu("PigConsoleDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "GetProcListenPortList"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetProcListenPortList")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input PID", Me.PID)
                    Dim abPort(0) As Integer
                    Me.Ret = Me.PigSysCmd.GetProcListenPortList(Me.PID, abPort)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine("ListenPortList is ")
                        For i = 0 To abPort.Length - 1
                            Console.WriteLine(abPort(i))
                        Next
                    End If
                Case "GetOSCaption"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetOSCaption")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.Ret = Me.PigSysCmd.GetOSCaption(Me.OSCaption)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine("OSCaption=" & Me.OSCaption)
                    End If
                Case "GetListenPortProcID"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetListenPortProcID")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input ListenPort", Me.ListenPort)
                    Me.Ret = Me.PigSysCmd.GetListenPortProcID(Me.ListenPort, Me.PID)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine("PID=" & PID)
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub Main()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "PigCmdAppDemo#PigCmdApp Demo|"
            Me.MenuDefinition &= "PigConsoleDemo#PigConsole Demo|"
            Me.MenuDefinition &= "PigSysCmdDemo#PigSysCmdDemo|"
            Me.PigConsole.SimpleMenu("Main menu", Me.MenuDefinition, Me.MenuKey)
            Select Case Me.MenuKey
                Case "PigCmdAppDemo"
                    Me.PigCmdAppDemo()
                Case "PigConsoleDemo"
                    Me.PigConsoleDemo()
                Case "PigSysCmdDemo"
                    Me.PigSysCmdDemo()
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


    Private Sub PigCmdApp_AsyncRet_CmdShell_FullString(SyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles PigCmdApp.AsyncRet_CmdShell_FullString
        Console.WriteLine("PigCmdApp_AsyncRet_CmdShell_FullString")
        With SyncRet
            Console.WriteLine("BeginTime=" & .AsyncBeginTime)
            Console.WriteLine("EndTime=" & .AsyncEndTime)
            Console.WriteLine("Ret=" & .AsyncRet)
            Console.WriteLine("ThreadID=" & .AsyncThreadID)
            Console.WriteLine("AsyncCmdPID=" & .AsyncCmdPID)
        End With
        Console.WriteLine("StandardOutput=" & StandardOutput)
        Console.WriteLine("StandardError=" & StandardError)
    End Sub

    Private Sub PigCmdApp_AsyncRet_CallFile_FullString(SyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles PigCmdApp.AsyncRet_CallFile_FullString
        Console.WriteLine("PigCmdApp_AsyncRet_CallFile_FullString")
        With SyncRet
            Console.WriteLine("BeginTime=" & .AsyncBeginTime)
            Console.WriteLine("EndTime=" & .AsyncEndTime)
            Console.WriteLine("Ret=" & .AsyncRet)
            Console.WriteLine("ThreadID=" & .AsyncThreadID)
            Console.WriteLine("AsyncCmdPID=" & .AsyncCmdPID)
        End With
        Console.WriteLine("StandardOutput=" & StandardOutput)
        Console.WriteLine("StandardError=" & StandardError)
    End Sub

    Public Sub PigConsoleDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "GetPwdStr#GetPwdStr|"
            Me.MenuDefinition &= "GetLine#GetLine|"
            Me.PigConsole.SimpleMenu("PigConsoleDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "GetPwdStr"
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
                Case "GetLine"
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
            Me.PigConsole.DisplayPause()
        Loop
    End Sub


End Class
