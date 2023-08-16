'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.6.2
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
'* 1.12   23/7/2022  Modify PigSysCmdDemo, add GetWmicSimpleXml
'* 1.13   17/11/2022  Add SelectControl demo
'* 2.1  29/12/2022  Modify PigSysCmdDemo, add GetBootUpTime
'* 2.2  17/1/2023  Add PigHostDemo
'* 2.3  22/5/2023  Modify PigCmdAppDemo
'* 2.5  23/5/2023  Add PigSudo
'* 2.6  5/6/2023  Modify PigCmdAppDemo
'************************************

Imports PigCmdLib
Imports PigToolsLiteLib


Public Class ConsoleDemo
    Public CmdOrFilePath As String
    Public CmdPara As String
    Public WithEvents PigCmdApp As New PigCmdApp
    Public PigSysCmd As New PigSysCmd
    Public PigHost As PigHost
    Public PID As String
    Public Cmd As String
    Public Para As String
    Public Line As String
    Public PigConsole As New PigConsole
    Public Pwd As String
    Public Ret As String
    Public PigFunc As New PigFunc
    Public MenuKey As String
    Public SelectKey As String
    Public MenuDefinition As String
    Public SelectDefinition As String
    Public OutThreadID As Integer
    Public ListenPort As Integer
    Public OSCaption As String
    Public WmicCmd As String
    Public UUID As String
    Public BootUpTime As Date
    Public SuDoUser As String
    Public OutFile As String
    Public IsBackRun As Boolean
    Public StandardOutputReadType As PigCmdApp.EnmStandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.FullString
    Public WithEvents PigSudo As PigSudo

    Public Sub PigCmdAppDemo()
        Me.PigCmdApp.SetDebug(Me.PigFunc.GetMyExePath & ".log")
        Do While True
            Console.Clear()
            Me.MenuDefinition = "HideShell#HideShell|"
            Me.MenuDefinition &= "CallFile#CallFile|"
            Me.MenuDefinition &= "CmdShell#CmdShell|"
            Me.MenuDefinition &= "GetParentProc#GetParentProc|"
            Me.MenuDefinition &= "CallFileWaitForExit#CallFileWaitForExit|"
            Me.PigConsole.SimpleMenu("PigCmdAppDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "CallFileWaitForExit"
                    Console.WriteLine("*******************")
                    Console.WriteLine("CallFileWaitForExit")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input exec file path", Me.CmdOrFilePath)
                    Console.WriteLine("CallFileWaitForExit" & "->" & Me.CmdOrFilePath)
                    Me.Ret = Me.PigCmdApp.CallFileWaitForExit(Me.CmdOrFilePath)
                    Console.WriteLine(Me.Ret)
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
                    If Me.PigConsole.IsYesOrNo("Is FullString") = True Then
                        Me.StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.FullString
                    Else
                        Me.StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.StringArray
                    End If
                    If Me.PigConsole.IsYesOrNo("Is asynchronous processing") = True Then
                        Me.Ret = PigCmdApp.AsyncCmdShell(Me.Cmd, Me.OutThreadID)
                        Console.WriteLine(vbCrLf & "OutThreadID=" & Me.OutThreadID)
                        Console.WriteLine("Delay(1000)")
                        Me.PigFunc.Delay(1000)
                    Else
                        Me.Ret = PigCmdApp.CmdShell(Me.Cmd, Me.StandardOutputReadType)
                        Console.WriteLine("CmdShell=" & Me.Ret)
                        If Me.Ret = "OK" Then
                            Console.WriteLine("PID=" & Me.PigCmdApp.PID)
                            If Me.StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.FullString Then
                                Console.WriteLine("StandardOutput=" & Me.PigCmdApp.StandardOutput)
                                Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                            Else
                                Console.WriteLine("StandardOutputArray.Count=" & Me.PigCmdApp.StandardOutputArray.Length)
                                For i = 0 To Me.PigCmdApp.StandardOutputArray.Length - 1
                                    Console.WriteLine("StandardOutputArray(" & i & ")=" & Me.PigCmdApp.StandardOutputArray(i))
                                Next
                                Console.WriteLine("StandardError=" & Me.PigCmdApp.StandardError)
                            End If
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
            Me.MenuDefinition &= "GetUUID#GetUUID|"
            Me.MenuDefinition &= "GetBootUpTime#GetBootUpTime|"
            Me.MenuDefinition &= "GetWmicSimpleXml#GetWmicSimpleXml|"
            Me.MenuDefinition &= "ReBootHost#ReBootHost|"
            Me.PigConsole.SimpleMenu("PigConsoleDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "ReBootHost"
                    If Me.PigConsole.IsYesOrNo("Is reboot host now?") = True Then
                        Me.Ret = Me.PigSysCmd.ReBootHost
                        Console.WriteLine(Me.Ret)
                    End If
                Case "GetUUID"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetUUID")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.Ret = Me.PigSysCmd.GetUUID(Me.UUID)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine(Me.UUID)
                    End If
                Case "GetBootUpTime"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetBootUpTime")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.Ret = Me.PigSysCmd.GetBootUpTime(Me.BootUpTime)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine(Me.BootUpTime)
                    End If
                Case "GetWmicSimpleXml"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetWmicSimpleXml")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input WmicCmd", Me.WmicCmd)
                    Dim strXml As String = ""
                    Me.Ret = Me.PigSysCmd.GetWmicSimpleXml(Me.WmicCmd, strXml)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine(strXml)
                    End If
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
            Me.MenuDefinition &= "PigSysCmdDemo#PigSysCmd Demo|"
            Me.MenuDefinition &= "PigHostDemo#PigHost Demo|"
            Me.MenuDefinition &= "PigSudo#PigSudo|"
            Me.MenuDefinition &= "PigNohup#PigNohup|"
            Me.PigConsole.SimpleMenu("Main menu", Me.MenuDefinition, Me.MenuKey)
            Select Case Me.MenuKey
                Case "PigSudo"
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigSudo")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input Command", Me.Cmd)
                    Me.PigConsole.GetLine("Input sudo user", Me.SuDoUser)
                    Me.IsBackRun = Me.PigConsole.IsYesOrNo("Is back run")
                    Me.PigSudo = New PigSudo(Me.Cmd, Me.SuDoUser, Me.IsBackRun)
                    If Me.PigConsole.IsYesOrNo("Is execute asynchronously") = True Then
                        Me.Ret = Me.PigSudo.AsyncRun()
                        Console.WriteLine(Me.Ret)
                    Else
                        Me.Ret = Me.PigSudo.Run()
                        Console.WriteLine(Me.Ret)
                        Console.WriteLine("StandardOutput=" & Me.PigSudo.StandardOutput)
                        Console.WriteLine("StandardError=" & Me.PigSudo.StandardError)
                    End If
                    Me.PigConsole.DisplayPause()
                Case "PigHostDemo"
                    Me.PigHostDemo()
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
        If Me.PigConsole.IsYesOrNo("Whether to use multiple languages") = True Then
            Console.WriteLine("Copy the multilingual file to the same directory of the class library, with the file name such as PigCmdLib.zh-CN. See https://github.com/PhongSeow/PigTools/tree/main/Release/PigMLang")
            Me.PigConsole.RefMLang()
            Me.PigConsole.DisplayPause()
        End If
        Do While True
            Console.Clear()
            Me.MenuDefinition = "GetPwdStr#GetPwdStr|"
            Me.MenuDefinition &= "GetLine#GetLine|"
            Me.MenuDefinition &= "SelectControl#SelectControl|"
            Me.PigConsole.SimpleMenu("PigConsoleDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "SelectControl"
                    Me.SelectDefinition = ""
                    Me.SelectDefinition &= "China#China|"
                    Me.SelectDefinition &= "USA#USA|"
                    Me.SelectDefinition &= "Japan#Japan|"
                    Me.SelectDefinition &= "Argentina#Argentina|"
                    Me.Ret = Me.PigConsole.SelectControl("Select country", Me.SelectDefinition, Me.SelectKey, True)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine("You chose " & Me.SelectKey)
                        Console.WriteLine(Console.WindowWidth)
                    End If
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

    Public Sub PigHostDemo()
        Me.PigHost = New PigHost
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "DisplayProperties#Display Properties|"
            Me.PigConsole.SimpleMenu("PigHostDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "DisplayProperties"
                    Console.WriteLine("*******************")
                    Console.WriteLine("Display Properties")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    With Me.PigHost
                        Console.WriteLine("HostID=" & .HostID)
                        Console.WriteLine("HostName=" & .HostName)
                        Console.WriteLine("UUID=" & .UUID)
                        Console.WriteLine("OSCaption=" & .OSCaption)
                        '--------
                        Console.WriteLine("CPU.Model=" & .CPU.Model)
                        Console.WriteLine("CPU.CPUs=" & .CPU.CPUs)
                        Console.WriteLine("CPU.CPUCores=" & .CPU.CPUCores)
                        Console.WriteLine("CPU.Processors=" & .CPU.Processors)
                    End With
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Private Sub PigSudo_AsyncRet_FullString(AsyncRet As PigAsync, StandardOutput As String, StandardError As String) Handles PigSudo.AsyncRet_FullString
        Console.WriteLine("PigSudo_AsyncRet_FullString")
        Console.WriteLine("AsyncRet.AsyncRet=" & AsyncRet.AsyncRet)
        Console.WriteLine("AsyncRet.AsyncThreadID=" & AsyncRet.AsyncThreadID)
        Console.WriteLine("StandardOutput=" & StandardOutput)
        Console.WriteLine("StandardError=" & StandardError)
    End Sub
End Class
