'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 2.9.6
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
'* 2.7  16/9/2023  Add PigService
'* 2.8  20/10/2023  Modify PigSysCmdDemo
'* 2.9  16/11/2023  Add CmdZipDemo
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
    Public CmdZip As CmdZip
    Public ZipExecPath As String
    Public ZipFilePath As String
    Public ZipType As CmdZip.EnmZipType
    Public SrcDir As String
    Public TargetDir As String
    Public TarFilePath As String
    Public Pwd As String
    Public Ret As String
    Public PigFunc As New PigFunc
    Public MenuKey As String
    Public MenuKey2 As String
    Public SelectKey As String
    Public SelectKey2 As String
    Public MenuDefinition As String
    Public MenuDefinition2 As String
    Public SelectDefinition As String
    Public OutThreadID As Integer
    Public ListenPort As Integer
    Public OSCaption As String
    Public WmicCmd As String
    Public UUID As String
    Public DefaultIPGateway
    Public BootUpTime As Date
    Public SuDoUser As String
    Public OutFile As String
    Public IsBackRun As Boolean
    Public StandardOutputReadType As PigCmdApp.EnmStandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.FullString
    Public WithEvents PigSudo As PigSudo
    Public PigService As PigService
    Public ServiceName As String
    Public DisplayName As String
    Public PathName As String
    Public StartUser As String
    Public StartUserPwd As String
    Public StartMode As PigService.EnmStartMode

    Public Sub PigCmdAppDemo()
        Me.PigCmdApp = New PigCmdApp
        Me.PigCmdApp.SetDebug(Me.PigFunc.GetMyExePath & ".log")
        Do While True
            Console.Clear()
            Me.MenuDefinition = "HideShell#HideShell|"
            Me.MenuDefinition &= "CallFile#CallFile|"
            Me.MenuDefinition &= "CmdShell#CmdShell|"
            Me.MenuDefinition &= "GetParentProc#GetParentProc|"
            Me.MenuDefinition &= "GetSubProcs#GetSubProcs|"
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
                    Me.PigConsole.GetLine("Input command parameters", Me.CmdPara)
                    Dim bolIsRunAsAdmin As Boolean
                    bolIsRunAsAdmin = Me.PigConsole.IsYesOrNo("Is run at Admin or root")
                    If Me.CmdPara = "" Then
                        Me.Ret = Me.PigCmdApp.CallFileWaitForExit(Me.CmdOrFilePath, bolIsRunAsAdmin)
                    Else
                        Me.Ret = Me.PigCmdApp.CallFileWaitForExit(Me.CmdOrFilePath, Me.CmdPara, bolIsRunAsAdmin)
                    End If
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
                Case "GetSubProcs"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetSubProcs")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input PID", Me.PID)
                    If IsNumeric(Me.PID) = False Then
                        Console.WriteLine("PID is not Numeric")
                    Else
                        Dim oSubPigProcs As PigProcs = PigCmdApp.GetSubProcs(Me.PID)
                        If PigCmdApp.LastErr <> "" Then
                            Console.WriteLine(PigCmdApp.LastErr)
                        Else
                            For Each oPigProc As PigProc In oSubPigProcs
                                With oPigProc
                                    Console.WriteLine("ProcessID=" & .ProcessID)
                                    Console.WriteLine("ProcessName=" & .ProcessName)
                                    Console.WriteLine("FilePath=" & .FilePath)
                                    Console.WriteLine("TotalProcessorTime=" & .TotalProcessorTime.ToString)
                                    Console.WriteLine("UserProcessorTime=" & .UserProcessorTime.ToString)
                                    Console.WriteLine("MemoryUse=" & Me.PigFunc.GetHumanSize(.MemoryUse))
                                    Console.WriteLine("StartTime=" & .StartTime.ToString)
                                End With
                            Next
                        End If
                    End If
                Case "GetParentProc"
                    Console.WriteLine("*******************")
                    Console.WriteLine("GetParentProc")
                    Console.WriteLine("*******************")
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
            Me.MenuDefinition &= "GetDefaultIPGateway#GetDefaultIPGateway|"
            Me.PigConsole.SimpleMenu("PigConsoleDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "GetDefaultIPGateway"
                    Console.WriteLine("GetDefaultIPGateway")
                    Me.Ret = Me.PigSysCmd.GetDefaultIPGateway(Me.DefaultIPGateway)
                    Console.WriteLine(Me.Ret)
                    Console.WriteLine("DefaultIPGateway=" & Me.DefaultIPGateway)
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
            Me.MenuDefinition &= "PigService#PigService|"
            Me.MenuDefinition &= "CmdZip#CmdZip|"
            Me.MenuDefinition &= "TestAllMenu#TestAllMenu|"
            Me.PigConsole.SimpleMenu("Main menu", Me.MenuDefinition, Me.MenuKey)
            Select Case Me.MenuKey
                Case "TestAllMenu"
                    Me.TestAllMenu()
                Case "CmdZip"
                    Me.CmdZipDemo()
                Case "PigService"
                    Me.PigSrvcieDemo()
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

    Public Sub CmdZipDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "New#New|"
            Me.MenuDefinition &= "AddArchive#AddArchive|"
            Me.MenuDefinition &= "ExtractArchive#ExtractArchive|"
            Me.PigConsole.SimpleMenu("CmdZipDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "New"
                    Dim intZipType As CmdZip.EnmZipType = CmdZip.EnmZipType.Tar
                    Me.MenuDefinition2 = intZipType & "#" & intZipType.ToString & "|"
                    intZipType = CmdZip.EnmZipType.WinRar : Me.MenuDefinition2 &= intZipType & "#" & intZipType.ToString & "|"
                    intZipType = CmdZip.EnmZipType._7_Zip : Me.MenuDefinition2 &= intZipType & "#" & intZipType.ToString & "|"
                    Me.PigConsole.SimpleMenu("Select ZipType", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.Null)
                    Me.ZipType = CInt(Me.MenuKey2)
                    Me.PigConsole.GetLine("Input ZipExecPath", Me.ZipExecPath)
                    Me.CmdZip = New CmdZip(Me.ZipType, Me.ZipExecPath)
                    If Me.CmdZip.LastErr <> "" Then Console.WriteLine(Me.CmdZip.LastErr)
                Case "AddArchive"
                    Me.PigConsole.GetLine("Enter the source directory to compress", Me.SrcDir)
                    Me.PigConsole.GetLine("Enter compressed package path", Me.ZipFilePath)
                    If Me.PigConsole.IsYesOrNo("Is use sudo") = True Then
                        Me.PigConsole.GetLine("Input sudo user", Me.SuDoUser)
                        Me.Ret = Me.CmdZip.AddArchive(Me.SrcDir, Me.ZipFilePath, Me.SuDoUser)
                    Else
                        Me.Ret = Me.CmdZip.AddArchive(Me.SrcDir, Me.ZipFilePath)
                    End If
                    Console.WriteLine(Me.Ret)
                Case "ExtractArchive"
                    Me.PigConsole.GetLine("Enter the target directory for decompression", Me.TargetDir)
                    Me.PigConsole.GetLine("Enter compressed package path", Me.ZipFilePath)
                    Me.Ret = Me.CmdZip.ExtractArchive(Me.TargetDir, Me.ZipFilePath)
                    Console.WriteLine(Me.Ret)
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub


    Public Sub PigSrvcieDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "New#New|"
            Me.MenuDefinition &= "Refresh#Refresh and display|"
            Me.MenuDefinition &= "Create#Create|"
            Me.MenuDefinition &= "Delete#Delete|"
            Me.MenuDefinition &= "StartService#StartService|"
            Me.MenuDefinition &= "StopService#StopService|"
            Me.PigConsole.SimpleMenu("PigSrvcieDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "New"
                    Me.PigConsole.GetLine("Input Service Name", Me.ServiceName)
                    Me.PigService = New PigService(Me.ServiceName)
                    If Me.PigService.LastErr <> "" Then Console.WriteLine(Me.PigService.LastErr)
                Case "StopService"
                    Me.Ret = Me.PigService.StopService
                    Console.WriteLine(Me.Ret)
                Case "StartService"
                    Me.Ret = Me.PigService.StartService
                    Console.WriteLine(Me.Ret)
                Case "Delete"
                    Me.Ret = Me.PigService.Delete
                    Console.WriteLine(Me.Ret)
                Case "Create"
                    Me.PigConsole.GetLine("Input DisplayName", Me.DisplayName)
                    Me.PigConsole.GetLine("Input PathName", Me.PathName)
                    Dim intStartMode As PigService.EnmStartMode
                    Me.SelectDefinition = ""
                    intStartMode = PigService.EnmStartMode.Automatic : Me.SelectDefinition &= intStartMode & "#" & intStartMode.ToString & "|"
                    intStartMode = PigService.EnmStartMode.Manual : Me.SelectDefinition &= intStartMode & "#" & intStartMode.ToString & "|"
                    intStartMode = PigService.EnmStartMode.Disabled : Me.SelectDefinition &= intStartMode & "#" & intStartMode.ToString & "|"
                    Me.PigConsole.SelectControl("Select StartMode", Me.SelectDefinition, Me.SelectKey, True)
                    intStartMode = CInt(Me.SelectKey)
                    If Me.PigConsole.IsYesOrNo("Is set StartUser") = True Then
                        Me.PigConsole.GetLine("Input StartUser", Me.StartUser)
                        Me.StartUserPwd = Me.PigConsole.GetPwdStr("Input StartUser password")
                        Me.Ret = Me.PigService.Create(Me.DisplayName, Me.PathName, intStartMode, Me.StartUser, Me.StartUserPwd)
                    Else
                        Me.Ret = Me.PigService.Create(Me.DisplayName, Me.PathName, intStartMode)
                    End If
                    Console.WriteLine(Me.Ret)
                Case "Refresh"
                    If Me.PigService Is Nothing Then
                        Console.WriteLine("PigService not new")
                    Else
                        Console.WriteLine("Refresh")
                        Me.Ret = Me.PigService.Refresh()
                        Console.WriteLine(Me.Ret)
                        If Me.Ret = "OK" Then
                            With Me.PigService
                                Console.WriteLine("ServiceName=" & .ServiceName)
                                Console.WriteLine("DisplayName=" & .DisplayName)
                                Console.WriteLine("PathName=" & .PathName)
                                Console.WriteLine("StartMode=" & .StartMode.ToString)
                                Console.WriteLine("Description=" & .Description)
                                Console.WriteLine("StartUser=" & .StartUser)
                                Console.WriteLine("ServiceState=" & .ServiceState.ToString)
                                Console.WriteLine("ProcessId=" & .ProcessId)
                            End With
                        End If
                    End If
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

    Public Sub TestAllMenu()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            For i = 1 To 25
                Me.MenuDefinition &= "Menu" & i & "#Menu" & i & "|"
            Next
            Me.PigConsole.SimpleMenu("TestAllMenu", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case Else
                    Console.WriteLine("You select " & Me.MenuKey)
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
            Me.MenuDefinition &= "DisplayActInf#Display Activities|"
            Me.PigConsole.SimpleMenu("PigHostDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "DisplayActInf"
                    Console.WriteLine("*******************")
                    Console.WriteLine("Display Activities")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    With Me.PigHost
                        '--------
                        Me.Ret = Me.PigHost.CPU.RefCpuActInf()
                        If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                        Console.WriteLine("CPU.HostCpuUseRate=" & .CPU.HostCpuUseRate)
                    End With
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
