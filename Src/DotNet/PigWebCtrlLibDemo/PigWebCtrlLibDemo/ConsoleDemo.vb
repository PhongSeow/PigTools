'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3.5
'* Create Time: 15/2/2022
'* 1.1    26/5/2022   Add Main
'* 1.2    30/5/2022   Add WebLogicDemo
'* 1.3    1/6/2022   Modify WebLogicDemo
'************************************

Imports PigCmdLib
Imports PigToolsLiteLib
Imports PigWebCtrlLib


Public Class ConsoleDemo
    Public WebLogicApp As WebLogicApp
    Public WebLogicDomain As WebLogicDomain
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
    Public HomeDirPath As String
    Public WorkTmpDirPath As String
    Public DomainHomeDirPath As String
    Public ListPort As Integer = 8888
    Public AdminUserName As String = "weblogic"
    Public AdminUserPassword As String = "PigWebCtrl8888"

    Public Sub New()
        If Me.PigFunc.IsOsWindows = True Then
            Me.HomeDirPath = "C:\WeblogicHome12c"
            Me.WorkTmpDirPath = "C:\Temp"
            Me.DomainHomeDirPath = "C:\Temp\WebLogic8888"
        Else
            Me.HomeDirPath = "/home/mw/weblogic/wls1221"
            Me.WorkTmpDirPath = "/tmp"
            Me.WorkTmpDirPath = "/tmp/WebLogic8888"
        End If
    End Sub

    Public Sub WebLogicDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "NewWebLogicApp#New WebLogicApp|"
            Me.MenuDefinition &= "AddNewWebLogicDomain#Add New WebLogicDomain|"
            Me.MenuDefinition &= "ShowParas#Display parameters|"
            Me.MenuDefinition &= "SetWebLogicDomain#Select WebLogicDomain|"
            Me.MenuDefinition &= "WebLogicDomainCreateDomain#WebLogicDomain.CreateDomain|"
            Me.MenuDefinition &= "WebLogicDomainSaveSecurityBoot#WebLogicDomain.SaveSecurityBoot|"
            Me.MenuDefinition &= "WebLogicDomainStartDomain#WebLogicDomain.StartDomain|"
            Me.MenuDefinition &= "WebLogicDomainStopDomain#WebLogicDomain.StopDomain|"
            Me.PigConsole.SimpleMenu("WebLogicDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "WebLogicDomainStopDomain"
                    Console.WriteLine("*******************")
                    Console.WriteLine("WebLogicDomain.StopDomain")
                    Console.WriteLine("*******************")
                    If Me.WebLogicDomain Is Nothing Then
                        Console.WriteLine("WebLogicDomain Is Nothing")
                    Else
                        Me.Ret = Me.WebLogicDomain.StopDomain
                        Console.WriteLine(Me.Ret)
                    End If
                Case "WebLogicDomainStartDomain"
                    Console.WriteLine("*******************")
                    Console.WriteLine("WebLogicDomain.StartDomain")
                    Console.WriteLine("*******************")
                    If Me.WebLogicDomain Is Nothing Then
                        Console.WriteLine("WebLogicDomain Is Nothing")
                    Else
                        Me.Ret = Me.WebLogicDomain.StartDomain
                        Console.WriteLine(Me.Ret)
                    End If
                Case "WebLogicDomainSaveSecurityBoot"
                    Console.WriteLine("*******************")
                    Console.WriteLine("WebLogicDomain.SaveSecurityBoot")
                    Console.WriteLine("*******************")
                    If Me.WebLogicDomain Is Nothing Then
                        Console.WriteLine("WebLogicDomain Is Nothing")
                    Else
                        Me.PigConsole.GetLine("AdminUserName", Me.AdminUserName)
                        Me.PigConsole.GetLine("AdminUserPassword", Me.AdminUserPassword)
                        Console.WriteLine("SaveSecurityBoot")
                        Me.Ret = Me.WebLogicDomain.SaveSecurityBoot(Me.AdminUserName, Me.AdminUserPassword)
                        Console.WriteLine(Me.Ret)
                    End If
                Case "WebLogicDomainCreateDomain"
                    Console.WriteLine("*******************")
                    Console.WriteLine("WebLogicDomain.CreateDomain")
                    Console.WriteLine("*******************")
                    If Me.WebLogicDomain Is Nothing Then
                        Console.WriteLine("WebLogicDomain Is Nothing")
                    Else
                        Me.PigConsole.GetLine("ListPort", Me.ListPort)
                        Me.PigConsole.GetLine("AdminUserName", Me.AdminUserName)
                        Me.PigConsole.GetLine("AdminUserPassword", Me.AdminUserPassword)
                        Console.WriteLine("CreateDomain")
                        Me.Ret = Me.WebLogicDomain.CreateDomain(Me.ListPort, Me.AdminUserName, Me.AdminUserPassword)
                        Console.WriteLine(Me.Ret)
                    End If
                Case "SetWebLogicDomain"
                    Console.WriteLine("*******************")
                    Console.WriteLine("Select WebLogicDomain")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input weblogic domain home folder path", Me.DomainHomeDirPath)
                    If Me.WebLogicApp.WebLogicDomains.IsItemExists(Me.DomainHomeDirPath) = False Then
                        Console.WriteLine("DomainHomeDirPath not found.")
                    Else
                        Me.WebLogicDomain = Me.WebLogicApp.WebLogicDomains.Item(Me.DomainHomeDirPath)
                        Console.WriteLine("Set WebLogicDomain to " & Me.DomainHomeDirPath)
                    End If
                Case "NewWebLogicApp"
                    Console.WriteLine("*******************")
                    Console.WriteLine("New WebLogicApp")
                    Console.WriteLine("*******************")

                    Me.PigConsole.GetLine("Input weblogic home folder path", Me.HomeDirPath)
                    Me.PigConsole.GetLine("Input work temporary folder path", Me.WorkTmpDirPath)
                    Me.WebLogicApp = New WebLogicApp(Me.HomeDirPath, Me.WorkTmpDirPath)
                    If Me.WebLogicApp.LastErr <> "" Then Console.WriteLine(Me.WebLogicApp.LastErr)
                    Console.WriteLine("WebLogicApp.GetJavaVersion")
                    Me.Ret = Me.WebLogicApp.GetJavaVersion()
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                Case "AddNewWebLogicDomain"
                    Console.WriteLine("*******************")
                    Console.WriteLine("Add New WebLogicDomain")
                    Console.WriteLine("*******************")
                    Me.PigConsole.GetLine("Input weblogic domain home folder path", Me.DomainHomeDirPath)
                    Me.WebLogicApp.WebLogicDomains.AddOrGet(Me.DomainHomeDirPath)
                Case "ShowParas"
                    Console.WriteLine("*******************")
                    Console.WriteLine("WebLogicApp.JavaVersion=" & Me.WebLogicApp.JavaVersion)
                    Console.WriteLine("WebLogicApp.WlstPath=" & Me.WebLogicApp.WlstPath)
                    For Each oWebLogicDomain As WebLogicDomain In Me.WebLogicApp.WebLogicDomains
                        Console.WriteLine("*******************")
                        Console.WriteLine("WebLogicDomain")
                        Console.WriteLine("*******************")
                        With oWebLogicDomain
                            Console.WriteLine("ConfPath=" & .ConfPath)
                            Console.WriteLine("ConsolePath=" & .ConsolePath)
                            Console.WriteLine("HomeDirPath=" & .HomeDirPath)
                            Console.WriteLine("LogDirPath=" & .LogDirPath)
                            Console.WriteLine("SecurityBootPath=" & .SecurityBootPath)
                            Console.WriteLine("SecurityDirPath=" & .SecurityDirPath)
                            Console.WriteLine("startWebLogicPath=" & .startWebLogicPath)
                            Console.WriteLine("stopWebLogicPath=" & .stopWebLogicPath)
                            Console.WriteLine("###############")
                            Console.WriteLine("RefConf")
                            Me.Ret = .RefConf()
                            Console.WriteLine(Me.Ret)
                            Console.WriteLine("RefDeployStatus")
                            Me.Ret = .RefDeployStatus()
                            Console.WriteLine(Me.Ret)
                            Console.WriteLine("RefRunStatus")
                            Me.Ret = .RefRunStatus()
                            Console.WriteLine(Me.Ret)
                            Console.WriteLine("DeployStatus=" & .DeployStatus.ToString)
                            Console.WriteLine("RunStatus=" & .RunStatus.ToString)
                            Console.WriteLine("###############")
                            Console.WriteLine("DomainName=" & .DomainName)
                            Console.WriteLine("DomainVersion=" & .DomainVersion)
                            Console.WriteLine("ListenPort=" & .ListenPort)
                            Console.WriteLine("IsAdminPortEnable=" & .IsAdminPortEnable)
                            Console.WriteLine("IsIIopEnable=" & .IsIIopEnable)
                            Console.WriteLine("IsProdMode=" & .IsProdMode)
                            Console.WriteLine("CreateDomainRes=" & .CreateDomainRes)
                            Console.WriteLine("StartDomainRes=" & .StartDomainRes)
                            Console.WriteLine("StopDomainRes=" & .StopDomainRes)
                        End With
                        Console.WriteLine("*******************")
                    Next

            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub


    Public Sub Main()
        Do While True
            Console.Clear()
            Me.MenuDefinition = "WebLogic#WebLogic|Tomcat#Tomcat|Nginx#nginx"
            Me.PigConsole.SimpleMenu("Main menu", Me.MenuDefinition, Me.MenuKey)
            Select Case Me.MenuKey
                Case "WebLogic"
                    Me.WebLogicDemo()
                Case "Tomcat", "Nginx"
                    Console.WriteLine("Under development")
                    Me.PigConsole.DisplayPause()
                Case ""
                    If Me.PigConsole.IsYesOrNo("Is exit now?") = True Then
                        Exit Do
                    End If
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class
