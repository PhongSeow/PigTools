'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: ConsoleDemo for PigSQLSrv
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.28.1
'* Create Time: 17/4/2021
'* 1.2	23/9/2021	Add Test Cache Query
'* 1.3	5/10/2021	Imports PigKeyCacheLib
'* 1.4	8/10/2021	Add Test Cache Query -> CmdSQLSrvSp
'* 1.5	9/10/2021	Add Test Cache Query -> Print 
'* 1.6	5/12/2021	Add Test Cache Query -> Print 
'* 1.7	15/12/2021	Test the new class library
'* 1.8	23/1/2022	Refer to PigConsole.Getpwdstr of PigCmdLib  is used to hide the entered password.
'* 1.9	2/2/2022	Add Database connection management
'* 1.10	19/3/2022	Use PigCmdLib.GetLine
'* 1.11	23/3/2022	Modify MainSet
'* 1.13	29/4/2022	Modify MainSet,Main
'* 1.14	30/4/2022	Modify MainSet
'* 1.15	1/5/2022	Modify MainSet
'* 1.16	9/6/2022	Add SQLSrvToolsDemo
'* 1.17	23/6/2022	Modify SQLSrvToolsDemo
'* 1.18	24/6/2022	Modify SQLSrvToolsDemo
'* 1.19	27/7/2022	Modify Imports
'* 1.20	29/7/2022	Modify Imports
'* 1.21	30/7/2022	Modify SQLSrvToolsDemo
'* 1.22	28/1/2023	Modify SQLSrvToolsDemo
'* 1.23	1/2/2023	Modify Imports
'* 1.25	6/3/2023	Add HostDemo
'* 1.26	5/4/2023	Modify Main,HostDemo
'* 1.27	17/4/2023	Modify MainFunc,HostDemo
'**********************************
Imports System.Data
'--------
'#If NETFRAMEWORK Then
'Imports PigSQLSrvLib
'Imports System.Data.SqlClient
'Imports PigHostLib
'#Else
'Imports PigSQLSrvCoreLib
'Imports Microsoft.Data.SqlClient
'Imports PigHostCoreLib
'#End If
'--------
Imports PigToolsLiteLib
'--.------
Imports PigSQLSrvLib
Imports System.Data.SqlClient
Imports PigHostLib
Imports System.Data.SqlTypes
'--------

Public Class ConsoleDemo
    Public ConnSQLSrv As ConnSQLSrv
    Public CmdSQLSrvSp As CmdSQLSrvSp
    Public CmdSQLSrvText As CmdSQLSrvText
    Public ConnStr As String
    Public SQL As String
    Public SpName As String
    Public RS As Recordset
    Public RS2 As Recordset
    Public DBSrv As String = "localhost"
    Public MirrDBSrv As String = "localhost"
    Public DBUser As String = "sa"
    Public DBPwd As String = ""
    Public CurrDB As String = "master"
    Public CurrConsoleKey As ConsoleKey
    Public InpStr As String
    Public AccessFilePath As String
    Public TableName As String
    Public ColName As String
    Public PigConsole As New PigConsole
    Public EncKey As String
    Public PigFunc As New PigFunc
    Public ConfFilePath As String = Me.PigFunc.GetMyExePath & ".conf"
    Public Ret As String
    Public DBConnName As String
    Public MenuKey As String
    Public MenuKey2 As String
    Public MenuDefinition As String
    Public MenuDefinition2 As String
    Public SQLSrvTools As SQLSrvTools
    Public VBCodeOrSQLFragment As String
    Public NotMathFillByRsList As String
    Public NotMathMD5List As String
    Public NotMathColList As String
    Public FilePath As String
    Public WhatFragment As SQLSrvTools.EnmWhatFragment
    Public PigHostApp As PigHostApp
    Public HostID As String
    Public FolderPath As String
    Public CurrHostFolder As HostFolder
    Public CurrHost As Host
    Public HostFolderID As String
    Public FolderType As HostFolder.EnmFolderType

    Public Sub MainFunc()
        Do While True
            Console.Clear()
            Console.WriteLine("*******************")
            Console.WriteLine("Main Function menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Up")
            Console.WriteLine("Press A to Set SQL Server Connection String")
            Console.WriteLine("Press B to OpenOrKeepActive Connection")
            'Console.WriteLine("Press C to Show Host Information")
            Console.CursorVisible = False
            Console.WriteLine("*******************")
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("Set Connection String")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    If Me.PigFunc.IsFileExists(Me.ConfFilePath) = True Then
                        Console.WriteLine("Load ConfFilePath:" & Me.ConfFilePath)
                        Dim strXml As String = ""
                        Me.Ret = Me.PigFunc.GetFileText(Me.ConfFilePath, strXml)
                        Console.WriteLine(Me.Ret)
                        If Me.Ret = "OK" Then
                            Dim oPigXml As New PigXml(False)
                            oPigXml.SetMainXml(strXml)
                            Me.DBSrv = oPigXml.XmlGetStr("DBSrv")
                            Me.CurrDB = oPigXml.XmlGetStr("CurrDB")
                            Me.DBUser = oPigXml.XmlGetStr("DBUser")
                            Me.DBPwd = oPigXml.XmlGetStr("DBPwd")
                        End If
                    End If
                    If Me.PigConsole.IsYesOrNo("Do you want to manually set up the database connection?") = True Then
                        Me.PigConsole.GetLine("Input SQL Server", Me.DBSrv)
                        Console.WriteLine("SQL Server=" & Me.DBSrv)
                        Console.WriteLine("Input Default DB:" & Me.CurrDB)
                        Me.CurrDB = Console.ReadLine()
                        If Me.CurrDB = "" Then Me.CurrDB = "master"
                        Console.WriteLine("Default DB=" & Me.CurrDB)
                        Console.WriteLine("Is Trusted Connection ? (Y/n)")
                        Me.InpStr = Console.ReadLine()
                        Select Case Me.InpStr
                            Case "Y", "y", ""
                                Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB)
                            Case Else
                                Console.WriteLine("Input DB User:" & Me.DBUser)
                                Me.DBUser = Console.ReadLine()
                                If Me.DBUser = "" Then Me.DBUser = "sa"
                                Console.WriteLine("DB User=" & Me.DBUser)
                                Console.WriteLine("Input DB Password:")
                                Me.DBPwd = Me.PigConsole.GetPwdStr
                        End Select
                        Dim oPigXml As New PigXml(False)
                        With oPigXml
                            .AddEle("DBSrv", Me.DBSrv)
                            .AddEle("CurrDB", Me.CurrDB)
                            .AddEle("DBUser", Me.DBUser)
                            .AddEle("DBPwd", Me.DBPwd)
                        End With
                        Me.PigFunc.SaveTextToFile(Me.ConfFilePath, oPigXml.MainXmlStr)
                    End If
                    Me.ConnSQLSrv = New ConnSQLSrv(Me.DBSrv, Me.CurrDB, Me.DBUser, Me.DBPwd)
                    Me.ConnSQLSrv.ConnectionTimeout = 5
                Case ConsoleKey.B
                    Console.WriteLine("#################")
                    Console.WriteLine("OpenOrKeepActive Connection")
                    Console.WriteLine("#################")
                    If Me.ConnSQLSrv Is Nothing Then
                        Console.WriteLine("ConnSQLSrv Is Nothing")
                    Else
                        With Me.ConnSQLSrv
                            Console.WriteLine("OpenOrKeepActive:")
                            .OpenOrKeepActive()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                            End If
                        End With
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub




    Public Sub Main()
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "MainFunc#Main Function|"
            Me.MenuDefinition &= "HostDemo#HostFile Demo|"
            Me.PigConsole.SimpleMenu("Main Menu", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoExit)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "HostDemo"
                    Me.HostDemo()
                Case "MainFunc"
                    Me.MainFunc()
            End Select
        Loop
    End Sub





    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub HostDemo()
        Do While True
            Dim intHitCache As ConnSQLSrv.HitCacheEnum = ConnSQLSrv.HitCacheEnum.Null
            Console.Clear()
            Me.MenuDefinition = ""
            Me.MenuDefinition &= "New#New|"
            Me.MenuDefinition &= "RefHosts#Refresh Hosts|"
            Me.MenuDefinition &= "ShowMyHost#Show My Host|"
            Me.MenuDefinition &= "SetCurrHost#Set current Host|"
            Me.MenuDefinition &= "RefHostFolders#Refresh Host Folders|"
            Me.MenuDefinition &= "AddNewHostFolder#Add New HostFolder|"
            Me.MenuDefinition &= "SetCurrHostFolder#Set current HostFolder|"
            Me.MenuDefinition &= "HostFolderBeginScan#HostFolder.BeginScan|"
            Me.PigConsole.SimpleMenu("HostDemo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "ShowMyHost"
                    If Me.PigHostApp Is Nothing Then
                        Console.WriteLine("PigHostApp Is Nothing")
                    Else
                        Me.ShowHost(Me.PigHostApp.MyHost)
                        'Console.WriteLine("RefMyHost...")
                        'Me.Ret = Me.PigHostApp.RefMyHost
                        'Console.WriteLine(Me.Ret)
                        'If Me.Ret = "OK" Then Me.ShowHost(Me.PigHostApp.MyHost)
                    End If
                Case "HostFolderBeginScan"
                    If Me.CurrHostFolder Is Nothing Then
                        Console.WriteLine("CurrHostFolder Is Nothing")
                    Else
                        Console.WriteLine("BeginScan...")
                        Me.Ret = Me.CurrHostFolder.BeginScan
                        Console.WriteLine(Me.Ret)
                        Console.WriteLine(Me.CurrHostFolder.ScanStatus.ToString)
                    End If
                Case "SetCurrHost"
                    Me.PigHostApp.RefHosts()
                    Me.PigConsole.GetLine("Input HostID", Me.HostID)
                    If Me.PigHostApp.Hosts.IsItemExists(Me.HostID) = False Then
                        Console.WriteLine("Invalid HostID")
                    Else
                        Me.CurrHost = Me.PigHostApp.Hosts.Item(Me.HostID)
                        Me.ShowHost(Me.CurrHost)
                    End If
                Case "SetCurrHostFolder"
                    If Me.CurrHost Is Nothing Then
                        Console.WriteLine("CurrHost Is Nothing")
                    Else
                        Me.CurrHost.RefHostFolders()
                        Me.PigConsole.GetLine("Input HostFolderID", Me.HostFolderID)
                        If Me.CurrHost.HostFolders.IsItemExists(Me.HostFolderID) = False Then
                            Console.WriteLine("Invalid HostFolderID")
                        Else
                            Me.CurrHostFolder = Me.CurrHost.HostFolders.Item(Me.HostFolderID)
                            Me.ShowHostFolder(Me.CurrHostFolder)
                        End If
                    End If
                Case "RefHosts"
                    If Me.PigHostApp Is Nothing Then
                        Console.WriteLine("PigHostApp Is Nothing")
                    Else
                        Console.WriteLine("Refresh RefHosts")
                        Me.Ret = Me.PigHostApp.RefHosts
                        Console.WriteLine(Me.Ret)
                        For Each oHost As Host In Me.PigHostApp.Hosts
                            Me.ShowHost(oHost)
                        Next
                    End If
                Case "RefHostFolders"
                    If Me.CurrHost Is Nothing Then
                        Console.WriteLine("CurrHost Is Nothing")
                    Else
                        Console.WriteLine("Refresh RefHosts")
                        Me.Ret = Me.CurrHost.RefHostFolders()
                        Console.WriteLine(Me.Ret)
                        For Each oHostFolder As HostFolder In Me.CurrHost.HostFolders
                            Me.ShowHostFolder(oHostFolder)
                        Next
                    End If
                Case "AddNewHostFolder"
                    If Me.CurrHost Is Nothing Then
                        Console.WriteLine("CurrHost Is Nothing")
                    Else
                        Me.PigConsole.GetLine("Input FolderPath", Me.FolderPath)
                        Dim bolIsLocalPath As Boolean = Me.PigConsole.IsYesOrNo("IsLocalPath")
                        Console.WriteLine("AddNewHostFolder")
                        Me.Ret = Me.PigHostApp.AddNewHostFolder(Me.CurrHost, Me.FolderPath, bolIsLocalPath)
                        Console.WriteLine(Me.Ret)
                    End If
                Case "New"
                    If Me.ConnSQLSrv Is Nothing Then
                        Console.WriteLine("ConnSQLSrv Is Nothing")
                    Else
                        Console.WriteLine("New PigHostApp")
                        Me.PigHostApp = New PigHostApp(Me.ConnSQLSrv)
                        If Me.PigHostApp.LastErr <> "" Then
                            Console.WriteLine(Me.PigHostApp.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                        If Me.PigConsole.IsYesOrNo("Is set debug") = True Then
                            Me.PigHostApp.SetDebug()
                        End If
                        Console.WriteLine("Set CurrHost")
                        Me.CurrHost = Me.PigHostApp.MyHost
                        If Me.PigHostApp.LastErr <> "" Then Console.WriteLine(Me.PigHostApp.LastErr)
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub ShowHost(oHost As Host)
        With oHost
            Console.WriteLine("HostID=" & .HostID)
            Console.WriteLine("HostName=" & .HostName)
            Console.WriteLine("HostMainIp=" & .HostMainIp)
        End With
    End Sub

    Public Sub ShowHostFolder(oHostFolder As HostFolder)
        With oHostFolder
            Console.WriteLine("HostID=" & .HostID)
            Console.WriteLine("FolderID=" & .FolderID)
            Console.WriteLine("FolderName=" & .FolderName)
            Console.WriteLine("FolderPath=" & .FolderPath)
            Console.WriteLine("FolderType=" & .FolderType.ToString)
        End With
    End Sub

End Class
