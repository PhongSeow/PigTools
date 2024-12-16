'**********************************
'* Name: PigSysCmd
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 系统操作的命令|Commands for system operation
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.31
'* Create Time: 2/6/2022
'* 1.1  3/6/2022  Add GetListenPortProcID
'* 1.2  7/6/2022  Add GetOSCaption
'* 1.3  17/6/2022 Add GetProcListenPortList
'* 1.4  23/7/2022 Add GetWmicSimpleXml
'* 1.5 26/7/2022  Modify Imports
'* 1.6 29/7/2022  Modify Imports
'* 1.7 17/8/2022  Add KillProc
'* 1.8 29/12/2022  Add GetBootUpTime
'* 1.9 16/1/2023  Add GetCmdRetRows
'* 1.10 17/1/2023  Modify GetCmdRetRows,GetWmicSimpleXml
'* 1.11 19/1/2023  Modify GetUUID,GetCmdRetRows
'* 1.12 20/1/2023  Modify GetUUID
'* 1.13 14/6/2023  Modify GetListenPortProcID
'* 1.15 25/6/2023  Modify GetListenPortProcID,GetWmicSimpleXml,GetOSCaption,GetBootUpTime,GetCmdRetRows,GetProcListenPortList
'* 1.16 23/7/2023  Modify GetBootUpTime
'* 1.17 16/8/2023  Add ReBootHost
'* 1.18 18/8/2023  Add mGetWmicSimpleXml,GetWmicSimpleXml modify ReBootHost
'* 1.19 20/10/2023 Add GetDefaultIPGateway
'* 1.20 23/11/2023 Add MoveDir,RmDirAndSubDir
'* 1.21 30/11/2023 Add GetDuSize, modify MoveDir
'* 1.22 21/4/2024  Modify GetListenPortProcID,GetProcListenPortList
'* 1.23 11/7/2024  Add GetTcpListenProcList,GetProcInfXml,GetProcListXml
'* 1.25 15/7/2024  Add GetProcUserName, modify GetTcpListenProcList
'* 1.26 18/7/2024  Modify GetListenPortProcID,GetProcListenPortList,mStruProcList
'* 1.27 21/7/2024  Modify PigFunc to PigFuncLite
'* 1.28  28/7/2024  Modify PigStepLog to StruStepLog
'* 1.29  16/10/2024 Modify mGetProcInfXml,GetProcListXml,GetTcpListenProcList,mGetWmicSimpleXml, add GetParentProcList,GetAllProcCmdList
'* 1.30  27/11/2024 Add GetLinuxServiceList,GetLinuxServicePID,GetWindowsServiceList
'* 1.31  13/12/2024 Modify GetAllProcCmdList
'**********************************
Imports System.Security.Cryptography
Imports PigCmdLib.PigSysCmd
Imports PigToolsLiteLib
''' <summary>
''' System command and WMIC processing class|系统命令及WMIC处理类
''' </summary>
Public Class PigSysCmd
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "31" & "." & "28"

    Private ReadOnly Property mPigFunc As New PigFunc
    Private ReadOnly Property mPigCmdApp As New PigCmdApp
    Private ReadOnly Property mPigProcApp As New PigProcApp

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub


    ''' <summary>
    ''' 获取进程的侦听端口数组|Get the listening port array of the process
    ''' </summary>
    ''' <param name="PID">进程号|Process number</param>
    ''' <param name="OutListPort">输出的侦听端口|Listening port for output</param>
    ''' <returns></returns>
    Public Function GetProcListenPortList(PID As Integer, ByRef OutListPort As Integer()) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetProcListenPortList"
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            If Me.IsWindows = True Then
                strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr " & PID.ToString
            Else
                strCmd = "netstat -apn|awk '{if($6==""LISTEN"" && $1==""tcp"") print $4"" ""$7}'|grep " & PID.ToString & """/"""
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = oPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            ReDim OutListPort(-1)
            Dim stTmp As New StruTmp, intPID As Integer, intListenPort As Integer
            With stTmp
                .Reset()
                .BeginStr = "<"
                .EndStr = ">"
            End With
            For i = 0 To oPigCmdApp.StandardOutputArray.Length - 1
                With stTmp
                    .ResetLine()
                    .LineIn = oPigCmdApp.StandardOutputArray(i)
                    .LineInStrSpaceMulti2One(True)
                    intPID = -1
                    intListenPort = -1
                    If Me.IsWindows = True Then
                        .CutLineIn2StrTmp()
                        .CutLineIn2StrTmp()
                        .StrTmp2 = .GetStr(.StrTmp, "[", "]")
                        If .StrTmp2 = "" Then
                            .StrTmp2 = .GetStr(.StrTmp, ":", "")
                            intListenPort = .IntTmp2
                            .CutLineIn2StrTmp()
                            .CutLineIn2StrTmp()
                            .CutLineIn2StrTmp()
                            If IsNumeric(.StrTmp) = True Then intPID = .IntTmp
                        End If
                    Else
                        .CutLineIn2StrTmp()
                        .StrTmp2 = .GetStr(.StrTmp, ":", "")
                        intListenPort = .IntTmp2
                        .CutLineIn2StrTmp()
                        .StrTmp2 = .GetStr(.StrTmp, "", "/")
                        If IsNumeric(.StrTmp2) = True Then intPID = .IntTmp2
                    End If
                    If intPID = PID And intListenPort > 0 Then
                        .AddItems()
                        ReDim Preserve .AsTmp(.Items - 1)
                        .AsTmp(.Items - 1) = intListenPort
                    End If
                End With
            Next
            LOG.StepName = "AsTmpDistinctString"
            LOG.Ret = stTmp.AsTmpDistinctString(stTmp.AsTmp2)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "AsTmp2 To OutListPort"
            ReDim OutListPort(stTmp.AsTmp2.Length - 1)
            For i = 0 To stTmp.AsTmp2.Length - 1
                OutListPort(i) = CInt(stTmp.AsTmp2(i))
            Next
            Return "OK"
        Catch ex As Exception
            ReDim OutListPort(-1)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' 获取侦听端口的进程号|Get the process number of the listening port
    ''' </summary>
    ''' <param name="ListenPort">侦听端口|Listening port</param>
    ''' <param name="OutPID">获取的进程号|Obtained process number</param>
    ''' <returns></returns>
    Public Function GetListenPortProcID(ListenPort As Integer, ByRef OutPID As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetListenPortProcID"
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr " & ListenPort.ToString
                Else
                    strCmd = "netstat -apn|awk '{if($6==""LISTEN"" && $1==""tcp"") print $4"" ""$7}'|grep "":" & ListenPort.ToString & """"
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                OutPID = -1
                Dim stTmp As New StruTmp, intListenPort As Integer
                With stTmp
                    .Reset()
                    .BeginStr = "<"
                    .EndStr = ">"
                End With
                For i = 0 To .StandardOutputArray.Length - 1
                    With stTmp
                        .ResetLine()
                        .LineIn = oPigCmdApp.StandardOutputArray(i)
                        .LineInStrSpaceMulti2One(True)
                        OutPID = -1
                        If Me.IsWindows = True Then
                            .CutLineIn2StrTmp()
                            .CutLineIn2StrTmp()
                            .StrTmp2 = .GetStr(.StrTmp, "[", "]")
                            If .StrTmp2 = "" Then
                                .StrTmp2 = .GetStr(.StrTmp, ":", "")
                                intListenPort = .IntTmp2
                                .CutLineIn2StrTmp()
                                .CutLineIn2StrTmp()
                                .CutLineIn2StrTmp()
                                If IsNumeric(.StrTmp) = True Then OutPID = .IntTmp
                            End If
                        Else
                            .CutLineIn2StrTmp()
                            .StrTmp2 = .GetStr(.StrTmp, ":", "")
                            intListenPort = .IntTmp2
                            .CutLineIn2StrTmp()
                            .StrTmp2 = .GetStr(.StrTmp, "", "/")
                            If IsNumeric(.StrTmp2) = True Then OutPID = .IntTmp2
                        End If
                        If intListenPort = ListenPort Then Exit For
                    End With
                Next
            End With
            Return "OK"
        Catch ex As Exception
            OutPID = -2
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function GetUUID(ByRef OutUUID As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetUUID"
        Try
            If Me.IsWindows = True Then
                LOG.StepName = "GetWmicSimpleXml"
                LOG.Ret = Me.GetWmicSimpleXml("csproduct get uuid", OutUUID)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Dim oPigXml As New PigXml(False)
                LOG.StepName = "SetMainXml"
                oPigXml.SetMainXml(OutUUID)
                If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
                LOG.StepName = "InitXmlDocument"
                LOG.Ret = oPigXml.InitXmlDocument()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "XmlDocGetStr"
                OutUUID = oPigXml.XmlDocGetStr("WmicXml.Row1.UUID")
                oPigXml = Nothing
            Else
                LOG.StepName = "GetProductUuid"
                OutUUID = Me.mPigFunc.GetProductUuid(LOG.Ret)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            Return "OK"
        Catch ex As Exception
            OutUUID = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 将Wmic查询结果转换成简单的XML，Wmic /Format:xml 输出的XML过于复杂|Convert wmic query results into simple XML. Wmic /format:xml output XML is too complex
    ''' </summary>
    ''' <param name="WmicCmd">Wmic 命令|Wmic command</param>
    ''' <param name="OutXml">输出的XML|Output XML</param>
    ''' <returns></returns>
    Public Function GetWmicSimpleXml(WmicCmd As String, ByRef OutXml As String, Optional IsCrLf As Boolean = False) As String
        Try
            Dim oPigXml As PigXml = Nothing
            Dim strRet As String = Me.mGetWmicSimpleXml(WmicCmd, oPigXml, IsCrLf)
            If strRet <> "OK" Then Throw New Exception(strRet)
            OutXml = oPigXml.MainXmlStr
            oPigXml = Nothing
            Return "OK"
        Catch ex As Exception
            OutXml = ""
            Return Me.GetSubErrInf("GetWmicSimpleXml", ex)
        End Try
    End Function

    ''' <summary>
    ''' 将Wmic查询结果转换成简单的XML，Wmic /Format:xml 输出的XML过于复杂|Convert wmic query results into simple XML. Wmic /format:xml output XML is too complex
    ''' </summary>
    ''' <param name="WmicCmd">Wmic 命令|Wmic command</param>
    ''' <param name="OutXml">输出的XML对象|Output XML object</param>
    ''' <returns></returns>
    Public Function GetWmicSimpleXml(WmicCmd As String, ByRef OutXml As PigXml, Optional IsCrLf As Boolean = False) As String
        Try
            Dim strRet As String = Me.mGetWmicSimpleXml(WmicCmd, OutXml, IsCrLf)
            If strRet <> "OK" Then Throw New Exception(strRet)
            strRet = OutXml.InitXmlDocument()
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return "OK"
        Catch ex As Exception
            OutXml = Nothing
            Return Me.GetSubErrInf("GetWmicSimpleXml", ex)
        End Try
    End Function

    Private Function mGetWmicSimpleXml(WmicCmd As String, ByRef OutXml As PigXml, Optional IsCrLf As Boolean = False, Optional IsChgCtrlChar As Boolean = False) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mGetWmicSimpleXml"
        Try
            If Me.IsWindows = False Then Throw New Exception("Only windows is supported")
            If Left(UCase(WmicCmd), 5) <> "WMIC " Then WmicCmd = "wmic " & WmicCmd
            If InStr(UCase(WmicCmd), "/FORMAT:") > 0 Then Throw New Exception("Wmiccmd cannot take /Format parameter")
            WmicCmd &= " /Format:textvaluelist"
            Dim oPigCmdApp As New PigCmdApp, intTotalRows As Integer = 0, bolIsNewRow As Boolean = False
            OutXml = New PigXml(IsCrLf, IsChgCtrlChar)
            With oPigCmdApp
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(WmicCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(WmicCmd)
                    Throw New Exception(LOG.Ret)
                ElseIf .StandardError <> "" Then
                    LOG.AddStepNameInf(WmicCmd)
                    LOG.Ret = .StandardError
                    Throw New Exception(LOG.Ret)
                End If
                bolIsNewRow = False
                OutXml.AddEleLeftSign("WmicXml")
                LOG.StepName = "Add Rows"
                Dim intRowNo As Integer = 0, intBlankCnt As Integer = 0
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strLine As String = Trim(.StandardOutputArray(i))
                    If strLine <> "" Then
                        If bolIsNewRow = False Then
                            bolIsNewRow = True
                            intRowNo += 1
                            OutXml.AddEleLeftSign("Row" & intRowNo.ToString, 1)
                            intTotalRows += 1
                        End If
                        Dim strColName As String = Me.mPigFunc.GetStr(strLine, "", "=")
                        OutXml.AddEle(strColName, strLine, 2)
                        intBlankCnt = 0
                    Else
                        If bolIsNewRow = True And intBlankCnt > 2 Then
                            OutXml.AddEleRightSign("Row" & intRowNo, 1)
                            bolIsNewRow = False
                        End If
                        intBlankCnt += 1
                    End If
                Next
                With OutXml
                    .AddEle("TotalRows", intTotalRows.ToString)
                    .AddEleRightSign("WmicXml")
                End With
            End With
            Return "OK"
        Catch ex As Exception
            OutXml = Nothing
            LOG.AddStepNameInf(WmicCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get operating system caption|获取操作系统概述
    ''' </summary>
    ''' <param name="OutOSCaption"></param>
    ''' <returns></returns>
    Public Function GetOSCaption(ByRef OutOSCaption As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetOSCaption"
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "wmic os get Caption"
                Else
                    strCmd = "cat /etc/os-release|grep ""PRETTY_NAME="""
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                OutOSCaption = ""
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strLine As String = Trim(.StandardOutputArray(i))
                    If Me.IsWindows = True Then
                        Select Case strLine
                            Case "", "Caption"
                            Case Else
                                OutOSCaption = strLine
                                Exit For
                        End Select
                    ElseIf InStr(strLine, "PRETTY_NAME=") > 0 Then
                        OutOSCaption = Me.mPigFunc.GetStr(strLine, """", """")
                        Exit For
                    End If
                Next
            End With
            Return "OK"
        Catch ex As Exception
            OutOSCaption = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function KillProc(PID As Integer) As String
        Dim LOG As New StruStepLog : LOG.SubName = "KillProc"
        Try
            Dim strCmd As String = ""
            If Me.IsWindows = True Then
                strCmd = "taskkill /pid " & PID.ToString & " /f"
            Else
                strCmd = "kill -9 " & PID.ToString
            End If
            Dim intOutPID As Integer = -1
            LOG.StepName = "AsyncCmdShell"
            LOG.Ret = Me.mPigCmdApp.AsyncCmdShell(strCmd, intOutPID)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the system boot up time|获取系统启动时间
    ''' </summary>
    ''' <param name="OutBootUpTime"></param>
    ''' <returns></returns>
    Public Function GetBootUpTime(ByRef OutBootUpTime As Date) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetBootUpTime"
        Try
            Dim strOutBootUpTime As String = ""
            If Me.IsWindows = False Then
                Dim oPigCmdApp As New PigCmdApp, strCmd As String
                With oPigCmdApp
                    strCmd = "cat /proc/uptime|awk '{print $1}'"
                    LOG.StepName = "CmdShell"
                    LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(strCmd)
                        Throw New Exception(LOG.Ret)
                    End If
                    strOutBootUpTime = oPigCmdApp.StandardOutput
                    If IsNumeric(strOutBootUpTime) = False Then Throw New Exception("Can not get BootUpTime.")
                    LOG.StepName = "DateAdd(" & strOutBootUpTime & ")"
                    OutBootUpTime = DateAdd(DateInterval.Second, 0 - CLng(strOutBootUpTime), Now)
                End With
            Else
                LOG.StepName = "GetWmicSimpleXml"
                LOG.Ret = Me.GetWmicSimpleXml("path Win32_OperatingSystem get LastBootUpTime", strOutBootUpTime)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Dim oPigXml As New PigXml(False)
                LOG.StepName = "SetMainXml"
                oPigXml.SetMainXml(strOutBootUpTime)
                If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
                LOG.StepName = "InitXmlDocument"
                LOG.Ret = oPigXml.InitXmlDocument()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "XmlDocGetStr"
                strOutBootUpTime = oPigXml.XmlDocGetStr("WmicXml.Row1.LastBootUpTime")
                oPigXml = Nothing
                strOutBootUpTime = Left(strOutBootUpTime, 4) & "-" & Mid(strOutBootUpTime, 5, 2) & "-" & Mid(strOutBootUpTime, 7, 2) & " " & Mid(strOutBootUpTime, 9, 2) & ":" & Mid(strOutBootUpTime, 11, 2) & ":" & Mid(strOutBootUpTime, 13, 2)
                LOG.StepName = "CDate(" & strOutBootUpTime & ")"
                OutBootUpTime = CDate(strOutBootUpTime)
            End If
            Return "OK"
        Catch ex As Exception
            OutBootUpTime = Date.MinValue
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the line content of the command return result|获取命令返回结果的行内容
    ''' </summary>
    ''' <param name="RetContent">Line content returned|返回的行内容</param>
    ''' <param name="RetRows">Returns the number of rows, 0 means all|返回行数，0表示全部</param>
    ''' <returns></returns>
    Public Function GetCmdRetRows(Cmd As String, ByRef RetContent As String, Optional RetRows As Integer = 0) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetCmdRetRows"
        Try
            Dim oPigCmdApp As New PigCmdApp
            With oPigCmdApp
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(Cmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(Cmd)
                    Throw New Exception(LOG.Ret)
                End If
                If RetRows < 0 Then RetRows = 0
                RetContent = ""
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strLine As String = Trim(.StandardOutputArray(i)) & Me.OsCrLf
                    RetContent &= strLine
                    If i >= RetRows - 1 And RetRows > 0 Then Exit For
                Next
            End With
            Return "OK"
        Catch ex As Exception
            RetContent = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' Restart the host|重启主机
    ''' </summary>
    ''' <returns></returns>
    Public Function ReBootHost() As String
        Dim LOG As New StruStepLog : LOG.SubName = "ReBootHost"
        Dim strCmd As String = ""
        Try
            Dim oPigCmdApp As New PigCmdApp
            If Me.IsWindows = True Then
                strCmd = "shutdown /r /t 10 /f"
                LOG.StepName = "CmdShell"
                LOG.Ret = oPigCmdApp.CmdShell(strCmd)
            Else
                strCmd = "/usr/sbin/reboot"
                LOG.StepName = "CallFile"
                LOG.Ret = oPigCmdApp.CallFile(strCmd, "", PigCmdApp.EnmStandardOutputReadType.FullString)
            End If
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If oPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(oPigCmdApp.StandardOutput)
                LOG.Ret = oPigCmdApp.StandardError
                Throw New Exception(LOG.Ret)
            End If
            oPigCmdApp = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetDefaultIPGateway(ByRef DefaultIPGateway As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetDefaultIPGateway"
        Try
            Dim strCmd As String
            If Me.IsWindows = True Then
                strCmd = "wmic path Win32_NetworkAdapterConfiguration get DefaultIPGateway"
            Else
                'strCmd = "route -n | grep '^0.0.0.0' | awk '{print $2}'"
                strCmd = "route -n|awk '{if($1==""0.0.0.0"") print $2}'"
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            If Me.IsWindows = True Then
                DefaultIPGateway = Me.mPigFunc.GetStr(Me.mPigCmdApp.StandardOutput, "{", "}")
            Else
                DefaultIPGateway = Me.mPigCmdApp.StandardOutput
            End If
            Return "OK"
        Catch ex As Exception
            DefaultIPGateway = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function RmDirAndSubDir(DirPath As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "RmDirAndSubDir"
        Dim strCmd As String = ""
        Try
            If Me.mPigFunc.IsFolderExists(DirPath) = False Then Throw New Exception("Directory does not exist.")
            If Me.IsWindows = True Then
                strCmd = "rmdir """ & DirPath & """ /S /Q"
            Else
                strCmd = "rm -rf """ & DirPath
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If Me.mPigCmdApp.StandardError <> "" Then Throw New Exception(Me.mPigCmdApp.StandardError)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Migrate the source directory to the target directory|把源目录迁移到目标目录下
    ''' </summary>
    ''' <param name="SrcDir"></param>
    ''' <param name="TargetDir"></param>
    ''' <param name="IsOverwrite"></param>
    ''' <returns></returns>
    Public Function MoveDir(SrcDir As String, TargetDir As String, Optional IsOverwrite As Boolean = False) As String
        Dim LOG As New StruStepLog : LOG.StepName = "MoveDir"
        Dim strCmd As String = ""
        Try
            Select Case Right(SrcDir, 1)
                Case "/", "\"
                    SrcDir = Left(SrcDir, Len(SrcDir) - 1)
            End Select
            Select Case Right(TargetDir, 1)
                Case "/", "\"
                    TargetDir = Left(TargetDir, Len(TargetDir) - 1)
            End Select
            If SrcDir = TargetDir Then
                LOG.AddStepNameInf(SrcDir)
                LOG.AddStepNameInf(TargetDir)
                Throw New Exception("The source directory and target directory cannot be the same.")
            End If
            If Me.mPigFunc.IsFolderExists(SrcDir) = False Then
                LOG.AddStepNameInf(SrcDir)
                Throw New Exception("The source directory does not exist.")
            End If
            If Me.mPigFunc.IsFolderExists(TargetDir) = False Then
                LOG.StepName = "CreateFolder"
                LOG.AddStepNameInf("The target directory does not exist, try creating it.")
                LOG.AddStepNameInf(TargetDir)
                LOG.Ret = Me.mPigFunc.CreateFolder(TargetDir)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If

            Dim strDirName As String = Me.mPigFunc.GetFilePart(SrcDir, PigFunc.EnmFilePart.FileTitle)
            Dim strTargetDir As String = TargetDir & Me.OsPathSep & strDirName
            If Me.mPigFunc.IsFolderExists(strTargetDir) = True Then
                If IsOverwrite = True Then
                    LOG.StepName = "The target directory already exists, please delete it first."
                    LOG.Ret = Me.RmDirAndSubDir(strTargetDir)
                    If LOG.Ret <> "OK" Then
                        LOG.AddStepNameInf(strTargetDir)
                        Throw New Exception(LOG.Ret)
                    End If
                    Me.mPigFunc.Delay(2000)
                Else
                    LOG.StepName = "Check target directory"
                    LOG.AddStepNameInf(strTargetDir)
                    Throw New Exception("The target directory already exists.")
                End If
            End If
            Dim lngSrcDirDuSize As Long = 0, strSrcDirFastPigMD5 As String = ""
            If Me.IsWindows = True Then
                LOG.StepName = "GetFastPigMD5.SrcDir"
                Dim oPigFolder As New PigFolder(SrcDir)
                LOG.Ret = oPigFolder.GetFastPigMD5(strSrcDirFastPigMD5, PigFolder.EnmGetFastPigMD5Type.CurrDirInfo)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(SrcDir)
                    Throw New Exception(LOG.Ret)
                End If
                strCmd = "move /Y """ & SrcDir & """ """ & TargetDir & """"
            Else
                LOG.StepName = "GetDuSize.SrcDir"
                LOG.Ret = Me.GetDuSize(SrcDir, lngSrcDirDuSize)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(SrcDir)
                    Throw New Exception(LOG.Ret)
                End If
                strCmd = "mv -f " & SrcDir & " " & TargetDir
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            If Me.IsWindows = True Then
                Dim strTargetDirFastPigMD5 As String = ""
                LOG.StepName = "GetFastPigMD5.TargetDir"
                Dim oPigFolder As New PigFolder(strTargetDir)
                LOG.Ret = oPigFolder.GetFastPigMD5(strTargetDirFastPigMD5, PigFolder.EnmGetFastPigMD5Type.CurrDirInfo)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strTargetDir)
                    Throw New Exception(LOG.Ret)
                End If
                If strSrcDirFastPigMD5 <> strTargetDirFastPigMD5 Then
                    LOG.StepName = "Check the FastPigMD5 of the source and target directories."
                    LOG.AddStepNameInf(SrcDir & ":" & strSrcDirFastPigMD5)
                    LOG.AddStepNameInf(strTargetDir & ":" & strTargetDirFastPigMD5)
                    Throw New Exception("Inconsistent")
                End If
            Else
                Dim lngTargetDirDuSize As Long
                LOG.StepName = "GetDuSize.TargetDir"
                LOG.Ret = Me.GetDuSize(strTargetDir, lngTargetDirDuSize)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strTargetDir)
                    Throw New Exception(LOG.Ret)
                End If
                If lngSrcDirDuSize <> lngTargetDirDuSize Then
                    LOG.StepName = "Check the size of the source and target directories."
                    LOG.AddStepNameInf(SrcDir & ":" & lngSrcDirDuSize)
                    LOG.AddStepNameInf(strTargetDir & ":" & lngTargetDirDuSize)
                    Throw New Exception("Inconsistent")
                End If
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetDuSize(DirPath As String, ByRef DuSize As Long) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetDuSize"
        Dim strCmd As String = ""
        Try
            If Me.IsWindows = True Then Throw New Exception("Can only support Linux platforms.")
            If Me.mPigFunc.IsFolderExists(DirPath) = False Then
                LOG.AddStepNameInf(DirPath)
                Throw New Exception("The directory does not exist.")
            End If
            strCmd = "du " & DirPath & "|tail -1|awk '{print $1}'"
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            ElseIf Me.mPigCmdApp.StandardError <> "" Then
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            LOG.StepName = "GECLng"
            DuSize = Me.mPigFunc.GECLng(Me.mPigCmdApp.StandardOutput)
            Return "OK"
        Catch ex As Exception
            DuSize = -1
            If strCmd <> "" Then LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the list of listening processes|获取侦听的进程列表
    ''' </summary>
    ''' <param name="OutList">Output StruTcpListenProcList|输出的 StruTcpListenProcList</param>
    ''' <param name="IsMergeIp">Whether to merge IP addresses|是否合并IP地址</param>
    ''' <returns></returns>
    Public Function GetTcpListenProcList(ByRef OutList As StruTcpListenProcList(), IsMergeIp As Boolean, Optional PriorityIpHead As String = "", Optional IsCrLf As Boolean = True) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetTcpListenProcList"
        Try
            ReDim OutList(-1)
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING"
                Else
                    strCmd = "netstat -apn|awk '{if($6==""LISTEN"" && $1==""tcp"") print $4"" ""$7}'"
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                ElseIf .StandardError <> "" Then
                    LOG.Ret = .StandardError
                    Throw New Exception(LOG.Ret)
                End If
                Dim strLocalIp As String = ""
                If IsMergeIp = True Then
                    strLocalIp = Me.mPigFunc.GetHostIp(PriorityIpHead)
                End If
                Dim stMain As New StruTmp
                With stMain
                    .Reset()
                    .BeginStr = "<"
                    .EndStr = ">"
                End With
                For i = 0 To .StandardOutputArray.Length - 1
                    With stMain
                        .ResetLine()
                        .LineIn = oPigCmdApp.StandardOutputArray(i)
                        If .LineIn <> "" Then
                            .LineInStrSpaceMulti2One(True)
                            If Me.IsWindows = True Then
                                If InStr(.LineIn, "[::]") = 0 Then
                                    .AddItems()
                                    ReDim Preserve OutList(.Items - 1)
                                    OutList(.Items - 1).ResetLine()
                                    .CutLineIn2StrTmp()
                                    .CutLineIn2StrTmp()
                                    OutList(.Items - 1).LocalIp = Me.mPigFunc.GetStr(.StrTmp, "", ":", True)
                                    OutList(.Items - 1).ListenPort = .IntTmp
                                    If IsMergeIp = True Then
                                        OutList(.Items - 1).LocalIp = strLocalIp
                                    End If
                                    .CutLineIn2StrTmp()
                                    .CutLineIn2StrTmp()
                                    .CutLineIn2StrTmp()
                                    OutList(.Items - 1).PID = .IntTmp
                                    ReDim Preserve .AsTmp(.Items - 1)
                                    .AsTmp(.Items - 1) = OutList(.Items - 1).FullStr
                                End If
                            Else
                                If InStr(.LineIn, ":::") = 0 Then
                                    .AddItems()
                                    ReDim Preserve OutList(.Items - 1)
                                    OutList(.Items - 1).ResetLine()
                                    .CutLineIn2StrTmp()
                                    OutList(.Items - 1).LocalIp = Me.mPigFunc.GetStr(.StrTmp, "", ":", True)
                                    OutList(.Items - 1).ListenPort = .IntTmp
                                    If IsMergeIp = True Then
                                        OutList(.Items - 1).LocalIp = strLocalIp
                                    End If
                                    .CutLineIn2StrTmp()
                                    .StrTmp = Me.mPigFunc.GetStr(.StrTmp, "", "/", True)
                                    OutList(.Items - 1).PID = .IntTmp
                                    ReDim Preserve .AsTmp(.Items - 1)
                                    .AsTmp(.Items - 1) = OutList(.Items - 1).FullStr
                                End If
                            End If
                        End If
                    End With
                Next
                LOG.StepName = "DistinctString"
                LOG.Ret = Me.mPigFunc.DistinctString(stMain.AsTmp, stMain.AsTmp2)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                ReDim OutList(stMain.AsTmp2.Length - 1)
                stMain.BeginStr = "<"
                stMain.EndStr = ">"
                For i = 0 To OutList.Length - 1
                    OutList(i).ResetLine()
                    With stMain
                        .LineIn = .AsTmp2(i)
                        .CutLineIn2StrTmp()
                        If .StrTmp = "" Then
                            OutList(i).PID = -1
                        Else
                            OutList(i).PID = stMain.IntTmp
                        End If
                        .CutLineIn2StrTmp()
                        OutList(i).LocalIp = stMain.StrTmp
                        .CutLineIn2StrTmp()
                        OutList(i).ListenPort = stMain.IntTmp
                    End With
                Next
            End With
            Return "OK"
        Catch ex As Exception
            ReDim OutList(-1)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Structure mStruProcList
        Friend Tmp As StruTmp
        '-----------
        Public PID As Integer
        Public ProcName As String
        Public TTYOrSession As String
        '-----------
        Public Sub ResetLine()
            PID = -1
            ProcName = ""
            TTYOrSession = ""
        End Sub
    End Structure

    ''' <summary>
    ''' Intermediate processing structure for strings|字符串中间处理结构
    ''' Version: 1.1
    ''' Create Time: 16/7/2024
    ''' 1.1  19/7/2024  Add ResetLine, modify Reset
    ''' </summary>
    Friend Structure StruTmp
        Public LineIn As String
        Public LineOut As String
        Public StrTmp As String
        Public StrTmp2 As String
        Public AsTmp As String()
        Public AsTmp2 As String()
        Public Items As Integer
        Public LeftTab As Integer
        Public BeginPos As Integer
        Public PigFunc As PigFunc
        Public BeginStr As String
        Public EndStr As String
        Public Sub Reset()
            Items = 0
            LeftTab = 0
            BeginPos = 0
            BeginStr = ""
            EndStr = ""
            If PigFunc Is Nothing Then
                PigFunc = New PigFunc
            End If
            ReDim AsTmp(-1)
            ReDim AsTmp2(-1)
        End Sub
        Public Sub ResetLine()
            LineIn = ""
            LineOut = ""
            StrTmp = ""
            StrTmp2 = ""
        End Sub
        Public Sub AddItems()
            Items += 1
        End Sub
        Public Sub CutLineIn2StrTmp()
            StrTmp = PigFunc.GetStr(LineIn, BeginStr, EndStr, True)
        End Sub
        Public Function IntTmp() As Integer
            If IsNumeric(StrTmp) Then
                Return CInt(StrTmp)
            Else
                Return 0
            End If
        End Function
        Public Function LngTmp() As Long
            If IsNumeric(StrTmp) Then
                Return CLng(StrTmp)
            Else
                Return 0
            End If
        End Function
        Public Function BoolTmp() As Boolean
            Return CBool(StrTmp)
        End Function
        Public Function DecTmp() As Decimal
            If IsNumeric(StrTmp) Then
                Return CDec(StrTmp)
            Else
                Return 0
            End If
        End Function
        Public Function DateTmp() As Date
            If IsDate(StrTmp) Then
                Return CDate(StrTmp)
            Else
                Return #2000/1/1#
            End If
        End Function
        Public Sub LineInStrSpaceMulti2One(Optional IsTrimConvert As Boolean = True)
            Dim strTmp As String = ""
            PigFunc.StrSpaceMulti2One(LineIn, strTmp, IsTrimConvert)
            LineIn = strTmp
        End Sub
        Public Function GetStr(ByRef SourceStr As String, BeginStr As String, EndStr As String) As String
            Return PigFunc.GetStr(SourceStr, BeginStr, EndStr, True)
        End Function
        Public Function IntTmp2() As Integer
            If IsNumeric(StrTmp2) Then
                Return CInt(StrTmp2)
            Else
                Return 0
            End If
        End Function
        Public Function LngTmp2() As Long
            If IsNumeric(StrTmp2) Then
                Return CLng(StrTmp2)
            Else
                Return 0
            End If
        End Function
        Public Function BoolTmp2() As Boolean
            Return CBool(StrTmp2)
        End Function
        Public Function DecTmp2() As Decimal
            If IsNumeric(StrTmp2) Then
                Return CDec(StrTmp2)
            Else
                Return 0
            End If
        End Function
        Public Function DateTmp2() As Date
            If IsDate(StrTmp2) Then
                Return CDate(StrTmp2)
            Else
                Return #2000/1/1#
            End If
        End Function
        Public Function AsTmpDistinctString(ByRef OutStrings As String()) As String
            Return PigFunc.DistinctString(AsTmp, OutStrings)
        End Function
    End Structure


    Public Structure StruTcpListenProcList
        Public PID As Integer
        Public LocalIp As String
        Public ListenPort As Integer
        '-----------
        Friend Sub ResetLine()
            PID = -1
            LocalIp = ""
            ListenPort = -1
        End Sub
        Public Function FullStr() As String
            FullStr = "<" & PID.ToString & ">"
            FullStr &= "<" & LocalIp & ">"
            FullStr &= "<" & ListenPort.ToString & ">"
        End Function
    End Structure


    Private Structure mStruProcInf
        Friend Tmp As StruTmp
        '-----------
        Public PID As Integer
        Public ParentPID As Integer
        Public ProcName As String
        Public TTYOrSession As String
        Public ProcUserName As String
        Public FilePath As String
        Public ProcCmd As String
        Public MemoryUse As Decimal
        Public TotalProcessorTime As String
        '-----------
        Public Sub ResetLine()
            PID = -1
            ParentPID = -1
            ProcName = ""
            TTYOrSession = ""
            ProcUserName = ""
            FilePath = ""
            ProcCmd = ""
            MemoryUse = 0
            TotalProcessorTime = ""
        End Sub
    End Structure

    ''' <summary>
    ''' Get process list|获取进程列表
    ''' </summary>
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="ByProcName">By process name|按进程名</param>
    ''' <param name="IsCrLf"></param>
    ''' <returns></returns>
    Public Function GetProcListXml(ByRef OutPigXml As PigXml, Optional ByProcName As String = "", Optional IsCrLf As Boolean = True) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetProcListXml"
        Dim strCmd As String = ""
        Try
            Dim oPigCmdApp As New PigCmdApp
            Dim msProcList As New mStruProcList
            msProcList.Tmp.Reset()
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "tasklist /FO CSV "
                    If ByProcName <> "" Then
                        strCmd &= " /FI ""IMAGENAME eq " & ByProcName & """"
                    End If
                    msProcList.Tmp.BeginPos = 1
                Else
                    strCmd = "ps -e|awk '{"
                    If ByProcName <> "" Then
                        strCmd &= "if($4==""" & ByProcName & """)"
                    Else
                        msProcList.Tmp.BeginPos = 1
                    End If
                    strCmd &= " print ""<"""
                    For i = 1 To 4
                        Select Case i
                            Case 1, 2, 4
                                strCmd &= "$" & i.ToString & """><"""
                        End Select
                    Next
                    strCmd &= "}'"
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    Throw New Exception(LOG.Ret)
                ElseIf .StandardError <> "" Then
                    LOG.Ret = .StandardError
                    Throw New Exception(LOG.Ret)
                End If
                OutPigXml = New PigXml(IsCrLf)
                If IsCrLf = True Then
                    msProcList.Tmp.LeftTab = 1
                End If
                OutPigXml.AddEleLeftSign("Root")
                OutPigXml.AddEle("ByProcName", ByProcName)
                For i = msProcList.Tmp.BeginPos To .StandardOutputArray.Length - 1
                    Dim strLineIn As String = oPigCmdApp.StandardOutputArray(i)
                    If strLineIn <> "" Then
                        With msProcList
                            .ResetLine()
                            .Tmp.AddItems()
                            .Tmp.StrTmp = ""","
                            .Tmp.LineIn = strLineIn
                            If Right(.Tmp.LineIn, 2) <> .Tmp.StrTmp Then
                                .Tmp.LineIn &= .Tmp.StrTmp
                            End If
                            If Me.IsWindows = True Then
                                .ProcName = Me.mPigFunc.GetStr(.Tmp.LineIn, """", """,")
                                .Tmp.StrTmp = Me.mPigFunc.GetStr(.Tmp.LineIn, """", """,")
                                If .Tmp.StrTmp = "" Then
                                    .PID = -1
                                Else
                                    .PID = Me.mPigFunc.GEInt(.Tmp.StrTmp)
                                End If
                                .TTYOrSession = Me.mPigFunc.GetStr(.Tmp.LineIn, """", """,")
                                .Tmp.StrTmp = Me.mPigFunc.GetStr(.Tmp.LineIn, """", """,")
                                .TTYOrSession &= "#" & .Tmp.StrTmp
                            Else
                                .Tmp.StrTmp = Me.mPigFunc.GetStr(.Tmp.LineIn, "<", ">")
                                If .Tmp.StrTmp = "" Then
                                    .PID = -1
                                Else
                                    .PID = Me.mPigFunc.GEInt(.Tmp.StrTmp)
                                End If
                                .TTYOrSession = Me.mPigFunc.GetStr(.Tmp.LineIn, "<", ">")
                                .ProcName = Me.mPigFunc.GetStr(.Tmp.LineIn, "<", ">")
                            End If
                            OutPigXml.AddEleLeftSign("Item")
                            OutPigXml.AddEle("ProcName", .ProcName, .Tmp.LeftTab)
                            OutPigXml.AddEle("PID", .PID.ToString, .Tmp.LeftTab)
                            OutPigXml.AddEle("TTYOrSession", .TTYOrSession, .Tmp.LeftTab)
                            OutPigXml.AddEleRightSign("Item")
                        End With
                    End If
                Next
                OutPigXml.AddEle("TotalItems", msProcList.Tmp.Items)
                OutPigXml.AddEleRightSign("Root")
            End With
            Return "OK"
        Catch ex As Exception
            If OutPigXml Is Nothing Then OutPigXml = New PigXml(IsCrLf)
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get detailed process information|获取进程详细信息
    ''' </summary>
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="ByPID">By process number|按进程号</param>
    ''' <param name="IsCrLf"></param>
    ''' <returns></returns>
    Public Function GetProcInfXml(ByRef OutPigXml As PigXml, ByPID As Integer, Optional IsCrLf As Boolean = True) As String
        Dim aiByPID(0) As Integer
        aiByPID(0) = ByPID
        Return Me.mGetProcInfXml(OutPigXml, aiByPID, False, IsCrLf)
    End Function

    ''' <summary>
    ''' Get detailed process information|获取进程详细信息
    ''' </summary>
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="ByPID">By process number|按进程号</param>
    ''' <param name="IsWinGetUserName">Whether to obtain the windows username of the process|是否获取 Windows 进程的用户名</param>
    ''' <param name="IsCrLf"></param>
    ''' <returns></returns>
    Public Function GetProcInfXml(ByRef OutPigXml As PigXml, ByPID As Integer, IsGetWinUserName As Boolean, Optional IsCrLf As Boolean = True) As String
        Dim aiByPID(0) As Integer
        aiByPID(0) = ByPID
        Return Me.mGetProcInfXml(OutPigXml, aiByPID, IsGetWinUserName, IsCrLf)
    End Function

    ''' <summary>
    ''' Get detailed process information|获取进程详细信息
    ''' </summary>
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="ByPIDs">By process number Array|按进程号数组</param>
    ''' <param name="IsWinGetUserName">Whether to obtain the windows username of the process|是否获取 Windows 进程的用户名</param>
    ''' <param name="IsCrLf"></param>
    ''' <returns></returns>
    Public Function GetProcInfXml(ByRef OutPigXml As PigXml, ByPIDs As Integer(), IsGetUserName As Boolean, Optional IsCrLf As Boolean = True) As String
        Return Me.mGetProcInfXml(OutPigXml, ByPIDs, IsGetUserName, IsCrLf)
    End Function

    ''' <summary>
    ''' Get detailed process information|获取进程详细信息
    ''' </summary>
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="ByPIDs">By process number Array|按进程号数组</param>
    ''' <param name="IsCrLf"></param>
    ''' <returns></returns>
    Public Function GetProcInfXml(ByRef OutPigXml As PigXml, ByPIDs As Integer(), Optional IsCrLf As Boolean = True) As String
        Return Me.mGetProcInfXml(OutPigXml, ByPIDs, False, IsCrLf)
    End Function

    Private Function mGetProcInfXml(ByRef OutPigXml As PigXml, ByPIDs As Integer(), Optional IsGetWinUserName As Boolean = False, Optional IsCrLf As Boolean = True) As String
        Dim LOG As New StruStepLog : LOG.SubName = "mGetProcInfXml"
        Dim strCmd As String = ""
        Try
            If ByPIDs.Length <= 0 Then
                LOG.Ret = "No process number."
                Throw New Exception(LOG.Ret)
            End If
            Dim msProcInf As New mStruProcInf
            msProcInf.Tmp.Reset()
            OutPigXml = New PigXml(IsCrLf)
            If IsCrLf = True Then
                msProcInf.Tmp.LeftTab = 1
            End If
            OutPigXml.AddEleLeftSign("Root")
            OutPigXml.AddEle("IsGetWinUserName", IsGetWinUserName)
            Dim strBadPIDs As String = "", strIf As String = ""
            Dim oPigProcs As New PigProcs
            For i = 0 To ByPIDs.Length - 1
                Dim bolIsBad As Boolean = False
                Dim intPID As Integer = ByPIDs(i)
                If Me.IsWindows = True Then
                    If intPID < 0 Then bolIsBad = True
                Else
                    If intPID <= 0 Then bolIsBad = True
                End If
                If bolIsBad = False Then
                    LOG.StepName = "PigProcApp.GetPigProc"
                    Dim oPigProc As PigProc = Me.mPigProcApp.GetPigProc(intPID)
                    If oPigProc Is Nothing Then
                        bolIsBad = True
                    ElseIf oPigProc.LastErr <> "" Then
                        bolIsBad = True
                    ElseIf oPigProc.ProcessName = "" Then
                        bolIsBad = True
                    End If
                End If
                If bolIsBad = True Then
                    If strBadPIDs = "" Then
                        strBadPIDs = intPID.ToString
                    Else
                        strBadPIDs &= "," & intPID.ToString
                    End If
                Else
                    oPigProcs.Add(intPID)
                    Dim strTmp As String = "$2==""" & intPID.ToString & """"
                    If strIf = "" Then
                        strIf = "if(" & strTmp
                    Else
                        strIf &= " || " & strTmp
                    End If
                End If
            Next
            If strIf <> "" Then
                strIf &= ")"
            End If
            Dim pxTmp As PigXml = Nothing
            If Me.IsWindows = False Then
                If strIf = "" Then
                    strIf = "if(""1""==""0"")"
                End If
                Dim oPigCmdApp As New PigCmdApp
                With oPigCmdApp
                    strCmd = "ps -ef|awk '{" & strIf & "{printf ""<""$1""><""$2""><""$3""><ProcCmd>""; for(i=8;i<=NF;i++){printf $i"" ""}; print ""</ProcCmd>"";}}'"
                    LOG.StepName = "CmdShell"
                    LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                    If LOG.Ret <> "OK" Then
                        Throw New Exception(LOG.Ret)
                    ElseIf .StandardError <> "" Then
                        LOG.Ret = .StandardError
                        Throw New Exception(LOG.Ret)
                    End If
                    pxTmp = New PigXml(False)
                    With pxTmp
                        .AddEleLeftSign("Root")
                        For i = 0 To oPigCmdApp.StandardOutputArray.Length - 1
                            With msProcInf
                                .ResetLine()
                                .Tmp.LineIn = oPigCmdApp.StandardOutputArray(i)
                                .Tmp.BeginStr = "<"
                                .Tmp.EndStr = ">"
                                .Tmp.CutLineIn2StrTmp()
                                .ProcUserName = .Tmp.StrTmp
                                .Tmp.CutLineIn2StrTmp()
                                If .Tmp.StrTmp = "" Then
                                    .PID = -1
                                Else
                                    .PID = .Tmp.IntTmp
                                End If
                                .Tmp.CutLineIn2StrTmp()
                                If .Tmp.StrTmp = "" Then
                                    .ParentPID = -1
                                Else
                                    .ParentPID = .Tmp.IntTmp
                                End If
                                .Tmp.BeginStr = "<ProcCmd>"
                                .Tmp.EndStr = "</ProcCmd>"
                                .Tmp.CutLineIn2StrTmp()
                                .ProcCmd = Trim(.Tmp.StrTmp)
                            End With
                            .AddEleLeftSign("P" & msProcInf.PID.ToString)
                            .AddEle("ProcUserName", msProcInf.ProcUserName)
                            .AddEle("ParentPID", msProcInf.ParentPID)
                            .AddEle("ProcCmd", msProcInf.ProcCmd)
                            .AddEleRightSign("P" & msProcInf.PID.ToString)
                        Next
                        .AddEleRightSign("Root")
                        LOG.StepName = "InitXmlDocument"
                        LOG.Ret = .InitXmlDocument
                        If LOG.Ret <> "OK" Then
                            LOG.AddStepNameInf(.MainXmlStr)
                            Throw New Exception(LOG.Ret)
                        End If
                    End With
                End With
            End If
            msProcInf.Tmp.Reset()
            For Each oPigProc As PigProc In oPigProcs
                msProcInf.Tmp.AddItems()
                With oPigProc
                    OutPigXml.AddEleLeftSign("Item")
                    OutPigXml.AddEle("PID", .ProcessID.ToString, msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("ProcName", .ProcessName, msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("FilePath", .FilePath, msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("StartTime", Me.mPigFunc.GetFmtDateTime(.StartTime), msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("ThreadsCount", .ThreadsCount.ToString, msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("TotalProcessorTime", .TotalProcessorTime.ToString, msProcInf.Tmp.LeftTab)
                    OutPigXml.AddEle("MemoryUse", .MemoryUse.ToString, msProcInf.Tmp.LeftTab)
                    If IsGetWinUserName = True And Me.IsWindows = True Then
                        Dim strProcUserName As String = ""
                        LOG.Ret = Me.GetProcUserName(.ProcessID, strProcUserName)
                        If LOG.Ret <> "OK" Then
                            LOG.AddStepNameInf("GetProcUserName")
                            LOG.AddStepNameInf("PID=" & .ProcessID.ToString)
                            Throw New Exception(LOG.Ret)
                        Else
                            OutPigXml.AddEle("ProcUserName", strProcUserName, msProcInf.Tmp.LeftTab)
                        End If
                    ElseIf Me.IsWindows = False Then
                        OutPigXml.AddEle("ParentPID", pxTmp.XmlDocGetStr("Root.P" & .ProcessID.ToString & ".ParentPID"), msProcInf.Tmp.LeftTab)
                        OutPigXml.AddEle("ProcCmd", pxTmp.XmlDocGetStr("Root.P" & .ProcessID.ToString & ".ProcCmd"), msProcInf.Tmp.LeftTab)
                        OutPigXml.AddEle("ProcUserName", pxTmp.XmlDocGetStr("Root.P" & .ProcessID.ToString & ".ProcUserName"), msProcInf.Tmp.LeftTab)
                    End If
                    OutPigXml.AddEleRightSign("Item")
                End With
            Next
            OutPigXml.AddEle("TotalItems", msProcInf.Tmp.Items)
            OutPigXml.AddEle("BadPIDs", strBadPIDs)
            OutPigXml.AddEleRightSign("Root")
            Return "OK"
        Catch ex As Exception
            If OutPigXml Is Nothing Then OutPigXml = New PigXml(IsCrLf)
            If strCmd <> "" Then LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Get the username of the process|获取进程的用户名|
    ''' </summary>
    ''' <param name="OutProcUserName">The username of the process|进程的用户名</param>
    ''' <returns></returns>
    Public Function GetProcUserName(PID As Integer, ByRef OutProcUserName As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetProcUserName"
        Dim strCmd As String = ""
        Try
            Dim oPigCmdApp As New PigCmdApp
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "TASKLIST /V /FO CSV /FI ""PID eq " & PID.ToString & " """
                Else
                    strCmd = "ps -ef|awk '{if($2==""" & PID.ToString & """) print $1}'"
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    Throw New Exception(LOG.Ret)
                ElseIf .StandardError <> "" Then
                    LOG.Ret = .StandardError
                    Throw New Exception(LOG.Ret)
                End If
                Dim bolIsGot As Boolean = False
                Select Case .StandardOutputArray.Length
                    Case 1
                        If Me.IsWindows = False Then bolIsGot = True
                    Case 2
                        If Me.IsWindows = True Then bolIsGot = True
                End Select
                If bolIsGot = False Then
                    LOG.Ret = "Unable to obtain process information"
                    Throw New Exception(LOG.Ret)
                End If
                If Me.IsWindows = True Then
                    Dim stMain As New StruTmp
                    With stMain
                        .Reset()
                        .BeginStr = """"
                        .EndStr = ""","
                        .ResetLine()
                        .LineIn = oPigCmdApp.StandardOutputArray(1)
                        For i = 1 To 7
                            .CutLineIn2StrTmp()
                        Next
                        OutProcUserName = .StrTmp
                    End With
                Else
                    OutProcUserName = Trim(.StandardOutputArray(0))
                    If OutProcUserName = "" Then
                        LOG.Ret = "Unable to obtain process information"
                        Throw New Exception(LOG.Ret)
                    End If
                End If
            End With
            Return "OK"
        Catch ex As Exception
            OutProcUserName = ""
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Structure StruServiceList
        Public PID As Integer
        Public IsWindows As Boolean
        Public ServiceName As String
        Public IsRunning As Boolean
        Public ServiceDesc As String
        Public IsActive As Boolean
        Public IsLoaded As Boolean
        Public PathName As String
        Public IsAutoStart As Boolean
        '-----------
        Friend Sub ResetLine()
            PID = -1
            ServiceName = ""
            IsRunning = False
            ServiceDesc = ""
            IsLoaded = False
            IsActive = False
            IsAutoStart = False
            PathName = ""
        End Sub
    End Structure


    Public Structure StruAllProcCmdList
        Public PID As Integer
        Public ParentPID As Integer
        Public ProcCmd As String
        Public ParentProcCmd As String
        '-----------
        Public ProcName As String
        Public FilePath As String
        Public StartTime As Date
        Public ThreadsCount As Integer
        Public TotalProcessorTime As String
        Public MemoryUse As Decimal
        Public ProcUserName As String
        Public ListenPort As Integer
        '-----------
        Friend Sub ResetLine()
            PID = -1
            ParentPID = -1
            ProcCmd = ""
            ParentProcCmd = ""
            '-----
            ProcName = ""
            FilePath = ""
            StartTime = #2000-1-1#
            ThreadsCount = -1
            TotalProcessorTime = ""
            MemoryUse = -1
            ProcUserName = ""
            ListenPort = -1
        End Sub
        Public Function ColNameLine(Optional IsDetailProcInf As Boolean = False) As String
            With Me
                ColNameLine = "PID,ParentPID,ProcCmd,ParentProcCmd"
                If IsDetailProcInf = True Then
                    ColNameLine &= ",ProcName,FilePath,StartTime,ThreadsCount,TotalProcessorTime,MemoryUse,ProcUserName,ListenPort"
                End If
            End With
        End Function
        Public Function BcpLine(Optional IsDetailProcInf As Boolean = False) As String
            With Me
                BcpLine = .PID.ToString & vbTab & .ParentPID.ToString & vbTab & .ProcCmd & vbTab & .ParentProcCmd
                If IsDetailProcInf = True Then
                    BcpLine &= vbTab & .ProcName & vbTab & .FilePath & vbTab & Format(.StartTime, "yyyy-MM-dd HH:mm:ss.fff") & vbTab & .ThreadsCount.ToString & vbTab & .TotalProcessorTime & vbTab & .MemoryUse.ToString & vbTab & .ProcUserName & vbTab & .ListenPort.ToString
                End If
            End With
        End Function
    End Structure

    ''' <summary>
    ''' 获取 Linux 服务的 PID|Get Linux service PID
    ''' </summary>
    ''' <param name="OutItem">Output StruServiceList item|输出的StruServiceList项</param>
    ''' <returns></returns>
    Public Function GetLinuxServicePID(ByRef OutItem As StruServiceList) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetLinuxServicePID"
        Try
            If Me.IsWindows = True Then
                LOG.Ret = "This is not a Linux system"
                Throw New Exception(LOG.Ret)
            End If
            Dim oPigCmdApp As New PigCmdApp, strCmd As String, strServiceName As String = Replace(OutItem.ServiceName, "&", "")
            LOG.StepName = "Check ServiceName"
            If strServiceName = "" Then
                LOG.Ret = "ServiceName is empty"
                Throw New Exception(LOG.Ret)
            ElseIf Len(strServiceName) > 64 Then
                LOG.AddStepNameInf("ServiceName=" & OutItem.ServiceName)
                LOG.Ret = "The service name is too long"
                Throw New Exception(LOG.Ret)
            End If
            strCmd = "systemctl status " & OutItem.ServiceName & "|grep ""Main PID:""|awk '{print $3}'"
            LOG.StepName = "CmdShell"
            LOG.Ret = oPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf oPigCmdApp.StandardError <> "" Then
                LOG.Ret = oPigCmdApp.StandardError
                Throw New Exception(LOG.Ret)
            End If
            If IsNumeric(oPigCmdApp.StandardOutput) = True Then
                OutItem.PID = CInt(oPigCmdApp.StandardOutput)
            Else
                OutItem.PID = -1
            End If
            Return "OK"
        Catch ex As Exception
            OutItem.PID = -1
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try

    End Function

    ''' <summary>
    ''' Get Windows service PID|获取 Windows 服务的 PID
    ''' </summary>
    ''' <param name="OutList">Output List|输出的列表</param>
    ''' <param name="IsGetRunningOnly">Only retrieve items that are currently running|只获取运行中的项|</param>
    ''' <param name="ByPID">Check by process number, if yes, ignore IsGetRunningOnly parameter|是否按进程号查，是则忽略IsGetRunningOnly参数</param>
    ''' <returns></returns>
    Public Function GetWindowsServiceList(ByRef OutList As StruServiceList(), IsGetRunningOnly As Boolean, Optional ByPID As Integer = -1) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetWindowsServiceList"
        Try
            If Me.IsWindows = False Then
                LOG.Ret = "This is not a Windows system"
                Throw New Exception(LOG.Ret)
            End If
            Dim strCmd As String = ""
            strCmd = "service"
            If ByPID >= 0 Then
                strCmd &= " where ProcessId=" & ByPID.ToString
            ElseIf IsGetRunningOnly = True Then
                strCmd &= " where State=""Running"""
            End If
            strCmd &= " get Caption,Name,PathName,ProcessId,Started,StartMode,State"
            Dim pxMain As PigXml = Nothing
            LOG.StepName = "GetWmicSimpleXml"
            LOG.Ret = Me.GetWmicSimpleXml(strCmd, pxMain)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim intTotalItems As Integer = pxMain.XmlDocGetInt("WmicXml.TotalRows")
            For i = 1 To intTotalItems
                Dim strXmlRowKey As String = "WmicXml.Row" & i.ToString & "."
                ReDim Preserve OutList(i - 1)
                With OutList(i - 1)
                    .ResetLine()
                    .IsWindows = True
                    .ServiceName = pxMain.XmlDocGetStr(strXmlRowKey & "Name")
                    If pxMain.XmlDocGetBool(strXmlRowKey & "Started") = True Then
                        .IsLoaded = True
                        .IsActive = True
                    Else
                        .IsLoaded = False
                        .IsActive = False
                    End If
                    If pxMain.XmlDocGetStr(strXmlRowKey & "StartMode") = "Auto" Then
                        .IsAutoStart = True
                    Else
                        .IsAutoStart = False
                    End If
                    .PathName = pxMain.XmlDocGetStr(strXmlRowKey & "PathName")
                    If pxMain.XmlDocGetStr(strXmlRowKey & "State") = "Running" Then
                        .IsRunning = True
                    Else
                        .IsRunning = False
                    End If
                    .ServiceDesc = pxMain.XmlDocGetStr(strXmlRowKey & "Caption")
                    .PID = pxMain.XmlDocGetInt(strXmlRowKey & "ProcessId")
                End With
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function




    ''' <summary>
    ''' Get Linux service list|获取 Linux 服务列表
    ''' </summary>
    ''' <param name="OutList">Output List|输出的列表</param>
    ''' <param name="IsGetRunningOnly">Only retrieve items that are currently running|只获取运行中的项|</param>
    ''' <returns></returns>
    Public Function GetLinuxServiceList(ByRef OutList As StruServiceList(), IsGetRunningOnly As Boolean) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetLinuxServiceList"
        Try
            If Me.IsWindows = True Then
                LOG.Ret = "This is not a Linux system"
                Throw New Exception(LOG.Ret)
            End If
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            strCmd = "systemctl --type=service"
            If IsGetRunningOnly = True Then
                strCmd &= "|awk '{if($4==""running"") print $0}'"
            Else
                strCmd &= "|awk '{if (NR>2) print $0}'"
            End If
            LOG.StepName = "CmdShell"
            LOG.Ret = oPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf oPigCmdApp.StandardError <> "" Then
                LOG.Ret = oPigCmdApp.StandardError
                Throw New Exception(LOG.Ret)
            End If
            ReDim OutList(-1)
            Dim stTmp As New StruTmp
            With stTmp
                .Reset()
                .BeginStr = "<"
                .EndStr = ">"
            End With
            For i = 0 To oPigCmdApp.StandardOutputArray.Length - 1
                ReDim Preserve OutList(i)
                With stTmp
                    .ResetLine()
                    .LineIn = oPigCmdApp.StandardOutputArray(i)
                    .LineInStrSpaceMulti2One(True)
                    .CutLineIn2StrTmp()
                    OutList(i).ServiceName = .StrTmp
                    .CutLineIn2StrTmp()
                    If .StrTmp = "loaded" Then
                        OutList(i).IsLoaded = True
                    Else
                        OutList(i).IsLoaded = False
                    End If
                    .CutLineIn2StrTmp()
                    If .StrTmp = "active" Then
                        OutList(i).IsActive = True
                    Else
                        OutList(i).IsActive = False
                    End If
                    .CutLineIn2StrTmp()
                    If .StrTmp = "running" Then
                        OutList(i).IsRunning = True
                    Else
                        OutList(i).IsRunning = False
                    End If
                    .CutLineIn2StrTmp()
                    OutList(i).ServiceDesc = .StrTmp
                End With
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function



    ''' <summary>
    ''' Retrieve command line information for all processes|获取全部进程的命令行信息
    ''' </summary>
    ''' <param name="OutList">Output StruAllProcCmdList|输出的 StruAllProcCmdList</param>
    ''' <returns></returns>
    Public Function GetAllProcCmdList(ByRef OutList As StruAllProcCmdList(), Optional IsDetailProcInf As Boolean = False) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetAllProcCmdList"
        Try
            Dim strCmd As String = ""
            ReDim OutList(-1)
            If Me.IsWindows = True Then
                strCmd = "process get ProcessId,ParentProcessId,CommandLine"
                Dim pxMain As PigXml = Nothing
                LOG.StepName = "GetWmicSimpleXml"
                LOG.Ret = Me.GetWmicSimpleXml(strCmd, pxMain)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Dim intTotalItems As Integer = pxMain.XmlDocGetInt("WmicXml.TotalRows")
                For i = 1 To intTotalItems
                    Dim strXmlRowKey As String = "WmicXml.Row" & i.ToString & "."
                    Dim intPID As Integer = pxMain.XmlDocGetInt(strXmlRowKey & "ProcessId")
                    ReDim Preserve OutList(i - 1)
                    With OutList(i - 1)
                        .ResetLine()
                        .PID = intPID
                        .ParentPID = pxMain.XmlDocGetInt(strXmlRowKey & "ParentProcessId")
                        .ProcCmd = pxMain.XmlDocGetStr(strXmlRowKey & "CommandLine")
                        .ProcUserName = ""
                    End With
                Next
            Else
                strCmd = "ps -ef|awk '{printf ""<""$1"">""<""$2""><""$3""><ProcCmd>""; for(i=8;i<=NF;i++){printf $i"" ""}; print ""</ProcCmd>"";}'"
                LOG.StepName = "CmdShell"
                LOG.Ret = mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                ElseIf mPigCmdApp.StandardError <> "" Then
                    LOG.Ret = mPigCmdApp.StandardError
                    Throw New Exception(LOG.Ret)
                End If
                Dim sTmp As New StruTmp
                With sTmp
                    .Reset()
                End With
                For i = 0 To mPigCmdApp.StandardOutputArray.Length - 1
                    Dim intPID As Integer, intParentPID As Integer, strProcCmd As String, strProcUserName As String
                    With sTmp
                        .ResetLine()
                        .BeginStr = "<" : .EndStr = ">"
                        .LineIn = mPigCmdApp.StandardOutputArray(i)
                        .CutLineIn2StrTmp() : strProcUserName = .StrTmp
                        .CutLineIn2StrTmp() : intPID = .IntTmp
                        .CutLineIn2StrTmp() : intParentPID = .IntTmp
                        .BeginStr = "<ProcCmd>" : .EndStr = "</ProcCmd>"
                        .CutLineIn2StrTmp() : strProcCmd = .StrTmp
                    End With
                    ReDim Preserve OutList(i)
                    With OutList(i)
                        .ResetLine()
                        .ProcUserName = strProcUserName
                        .PID = intPID
                        .ParentPID = intParentPID
                        .ProcCmd = strProcCmd
                    End With
                Next
            End If
            For i = 0 To OutList.Length - 1
                For Each oItem As StruAllProcCmdList In OutList
                    With oItem
                        If .PID = OutList(i).ParentPID Then
                            OutList(i).ParentProcCmd = .ProcCmd
                            Exit For
                        End If
                    End With
                Next
            Next
            If Me.IsWindows = True Then
                strCmd = "query process"
                LOG.StepName = "CmdShell"
                LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                Dim sTmp As New StruTmp
                With sTmp
                    .Reset()
                End With
                For i = 1 To mPigCmdApp.StandardOutputArray.Length - 1
                    Dim intPID As Integer, strProcUserName As String, bolIsDo As Boolean = False
                    With sTmp
                        .ResetLine()
                        .BeginStr = "{{" : .EndStr = "}}"
                        LOG.StepName = "StrSpaceMulti2One"
                        Me.mPigFunc.StrSpaceMulti2One(mPigCmdApp.StandardOutputArray(i), .LineIn, True)
                        If LOG.Ret <> "OK" Then
                            LOG.AddStepNameInf(mPigCmdApp.StandardOutputArray(i))
                        End If
                        .LineIn = .BeginStr & Replace(.LineIn, "><", "}}{{") & .EndStr
                        .CutLineIn2StrTmp() : strProcUserName = .StrTmp
                        strProcUserName = Replace(strProcUserName, ">", "")
                        strProcUserName = Replace(strProcUserName, "<", "")
                        .CutLineIn2StrTmp()
                        .CutLineIn2StrTmp()
                        .CutLineIn2StrTmp()
                        If IsNumeric(.StrTmp) = True Then
                            bolIsDo = True
                            intPID = .IntTmp
                        End If
                    End With
                    If bolIsDo = True Then
                        For j = 0 To OutList.Length - 1
                            With OutList(j)
                                If .PID = intPID Then
                                    .ProcUserName = strProcUserName
                                    Exit For
                                End If
                            End With
                        Next
                    End If
                Next
                'For i = 0 To OutList.Length - 1
                '    Dim strProcUserName As String = ""
                '    Me.GetProcUserName(OutList(i).PID, strProcUserName)
                '    OutList(i).ProcUserName = strProcUserName
                'Next
            End If
            If IsDetailProcInf = True Then
                For i = 0 To OutList.Length - 1
                    With OutList(i)
                        Dim oPigProc As New PigProc(.PID)
                        If oPigProc IsNot Nothing Then
                            .ProcName = oPigProc.ProcessName
                            .FilePath = oPigProc.FilePath
                            .TotalProcessorTime = oPigProc.TotalProcessorTime.ToString
                            .ThreadsCount = oPigProc.ThreadsCount
                            .StartTime = oPigProc.StartTime
                            .MemoryUse = oPigProc.MemoryUse
                        End If
                    End With
                Next
                LOG.StepName = ""
                Dim aTcpListenProcList(-1) As StruTcpListenProcList
                LOG.Ret = Me.GetTcpListenProcList(aTcpListenProcList, True,, False)
                If LOG.Ret = "OK" Then
                    For Each oItem As StruTcpListenProcList In aTcpListenProcList
                        For i = 0 To OutList.Length - 1
                            With OutList(i)
                                If .PID = oItem.PID Then
                                    .ListenPort = oItem.ListenPort
                                End If
                            End With
                        Next
                    Next
                End If
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Structure StruParentProcList
        Public PID As Integer
        Public ParentPID As Integer
        Public ParentProcCmd As String
        '-----------
        Friend Sub ResetLine()
            ParentPID = -1
            ParentProcCmd = ""
        End Sub
    End Structure

    ''' <summary>
    ''' Get the list of parent processes|获取父进程列表
    ''' </summary>
    ''' <param name="OutList">Output StruParentProcList|输出的 StruParentProcList</param>
    ''' <returns></returns>
    Public Function GetParentProcList(ByRef OutList As StruParentProcList()) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetParentProcList"
        Try
            Dim strCmd As String = "", strIDs As String = ""
            For i = 0 To OutList.Length - 1
                OutList(i).ResetLine()
                Dim intPID As Integer = OutList(i).PID
                If Me.IsWindows = True Then
                    If strIDs <> "" Then strIDs &= " or "
                    strIDs &= "ProcessId=" & intPID.ToString
                Else
                    If strIDs <> "" Then strIDs &= "||"
                    strIDs &= "$2==""" & intPID.ToString & """"
                End If
            Next
            If Me.IsWindows = True Then
                strCmd = "process where """ & strIDs & """ get ProcessId,ParentProcessId,CommandLine"
                Dim pxMain As PigXml = Nothing
                LOG.StepName = "GetWmicSimpleXml"
                LOG.Ret = Me.GetWmicSimpleXml(strCmd, pxMain)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Dim intTotalItems As Integer = pxMain.XmlDocGetInt("WmicXml.TotalRows")
                For i = 1 To intTotalItems
                    Dim strXmlRowKey As String = "WmicXml.Row" & i.ToString & "."
                    Dim intPID As Integer = pxMain.XmlDocGetInt(strXmlRowKey & "ProcessId")
                    For j = 0 To OutList.Length - 1
                        With OutList(j)
                            If .PID = intPID Then
                                .ParentPID = pxMain.XmlDocGetInt(strXmlRowKey & "ParentProcessId")
                                .ParentProcCmd = pxMain.XmlDocGetStr(strXmlRowKey & "CommandLine")
                                Exit For
                            End If
                        End With
                    Next
                Next
            Else
                strCmd = "ps -ef|awk '{if(" & strIDs & "){printf ""<""$2""><""$3""><ProcCmd>""; for(i=8;i<=NF;i++){printf $i"" ""}; print ""</ProcCmd>"";}}'"
                LOG.StepName = "CmdShell"
                LOG.Ret = mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                ElseIf mPigCmdApp.StandardError <> "" Then
                    LOG.Ret = mPigCmdApp.StandardError
                    Throw New Exception(LOG.Ret)
                End If
                Dim sTmp As New StruTmp
                With sTmp
                    .Reset()
                End With
                For i = 0 To mPigCmdApp.StandardOutputArray.Length - 1
                    Dim intPID As Integer, intParentPID As Integer, strParentProcCmd As String
                    With sTmp
                        .ResetLine()
                        .BeginStr = "<" : .EndStr = ">"
                        .LineIn = mPigCmdApp.StandardOutputArray(i)
                        .CutLineIn2StrTmp() : intPID = .IntTmp
                        .CutLineIn2StrTmp() : intParentPID = .IntTmp
                        .BeginStr = "<ProcCmd>" : .EndStr = "</ProcCmd>"
                        .CutLineIn2StrTmp() : strParentProcCmd = .StrTmp
                    End With
                    For j = 0 To OutList.Length - 1
                        With OutList(j)
                            If .PID = intPID Then
                                .ParentPID = intParentPID
                                .ParentProcCmd = strParentProcCmd
                                Exit For
                            End If
                        End With
                    Next
                Next
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
