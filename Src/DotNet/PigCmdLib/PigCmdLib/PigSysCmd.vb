'**********************************
'* Name: PigSysCmd
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 系统操作的命令|Commands for system operation
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.23
'* Create Time: 2/6/2022
'*1.1  3/6/2022  Add GetListenPortProcID
'*1.2  7/6/2022  Add GetOSCaption
'*1.3  17/6/2022 Add GetProcListenPortList
'*1.4  23/7/2022 Add GetWmicSimpleXml
'*1.5 26/7/2022  Modify Imports
'*1.6 29/7/2022  Modify Imports
'*1.7 17/8/2022  Add KillProc
'*1.8 29/12/2022  Add GetBootUpTime
'*1.9 16/1/2023  Add GetCmdRetRows
'*1.10 17/1/2023  Modify GetCmdRetRows,GetWmicSimpleXml
'*1.11 19/1/2023  Modify GetUUID,GetCmdRetRows
'*1.12 20/1/2023  Modify GetUUID
'*1.13 14/6/2023  Modify GetListenPortProcID
'*1.15 25/6/2023  Modify GetListenPortProcID,GetWmicSimpleXml,GetOSCaption,GetBootUpTime,GetCmdRetRows,GetProcListenPortList
'*1.16 23/7/2023  Modify GetBootUpTime
'*1.17 16/8/2023  Add ReBootHost
'*1.18 18/8/2023  Add mGetWmicSimpleXml,GetWmicSimpleXml modify ReBootHost
'*1.19 20/10/2023 Add GetDefaultIPGateway
'*1.20 23/11/2023 Add MoveDir,RmDirAndSubDir
'*1.21 30/11/2023 Add GetDuSize, modify MoveDir
'*1.22 21/4/2024  Modify GetListenPortProcID,GetProcListenPortList
'*1.23 11/7/2024  Add GetListenProcXml,GetProcInfXml,GetProcListXml
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' System command and WMIC processing class|系统命令及WMIC处理类
''' </summary>
Public Class PigSysCmd
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "23" & "." & "30"

    Private ReadOnly Property mPigFunc As New PigFunc
    Private ReadOnly Property mPigCmdApp As New PigCmdApp

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
        Dim LOG As New PigStepLog("GetProcListenPortList")
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr " & PID.ToString
                Else
                    strCmd = "netstat -apn|grep tcp|grep LISTEN|grep " & PID.ToString
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                ReDim OutListPort(-1)
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strLine As String = Trim(.StandardOutputArray(i)) & "}", strPort As String = "", strMat As String = ""
                    Do While True
                        If InStr(strLine, "  ") = 0 Then Exit Do
                        strLine = Replace(strLine, "  ", " ")
                    Loop
                    If Me.IsWindows = True Then
                        strMat = "LISTENING " & PID.ToString & "}"
                    Else
                        strMat = "LISTEN " & PID.ToString & "/"
                    End If
                    If InStr(strLine, strMat) > 0 Then strPort = Trim(mPigFunc.GetStr(strLine, ":", " "))
                    If IsNumeric(strPort) Then
                        Dim intPort As Integer = CInt(strPort)
                        If intPort > 0 Then
                            Dim bolIsFind As Boolean = False, intLen As Integer = 0
                            intLen = OutListPort.Length
                            For j = 0 To intLen - 1
                                If OutListPort(j) = intPort Then
                                    bolIsFind = True
                                    Exit For
                                End If
                            Next
                            If bolIsFind = False Then
                                ReDim Preserve OutListPort(intLen)
                                OutListPort(intLen) = intPort
                            End If
                        End If
                    End If
                Next
            End With
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
        Dim LOG As New PigStepLog("GetProcListenPort")
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr :" & ListenPort.ToString
                Else
                    strCmd = "netstat -apn|grep tcp|grep LISTEN|grep :" & ListenPort.ToString
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                OutPID = -1
                Dim strMat As String = ":" & ListenPort.ToString & " "
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strLine As String = Trim(.StandardOutputArray(i)) & "}", strPID As String = ""
                    If Me.IsWindows = True Then
                        If InStr(strLine, strMat) > 0 Then strPID = Trim(mPigFunc.GetStr(strLine, "LISTENING ", "}"))
                    Else
                        If InStr(strLine, strMat) > 0 Then strPID = Trim(mPigFunc.GetStr(strLine, "LISTEN ", "/"))
                    End If
                    If IsNumeric(strPID) Then
                        OutPID = CInt(strPID)
                        Exit For
                    End If
                Next
            End With
            Return "OK"
        Catch ex As Exception
            OutPID = -2
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function GetUUID(ByRef OutUUID As String) As String
        Dim LOG As New PigStepLog("GetUUID")
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

    Private Function mGetWmicSimpleXml(WmicCmd As String, ByRef OutXml As PigXml, Optional IsCrLf As Boolean = False) As String
        Dim LOG As New PigStepLog("mGetWmicSimpleXml")
        Try
            If Me.IsWindows = False Then Throw New Exception("Only windows is supported")
            If Left(UCase(WmicCmd), 5) <> "WMIC " Then WmicCmd = "wmic " & WmicCmd
            If InStr(UCase(WmicCmd), "/FORMAT:") > 0 Then Throw New Exception("Wmiccmd cannot take /Format parameter")
            WmicCmd &= " /Format:textvaluelist"
            Dim oPigCmdApp As New PigCmdApp, intTotalRows As Integer = 0, bolIsNewRow As Boolean = False
            OutXml = New PigXml(IsCrLf)
            With oPigCmdApp
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(WmicCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
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
        Dim LOG As New PigStepLog("GetOSCaption")
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
        Dim LOG As New PigStepLog("KillProc")
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
        Dim LOG As New PigStepLog("GetBootUpTime")
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
        Dim LOG As New PigStepLog("GetCmdRetRows")
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
        Dim LOG As New PigStepLog("ReBootHost")
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
        Dim LOG As New PigStepLog("GetDefaultIPGateway")
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
        Dim LOG As New PigStepLog("RmDirAndSubDir")
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
        Dim LOG As New PigStepLog("MoveDir")
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
        Dim LOG As New PigStepLog("GetDuSize")
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
    ''' <param name="OutPigXml">Output XML|输出的XML</param>
    ''' <param name="IsMergeIp">Whether to merge IP addresses|是否合并IP地址</param>
    ''' <returns></returns>
    Public Function GetListenProcXml(ByRef OutPigXml As PigXml, IsMergeIp As Boolean, Optional PriorityIpHead As String = "", Optional IsCrLf As Boolean = True) As String
        Dim LOG As New PigStepLog("GetListenProcXml")
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING"
                Else
                    strCmd = "netstat -apn|awk '{if($6==""LISTEN"" && $1==""tcp"") print ""<""$4""><""$7"">""}'"
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                Dim strLocalIp As String = ""
                If IsMergeIp = True Then
                    strLocalIp = Me.mPigFunc.GetHostIp(PriorityIpHead)
                End If
                Dim asTmp(-1) As String
                Dim strListenIp As String = "", intListenPort As Integer, intPID As Integer, strLine As String = "", intNo As Integer = 0
                For i = 0 To .StandardOutputArray.Length - 1
                    Dim strTmp As String, strLineOut As String = ""
                    If Me.IsWindows = True Then
                        LOG.StepName = "mClearDoubleSpace"
                        'LOG.Ret = Me.mPigFunc.ClearDoubleSpace(.StandardOutputArray(i), strLine)
                        If LOG.Ret <> "OK" Then
                            Throw New Exception(LOG.Ret)
                        ElseIf strLine = "" Then
                            Throw New Exception("Return an empty string")
                        End If
                        If InStr(strLine, "[::]") = 0 Then
                            strLine = "<" & Replace(strLine, " ", "><") & ">"
                            strTmp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                            strListenIp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                            strTmp = Me.mPigFunc.GetStr(strListenIp, ":", "", True)
                            intListenPort = Me.mPigFunc.GECLng(strTmp)
                            If IsMergeIp = True Then
                                strListenIp = strLocalIp
                            End If
                            strTmp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                            strTmp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                            strTmp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                            intPID = Me.mPigFunc.GECLng(strTmp) : If intPID <= 0 Then intPID = -1
                            strLineOut = "<" & strListenIp & "><" & intListenPort.ToString & "><" & intPID.ToString & ">"
                        End If
                    Else
                        strListenIp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                        strTmp = Me.mPigFunc.GetStr(strListenIp, ":", "", True)
                        intListenPort = Me.mPigFunc.GECLng(strTmp)
                        If IsMergeIp = True Then
                            strListenIp = strLocalIp
                        End If
                        strTmp = Me.mPigFunc.GetStr(strLine, "<", ">", True)
                        strTmp = Me.mPigFunc.GetStr(strTmp, "/", "", True)
                        intPID = Me.mPigFunc.GECLng(strTmp) : If intPID <= 0 Then intPID = -1
                    End If
                    If strLineOut <> "" Then
                        ReDim Preserve asTmp(intNo)
                        asTmp(intNo) = strLineOut
                        intNo += 1
                    End If
                Next
                Dim asOut(-1) As String
                LOG.StepName = "DistinctString"
                LOG.Ret = Me.mPigFunc.DistinctString(asTmp, asOut)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                OutPigXml = New PigXml(IsCrLf)
                Dim intLeftTab As Integer = 0
                If IsCrLf = True Then intLeftTab = 1
                With OutPigXml
                    .AddEleLeftSign("Root")
                    For i = 0 To asOut.Length - 1
                        strLine = asOut(i)
                        strListenIp = Me.mPigFunc.GetStr(strLine, "<", ">")
                        intListenPort = Me.mPigFunc.GECLng(Me.mPigFunc.GetStr(strLine, "<", ">"))
                        intPID = Me.mPigFunc.GECLng(Me.mPigFunc.GetStr(strLine, "<", ">"))
                        .AddEleLeftSign("Item")
                        .AddEle("LocalIp", strListenIp, intLeftTab)
                        .AddEle("LocalPort", intListenPort.ToString, intLeftTab)
                        .AddEle("PID", intPID.ToString, intLeftTab)
                        .AddEleRightSign("Item")
                    Next
                    .AddEle("TotalItems", asOut.Length)
                    .AddEleRightSign("Root")
                End With
            End With
            Return "OK"
        Catch ex As Exception
            If OutPigXml Is Nothing Then OutPigXml = New PigXml(IsCrLf)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Structure mStruProcList
        Public LineIn As String
        Public LineOut As String
        Public Tmp As String
        Public AsTmp As String()
        Public Items As Integer
        Public LeftTab As Integer
        Public BeginPos As Integer
        '-----------
        Public PID As Integer
        Public ProcName As String
        Public TTYOrSession As String
        '-----------
        Public Sub Reset()
            LineIn = ""
            LineOut = ""
            Tmp = ""
            ReDim AsTmp(-1)
            Items = 0
            LeftTab = 0
            BeginPos = 0
        End Sub
        Public Sub ResetLine()
            PID = -1
            ProcName = ""
            TTYOrSession = ""
        End Sub
    End Structure

    Private Structure mStruProcInf
        Public LineIn As String
        Public LineOut As String
        Public Tmp As String
        Public AsTmp As String()
        Public Items As Integer
        Public LeftTab As Integer
        '-----------
        Public PID As Integer
        Public ProcName As String
        Public TTYOrSession As String
        Public UserName As String
        Public FilePath As String
        Public ProcCmd As String
        Public MemoryUse As Decimal
        Public TotalProcessorTime As String
        '-----------
        Public Sub Reset()
            LineIn = ""
            LineOut = ""
            Tmp = ""
            ReDim AsTmp(-1)
            Items = 0
            LeftTab = 0
        End Sub
        Public Sub ResetLine()
            PID = -1
            ProcName = ""
            TTYOrSession = ""
            UserName = ""
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
        Dim LOG As New PigStepLog("GetProcListXml")
        Dim strCmd As String = ""
        Try
            Dim oPigCmdApp As New PigCmdApp
            Dim msProcList As New mStruProcList
            msProcList.Reset()
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "tasklist /FO CSV "
                    If ByProcName <> "" Then
                        strCmd &= " /FI ""IMAGENAME eq " & ByProcName & """"
                    End If
                    msProcList.BeginPos = 1
                Else
                    strCmd = "ps -e|awk '{"
                    If ByProcName <> "" Then
                        strCmd &= "if($4==""" & ByProcName & """)"
                    Else
                        msProcList.BeginPos = 1
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
                    msProcList.LeftTab = 1
                End If
                OutPigXml.AddEleLeftSign("Root")
                OutPigXml.AddEle("ByProcName", ByProcName)
                For i = msProcList.BeginPos To .StandardOutputArray.Length - 1
                    With msProcList
                        .ResetLine()
                        .Items += 1
                        .Tmp = ""","
                        .LineIn = oPigCmdApp.StandardOutputArray(i)
                        If Right(.LineIn, 2) <> .Tmp Then
                            .LineIn &= .Tmp
                        End If
                        If Me.IsWindows = True Then
                            .ProcName = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            .Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            If .Tmp = "" Then
                                .PID = -1
                            Else
                                .PID = Me.mPigFunc.GEInt(.Tmp)
                            End If
                            .TTYOrSession = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            .Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            .TTYOrSession &= "#" & .Tmp
                        Else
                            .Tmp = Me.mPigFunc.GetStr(.LineIn, "<", ">")
                            If .Tmp = "" Then
                                .PID = -1
                            Else
                                .PID = Me.mPigFunc.GEInt(.Tmp)
                            End If
                            .TTYOrSession = Me.mPigFunc.GetStr(.LineIn, "<", ">")
                            .ProcName = Me.mPigFunc.GetStr(.LineIn, "<", ">")
                        End If
                        OutPigXml.AddEleLeftSign("Item")
                        OutPigXml.AddEle("ProcName", .ProcName, .LeftTab)
                        OutPigXml.AddEle("PID", .PID.ToString, .LeftTab)
                        OutPigXml.AddEle("TTYOrSession", .TTYOrSession, .LeftTab)
                        OutPigXml.AddEleRightSign("Item")
                    End With
                Next
                OutPigXml.AddEle("TotalItems", msProcList.Items)
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
        Return Me.mGetProcInfXml(OutPigXml, aiByPID, IsCrLf)
    End Function
    Private Function mGetProcInfXml(ByRef OutPigXml As PigXml, ByPIDs As Integer(), Optional IsCrLf As Boolean = True) As String
        Dim LOG As New PigStepLog("mGetProcInfXml")
        Dim strCmd As String = ""
        Try
            Dim oPigCmdApp As New PigCmdApp
            With oPigCmdApp
                If Me.IsWindows = True Then
                    strCmd = "tasklist /FO CSV "
                    If ByPID >= 0 Then
                        strCmd &= " /FI ""PID eq " & ByPID.ToString & """"
                    ElseIf ByProcName <> "" Then
                        strCmd &= " /FI ""IMAGENAME eq " & ByProcName & """"
                    End If
                    If ByProcUserName <> "" Then
                        strCmd &= " /FI ""USERNAME eq " & ByProcUserName & """"
                    End If
                Else
                    If ByPID >= 0 Then
                        strCmd = "ps -e|awk '{if($2==" & ByPID.ToString & ") print $0}'"
                    ElseIf ByProcName <> "" Then
                        strCmd = "ps -e|grep " & ByProcName & "|grep -v grep"
                    Else
                        strCmd = "ps -e"
                    End If
                    If ByProcUserName <> "" Then
                        strCmd &= "|grep " & ByProcUserName & "|grep -v grep"
                    End If
                    strCmd &= "|awk '{print ""\"""""
                    For i = 1 To 8
                        strCmd &= "$" & i.ToString & """\"",\"""""
                    Next
                    strCmd &= """\""""}'"
                End If
                Console.WriteLine(strCmd)
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                Dim msProcInf As New mStruProcInf
                msProcInf.Reset()
                OutPigXml = New PigXml(IsCrLf)
                If IsCrLf = True Then
                    msProcInf.LeftTab = 1
                End If
                OutPigXml.AddEleLeftSign("Root")
                OutPigXml.AddEle("ByPID", ByPID)
                OutPigXml.AddEle("ByProcName", ByProcName)
                OutPigXml.AddEle("ByProcUserName", ByProcUserName)
                Dim intBegin As Integer = 0
                If Me.IsWindows = True Then
                    intBegin = 1
                End If
                For i = intBegin To .StandardOutputArray.Length - 1
                    With msProcInf
                        .ResetLine()
                        .Items += 1
                        .LineIn = oPigCmdApp.StandardOutputArray(i) & """"
                        If Me.IsWindows = True Then
                            '.ProcNameOrCmd = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            'If .Tmp = "" Then
                            '    .PID = -1
                            'Else
                            '    .PID = Me.mPigFunc.GEInt(.Tmp)
                            'End If
                            '.TTYOrSession = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.TTYOrSession &= "#" & .Tmp
                            '.MemUsed = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.UserName = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.CpuTime = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.WinTitle = Me.mPigFunc.GetStr(.LineIn, """", """,")
                        Else
                            '.UserName = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            'If .Tmp = "" Then
                            '    .PID = -1
                            'Else
                            '    .PID = Me.mPigFunc.GEInt(.Tmp)
                            'End If
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            'If .Tmp = "" Then
                            '    .ParentID = -1
                            'Else
                            '    .ParentID = Me.mPigFunc.GEInt(.Tmp)
                            'End If
                            '.Tmp = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.StartTime = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.TTYOrSession = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.CpuTime = Me.mPigFunc.GetStr(.LineIn, """", """,")
                            '.ProcNameOrCmd = Me.mPigFunc.GetStr(.LineIn, """", """,")
                        End If
                        OutPigXml.AddEleLeftSign("Item")
                        'OutPigXml.AddEle("ProcNameOrCmd", .ProcNameOrCmd, .LeftTab)
                        'OutPigXml.AddEle("PID", .PID.ToString, .LeftTab)
                        'OutPigXml.AddEle("ParentID", .ParentID.ToString, .LeftTab)
                        'OutPigXml.AddEle("TTYOrSession", .TTYOrSession, .LeftTab)
                        'OutPigXml.AddEle("MemUsed", .MemUsed, .LeftTab)
                        'OutPigXml.AddEle("UserName", .UserName, .LeftTab)
                        'OutPigXml.AddEle("CpuTime", .CpuTime, .LeftTab)
                        'OutPigXml.AddEle("StartTime", .StartTime, .LeftTab)
                        'OutPigXml.AddEle("WinTitle", .WinTitle, .LeftTab)
                        OutPigXml.AddEleRightSign("Item")
                    End With
                Next
                OutPigXml.AddEle("TotalItems", msProcInf.Items)
                OutPigXml.AddEleRightSign("Root")
            End With
            Return "OK"
        Catch ex As Exception
            If OutPigXml Is Nothing Then OutPigXml = New PigXml(IsCrLf)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
