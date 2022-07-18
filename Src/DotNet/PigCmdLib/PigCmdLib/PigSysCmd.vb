'**********************************
'* Name: PigSysCmd
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 系统操作的命令|Commands for system operation
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 2/6/2022
'*1.1  3/6/2022   Add GetListenPortProcID
'*1.2  7/6/2022   Add GetOSCaption
'*1.3  17/6/2022   Add GetProcListenPortList
'**********************************
Imports PigToolsLiteLib
Public Class PigSysCmd
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.6"

    Private ReadOnly Property mPigFunc As New PigFunc

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
                .StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.StringArray
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr " & PID.ToString
                Else
                    strCmd = "netstat -apn|grep tcp|grep LISTEN|grep " & PID.ToString
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd)
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strCmd)
                    Throw New Exception(LOG.Ret)
                End If
                ReDim OutListPort(0)
                OutListPort(0) = 0
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
                                If OutListPort(0) = 0 Then
                                    OutListPort(0) = intPort
                                Else
                                    ReDim Preserve OutListPort(intLen)
                                    OutListPort(intLen) = intPort
                                End If
                            End If
                        End If
                    End If
                Next
                If OutListPort(0) = 0 Then Throw New Exception("Process has no listening port")
            End With
            Return "OK"
        Catch ex As Exception
            ReDim OutListPort(0)
            OutListPort(0) = 0
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
                .StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.StringArray
                If Me.IsWindows = True Then
                    strCmd = "netstat -ano|findstr TCP|findstr LISTENING|findstr :" & ListenPort.ToString
                Else
                    strCmd = "netstat -apn|grep tcp|grep LISTEN|grep :" & ListenPort.ToString
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd)
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
                If OutPID < 0 Then
                    Throw New Exception("Cannot get process number.")
                End If
            End With
            Return "OK"
        Catch ex As Exception
            OutPID = -1
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetOSCaption(ByRef OutOSCaption As String) As String
        Dim LOG As New PigStepLog("GetOSCaption")
        Try
            Dim oPigCmdApp As New PigCmdApp, strCmd As String
            With oPigCmdApp
                .StandardOutputReadType = PigCmdApp.EnmStandardOutputReadType.StringArray
                If Me.IsWindows = True Then
                    strCmd = "wmic os get Caption"
                Else
                    strCmd = "cat /etc/os-release|grep ""PRETTY_NAME="""
                End If
                LOG.StepName = "CmdShell"
                LOG.Ret = .CmdShell(strCmd)
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


End Class
