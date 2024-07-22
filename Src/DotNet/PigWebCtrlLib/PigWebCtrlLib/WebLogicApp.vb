'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2022-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Application of dealing with Weblogic
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.17
'* Create Time: 31/1/2022
'* 1.1  5/2/2022   Add GetJavaVersion 
'* 1.2  6/3/2022   Add WlstPath 
'* 1.3  26/5/2022  Add SaveSecurityBoot ,modify New
'* 1.4  27/5/2022  Modify SaveSecurityBoot
'* 1.5  1/6/2022  Add IsWindows
'* 1.6  5/6/2022  Add StartOrStopTimeout
'* 1.7  26/7/2022 Modify Imports
'* 1.8  29/7/2022 Modify Imports
'* 1.9  7/9/2022  Add RunOpatch
'* 1.10 28/9/2022 Modify StartOrStopTimeout
'* 1.11 24/6/2023 Change the reference to PigObjFsLib to PigToolsLiteLib
'* 1.12 7/12/2023 Add RunOpatch
'* 1.13 7/6/2024 Add JdkHomeDir,mSetJdkHomeDir, modify GetJavaVersion
'* 1.15 24/6/2024 Modify GetJavaVersion,add GetJMapHeapXml
'* 1.16 26/6/2024 Modify GetJMapHeapXml
'* 1.17 27/6/2024 Modify GetJMapHeapXml, add GetJStatGcUtilXml
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib

''' <summary>
''' WebLogic Processing Class|WebLogic处理类
''' </summary>
Public Class WebLogicApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "17" & "." & "6"
    Public ReadOnly Property HomeDirPath As String
    Public ReadOnly Property WorkTmpDirPath As String
    Public ReadOnly Property CallWlstTimeout As Integer = 300
    Public ReadOnly Property StartOrStopTimeout As Integer = 300

    Public Property WebLogicDomains As WebLogicDomains

    Private WithEvents mPigCmdApp As New PigCmdApp
    Private mFS As New PigFileSystem
    Private mPigFunc As New PigFuncLite

    Private mGetJavaVersionThreadID As Integer

    Private mJdkHomeDir As String = ""
    Public Property JdkHomeDir As String
        Get
            Return mJdkHomeDir
        End Get
        Friend Set(value As String)
            Try
                If value = "" Then
                    Dim strPath As String = Me.mPigFunc.GetEnvVar("PATH")
                    Dim strItemSep As String
                    If Me.IsWindows = True Then
                        strItemSep = ";"
                    Else
                        strItemSep = ":"
                    End If
                    strPath = strItemSep & strPath & strItemSep
                    Dim strLeft As String = Me.OsPathSep & "jdk", strRight As String = Me.OsPathSep & "bin" & strItemSep
                    value = Me.mPigFunc.GetStr(strPath, strLeft, strRight, 1)
                    If value <> "" Then
                        value = Me.mPigFunc.GetStr(strPath, strItemSep, strRight, 1)
                        value &= Me.OsPathSep & "bin"
                    End If
                End If
                mJdkHomeDir = value
            Catch ex As Exception
                Me.SetSubErrInf("JdkHomeDir.Set", ex)
            End Try
        End Set
    End Property

    Public Sub New(HomeDirPath As String, WorkTmpDirPath As String)
        MyBase.New(CLS_VERSION)
        Try
            Me.WebLogicDomains = New WebLogicDomains
            Me.WebLogicDomains.fParent = Me
            Me.HomeDirPath = HomeDirPath
            Me.WorkTmpDirPath = WorkTmpDirPath
            Me.CallWlstTimeout = CallWlstTimeout
            Me.JdkHomeDir = ""
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub


    Public Sub New(HomeDirPath As String, WorkTmpDirPath As String, JdkHomeDir As String)
        MyBase.New(CLS_VERSION)
        Try
            Me.WebLogicDomains = New WebLogicDomains
            Me.WebLogicDomains.fParent = Me
            Me.HomeDirPath = HomeDirPath
            Me.WorkTmpDirPath = WorkTmpDirPath
            Me.CallWlstTimeout = CallWlstTimeout
            Me.JdkHomeDir = JdkHomeDir
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Private Function mGetExeHead(CmdName As String) As String
        Try
            mGetExeHead = ""
            If Me.JdkHomeDir <> "" Then
                mGetExeHead = Me.JdkHomeDir & Me.OsPathSep & CmdName
                If Me.IsWindows = True Then
                    mGetExeHead = """" & mGetExeHead & ".exe"""
                End If
            Else
                mGetExeHead = CmdName
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetExeHead", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get java version information|获取java版本信息
    ''' </summary>
    ''' <returns></returns>
    Public Function GetJavaVersion() As String
        Dim LOG As New PigStepLog("GetJavaVersion")
        Try
            Dim strCmd As String = ""
            LOG.StepName = "mGetExeHead"
            strCmd = Me.mGetExeHead("java")
            If strCmd = "" Then
                Throw New Exception("Unable to obtain command head.")
            End If
            strCmd &= " -version"
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            Me.JavaVersion = Me.mPigCmdApp.StandardError
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private mstrJavaVersion As String
    Public Property JavaVersion As String
        Get
            Return mstrJavaVersion
        End Get
        Friend Set(value As String)
            mstrJavaVersion = value
        End Set
    End Property

    ''' <summary>
    ''' 操作超时时间，单位为秒|Operation timeout in seconds
    ''' </summary>
    Private mintOperationTimeout As Integer = 60
    Public Property OperationTimeout As Decimal
        Get
            Return mintOperationTimeout
        End Get
        Friend Set(value As Decimal)
            mintOperationTimeout = value
        End Set
    End Property

    Public ReadOnly Property WlsJarPath() As String
        Get
            WlsJarPath = Me.HomeDirPath & Me.OsPathSep & "wlserver" & Me.OsPathSep & "common" & Me.OsPathSep & "templates" & Me.OsPathSep & "wls" & Me.OsPathSep & "wls.jar"
        End Get
    End Property

    Public ReadOnly Property WlstPath() As String
        Get
            WlstPath = Me.HomeDirPath & Me.OsPathSep & "oracle_common" & Me.OsPathSep & "common" & Me.OsPathSep & "bin" & Me.OsPathSep & "wlst."
            If Me.IsWindows = True Then
                WlstPath &= "cmd"
            Else
                WlstPath &= "sh"
            End If
        End Get
    End Property



    Public Overloads ReadOnly Property IsWindows() As String
        Get
            Return MyBase.IsWindows
        End Get
    End Property

    Public Overloads Sub SetDebug(DebugFilePath)
        MyBase.SetDebug(DebugFilePath)
    End Sub

    Public Overloads Sub PrintDebugLog(SubName As String, StepName As String, LogInf As String)
        MyBase.PrintDebugLog(SubName, StepName, LogInf)
    End Sub

    ''' <summary>
    ''' Run patch script|运行补丁脚本
    ''' </summary>
    ''' <param name="Cmd">script|脚本</param>
    ''' <param name="ResInf">Return Results|返回结果</param>
    ''' <returns></returns>
    Public Function RunOpatch(Cmd As String, ByRef ResInf As String) As String
        Dim LOG As New PigStepLog("RunOpatch")
        Try
            Dim strCmd As String = Me.HomeDirPath & Me.OsPathSep & "OPatch" & Me.OsPathSep & "opatch " & Cmd
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            ResInf = Me.mPigCmdApp.StandardOutput & Me.OsCrLf & Me.mPigCmdApp.StandardError
            Return "OK"
        Catch ex As Exception
            ResInf = ""
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

    ''' <summary>
    ''' Run patch script|运行补丁脚本
    ''' </summary>
    ''' <param name="SudoUser">sudo user|sudo 用户</param>
    ''' <param name="Cmd">script|脚本</param>
    ''' <param name="ResInf">Return Results|返回结果</param>
    ''' <returns></returns>
    Public Function RunOpatch(SudoUser As String, Cmd As String, ByRef ResInf As String) As String
        Dim LOG As New PigStepLog("RunOpatch")
        Try
            If Me.IsWindows = True Then Throw New Exception("Can only be executed on the Linux platform")
            Dim strCmd As String = Me.HomeDirPath & Me.OsPathSep & "OPatch" & Me.OsPathSep & "opatch " & Cmd
            LOG.StepName = "PigSudo.Run"
            Dim oPigSudo As New PigSudo(strCmd, SudoUser)
            LOG.Ret = oPigSudo.Run()
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            End If
            ResInf = oPigSudo.StandardOutput & Me.OsCrLf & oPigSudo.StandardError
            oPigSudo = Nothing
            Return "OK"
        Catch ex As Exception
            ResInf = ""
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function



    ''' <summary>
    ''' Obtain the XML output result of jstat -gcutil|获取 jstat -gcutil 的xml输出结果
    ''' </summary>
    ''' <param name="JavaPID">Java process number|Java进程号</param>
    ''' <param name="OutPigXml">Output PigXml object|输出的PigXml对象</param>
    ''' <param name="RootKeyName">The name of the root node|根结点的名称</param>
    ''' <returns></returns>
    Public Function GetJStatGcUtilXml(JavaPID As Integer, ByRef OutPigXml As PigXml, Optional RootKeyName As String = "Root") As String
        Dim LOG As New PigStepLog("GetJStatGcUtilXml")
        Try
            Dim strCmd As String = ""
            LOG.StepName = "mGetExeHead"
            strCmd = Me.mGetExeHead("jstat")
            If strCmd = "" Then
                Throw New Exception("Unable to obtain command head.")
            End If
            strCmd &= " -gcutil " & JavaPID
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            Dim strTitle As String = Trim(Me.mPigCmdApp.StandardOutputArray(0))
            Dim strValue As String = Trim(Me.mPigCmdApp.StandardOutputArray(1))
            Do While InStr(strTitle, "  ") > 0
                strTitle = Replace(strTitle, "  ", " ")
            Loop
            strTitle = "<" & Replace(strTitle, " ", "><") & ">"
            Do While InStr(strValue, "  ") > 0
                strValue = Replace(strValue, "  ", " ")
            Loop
            strValue = "<" & Replace(strValue, " ", "><") & ">"
            LOG.StepName = "New PigXml"
            OutPigXml = New PigXml(True)
            OutPigXml.AddEleLeftSign(RootKeyName)
            Do While True
                Dim strKey As String = Me.mPigFunc.GetStr(strTitle, "<", ">")
                If strKey = "" Then Exit Do
                Dim strKeyValue As String = Me.mPigFunc.GetStr(strValue, "<", ">")
                OutPigXml.AddEle(strKey, strKeyValue)
            Loop
            OutPigXml.AddEle("GcUtilTime", Me.mPigFunc.GetFmtDateTime(Now))
            OutPigXml.AddEle("JavaPID", JavaPID)
            OutPigXml.AddEleRightSign(RootKeyName)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' Obtain the XML output result of jmap -heap|获取 jmap -heap 的xml输出结果
    ''' </summary>
    ''' <param name="JavaPID">Java process number|Java进程号</param>
    ''' <param name="OutPigXml">Output PigXml object|输出的PigXml对象</param>
    ''' <param name="RootKeyName">The name of the root node|根结点的名称</param>
    ''' <returns></returns>
    Public Function GetJMapHeapXml(JavaPID As Integer, ByRef OutPigXml As PigXml, Optional RootKeyName As String = "Root") As String
        Dim LOG As New PigStepLog("GetJMapHeapXml")
        Try
            Dim strCmd As String = ""
            LOG.StepName = "mGetExeHead"
            strCmd = Me.mGetExeHead("jmap")
            If strCmd = "" Then
                Throw New Exception("Unable to obtain command head.")
            End If
            strCmd &= " -heap " & JavaPID
            LOG.StepName = "CmdShell"
            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.StringArray)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(LOG.Ret)
            ElseIf Me.mPigCmdApp.StandardError <> "" Then
                LOG.AddStepNameInf(strCmd)
                Throw New Exception(Me.mPigCmdApp.StandardError)
            End If
            LOG.StepName = "New PigXml"
            OutPigXml = New PigXml(True)
            Dim bolIsOK As Boolean = False, bolHeapConfiguration As Boolean = False, bolIsYoungGeneration As Boolean = False, bolOldGeneration As Boolean = False
            OutPigXml.AddEleLeftSign(RootKeyName)
            For i = 0 To Me.mPigCmdApp.StandardOutputArray.Length - 1
                Dim strLine As String = Me.mPigCmdApp.StandardOutputArray(i)
                If bolIsOK = True Then
                    Dim strLeftKey As String = ""
                    Dim strXmlKey As String = "", strXmlValue As String = ""
                    If bolHeapConfiguration = False Then
                        strLeftKey = "JVM version is "
                        If Left(strLine, Len(strLeftKey)) = strLeftKey Then
                            strLine &= Me.OsCrLf
                            Dim strJvmVer As String = Me.mPigFunc.GetStr(strLine, strLeftKey, Me.OsCrLf)
                            OutPigXml.AddEle("JVM_Version", strJvmVer)
                        End If
                        If strLine = "Heap Configuration:" Then
                            OutPigXml.AddEleLeftSign("HeapConfiguration")
                            bolHeapConfiguration = True
                        End If
                    ElseIf bolIsYoungGeneration = False And bolHeapConfiguration = True Then
                        If strLine = "PS Young Generation" Then
                            OutPigXml.AddEleRightSign("HeapConfiguration")
                            OutPigXml.AddEleLeftSign("PSYoungGeneration")
                            bolIsYoungGeneration = True
                        Else
                            strLeftKey = " = "
                            strLine &= Me.OsCrLf
                            If InStr(strLine, strLeftKey) > 0 Then
                                strXmlKey = Trim(Me.mPigFunc.GetStr(strLine, "", strLeftKey, 1))
                                strXmlValue = Trim(Me.mPigFunc.GetStr(strLine, strLeftKey, Me.OsCrLf))
                                If InStr(strXmlValue, "(") > 0 Then strXmlValue = Me.mPigFunc.GetStr(strXmlValue, "", " (")
                                OutPigXml.AddEle(strXmlKey, strXmlValue, 1)
                            End If
                        End If
                    ElseIf bolIsYoungGeneration = True And bolOldGeneration = False Then
                        If strLine = "Eden Space:" Then
                            OutPigXml.AddEleLeftSign("EdenSpace")
                        ElseIf strLine = "From Space:" Then
                            OutPigXml.AddEleRightSign("EdenSpace")
                            OutPigXml.AddEleLeftSign("FromSpace")
                        ElseIf strLine = "To Space:" Then
                            OutPigXml.AddEleRightSign("FromSpace")
                            OutPigXml.AddEleLeftSign("ToSpace")
                        ElseIf strLine = "PS Old Generation" Then
                            OutPigXml.AddEleRightSign("ToSpace")
                            OutPigXml.AddEleRightSign("PSYoungGeneration")
                            OutPigXml.AddEleLeftSign("PSOldGeneration")
                            bolOldGeneration = True
                        Else
                            strLeftKey = " = "
                            strLine &= Me.OsCrLf
                            If InStr(strLine, strLeftKey) > 0 Then
                                strXmlKey = Trim(Me.mPigFunc.GetStr(strLine, "", strLeftKey, 1))
                                strXmlValue = Trim(Me.mPigFunc.GetStr(strLine, strLeftKey, Me.OsCrLf))
                                If InStr(strXmlValue, "(") > 0 Then strXmlValue = Me.mPigFunc.GetStr(strXmlValue, "", " (")
                                OutPigXml.AddEle(strXmlKey, strXmlValue, 1)
                            ElseIf InStr(strLine, "% used") > 0 Then
                                strXmlKey = "UsedRate"
                                strXmlValue = Trim(Me.mPigFunc.GetStr(strLine, "", "% used"))
                                OutPigXml.AddEle(strXmlKey, strXmlValue, 1)
                            End If
                        End If
                    ElseIf bolOldGeneration = True Then
                        strLeftKey = " = "
                        strLine &= Me.OsCrLf
                        If InStr(strLine, strLeftKey) > 0 Then
                            strXmlKey = Trim(Me.mPigFunc.GetStr(strLine, "", strLeftKey, 1))
                            strXmlValue = Trim(Me.mPigFunc.GetStr(strLine, strLeftKey, Me.OsCrLf))
                            If InStr(strXmlValue, "(") > 0 Then strXmlValue = Me.mPigFunc.GetStr(strXmlValue, "", " (")
                            OutPigXml.AddEle(strXmlKey, strXmlValue, 1)
                        ElseIf InStr(strLine, "% used") > 0 Then
                            strXmlKey = "UsedRate"
                            strXmlValue = Trim(Me.mPigFunc.GetStr(strLine, "", "% used"))
                            OutPigXml.AddEle(strXmlKey, strXmlValue, 1)
                            OutPigXml.AddEleRightSign("PSOldGeneration")
                        End If
                    End If
                ElseIf strLine = "Debugger attached successfully." Then
                    bolIsOK = True
                    OutPigXml.AddEle("RunRes", "OK")
                ElseIf Left(strLine, 6) = "Error " Then
                    OutPigXml.AddEle("RunRes", strLine)
                End If
            Next
            LOG.StepName = "Get MemoryUse"
            Dim oPigProc As New PigProc(JavaPID)
            OutPigXml.AddEle("ProcessMemory", oPigProc.MemoryUse)
            OutPigXml.AddEle("TotalProcessorTime", oPigProc.TotalProcessorTime.ToString)
            OutPigXml.AddEle("ThreadsCount", oPigProc.ThreadsCount)
            oPigProc.Close()
            OutPigXml.AddEle("LeapTime", Me.mPigFunc.GetFmtDateTime(Now))
            OutPigXml.AddEle("JavaPID", JavaPID)
            OutPigXml.AddEleRightSign(RootKeyName)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
