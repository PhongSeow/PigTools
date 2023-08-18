'**********************************
'* Name: PigHost
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Host class|主机类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 1/16/2023
'* 1.1  1/18/2023  Modify GetCPUBaseInf,New
'* 1.2  1/19/2023  Modify GetCPUBaseInf,New
'* 1.3  1/20/2023  Modify New
'* 1.5  18/8/2023  Add fRefCpuActInf
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' Host class|主机类
''' </summary>
Public Class PigHost
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.5.2"

    Friend Enum EnmGetCPUBaseWhat
        Model = 0
        CPUs = 10
        CPUCores = 20
        Processors = 30
        WmicGetAll = 50
    End Enum


    Public ReadOnly Property IsInit As Boolean
    Public ReadOnly Property HostName As String
    Public ReadOnly Property HostID As String
    Public ReadOnly Property UUID As String
    Public ReadOnly Property CPU As PigCPU
    Public ReadOnly Property OSCaption As String
    Private ReadOnly Property mPigFunc As New PigFunc
    Private ReadOnly Property mPigSysCmd As New PigSysCmd
    Private ReadOnly Property mPigCmdApp As New PigCmdApp


    Public Sub New(Optional HostID As String = "")
        MyBase.New(CLS_VERSION)
        Dim LOG As New PigStepLog("New")
        Try
            Dim strTmp As String = ""
            If HostID <> "" Then
                Me.HostID = HostID
                LOG.Ret = Me.mPigSysCmd.GetUUID(strTmp)
                Me.UUID = strTmp
            Else
                LOG.StepName = "PigSysCmd.GetUUID"
                LOG.Ret = Me.mPigSysCmd.GetUUID(strTmp)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If strTmp = "" Then Throw New Exception("Unable to get UUID")
                Me.UUID = strTmp
                LOG.StepName = "GetTextPigMD5(HostID)"
                LOG.Ret = Me.mPigFunc.GetTextPigMD5(Me.UUID, PigMD5.enmTextType.UTF8, strTmp)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Me.HostID = UCase(strTmp)
            End If
            Me.HostName = Me.mPigFunc.GetComputerName
            LOG.StepName = "GetOSCaption"
            LOG.Ret = Me.mPigSysCmd.GetOSCaption(strTmp)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Me.OSCaption = strTmp
            Me.CPU = New PigCPU(Me)
            Me.IsInit = True
        Catch ex As Exception
            Me.IsInit = False
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Sub

    Friend Function fRefCpuActInf(ByRef InPigCPU As PigCPU) As String
        Dim LOG As New PigStepLog("fGetCPUBaseInf")
        Dim strCmd As String = ""
        Try
            If Me.IsWindows = True Then
                Dim pxMain As PigXml = Nothing
                strCmd = "cpu get loadpercentage"
                LOG.StepName = "GetWmicSimpleXml"
                LOG.Ret = Me.mPigSysCmd.GetWmicSimpleXml(strCmd, pxMain)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                InPigCPU.HostCpuUseRate = pxMain.XmlDocGetDec("WmicXml.Row1.LoadPercentage") / 100
            Else
                strCmd = "top -b -n 1|grep %Cpu|awk '{print $4}'"
                LOG.StepName = "CmdShell"
                LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If Me.mPigCmdApp.StandardError <> "" Then
                    LOG.AddStepNameInf(Me.mPigCmdApp.StandardOutput)
                    LOG.Ret = Me.mPigCmdApp.StandardError
                    Throw New Exception(LOG.Ret)
                End If
                InPigCPU.HostCpuUseRate = Me.mPigFunc.GEDec(Me.mPigCmdApp.StandardOutput) / 100
            End If
            Return "OK"
        Catch ex As Exception
            InPigCPU.HostCpuUseRate = -1
            LOG.AddStepNameInf(strCmd)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Friend Function fGetCPUBaseInf(ByRef InPigCPU As PigCPU, GetCPUBaseWhat As EnmGetCPUBaseWhat) As String
        Dim LOG As New PigStepLog("fGetCPUBaseInf")
        Try
            If InPigCPU Is Nothing Then Throw New Exception("InPigCPU is Nothing")
            Dim strRetContent As String = ""
            If Me.IsWindows = True Then
                Dim strXml As String = ""
                If GetCPUBaseWhat = EnmGetCPUBaseWhat.WmicGetAll Then
                    LOG.StepName = "cpu get Name,NumberOfCores,NumberOfLogicalProcessors"
                    LOG.Ret = Me.mPigSysCmd.GetWmicSimpleXml(LOG.StepName, strXml)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                    Dim oPigXml As New PigXml(False)
                    LOG.StepName = "SetMainXml"
                    oPigXml.SetMainXml(strXml)
                    If oPigXml.LastErr <> "" Then Throw New Exception(oPigXml.LastErr)
                    LOG.StepName = "InitXmlDocument"
                    LOG.Ret = oPigXml.InitXmlDocument()
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                    LOG.StepName = "XmlDocGetStr"
                    With InPigCPU
                        .Model = oPigXml.XmlDocGetStr("WmicXml.Row1.Name")
                        .CPUCores = oPigXml.XmlDocGetInt("WmicXml.Row1.NumberOfCores")
                        .Processors = oPigXml.XmlDocGetInt("WmicXml.Row1.NumberOfLogicalProcessors")
                        .CPUs = oPigXml.XmlDocGetInt("WmicXml.TotalRows")
                    End With
                    oPigXml = Nothing
                Else
                    Throw New Exception("Invalid GetCPUBaseWhat")
                End If
            Else
                Select Case GetCPUBaseWhat
                    Case EnmGetCPUBaseWhat.CPUCores
                        LOG.StepName = "cat /proc/cpuinfo|grep -w ""cpu cores""|uniq|awk -F"":"" '{print $2}'"
                    Case EnmGetCPUBaseWhat.CPUs
                        LOG.StepName = "cat /proc/cpuinfo|grep ""physical id""|sort|uniq| wc -l"
                    Case EnmGetCPUBaseWhat.Model
                        LOG.StepName = "cat /proc/cpuinfo|grep -w ""model name""|uniq -c|awk -F"":"" '{print $2}'"
                    Case EnmGetCPUBaseWhat.Processors
                        LOG.StepName = "cat /proc/cpuinfo|grep -w ""processor""|wc -l"
                    Case Else
                        Throw New Exception("Invalid GetCPUBaseWhat")
                End Select
                LOG.Ret = Me.mPigSysCmd.GetCmdRetRows(LOG.StepName, strRetContent, 1)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                strRetContent = Replace(Trim(strRetContent), vbCrLf, "")
                strRetContent = Replace(strRetContent, Me.OsCrLf, "")
                Select Case GetCPUBaseWhat
                    Case EnmGetCPUBaseWhat.CPUCores
                        InPigCPU.CPUCores = Me.mPigFunc.GEInt(strRetContent)
                    Case EnmGetCPUBaseWhat.CPUs
                        InPigCPU.CPUs = Me.mPigFunc.GEInt(strRetContent)
                    Case EnmGetCPUBaseWhat.Model
                        InPigCPU.Model = strRetContent
                    Case EnmGetCPUBaseWhat.Processors
                        InPigCPU.Processors = Me.mPigFunc.GEInt(strRetContent)
                End Select
            End If
            Return "OK"
        Catch ex As Exception
            Dim strRet As String = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return strRet
        End Try
    End Function


End Class
