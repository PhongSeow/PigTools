'**********************************
'* Name: PigCPU
'* Author: Seow Phong
'* License: Copyright (c) 2023-2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: CPU information class|CPU信息类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 16/1/2023
'* 1.1  19/1/2023  Modify Model,Processors,CPUCores
'* 1.2  10/1/2023  Add TotalProcessors
'* 1.3  18/8/2023  Add HostCpuUseRate,RefCpuActInf
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' CPU information class|CPU信息类
''' </summary>
Public Class PigCPU
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "18"

    Public ReadOnly Property Parent As PigHost

    Public Sub New(Parent As PigHost)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
    End Sub

    Private mHostCpuUseRate As Decimal
    Public Property HostCpuUseRate As Decimal
        Get
            Return mHostCpuUseRate
        End Get
        Friend Set(value As Decimal)
            mHostCpuUseRate = value
        End Set
    End Property


    ''' <summary>
    ''' CPU model|CPU型号
    ''' </summary>
    Private mModel As String
    Public Property Model As String
        Get
            Dim strRet As String = ""
            Try
                If mModel = "" Then
                    If Me.IsWindows = True Then
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.Model)
                    End If
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return mModel
            Catch ex As Exception
                Return ""
                Me.SetSubErrInf("Model.Get", ex)
            End Try
        End Get
        Friend Set(value As String)
            mModel = value
        End Set
    End Property

    ''' <summary>
    ''' Number of CPUs|CPU个数
    ''' </summary>
    Private mCPUs As Integer
    Public Property CPUs As Integer
        Get
            Dim strRet As String = ""
            Try
                If mCPUs <= 0 Then
                    If Me.IsWindows = True Then
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.CPUs)
                    End If
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return mCPUs
            Catch ex As Exception
                Return -1
                Me.SetSubErrInf("CPUs.Get", ex)
            End Try
        End Get
        Friend Set(value As Integer)
            mCPUs = value
        End Set
    End Property

    ''' <summary>
    ''' CPU cores|CPU核数
    ''' </summary>
    Private mCPUCores As Integer
    Public Property CPUCores As Integer
        Get
            Dim strRet As String = ""
            Try
                If mCPUCores <= 0 Then
                    If Me.IsWindows = True Then
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.CPUCores)
                    End If
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return mCPUCores
            Catch ex As Exception
                Return -1
                Me.SetSubErrInf("CPUCores.Get", ex)
            End Try
        End Get
        Friend Set(value As Integer)
            mCPUCores = value
        End Set
    End Property

    ''' <summary>
    ''' Total number of logical CPUs|总的逻辑CPU个数
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property TotalProcessors As Integer
        Get
            Return Me.CPUs * Me.Processors
        End Get
    End Property

    ''' <summary>
    ''' Number of logical CPUs|逻辑CPU的个数
    ''' </summary>
    ''' <returns></returns>
    Private mProcessors As Integer
    Public Property Processors As Integer
        Get
            Dim strRet As String = ""
            Try
                If mProcessors <= 0 Then
                    If Me.IsWindows = True Then
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.fGetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.Processors)
                    End If
                    If strRet <> "OK" Then Throw New Exception(strRet)
                End If
                Return mProcessors
            Catch ex As Exception
                Return -1
                Me.SetSubErrInf("Processors.Get", ex)
            End Try
        End Get
        Friend Set(value As Integer)
            mProcessors = value
        End Set
    End Property

    Public Function RefCpuActInf() As String
        Return Me.Parent.fRefCpuActInf(Me)
    End Function

    '''' <summary>
    '''' Obtain the current host CPU usage rate|获取当前主机CPU使用率
    '''' </summary>
    '''' <returns></returns>
    'Public Function GetHostCpuUseRate() As Decimal
    '    Dim LOG As New PigStepLog("GetHostCpuUseRate")
    '    Dim strCmd As String = ""
    '    Try
    '        If Me.IsWindows = True Then
    '            Dim pxMain As PigXml = Nothing
    '            strCmd = "cpu get loadpercentage"
    '            LOG.StepName = "GetWmicSimpleXml"
    '            LOG.Ret = Me.mPigSysCmd.GetWmicSimpleXml(strCmd, pxMain)
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            GetHostCpuUseRate = pxMain.XmlDocGetDec("Root.LoadPercentage") / 100
    '        Else
    '            strCmd = "top -b -n 1|grep %Cpu|awk '{print $4}'"
    '            LOG.StepName = "CmdShell"
    '            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            If Me.mPigCmdApp.StandardError <> "" Then
    '                LOG.AddStepNameInf(Me.mPigCmdApp.StandardOutput)
    '                LOG.Ret = Me.mPigCmdApp.StandardError
    '                Throw New Exception(LOG.Ret)
    '            End If
    '            GetHostCpuUseRate = Me.mPigFunc.GEDec(Me.mPigCmdApp.StandardOutput) / 100
    '        End If
    '        Return "OK"
    '    Catch ex As Exception
    '        GetHostCpuUseRate = -1
    '        LOG.AddStepNameInf(strCmd)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

    '''' <summary>
    '''' Obtain the current host memory usage rate|获取当前主机内存使用率
    '''' </summary>
    '''' <returns></returns>
    'Public Function GetHostMemUseRate() As Decimal
    '    Dim LOG As New PigStepLog("GetHostMemUseRate")
    '    Dim strCmd As String = ""
    '    Try
    '        If Me.IsWindows = True Then
    '            Dim pxMain As PigXml = Nothing, decTotalMem As Decimal = 0, deFreeMem As Decimal = 0
    '            strCmd = "os get freephysicalmemory,MaxProcessMemorySize,TotalVisibleMemorySize"
    '            LOG.StepName = "GetWmicSimpleXml"
    '            LOG.Ret = Me.mPigSysCmd.GetWmicSimpleXml(strCmd, pxMain)
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            deFreeMem = pxMain.XmlDocGetDec("Root.MaxProcessMemorySize")
    '            decTotalMem = pxMain.XmlDocGetDec("Root.MaxProcessMemorySize")
    '        Else
    '            strCmd = "top -b -n 1|grep %Cpu|awk '{print $4}'"
    '            LOG.StepName = "CmdShell"
    '            LOG.Ret = Me.mPigCmdApp.CmdShell(strCmd, PigCmdApp.EnmStandardOutputReadType.FullString)
    '            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
    '            If Me.mPigCmdApp.StandardError <> "" Then
    '                LOG.AddStepNameInf(Me.mPigCmdApp.StandardOutput)
    '                LOG.Ret = Me.mPigCmdApp.StandardError
    '                Throw New Exception(LOG.Ret)
    '            End If
    '            GetHostMemUseRate = Me.mPigFunc.GEDec(Me.mPigCmdApp.StandardOutput) / 100
    '        End If
    '        Return "OK"
    '    Catch ex As Exception
    '        GetHostMemUseRate = -1
    '        LOG.AddStepNameInf(strCmd)
    '        Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
    '    End Try
    'End Function

End Class
