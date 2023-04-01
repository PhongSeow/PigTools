'**********************************
'* Name: PigCPU
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: CPU information class|CPU信息类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 1/16/2023
'* 1.1  1/19/2023  Modify Model,Processors,CPUCores
'* 1.2  1/20/2023  Add TotalProcessors
'**********************************
Imports PigToolsLiteLib
''' <summary>
''' CPU information class|CPU信息类
''' </summary>
Public Class PigCPU
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.2.6"

    Public ReadOnly Property Parent As PigHost

    Public Sub New(Parent As PigHost)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
    End Sub

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
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.Model)
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
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.CPUs)
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
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.CPUCores)
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
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.WmicGetAll)
                    Else
                        strRet = Me.Parent.GetCPUBaseInf(Me, PigHost.EnmGetCPUBaseWhat.Processors)
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

End Class
