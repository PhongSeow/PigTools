'**********************************
'* Name: PigHttpContext
'* Author: Seow Phong
'* License: Copyright (c) 2019-2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Registry Processing Class|注册表处理类
'* Home Url: https://en.seowphong.com
'* Version: 1.3
'* Create Time: 5/11/2019
'* 1.0.2  2019-11-7   修改BUG
'* 1.0.3  15/4/2021   Add to PigToolsWinLib
'* 1.1  15/4/2021   Modify mOpenRegPath
'* 1.2  20/8/2021   Use Throw New Exception
'* 1.3  19/5/2024   Add mSaveRegValue,SaveRegValue
'* 1.5  19/5/2024   Modify mSaveRegValue,SaveRegValue
'************************************
Imports Microsoft.Win32
''' <summary>
''' Registry Processing Class|注册表处理类
''' </summary>
Public Class PigReg
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "28"
    ''' <summary>The root of the registry|注册表的根区</summary>
    Public Enum EmnRegRoot
        ''' <summary>HKEY_CLASSES_ROOT</summary>
        CLASSES_ROOT = &H80000000
        ''' <summary>HKEY_CURRENT_USER</summary>
        CURRENT_USER = &H80000001
        ''' <summary>HKEY_LOCAL_MACHINE</summary>
        LOCAL_MACHINE = &H80000002
        ''' <summary>HKEY_USERS</summary>
        USERS = &H80000003
        ''' <summary>HKEY_PERFORMANCE_DATA</summary>
        PERFORMANCE_DATA = &H80000004
    End Enum

    ''' <summary>What registry key to obtain|获取什么注册表项</summary>
    Public Enum EmnGetWhatRegItem
        ''' <summary>Product ID for Windows|Windows的产品ID</summary>
        WinProductId = 10
        ''' <summary>The machine Guid and ID should be different after reinstalling the Windows system|机器GUID，ID在重装Windows系统后应该不一样了</summary>
        MachineGUID = 20
    End Enum

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Try
            If Me.IsWindows = False Then Throw New Exception("This class only supports windows.")
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    ''' <summary>Is it a 64 bit program|是否64位程序</summary>
    Public ReadOnly Property Is64Bit As Boolean
        Get
            If System.Runtime.InteropServices.Marshal.SizeOf(IntPtr.Zero) * 8 = 64 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>Read specific registry values|读取特定注册表值</summary>
    ''' <param name="WhatReg">What registry key to obtain|获取什么注册表项</param>
    Public Overloads Function GetSomeRegValue(WhatReg As EmnGetWhatRegItem) As String
        Dim strStepName As String = ""
        Try
            Dim strRet As String = ""
            Dim intRegRoot As EmnRegRoot, strRegPath As String = "", strRegName As String = ""
            Select Case WhatReg
                Case EmnGetWhatRegItem.MachineGUID
                    intRegRoot = EmnRegRoot.LOCAL_MACHINE
                    strRegPath = "SOFTWARE\Microsoft\Cryptography"
                    strRegName = "MachineGuid"
                Case EmnGetWhatRegItem.WinProductId
                    intRegRoot = EmnRegRoot.LOCAL_MACHINE
                    strRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                    strRegName = "ProductId"
                Case Else
                    Throw New Exception("Invalid WhatReg" & WhatReg.ToString)
            End Select
            strStepName = "mGetRegValue(" & strRegPath & "," & strRegName & ")"
            GetSomeRegValue = Me.mGetRegValue(intRegRoot, strRegPath, strRegName, "", strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetSomeRegValue", strStepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>Read registry values|读取注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Overloads Function GetRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As String) As String
        Try
            Dim strRet As String = ""
            GetRegValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegValue", ex)
            Return ""
        End Try
    End Function

    ''' <summary>Read registry values|读取注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Overloads Function GetRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Byte()
        Try
            Dim strRet As String = ""
            GetRegValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values|读取注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Overloads Function GetRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Object) As Object
        Try
            Dim strRet As String = ""
            GetRegValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegValue", ex)
            Return Nothing
        End Try
    End Function


    Private Function mGetRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Object, ByRef TxRes As String) As Object
        Dim strStepName As String = ""
        Try
            strStepName = "mOpenRegPath"
            Dim rkAny As RegistryKey = Me.mOpenRegPath(RegRoot, RegPath, True, TxRes)
            If TxRes <> "OK" Then Throw New Exception(TxRes)
            If rkAny Is Nothing Then Throw New Exception("Failed to get registry key")
            strStepName = "GetValue(" & RegName & ")"
            If DefaValue Is Nothing Then
                mGetRegValue = rkAny.GetValue(RegName)
            Else
                mGetRegValue = rkAny.GetValue(RegName, DefaValue)
            End If
            TxRes = "OK"
            rkAny = Nothing
        Catch ex As Exception
            TxRes = Me.GetSubErrInf("GetRegValue", strStepName, ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Check if the registry key exists|判断注册表键是否存在</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    Public ReadOnly Property IsRegKeyExists(RegRoot As EmnRegRoot, RegPath As String) As Boolean
        Get
            Try
                Dim strRet As String = ""
                Dim rkAny As RegistryKey = Me.mOpenRegPath(RegRoot, RegPath, True, strRet)
                If rkAny Is Nothing Then
                    Return False
                ElseIf strRet <> "OK" Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property

    ''' <summary>Create registry key|创建注册表键</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    Public Function CreateRegKey(RegRoot As EmnRegRoot, RegPath As String) As String
        Dim strStepName As String = ""
        Try
            Dim rkAny As RegistryKey
            Select Case RegRoot
                Case EmnRegRoot.CLASSES_ROOT
                    rkAny = Registry.ClassesRoot
                Case EmnRegRoot.CURRENT_USER
                    rkAny = Registry.CurrentUser
                Case EmnRegRoot.LOCAL_MACHINE
                    rkAny = Registry.LocalMachine
                Case EmnRegRoot.PERFORMANCE_DATA
                    rkAny = Registry.PerformanceData
                Case EmnRegRoot.USERS
                    rkAny = Registry.Users
                Case Else
                    rkAny = Nothing
                    Throw New Exception("Invalid RegRoot")
            End Select
            strStepName = "CreateSubKey(" & RegPath & ")"
            rkAny.CreateSubKey(RegPath)
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CreateRegKey", strStepName, ex)
        End Try
    End Function

    Private Function mOpenRegPath(RegRoot As EmnRegRoot, RegPath As String, IsReadOnly As Boolean, ByRef TxRes As String) As RegistryKey
        Dim strStepName As String = ""
        Try
            Dim rkRoot As RegistryKey
            Select Case RegRoot
                Case EmnRegRoot.CLASSES_ROOT
                    rkRoot = Registry.ClassesRoot
                Case EmnRegRoot.CURRENT_USER
                    rkRoot = Registry.CurrentUser
                Case EmnRegRoot.LOCAL_MACHINE
                    rkRoot = Registry.LocalMachine
                Case EmnRegRoot.PERFORMANCE_DATA
                    rkRoot = Registry.PerformanceData
                Case EmnRegRoot.USERS
                    rkRoot = Registry.Users
                Case Else
                    rkRoot = Nothing
                    Throw New Exception("Invalid RegRoot")
            End Select
            strStepName = "OpenSubKey(" & RegPath & ")"
            mOpenRegPath = rkRoot.OpenSubKey(RegPath, IsReadOnly)
            TxRes = "OK"
        Catch ex As Exception
            TxRes = Me.GetSubErrInf("mOpenRegPath", strStepName, ex)
            Return Nothing
        End Try
    End Function

    Private Function mSaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, RegValue As Object) As String
        Dim LOG As New PigStepLog("mSaveRegValue")
        Try
            LOG.StepName = "mOpenRegPath"
            Dim rkAny As RegistryKey = Me.mOpenRegPath(RegRoot, RegPath, False, LOG.Ret)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If rkAny Is Nothing Then Throw New Exception("Failed to get registry key")
            LOG.StepName = "SetValue"
            rkAny.SetValue(RegName, RegValue)
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf("RegRoot=" & RegRoot.ToString)
            LOG.AddStepNameInf("RegPath=" & RegPath)
            LOG.AddStepNameInf("RegName=" & RegName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="StrValue">String value|字符串值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, StrValue As String) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, StrValue)
    End Function


    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="IntValue">Integer value|整型值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, IntValue As Integer) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, IntValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="LongValue">Long value|长整型值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, LongValue As Long) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, LongValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="BooleanValue">Boolean value|布尔值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, BooleanValue As Boolean) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, BooleanValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DecValue">Decimal value|小数值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DecValue As Decimal) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, DecValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="BytesValue">Byte array value|字节数组值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, BytesValue As Byte()) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, BytesValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DateValue">Date value|日期值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DateValue As Date) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, DateValue)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="MultiStringValue">Multi string value|字符串数组 </param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, MultiStringValue As String()) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, MultiStringValue)
    End Function

End Class
