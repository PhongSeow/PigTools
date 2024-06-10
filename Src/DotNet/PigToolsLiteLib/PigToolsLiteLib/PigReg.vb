'**********************************
'* Name: PigHttpContext
'* Author: Seow Phong
'* License: Copyright (c) 2019-2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Registry Processing Class|注册表处理类
'* Home Url: https://en.seowphong.com
'* Version: 1.6
'* Create Time: 5/11/2019
'* 1.0.2  2019-11-7   修改BUG
'* 1.0.3  15/4/2021   Add to PigToolsWinLib
'* 1.1  15/4/2021   Modify mOpenRegPath
'* 1.2  20/8/2021   Use Throw New Exception
'* 1.3  19/5/2024   Add mSaveRegValue,SaveRegValue
'* 1.5  19/5/2024   Modify mSaveRegValue,SaveRegValue
'* 1.6  9/6/2024   Modify mSaveRegValue,SaveRegValue, add DeleteRegKey,GetRegValue...
'************************************
Imports Microsoft.Win32
''' <summary>
''' Registry Processing Class|注册表处理类
''' </summary>
Public Class PigReg
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "6" & "." & "68"
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

    Private ReadOnly Property mPigFunc As New PigFunc

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



    ''' <summary>Read registry values(string)|读取注册表值(字符串)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegStrValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As String) As String
        Try
            Dim strRet As String = ""
            GetRegStrValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegStrValue", ex)
            Return ""
        End Try
    End Function

    ''' <summary>Read registry values(string)|读取注册表值(字符串)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegStrValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As String
        Try
            Dim strRet As String = ""
            GetRegStrValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegStrValue", ex)
            Return ""
        End Try
    End Function

    ''' <summary>Read registry values(string array)|读取注册表值(字符串数组)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegMultiStrValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As String()) As String()
        Try
            Dim strRet As String = ""
            GetRegMultiStrValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegMultiStrValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values(string array)|读取注册表值(字符串数组)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegMultiStrValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As String()
        Try
            Dim strRet As String = ""
            GetRegMultiStrValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegMultiStrValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values(byte array)|读取注册表值(字节数组)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegBytesValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Byte()) As Byte()
        Try
            Dim strRet As String = ""
            GetRegBytesValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegBytesValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values(byte array)|读取注册表值(字节数组)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegBytesValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Byte()
        Try
            Dim strRet As String = ""
            GetRegBytesValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegBytesValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values(decimal)|读取注册表值(浮点型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegDecValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Decimal) As Decimal
        Try
            Dim strRet As String = ""
            GetRegDecValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegDecValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Read registry values(decimal)|读取注册表值(浮点型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegDecValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Decimal
        Try
            Dim strRet As String = ""
            GetRegDecValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegDecValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Read registry values(integer)|读取注册表值(整型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegLongValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Long) As Long
        Try
            Dim strRet As String = ""
            GetRegLongValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegLongValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Read registry values(integer)|读取注册表值(整型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegLongValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Long
        Try
            Dim strRet As String = ""
            GetRegLongValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegLongValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Read registry values(long integer)|读取注册表值(长整型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegIntValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Integer) As Integer
        Try
            Dim strRet As String = ""
            GetRegIntValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegIntValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Read registry values(long integer)|读取注册表值(长整型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegIntValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Integer
        Try
            Dim strRet As String = ""
            GetRegIntValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegIntValue", ex)
            Return 0
        End Try
    End Function


    ''' <summary>Read registry values(boolean)|读取注册表值(布尔型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegBooleanValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Boolean) As Boolean
        Try
            Dim strRet As String = ""
            GetRegBooleanValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegBooleanValue", ex)
            Return False
        End Try
    End Function

    ''' <summary>Read registry values(boolean)|读取注册表值(布尔型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegBooleanValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Boolean
        Try
            Dim strRet As String = ""
            GetRegBooleanValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegBooleanValue", ex)
            Return False
        End Try
    End Function

    ''' <summary>Read registry values(date)|读取注册表值(日期型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DefaValue">Default value, if not found, return this value|默认值，取不到则返回这个</param>
    Public Function GetRegDateValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DefaValue As Date) As Date
        Try
            Dim strRet As String = ""
            GetRegDateValue = Me.mGetRegValue(RegRoot, RegPath, RegName, DefaValue, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegDateValue", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>Read registry values(date)|读取注册表值(日期型)</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function GetRegDateValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As Date
        Try
            Dim strRet As String = ""
            GetRegDateValue = Me.mGetRegValue(RegRoot, RegPath, RegName, Nothing, strRet)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetRegDateValue", ex)
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
        Dim LOG As New PigStepLog("CreateRegKey")
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
            LOG.StepName = "CreateSubKey"
            rkAny.CreateSubKey(RegPath)
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(RegRoot.ToString)
            LOG.AddStepNameInf(RegPath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
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
            mOpenRegPath = rkRoot.OpenSubKey(RegPath, Not IsReadOnly)
            TxRes = "OK"
        Catch ex As Exception
            TxRes = Me.GetSubErrInf("mOpenRegPath", strStepName, ex)
            Return Nothing
        End Try
    End Function

    Private Function mSaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, RegValue As Object, RegValueKind As RegistryValueKind) As String
        Dim LOG As New PigStepLog("mSaveRegValue")
        Try
            LOG.StepName = "mOpenRegPath"
            Dim rkAny As RegistryKey = Me.mOpenRegPath(RegRoot, RegPath, False, LOG.Ret)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            If rkAny Is Nothing Then Throw New Exception("Failed to get registry key")
            LOG.StepName = "SetValue"
            '如何判断 RegValue 的数据类型
            rkAny.SetValue(RegName, RegValue, RegValueKind)
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
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, StrValue, RegistryValueKind.String)
    End Function


    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="IntValue">Integer value|整型值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, IntValue As Integer) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, IntValue, RegistryValueKind.DWord)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="LongValue">Long value|长整型值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, LongValue As Long) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, LongValue, RegistryValueKind.QWord)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="BooleanValue">Boolean value|布尔值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, BooleanValue As Boolean) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, CInt(BooleanValue), RegistryValueKind.DWord)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DecValue">Decimal value|小数值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DecValue As Decimal) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, CStr(DecValue), RegistryValueKind.String)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="BytesValue">Byte array value|字节数组值</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, BytesValue As Byte()) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, BytesValue, RegistryValueKind.Binary)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="DateValue">Date value|日期值</param>
    ''' <param name="TimeFmt">Date format|日期格式</param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, DateValue As Date, Optional TimeFmt As String = "yyyy-MM-dd HH:mm:ss.fff") As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, Me.mPigFunc.GetFmtDateTime(DateValue, TimeFmt), RegistryValueKind.String)
    End Function

    ''' <summary>Save registry values|保存注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    ''' <param name="MultiStringValue">Multi string value|字符串数组 </param>
    Public Function SaveRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String, MultiStringValue As String()) As String
        Return Me.mSaveRegValue(RegRoot, RegPath, RegName, MultiStringValue, RegistryValueKind.MultiString )
    End Function

    ''' <summary>Delete registry key|删除注册表键</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    Public Function DeleteRegKey(RegRoot As EmnRegRoot, RegPath As String) As String
        Dim LOG As New PigStepLog("DeleteRegKey")
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
            LOG.StepName = "DeleteSubKeyTree"
            rkAny.DeleteSubKey(RegPath)
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(RegRoot.ToString)
            LOG.AddStepNameInf(RegPath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try

    End Function

    ''' <summary>Get the value of the registry|获取注册表值</summary>
    ''' <param name="RegRoot">Registry Root|根区</param>
    ''' <param name="RegPath">Registry key name path|键名路径</param>
    ''' <param name="RegName">Registry Name|项名</param>
    Public Function DeleteRegValue(RegRoot As EmnRegRoot, RegPath As String, RegName As String) As String
        Dim LOG As New PigStepLog("DeleteRegValue")
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
            LOG.StepName = "OpenSubKey"
            Dim rkSub As RegistryKey = rkAny.OpenSubKey(RegPath, True)
            If rkSub Is Nothing Then
                rkAny.Close()
                Return "OK"
            End If
            LOG.StepName = "DeleteValue"
            rkSub.DeleteValue(RegName)
            rkSub.Close()
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(RegRoot.ToString)
            LOG.AddStepNameInf(RegPath)
            LOG.AddStepNameInf(RegName)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
