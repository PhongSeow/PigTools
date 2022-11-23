'**********************************
'* Name: PigHttpContext
'* Author: Seow Phong
'* License: Copyright (c) 2019-2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Registry Processing Class|注册表处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 5/11/2019
'* 1.0.2  2019-11-7   修改BUG
'* 1.0.3  15/4/2021   Add to PigToolsWinLib
'* 1.1  15/4/2021   Modify mOpenRegPath
'* 1.2  20/8/2021   Use Throw New Exception
'************************************
Imports Microsoft.Win32
''' <summary>
''' Registry Processing Class|注册表处理类
''' </summary>
Public Class PigReg
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.3"
    ''' <summary>注册表的根区</summary>
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

    ''' <summary>获取什么注册表项</summary>
    Public Enum EmnGetWhatRegItem
        ''' <summary>Windows的产品ID</summary>
        WinProductId = 10
        ''' <summary>机器GUID，ID在重装Windows系统后应该不一样了</summary>
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

    ''' <summary>是否64位程序</summary>
    Public ReadOnly Property Is64Bit As Boolean
        Get
            If System.Runtime.InteropServices.Marshal.SizeOf(IntPtr.Zero) * 8 = 64 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ''' <summary>读取特定注册表值</summary>
    ''' <param name="WhatReg">获取什么注册表项</param>
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
                    Err.Raise(-1, , "无效WhatReg" & WhatReg.ToString)
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

    ''' <summary>读取注册表值</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <param name="RegName">项名</param>
    ''' <param name="DefaValue">默认值，取不到则返回这个</param>
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

    ''' <summary>读取注册表值</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <param name="RegName">项名</param>
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

    ''' <summary>读取注册表值</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <param name="RegName">项名</param>
    ''' <param name="DefaValue">默认值，取不到则返回这个，传空则不取</param>
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


    ''' <summary>读取注册表值（内部）</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <param name="RegName">项名</param>
    ''' <param name="DefaValue">默认值，取不到则返回这个，传空则不取</param>
    ''' <param name="TxRes">键名路径</param>
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

    ''' <summary>判断注册表键是否存在</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
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

    ''' <summary>创建注册表键</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <returns>操作结果</returns>
    ''' <remarks></remarks>
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
                    Err.Raise(-1, , "Invalid RegRoot")
            End Select
            strStepName = "CreateSubKey(" & RegPath & ")"
            rkAny.CreateSubKey(RegPath)
            rkAny.Close()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CreateRegKey", strStepName, ex)
        End Try
    End Function

    ''' <summary>打开注册表键（内部）</summary>
    ''' <param name="RegRoot">根区</param>
    ''' <param name="RegPath">键名路径</param>
    ''' <param name="IsReadOnly">是否只读</param>
    ''' <param name="TxRes">操作结果</param>
    ''' <returns>注册表键</returns>
    ''' <remarks></remarks>
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

    ''' <summary>设置注册表值（内部）</summary>
    ''' <param name="RegKey">注册表键</param>
    ''' <param name="RegName">项名</param>
    ''' <param name="RegValue">值</param>
    ''' <param name="ValueType">值类型</param>
    ''' <returns>操作结果</returns>
    ''' <remarks></remarks>
    Private Function mSetRegValue(RegKey As RegistryKey, RegName As String, RegValue As Object, ValueType As RegistryValueKind) As String
        Dim strStepName As String = ""
        Try
            strStepName = "SetValue(" & RegName & ")"
            RegKey.SetValue(RegName, RegValue, ValueType)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSetRegValue", strStepName, ex)
        End Try
    End Function

End Class
