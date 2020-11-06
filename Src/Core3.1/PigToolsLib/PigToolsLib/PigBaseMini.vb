'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Basic lightweight Edition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.10
'* Create Time: 31/8/2019
'*1.0.2  1/10/2019   Add mGetSubErrInf 
'*1.0.3  4/11/2019   Add LastErr
'*1.0.4  5/11/2019-11-5   Add SetSubErrInf
'*1.0.5  6/2/2020    Add GetSubStepDebugInf
'*1.0.6  3/3/2020    Add Debug function, in debug mode, you can also print log
'*1.0.7  3/4/2020    Add Hard debug function, modify GetSubStepDebugInf
'*1.0.8  3/6/2020    Add KeyInf
'*1.0.9  6/17/2020   modify mPrintDebugLog Add StepName
'*1.0.10 6/25/2020   Not used My.Application , better compatibility, Add MyClassName
'************************************
Public Class PigBaseMini
    Private mstrClsName As String    '类名
    Private mstrClsVersion As String    '类版本
    Private mstrLastErr As String
    Private mbolIsDebug As Boolean
    Private mbolIsHardDebug As Boolean
    Private mstrDebugFilePath As String
    Public KeyInf As String


    Public Sub New(Version As String)
        mstrClsName = Me.GetType.Name.ToString()
        mstrClsVersion = Version
    End Sub

    ''' <remarks>返回成功</remarks>
    Friend Sub ClearErr()
        mstrLastErr = ""
    End Sub

    ''' <summary>设置调试</summary>
    Public Overloads Sub SetDebug(DebugFilePath As String)
        mbolIsDebug = True
        mbolIsHardDebug = False
        mstrDebugFilePath = DebugFilePath
    End Sub

    ''' <summary>设置调试</summary>
    Public Overloads Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
        mbolIsDebug = True
        mbolIsHardDebug = IsHardDebug
        mstrDebugFilePath = DebugFilePath
    End Sub

    ''' <summary>调试</summary>
    Public Overloads Function PrintDebugLog(LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog("", LogInf, IsHardDebug)
    End Function

    ''' <summary>调试</summary>
    Public Overloads Function PrintDebugLog(StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog(StepName, LogInf, IsHardDebug)
    End Function

    ''' <summary>调试</summary>
    Public Overloads Function PrintDebugLog(StepName As String, LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog(StepName, LogInf, False)
    End Function

    ''' <summary>调试</summary>
    Public Overloads Function PrintDebugLog(LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog("", LogInf, False)
    End Function

    ''' <summary>调试</summary>
    Private Function mPrintDebugLog(StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        Try
            If IsHardDebug = True And mbolIsHardDebug = False Then Err.Raise(-1, , "只有硬调试模式才能打印日志")
            If mbolIsDebug = False Then Err.Raise(-1, , "只有调试模式才能打印日志")
            Dim sfAny As New System.IO.FileStream(Me.mstrDebugFilePath, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.Write, 10240, False)
            Dim swAny = New System.IO.StreamWriter(sfAny)
            LogInf = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss.fff") & "][" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString & "]" & LogInf
            If StepName <> "" Then LogInf &= "(" & StepName & ")"
            swAny.WriteLine(LogInf)
            swAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>我的类名</summary>
    Public Overloads ReadOnly Property MyClassName() As String
        Get
            MyClassName = mstrClsName
        End Get
    End Property

    ''' <summary>我的类名</summary>
    Public Overloads ReadOnly Property MyClassName(IsIncAppName As Boolean) As String
        Get
            If IsIncAppName = True Then
                MyClassName = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name & "."
            End If
            MyClassName = mstrClsName
        End Get
    End Property

    ''' <summary>是否硬调试</summary>
    Public ReadOnly Property IsHardDebug() As Boolean
        Get
            IsHardDebug = mbolIsHardDebug
        End Get
    End Property

    ''' <summary>是否调试</summary>
    Public ReadOnly Property IsDebug() As Boolean
        Get
            IsDebug = mbolIsDebug
        End Get
    End Property

    '''' <summary>应用名称</summary>
    'Public ReadOnly Property AppName() As String
    '    Get
    '        AppName = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name
    '    End Get
    'End Property

    '''' <summary>应用名称</summary>
    'Public ReadOnly Property AppName(IsIncClass As Boolean) As String
    '    Get
    '        'AppName = My.Application.Info.ProductName
    '        AppName = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name
    '        If IsIncClass = True Then AppName &= "." & mstrClsName
    '    End Get
    'End Property


    ''' <remarks>最近的错误信息</remarks>
    Public ReadOnly Property LastErr() As String
        Get
            LastErr = mstrLastErr
        End Get
    End Property

    ''' <remarks>完整过程名</remarks>
    Public ReadOnly Property FullSubName(SubName As String) As String
        Get
            FullSubName = mstrClsName & "." & SubName
        End Get
    End Property

    Private Function mGetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False, Optional IsSetLastErr As Boolean = False) As String
        Try
            Dim sbAny As New System.Text.StringBuilder("")
            sbAny.Append(Me.FullSubName(SubName))
            If Len(StepName) > 0 Then
                sbAny.Append("(")
                sbAny.Append(StepName)
                sbAny.Append(")")
            End If
            If Len(Me.KeyInf) > 0 Then sbAny.Append(";键值:" & Me.KeyInf)
            sbAny.Append(";错误信息:")
            sbAny.Append(exIn.Message)
            If IsStackTrace = True Then
                Dim strExStackTrace As String = exIn.StackTrace
                With strExStackTrace
                    If .Length > 0 Then
                        If .LastIndexOf(vbCrLf) >= 0 Then .Replace(vbCrLf, "")
                        If .LastIndexOf(vbTab) >= 0 Then .Replace(vbTab, " ")
                        .Trim()
                    End If
                End With
                sbAny.Append(";跟踪信息:")
                sbAny.Append(strExStackTrace)
            End If
            If IsSetLastErr = True Then mstrLastErr = sbAny.ToString
            mGetSubErrInf = sbAny.ToString
            sbAny = Nothing
        Catch ex As Exception
            If IsSetLastErr = True Then mstrLastErr = ex.Message.ToString
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>设置过程错误信息,用于不返回结果的调用，结果会保存在LastErr，但如果执行成功，要调用ClearErr</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Public Overloads Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace, True)
    End Sub

    ''' <summary>设置过程错误信息,带步骤名称,用于不返回结果的调用，结果会保存在LastErr，但如果执行成功，要调用ClearErr</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="StepName">步骤名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Public Overloads Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace, True)
    End Sub


    ''' <summary>获取过程错误信息</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Public Overloads Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace)
    End Function

    ''' <remarks>获取过程错误信息，带步骤名称</remarks>
    Public Overloads Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function

    ''' <remarks>获取未出错时过程调试信息</remarks>
    Public Function GetSubStepDebugInf(SubName As String, StepName As String, DebugInf As String) As String
        Try
            Dim sbAny As New System.Text.StringBuilder("")
            sbAny.Append(Me.FullSubName(SubName))
            If Len(StepName) > 0 Then
                sbAny.Append("(")
                sbAny.Append(StepName)
                sbAny.Append(")")
            End If
            If Len(Me.KeyInf) > 0 Then sbAny.Append(";键值:" & Me.KeyInf)
            sbAny.Append(";调试信息:")
            sbAny.Append(DebugInf)
            Return sbAny.ToString
            sbAny = Nothing
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

End Class
