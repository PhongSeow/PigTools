'**********************************
'* Name: AnyJava
'* Author: Seow Phong
'* License: Copyright (c) 2025 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: A universal JAVA process
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 19/4/2025
'************************************
Friend Class AnyJava
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "2"

    Public Enum EnmRunCmdType
        Unkown = 0
        CmdShell = 1
        AsyncCmdShell = 2
        CallFile = 3
        AsyncCallFile = 7
        CallFileWaitForExit = 5
        CmdShellWaitForExit = 8
        Sudo = 6
        NoHup = 9
    End Enum

    Public ReadOnly Property Parent As AnyJavaApp
    Public ReadOnly Property AppName As String

    Public Sub New(AppName As String, Parent As AnyJavaApp)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
        Me.AppName = AppName
        Me.RunUser = ""
        Me.ListenPort = -1
        Me.IsForceStop = True
    End Sub

    Private mRunUser As String
    Public Property RunUser As String
        Get
            Return mRunUser
        End Get
        Friend Set(value As String)
            mRunUser = value
        End Set
    End Property

    Private mListenPort As Integer
    Public Property ListenPort As Integer
        Get
            Return mListenPort
        End Get
        Friend Set(value As Integer)
            mListenPort = value
        End Set
    End Property

    Private mAppCmdParaRegEx As String
    Public Property AppCmdParaRegEx As String
        Get
            Return mAppCmdParaRegEx
        End Get
        Friend Set(value As String)
            mAppCmdParaRegEx = value
        End Set
    End Property

    Private mAppParentCmdParaRegEx As String
    Public Property AppParentCmdParaRegEx As String
        Get
            Return mAppParentCmdParaRegEx
        End Get
        Friend Set(value As String)
            mAppParentCmdParaRegEx = value
        End Set
    End Property

    Private mProcCmd As String
    Public Property ProcCmd As String
        Get
            Return mProcCmd
        End Get
        Friend Set(value As String)
            mProcCmd = value
        End Set
    End Property

    Private mParentProcCmd As String
    Public Property ParentProcCmd As String
        Get
            Return mParentProcCmd
        End Get
        Friend Set(value As String)
            mParentProcCmd = value
        End Set
    End Property

    Private mPID As Integer
    Public Property PID As Integer
        Get
            Return mPID
        End Get
        Friend Set(value As Integer)
            mPID = value
        End Set
    End Property

    Private mIsRunning As Boolean
    Public Property IsRunning As Boolean
        Get
            Return mIsRunning
        End Get
        Friend Set(value As Boolean)
            mIsRunning = value
        End Set
    End Property

    Public Function Refresh() As String

        Try

        Catch ex As Exception

        End Try
    End Function

    Private mStartCmdType As EnmRunCmdType = EnmRunCmdType.AsyncCmdShell
    Public Property StartCmdType As EnmRunCmdType
        Get
            Return mStartCmdType
        End Get
        Friend Set(value As EnmRunCmdType)
            mStartCmdType = value
        End Set
    End Property

    Private mStopCmdType As EnmRunCmdType = EnmRunCmdType.AsyncCmdShell
    Public Property StopCmdType As EnmRunCmdType
        Get
            Return mStopCmdType
        End Get
        Friend Set(value As EnmRunCmdType)
            mStopCmdType = value
        End Set
    End Property

    Private mStartPathOrCmd As String
    Public Property StartPathOrCmd As String
        Get
            Return mStartPathOrCmd
        End Get
        Friend Set(value As String)
            mStartPathOrCmd = value
        End Set
    End Property

    Public ReadOnly Property IsStartCmd As Boolean
        Get
            Select Case Me.StartCmdType
                Case EnmRunCmdType.AsyncCallFile, EnmRunCmdType.CallFile, EnmRunCmdType.CallFileWaitForExit
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property

    Private mStartShellPara As String
    Public Property StartShellPara As String
        Get
            Return mStartShellPara
        End Get
        Set(value As String)
            mStartShellPara = value
        End Set
    End Property


    Private mStopPathOrCmd As String
    Public Property StopPathOrCmd As String
        Get
            Return mStopPathOrCmd
        End Get
        Friend Set(value As String)
            mStopPathOrCmd = value
        End Set
    End Property

    Public ReadOnly Property IsStopCmd As Boolean
        Get
            Select Case Me.StopCmdType
                Case EnmRunCmdType.AsyncCallFile, EnmRunCmdType.CallFile, EnmRunCmdType.CallFileWaitForExit
                    Return False
                Case Else
                    Return True
            End Select
        End Get
    End Property

    Private mStopShellPara As String
    Public Property StopShellPara As String
        Get
            Return mStopShellPara
        End Get
        Set(value As String)
            mStopShellPara = value
        End Set
    End Property

    Private mIsForceStop As Boolean
    Public Property IsForceStop As Boolean
        Get
            Return mIsForceStop
        End Get
        Friend Set(value As Boolean)
            mIsForceStop = value
        End Set
    End Property

    Private mVersion As String
    Public Property Version As String
        Get
            Return mVersion
        End Get
        Friend Set(value As String)
            mVersion = value
        End Set
    End Property

    Private mClassPath As String
    Public Property ClassPath As String
        Get
            Return mClassPath
        End Get
        Friend Set(value As String)
            mClassPath = value
        End Set
    End Property

    Private mHeapDumpOnOutOfMemoryError As Boolean
    Public Property HeapDumpOnOutOfMemoryError As Boolean
        Get
            Return mHeapDumpOnOutOfMemoryError
        End Get
        Friend Set(value As Boolean)
            mHeapDumpOnOutOfMemoryError = value
        End Set
    End Property

    Private mHeapDumpPath As String
    Public Property HeapDumpPath As String
        Get
            Return mHeapDumpPath
        End Get
        Friend Set(value As String)
            mHeapDumpPath = value
        End Set
    End Property


End Class
