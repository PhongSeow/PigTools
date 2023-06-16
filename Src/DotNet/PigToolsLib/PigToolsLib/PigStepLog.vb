'**********************************
'* Name: PigStepLog
'* Author: Seow Phong
'* License: Copyright (c) 2020-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigStepLog is for logging and error handling in the process.
'* Home Url: https://en.seowphong.com
'* Version: 1.7.2
'* Create Time: 8/12/2019
'1.1    18/12/2021  Add TrackID,ErrInf2User, modify mNew,StepName
'1.2    21/12/2021  Modify TrackID
'1.3    10/2/2022  Add StepLogInf
'1.5    30/8/2022  Modify New
'1.6    8/12/2022  Modify mNew,StepName
'1.7    1/6/2023  Modify ErrInf2User
'************************************
''' <summary>
''' Error tracking processing class|错误跟踪处理类
''' </summary>
Public Class PigStepLog
    Public ReadOnly Property SubName As String

    Private Property mHostInf As String = ""
    Private mbolIsLogUseTime As Boolean
    Public Property IsLogUseTime As Boolean
        Get
            Return mbolIsLogUseTime
        End Get
        Friend Set(value As Boolean)
            mbolIsLogUseTime = value
        End Set
    End Property

    Public Ret As String = ""

    Private moUseTime As UseTime

    ''' <summary>
    ''' 显示给用户看的错误信息，不能显示内部错误内容。|The error message displayed to the user cannot display internal error content.
    ''' </summary>
    Private mstrErrInf2User As String
    Public Property ErrInf2User(Optional IsTrackID As Boolean = True) As String
        Get
            If Me.TrackID = "" Then
                Return mstrErrInf2User
            ElseIf IsTrackID = True Then
                Return mstrErrInf2User & "(" & Me.TrackID & ")"
            Else
                Return mstrErrInf2User
            End If
        End Get
        Set(value As String)
            mstrErrInf2User = value
        End Set
    End Property

    Private Sub mNew(Optional IsTrack As Boolean = False, Optional IsLogUseTime As Boolean = False, Optional TrackIDHead As String = "")
        Me.IsLogUseTime = IsLogUseTime
        If IsTrack = True Then
            If TrackIDHead = "" Then
                Dim strID As String = System.Net.Dns.GetHostName() & "." & Me.mGetProcThreadID
                TrackIDHead = Me.mGEMD5(strID)
            End If
            Me.TrackID = TrackIDHead & "." & Format(Now, "yyyy-MM-dd HH:mm:ss.fff") & "."
            'For i = 0 To 5
            '    Me.TrackID &= Me.mGetRandNum(0, 255)
            'Next
            Me.TrackID = Mid(Me.mGEMD5(Me.TrackID), 9, 16)
        End If
        If Me.IsLogUseTime = True Then
            moUseTime = New UseTime
            moUseTime.GoBegin()
        End If
    End Sub

    Private Function mGetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()   '初始化随机数生成器
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function


    Private Function mGEMD5(SrcStr As String) As String
        Dim bytSrc2Hash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(SrcStr)
        Dim bytHashValue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(bytSrc2Hash)
        Dim i As Integer
        mGEMD5 = ""
        For i = 0 To 15 '选择32位字符的加密结果
            mGEMD5 += Right("00" & Hex(bytHashValue(i)).ToLower, 2)
        Next
    End Function

    Private Function mGetProcThreadID() As String
        Return System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString
    End Function


    Public Sub New(SubName As String)
        Me.SubName = SubName
        Me.mNew()
    End Sub

    Public Sub New(SubName As String, IsLogUseTime As Boolean)
        Me.SubName = SubName
        Me.mNew(, IsLogUseTime)
    End Sub

    Public Sub New(SubName As String, TrackIDHead As String)
        Me.SubName = SubName
        Me.mNew(True, IsLogUseTime, TrackIDHead)
    End Sub

    Public Sub New(SubName As String, IsLogUseTime As Boolean, IsTrack As Boolean)
        Me.SubName = SubName
        Me.mNew(IsTrack, IsLogUseTime)
    End Sub

    Private mstrTrackID As String = ""
    Public Property TrackID As String
        Get
            Return mstrTrackID
        End Get
        Friend Set(value As String)
            mstrTrackID = value
        End Set
    End Property

    Private mstrStepName As String = ""
    Public Property StepName As String
        Get
            Return mstrStepName
        End Get
        Set(value As String)
            If Me.TrackID <> "" Then
                mstrStepName = value & "[TrackID:" & Me.TrackID & "]"
            Else
                mstrStepName = value
            End If
        End Set
    End Property

    Public Sub AddStepNameInf(AddInf As String)
        Me.StepName &= "(" & AddInf & ")"
    End Sub

    Public ReadOnly Property BeginTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.BeginTime
            Else
                Return DateTime.MinValue
            End If
        End Get
    End Property

    Public ReadOnly Property EndTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.EndTime
            Else
                Return DateTime.MaxValue
            End If
        End Get
    End Property

    Public ReadOnly Property AllDiffSeconds As Decimal
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.AllDiffSeconds
            Else
                Return -1
            End If
        End Get
    End Property


    Public Sub ToEnd()
        If Me.IsLogUseTime = True Then
            moUseTime.ToEnd()
        End If
    End Sub

    ''' <summary>
    ''' 当前步骤的日志信息|Log information of the current step
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property StepLogInf As String
        Get
            With Me
                StepLogInf = "[" & Me.SubName & "]"
                StepLogInf &= "[" & Me.StepName & "]"
                If Me.Ret <> "" Then StepLogInf &= "[" & Me.Ret & "]"
                If Me.TrackID <> "" Then StepLogInf &= "[TrackID:" & Me.TrackID & "]"
                If Me.IsLogUseTime = True Then StepLogInf &= "[" & Me.AllDiffSeconds.ToString & "]"
            End With
        End Get
    End Property

    Private Function mGetHostIp(IsIPv6 As Boolean, Optional IpHead As String = "") As String
        mGetHostIp = ""
        For Each oIPAddress As System.Net.IPAddress In System.Net.Dns.GetHostAddresses(System.Net.Dns.GetHostName)
            With oIPAddress
                mGetHostIp = .ToString()
                If IsIPv6 = True Then
                    If InStr(mGetHostIp, ":") > 0 Then
                        If IpHead = "" Then
                            Exit For
                        ElseIf UCase(Left(mGetHostIp, Len(IpHead))) = UCase(IpHead) Then
                            Exit For
                        End If
                    End If
                ElseIf InStr(mGetHostIp, ".") > 0 Then
                    If IpHead = "" Then
                        Exit For
                    ElseIf Left(mGetHostIp, Len(IpHead)) = IpHead Then
                        Exit For
                    End If
                End If
            End With
            mGetHostIp = ""
        Next
    End Function

End Class
