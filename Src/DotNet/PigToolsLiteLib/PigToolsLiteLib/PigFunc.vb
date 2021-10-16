'**********************************
'* Name: PigFunc
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Some common functions|一些常用的功能函数
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 2/2/2021
'*1.0.2  1/3/2021   Add UrlEncode,UrlDecode
'*1.0.3  20/7/2021   Add GECBool,GECLng
'*1.0.4  26/7/2021   Modify UrlEncode
'*1.0.5  26/7/2021   Modify UrlEncode
'*1.0.6  24/8/2021   Modify GetIpList
'*1.1    4/9/2021    Add Date2Lng,Lng2Date,Src2CtlStr,CtlStr2Src,AddMultiLineText
'**********************************

Public Class PigFunc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.5"

    ''' <summary>文件的部分</summary>
    Public Enum enmFilePart
        Path = 1         '路径
        FileTitle = 2    '文件名
        ExtName = 3      '扩展名
        DriveNo = 4      '驱动器名
    End Enum

    ''' <summary>获取随机字符串的方式</summary>
    Public Enum enmGetRandString
        NumberOnly = 1      '只有数字
        NumberAndLetter = 2 '只有数字和字母(包括大小写)
        DisplayChar = 3     '全部可显示字符(ASCII 33-126)
        AllAsciiChar = 4    '全部ASCII码(返回结果以16进制方式显示)
    End Enum

    Sub New()
        MyBase.New(CLS_VERSION)
    End Sub


    ''' <remarks>获取比率的描述</remarks>
    Public Overloads Function GetRateDesc(Rate As Decimal) As String
        GetRateDesc = Math.Round(Rate * 100, 2) & " %"
    End Function

    ''' <remarks>获取比率的描述</remarks>
    Public Overloads Function GetRateDesc(Rate As Decimal, Optional RoundNo As Integer = 2) As String
        GetRateDesc = Math.Round(Rate * 100, RoundNo) & " %"
    End Function

    ''' <remarks>同步等待</remarks>
    Public Sub Delay(DelayMs As Single)
        System.Threading.Thread.Sleep(DelayMs)
    End Sub

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String) As String
        OptLogInf = Me.OptLogInfMain(OptStr, LogFilePath, True)
    End Function

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean) As String
        OptLogInf = Me.OptLogInfMain(OptStr, LogFilePath, IsShowLocalInf)
    End Function

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, LineBufSize As Integer) As String
        OptLogInf = Me.OptLogInfMain(OptStr, LogFilePath, True, LineBufSize)
    End Function

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean, LineBufSize As Integer) As String
        OptLogInf = Me.OptLogInfMain(OptStr, LogFilePath, IsShowLocalInf, LineBufSize)
    End Function

    Private Function OptLogInfMain(OptStr As String, LogFilePath As String, Optional IsShowLocalInf As Boolean = True, Optional LineBufSize As Integer = 10240) As String
        '写日志文件
        Try
            Dim sfAny As New System.IO.FileStream(LogFilePath, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.Write, LineBufSize, False)
            Dim swAny = New System.IO.StreamWriter(sfAny)
            If IsShowLocalInf = True Then OptStr = "[" & GENow() & "][" & GetProcThreadID() & "]" & OptStr
            swAny.WriteLine(OptStr)
            swAny.Close()
            sfAny.Close()
            OptLogInfMain = "OK"
        Catch ex As Exception
            OptLogInfMain = ex.Message.ToString
        End Try
    End Function

    ''' <remarks>获取进程及线程标识</remarks>
    Public Function GetProcThreadID() As String
        GetProcThreadID = System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString
    End Function


    ''' <remarks>获取32位MD5字符串</remarks>
    Public Function GEMD5(SrcStr As String) As String
        Dim bytSrc2Hash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(SrcStr)
        Dim bytHashValue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(bytSrc2Hash)
        Dim i As Integer
        GEMD5 = ""
        For i = 0 To 15 '选择32位字符的加密结果
            GEMD5 += Right("00" & Hex(bytHashValue(i)).ToLower, 2)
        Next
    End Function

    ''' <remarks>获取16位MD5字符串</remarks>
    Public Function Get16BitMD5(SrcStr As String) As String
        Get16BitMD5 = Mid(GEMD5(SrcStr), 9, 16)
    End Function

    ''' <remarks>显示精确到毫秒的时间</remarks>
    Public Function GENow() As String
        GENow = Format(Now, "yyyy-MM-dd HH:mm:ss.fff")
    End Function

    ''' <remarks>获取随机数</remarks>
    Public Function GetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()   '初始化随机数生成器
            GetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            GetRandNum = 0
        End Try
    End Function

    ''' <remarks>产生随机字符串</remarks>
    Public Function GetRandString(StrLen As Integer, Optional MemberType As enmGetRandString = enmGetRandString.DisplayChar) As String
        Dim i As Integer
        Dim intChar As Integer
        Try
            GetRandString = ""
            For i = 1 To StrLen
                Select Case MemberType
                    Case enmGetRandString.AllAsciiChar
                        intChar = GetRandNum(0, 255)
                    Case enmGetRandString.DisplayChar    '!-~
                        intChar = GetRandNum(33, 126)
                    Case enmGetRandString.NumberAndLetter
                        intChar = GetRandNum(1, 3)
                        Select Case intChar
                            Case 1  '0-9
                                intChar = GetRandNum(48, 57)
                            Case 2  'A-Z
                                intChar = GetRandNum(65, 90)
                            Case 3  'a-z
                                intChar = GetRandNum(97, 122)
                        End Select
                    Case enmGetRandString.NumberOnly
                        intChar = GetRandNum(48, 57)
                End Select
                If MemberType = enmGetRandString.AllAsciiChar Then
                    GetRandString = GetRandString & Right("0" & Hex(intChar), 2)
                Else
                    GetRandString = GetRandString & Chr(intChar)
                End If
            Next
        Catch ex As Exception
            GetRandString = ""
        End Try

    End Function

    ''' <remarks>生成主键值，唯一</remarks>
    Public Function GetPKeyValue(SrcKey As String, Is32Bit As Boolean) As String
        Dim strRndStr As String, strRndKey As String, i As Long, strTmp As String, strChar As String, strChar2 As String
        Dim intCharAscAdd As Integer, intStrLen As Integer
        GetPKeyValue = ""
        Try
            If Is32Bit = True Then
                intStrLen = 32
            Else
                intStrLen = 16
            End If
            strRndStr = GetRandString(64, enmGetRandString.DisplayChar)
            strRndKey = GetRandString(intStrLen, enmGetRandString.NumberOnly)
            strTmp = GENow() & "-" & SrcKey & "-" & strRndStr
            If Is32Bit = True Then
                strTmp = GEMD5(strTmp)
            Else
                strTmp = Get16BitMD5(strTmp)
            End If
            strTmp = LCase(strTmp)
            For i = 1 To intStrLen
                strChar = Mid(strTmp, i, 1)
                strChar2 = Mid(strRndKey, i, 1)
                intCharAscAdd = 0
                Select Case strChar2
                    Case "0"    '不变
                    Case "1"    '数字不变，字母变大写
                        Select Case strChar
                            Case "0" To "9"
                            Case Else
                                intCharAscAdd = -32
                        End Select
                    Case "2" To "5"    '位移到 g-j
                        Select Case strChar
                            Case "0" To "9"
                                intCharAscAdd = (Asc("g") + Int(strChar2) - 1) - Asc("0")
                            Case Else
                                intCharAscAdd = (Asc("g") + Int(strChar2) - 1) - Asc("a")
                        End Select
                    Case "6" To "9"    '位移到 G-J
                        Select Case strChar
                            Case "0" To "9"
                                intCharAscAdd = (Asc("G") + Int(strChar2) - 1) - Asc("0")
                            Case Else
                                intCharAscAdd = (Asc("G") + Int(strChar2) - 1) - Asc("a")
                        End Select
                    Case Else   '不变
                End Select
                strChar2 = Chr(Asc(strChar) + intCharAscAdd)
                Select Case strChar2
                    Case "0" To "9"
                        strChar = strChar2
                    Case "a" To "z"
                        strChar = strChar2
                    Case "A" To "Z"
                        strChar = strChar2
                    Case Else
                End Select
                GetPKeyValue &= strChar
            Next
        Catch ex As Exception
            GetPKeyValue = ""
        End Try
    End Function

    ''' <remarks>获取IP列表，主IP</remarks>
    Public Sub GetIpList(ByRef IpList As String, ByRef MainIp As String, MainIpRole As String)
        'IpList:本机IP列表
        'MainIp:主IP
        'MainIpRole:主IP规则，如：56.0.，如果IP列表中有第一个匹配的就是主IP
        Try
            Dim aipaAny() As System.Net.IPAddress, strIp As String = "", intFindCnt As Integer = 0, intLen As Integer = 0
            aipaAny = System.Net.Dns.GetHostAddresses(System.Environment.MachineName.ToString)
            IpList = ""
            MainIp = ""
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            Dim i As Integer 
            For i = 0 To aipaAny.Count - 1
                strIp = aipaAny(i).ToString
                If InStr(strIp, "::") = 0 Then
                    '这种IP没有有效的IP4地址，如：fe80::95fa:649f:cbab:5912%21，不要

                    IpList &= strIp & ";"
                    intFindCnt += 1
                    If intFindCnt = 1 Then
                        '第一个先假定为主IP
                        MainIp = strIp
                    Else
                        intLen = Len(MainIpRole)
                        If intLen > 0 Then
                            If Microsoft.VisualBasic.Left(strIp, Len(MainIpRole)) = MainIpRole Then
                                'IP的前N位与主IP规则匹配
                                MainIp = strIp
                            End If
                        End If
                    End If
                End If
            Next
#End If
        Catch ex As Exception
            IpList = ""
            MainIp = ""
        End Try
    End Sub

    ''' <remarks>获取文件路径的组成部分</remarks>
    Public Function GetFilePart(ByVal FilePath As String, Optional FilePart As enmFilePart = enmFilePart.FileTitle) As String
        Dim strTemp As String, i As Long, lngLen As Long
        Dim strPath As String = "", strFileTitle As String = ""
        Try
            GetFilePart = ""
            Select Case FilePart
                Case enmFilePart.DriveNo
                    GetFilePart = GetStr(FilePath, "", ":", False)
                    If GetFilePart = "" Then
                        GetFilePart = GetStr(FilePath, "", "$", False)
                        If GetFilePart <> "" Then GetFilePart = GetFilePart & "$"
                    End If
                Case enmFilePart.ExtName
                    'GetFilePart = GetStr(FilePath, ".", "", False)
                    lngLen = Len(FilePath)
                    For i = lngLen To 1 Step -1
                        Select Case Mid(FilePath, i, 1)
                            Case "/", ":", "$"
                                Exit For
                            Case "."
                                GetFilePart = Mid(FilePath, i + 1)
                                Exit For
                        End Select

                    Next
                Case enmFilePart.FileTitle, enmFilePart.Path
                    Do While True
                        'My.Application.DoEvents()
                        strTemp = GetStr(FilePath, "", "\", True)
                        If Len(strTemp) = 0 Then
                            If Right(strPath, 1) = "\" Then
                                If Right(strPath, 2) <> ":\" Then
                                    strPath = Left(strPath, Len(strPath) - 1)
                                End If
                            ElseIf Left(FilePath, 1) = "\" Then
                                strPath = "\"
                                FilePath = Mid(FilePath, 2)
                            End If
                            If FilePath <> "" Then
                                strFileTitle = FilePath
                            Else
                                strFileTitle = "."
                            End If
                            Exit Do
                        End If
                        strPath = strPath & strTemp & "\"
                    Loop
                    If FilePart = enmFilePart.FileTitle Then
                        GetFilePart = strFileTitle
                    Else
                        GetFilePart = strPath
                    End If
                Case Else
                    GetFilePart = ""
            End Select
        Catch ex As Exception
            GetFilePart = ""
        End Try
    End Function

    ''' <remarks>截取字符串</remarks>
    Public Function GetStr(ByRef SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
        Dim lngBegin As Long
        Dim lngEnd As Long
        Dim lngBeginLen As Long
        Dim lngEndLen As Long
        Try
            lngBeginLen = Len(strBegin)
            lngBegin = InStr(SourceStr, strBegin, CompareMethod.Text)
            lngEndLen = Len(strEnd)
            If lngEndLen = 0 Then
                lngEnd = Len(SourceStr) + 1
            Else
                lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd, CompareMethod.Text)
                If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0")
            End If
            If lngEnd <= lngBegin Then Err.Raise(-1, , "lngEnd <= lngBegin")
            If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0[2]")

            GetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
            If IsCut = True Then
                SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
            End If
            Me.ClearErr()
        Catch ex As Exception
            GetStr = ""
            Me.SetSubErrInf("GetStr", ex)
        End Try
    End Function



    Public Function UrlEncode(SrcUrl As String) As String
        Try
#If NETCOREAPP Or NET5_0_OR_GREATER Then
            UrlEncode = System.Web.HttpUtility.UrlEncode(SrcUrl)
#Else
            UrlEncode = System.Uri.EscapeDataString(SrcUrl)
#End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("UrlEncode", ex)
            Return Nothing
        End Try
    End Function

    Public Function UrlDecode(DecodeUrl As String) As String
        Try
#If NETCOREAPP Or NET5_0_OR_GREATER Then
            UrlDecode = System.Web.HttpUtility.UrlDecode(DecodeUrl)
#Else
            UrlDecode = System.Uri.UnescapeDataString(DecodeUrl)
#End If

            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("UrlDecode", ex)
            Return Nothing
        End Try
    End Function

    Public Function IsRegexMatch(SrcStr As String, RegularExp As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(SrcStr, RegularExp)
    End Function

    Public Function GECBool(vData As String) As Boolean
        Try
            Return CBool(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECBool", ex)
            Return False
        End Try
    End Function

    Public Function GECLng(vData As String) As Long
        Try
            Return GECLng(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
        End Try
    End Function

    Public Function GEDec(vData As String) As Decimal
        Try
            Return CDec(vData)
        Catch ex As Exception
            Me.SetSubErrInf("CDec", ex)
            Return 0
        End Try
    End Function

    Public Function GECDate(vData As String) As DateTime
        Try
            Return Date.Parse(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECDate", ex)
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>
    ''' Gets the number of microseconds of Greenwich mean time for the current time|获取当前时间的格林威治时间微秒数
    ''' </summary>
    ''' <param name="DateValue"></param>
    ''' <returns></returns>
    Public Function Date2Lng(DateValue As DateTime) As Long
        Dim dteStart As New DateTime(1970, 1, 1)
        Dim mtsTimeDiff As TimeSpan = DateValue - dteStart
        Try
            Return mtsTimeDiff.TotalMilliseconds
        Catch ex As Exception
            Me.SetSubErrInf("Date2Lng", ex)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="LngValue">The number of milliseconds since 1970-1-1</param>
    ''' <param name="IsLocalTime">Convert to local time</param>
    ''' <returns></returns>
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Public  Function Lng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            Dim intHourAdd As Integer = 0
            If IsLocalTime = False Then
                Dim oTimeZoneInfo As System.TimeZoneInfo
                oTimeZoneInfo = System.TimeZoneInfo.Local
                intHourAdd = oTimeZoneInfo.GetUtcOffset(Now).Hours
            End If

            Return dteStart.AddSeconds(LngValue + intHourAdd * 3600)
            Me.ClearErr()
        Catch ex As Exception
            Return dteStart
            Me.SetSubErrInf("Lng2Date", ex)
        End Try
    End Function
#Else
    Public Function Lng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            If IsLocalTime = False Then
                Lng2Date = dteStart.AddMilliseconds(LngValue - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000)
            Else
                Lng2Date = dteStart.AddMilliseconds(LngValue)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("Lng2Date", ex)
            Return DateTime.MinValue
        End Try
    End Function
#End If

    Public Function Src2CtlStr(ByRef SrcStr As String) As String
        Try
            If SrcStr.IndexOf(vbCrLf) > 0 Then SrcStr = Replace(SrcStr, vbCrLf, "\r\n")
            If SrcStr.IndexOf(vbCr) > 0 Then SrcStr = Replace(SrcStr, vbCr, "\r")
            If SrcStr.IndexOf(vbTab) > 0 Then SrcStr = Replace(SrcStr, vbTab, "\t")
            If SrcStr.IndexOf(vbBack) > 0 Then SrcStr = Replace(SrcStr, vbBack, "\b")
            If SrcStr.IndexOf(vbFormFeed) > 0 Then SrcStr = Replace(SrcStr, vbFormFeed, "\f")
            If SrcStr.IndexOf(vbVerticalTab) > 0 Then SrcStr = Replace(SrcStr, vbVerticalTab, "\v")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Src2CtlStr", ex)
        End Try
    End Function

    Public Function CtlStr2Src(ByRef CtlStr As String) As String
        Try
            If CtlStr.IndexOf("\r\n") > 0 Then CtlStr = Replace(CtlStr, "\r\n", vbCrLf)
            If CtlStr.IndexOf("\r") > 0 Then CtlStr = Replace(CtlStr, "\r", vbCr)
            If CtlStr.IndexOf("\t") > 0 Then CtlStr = Replace(CtlStr, "\t", vbTab)
            If CtlStr.IndexOf("\b") > 0 Then CtlStr = Replace(CtlStr, "\b", vbBack)
            If CtlStr.IndexOf(vbFormFeed) > 0 Then CtlStr = Replace(CtlStr, "\f", vbFormFeed)
            If CtlStr.IndexOf(vbVerticalTab) > 0 Then CtlStr = Replace(CtlStr, "\v", vbVerticalTab)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CtlStr2Src", ex)
        End Try
    End Function

    Public Function AddMultiLineText(ByRef MainText As String, NewLine As String, Optional LeftTabs As Integer = 0) As String
        Try
            Dim strTabs As String = ""
            If LeftTabs > 0 Then
                For i = 1 To LeftTabs
                    strTabs &= vbTab
                Next
            End If
            MainText &= strTabs & NewLine & vbCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddMultiLineText", ex)
        End Try
    End Function

End Class
