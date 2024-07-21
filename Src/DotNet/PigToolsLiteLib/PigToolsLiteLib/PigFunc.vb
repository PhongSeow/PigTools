'**********************************
'* Name: PigFunc
'* Author: Seow Phong
'* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Some common functions|一些常用的功能函数
'* Home Url: https://en.seowphong.com
'* Version: 1.70
'* Create Time: 2/2/2021
'*1.0.2  1/3/2021   Add UrlEncode,UrlDecode
'*1.0.3  20/7/2021   Add GECBool,GECLng
'*1.0.4  26/7/2021   Modify UrlEncode
'*1.0.5  26/7/2021   Modify UrlEncode
'*1.0.6  24/8/2021   Modify GetIpList
'*1.1    4/9/2021    Add Date2Lng,Lng2Date,Src2CtlStr,CtlStr2Src,AddMultiLineText
'*1.2    2/1/2022    Modify IsFileExists
'*1.3    12/1/2022   Add GetHostName,GetHostIp,mGetHostIp,GetEnvVar,GetUserName,GetComputerName,mGetHostIpList
'*1.4    23/1/2022   Add IsOsWindows,MyOsCrLf,MyOsPathSep
'*1.5    3/2/2022   Add GetFileText,SaveTextToFile, modify GetFilePart
'*1.6    3/2/2022   Add GetFmtDateTime, modify GENow
'*1.7    13/2/2022   Add DeleteFolder,DeleteFile
'*1.8    23/2/2022   Add PLSqlCsv2Bcp
'*1.9    20/3/2022   Modify GetProcThreadID
'*1.10   2/4/2022   Add GetHumanSize
'*1.11   9/4/2022   Add DigitalToChnName,ConvertHtmlStr,GetAlignStr,GetRepeatStr
'*1.12   14/5/2022  Rename and modify OptLogInfMain to mOptLogInfMain,SaveTextToFile to mSaveTextToFile, add ASyncOptLogInf,ASyncSaveTextToFile
'*1.13   15/5/2022  Add ASyncRet_SaveTextToFile, modify mASyncSaveTextToFile,ASyncSaveTextToFile
'*1.14   16/5/2022  Modify mASyncSaveTextToFile,ASyncSaveTextToFile
'*1.15   31/5/2022  Add GEInt, modify GECBool
'*1.16   5/7/2022   Modify GetHostIp
'*1.17   6/7/2022   Add GetFileVersion
'*1.18   19/7/2022  Add GetFileUpdateTime,GetFileCreateTime,GetFileMD5
'*1.19   4/8/2022   Add GetShareMem,SaveShareMem,GetTextPigMD5,GetBytesPigMD5
'*1.20   16/8/2022  Add GetMyExeName,GetMyPigProc,GetMyExePath
'*1.21   17/8/2022  Add GetMyPigProc,GetMyExePath
'*1.22   20/8/2022  Add Is64Bit,GetMachineGUID
'*1.23   25/8/2022  Modify CtlStr2Src,Src2CtlStr,IsFolderExists
'*1.25   2/9/2022   Add CheckFileDiff
'*1.26   6/9/2022   Add IsMathDate,IsMathDecimal
'*1.27   12/9/2022  Add GetEnmDispStr
'*1.28   17/9/2022  Add IsStrongPassword,GetCompMinutePart
'*1.29   17/10/2022  Add EscapeStr,UnEscapeStr
'*1.30   18/10/2022  Modify EscapeStr,UnEscapeStr
'*1.31   2/11/2022  Modify AddMultiLineText
'*1.32   14/11/2022  Add GetTextBase64,ClearErr
'*1.33   15/11/2022  Add GetTextSHA1
'*1.35   18/1/2023  Modify GetUUID,GetMachineGUID, add mGetUUID,mGetMachineGUID
'*1.36   19/1/2023  Modify mGetUUID,GetUUID,GetMachineGUID, add GetBootID,GetProductUuid
'*1.37   31/3/2023  Modify GetFilePart,GetStr
'*1.38   7/4/2023   Add GetPathPart
'*1.39   12/4/2023  Add SQLCDate
'*1.50   26/4/2023  Add IsFileDiff
'*1.51   28/4/2023  Modify GetWindowsProductId
'*1.52   21/6/2023  Add IsDeviationTime,IsTimeout
'*1.53   23/7/2023  Add IsNewVersion
'*1.55   21/10/2023  Add ClipboardGetText,ClipboardSetText
'*1.56   4/11/2023  Add OpenUrl,GetDefaultBrowser
'*1.57   7/11/2023  Add GetTimeSlot
'*1.58   10/11/2023  Add GetTextFileEncCode
'*1.59   23/12/2023  Add AddMultiLineText
'*1.60   29/12/2023  Add CombinePath,GetStr,GetStrAndReplace
'*1.61   30/12/2023  Add IsAbsolutePath
'*1.62   16/1/2024  Modify GetStr
'*1.63   22/1/2024  Modify GetStrAndReplace,IsRegexMatch
'*1.65   13/3/2024  Add LenA,LeftA,RightA,MidA,GetAlignStrA
'*1.66   21/3/2024  Add GetDateUnixTimestamp,GetNowUnixTimestamp, modify GetWindowsProductId,mGetMachineGUID
'*1.67   10/6/2024  Modify GetWindowsProductId,mGetMachineGUID
'*1.68   2/7/2024  Add DistinctString
'*1.69   2/7/2024  Add StrSpaceMulti2Double,GECInt
'*1.70   16/7/2024  Add StrSpaceMulti2One
'**********************************
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Environment
Imports System.Threading
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic
Imports System.Text
Imports System.Runtime.InteropServices.ComTypes

''' <summary>
''' Function set|功能函数集
''' </summary>
Public Class PigFunc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "70" & "." & "2"

    Public Event ASyncRet_SaveTextToFile(SyncRet As StruASyncRet)

    Public Enum EnmTimeSlot
        CurrentWeek = 0
        CurrentMonth = 1
        CurrentQuarter = 2
        CurrentYear = 3
        LastWeek = 10
        LastMonth = 11
        LastQuarter = 12
        LastYear = 13
    End Enum

    ''' <summary>对齐方式|Alignment</summary>
    Public Enum EnmAlignment
        Left = 1
        Right = 2
        Center = 3
    End Enum

    ''' <summary>
    ''' 检查不同的方式|Check different ways
    ''' </summary>
    Public Enum EnmCheckDiffType
        Size_Date = 1
        Size_Date_FastPigMD5 = 2
        Size_FullPigMD5 = 3
    End Enum

    ''' <summary>文件的部分</summary>
    Public Enum EnmFilePart
        Path = 1         '路径
        FileTitle = 2    '文件名
        ExtName = 3      '扩展名
        DriveNo = 4      '驱动器名
    End Enum

    ''' <summary>
    ''' 路径的部分|Part of the path
    ''' </summary>
    Public Enum EnmFathPart
        ''' <summary>
        ''' 父路径|Parent Path
        ''' </summary>
        ParentPath = 1
        ''' <summary>
        ''' 文件或目录名|File or directory name
        ''' </summary>
        FileOrDirTitle = 2
        ''' <summary>
        ''' 扩展名|File extensions
        ''' </summary>
        ExtName = 3
        ''' <summary>
        ''' 驱动器名(Windows)|Drive Letter (Windows)
        ''' </summary>
        DriveLetter = 4
        ''' <summary>
        ''' 文件或目录名（不含扩展名）|File or directory name(without extension)
        ''' </summary>
        FileOrDirTitleWithoutExtName = 5
    End Enum

    ''' <summary>获取随机字符串的方式</summary>
    Public Enum EnmGetRandString
        NumberOnly = 1      '只有数字
        NumberAndLetter = 2 '只有数字和字母(包括大小写)
        DisplayChar = 3     '全部可显示字符(ASCII 33-126)
        AllAsciiChar = 4    '全部ASCII码(返回结果以16进制方式显示)
    End Enum

    ''' <summary>
    ''' 转换HTML标记的方式|How to convert HTML tags
    ''' </summary>
    Public Enum EmnHowToConvHtml
        DisableHTML = 1            '使HTML标记失效(空格、制表符及回车符都会被转换)
        EnableHTML = 2             '使HTML标记生效(空格会被还原，但制表符及回车符不会被还原)
        DisableHTMLOnlySymbol = 3  '使HTML标记失效(只有标记的符号本身受影响，空格、制表符及回车符不会被转换)
        EnableHTMLOnlySymbol = 4   '使HTML标记生效(只有标记的符号本身被还原)
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

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath)
    End Function


    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = True
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>写日志</remarks>
    Private Function mASyncOptLogInfMain(OptStr As String, LogFilePath As String, Optional IsShowLocalInf As Boolean = True, Optional LineBufSize As Integer = 10240) As String
        Try
            Dim struMain As mStruOptLogInfMain
            With struMain
                .OptStr = OptStr
                .LogFilePath = LogFilePath
                .IsShowLocalInf = IsShowLocalInf
                .LineBufSize = LineBufSize
            End With
            Dim oThread As New Thread(AddressOf mOptLogInfMain)
            oThread.Start(struMain)
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mASyncOptLogInfMain", ex)
        End Try
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, IsShowLocalInf)
    End Function

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = IsShowLocalInf
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, LineBufSize As Integer) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, , LineBufSize)
    End Function


    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, LineBufSize As Integer) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = True
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean, LineBufSize As Integer) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, IsShowLocalInf, LineBufSize)
    End Function
    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean, LineBufSize As Integer) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = IsShowLocalInf
            .LineBufSize = LineBufSize
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    Private Structure mStruOptLogInfMain
        Public OptStr As String
        Public LogFilePath As String
        Public IsShowLocalInf As Boolean
        Public LineBufSize As Integer
    End Structure

    Private Function mOptLogInfMain(StruMain As mStruOptLogInfMain) As String
        Dim LOG As New PigStepLog("mOptLogInfMain")
        Try
            LOG.StepName = "Check StruOptLogInfMain"
            With StruMain
                If .LineBufSize <= 0 Then .LineBufSize = 10240
                If .LogFilePath = "" Then Throw New Exception("LogFilePath invalid")
            End With
            LOG.StepName = "New FileStream"
            Dim sfAny As New FileStream(StruMain.LogFilePath, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.Write, StruMain.LineBufSize, False)
            LOG.StepName = "New StreamWriter"
            Dim swAny = New StreamWriter(sfAny)
            If StruMain.IsShowLocalInf = True Then StruMain.OptStr = "[" & GENow() & "][" & GetProcThreadID() & "]" & StruMain.OptStr
            LOG.StepName = "WriteLine"
            swAny.WriteLine(StruMain.OptStr)
            LOG.StepName = "Close"
            swAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(StruMain.LogFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <remarks>获取进程及线程标识</remarks>
    Public Function GetProcThreadID() As String
        Return System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString
    End Function

    ''' <remarks>获取进程号</remarks>
    Public Function GetProcID() As Long
        Return Me.fMyPID
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
        Return Format(Now, "yyyy-MM-dd HH:mm:ss.fff")
    End Function

    Public Function GetFmtDateTime(SrcTime As DateTime, Optional TimeFmt As String = "yyyy-MM-dd HH:mm:ss.fff") As String
        Return Format(SrcTime, TimeFmt)
    End Function

    Public Function GetHostName() As String
        Return Dns.GetHostName()
    End Function

    Public Function GetHostIp() As String
        Return mGetHostIp(False)
    End Function

    Public Function GetHostIpList() As String
        Return mGetHostIpList(False)
    End Function

    Public Function GetHostIpList(IsIPv6 As Boolean) As String
        Return mGetHostIpList(IsIPv6)
    End Function

    ''' <summary>
    ''' 获取主机的IP|Get the IP of the host
    ''' </summary>
    ''' <param name="PriorityIpHead">优先匹配的IP地址开头，如果匹配不到则取第一个IP|The first IP address to be matched first. If it cannot be matched, the first IP will be taken</param>
    ''' <returns></returns>
    Public Function GetHostIp(PriorityIpHead As String) As String
        Try
            Dim strIpList As String = Me.mGetHostIpList(False)
            Dim strFirstIp As String = ""
            GetHostIp = ""
            Do While True
                Dim strIp As String = Me.GetStr(strIpList, "", ";")
                If strIp = "" Then Exit Do
                If strIp <> "127.0.0.1" Then
                    If strFirstIp = "" Then strFirstIp = strIp
                    If Left(strIp, Len(PriorityIpHead)) = PriorityIpHead Then
                        GetHostIp = strIp
                        Exit Do
                    End If
                End If
            Loop
            If GetHostIp = "" Then GetHostIp = strFirstIp
        Catch ex As Exception
            Me.SetSubErrInf("GetHostIp", ex)
            Return ""
        End Try
    End Function


    Public Function GetHostIp(IsIPv6 As Boolean) As String
        Return mGetHostIp(IsIPv6)
    End Function

    Public Function GetHostIp(IsIPv6 As Boolean, IpHead As String) As String
        Return mGetHostIp(IsIPv6, IpHead)
    End Function

    Private Function mGetHostIpList(IsIPv6 As Boolean) As String
        mGetHostIpList = ""
        For Each oIPAddress As IPAddress In Dns.GetHostAddresses(Dns.GetHostName)
            With oIPAddress
                Dim strHostIp = .ToString()
                If IsIPv6 = True Then
                    If InStr(strHostIp, ":") > 0 Then mGetHostIpList &= strHostIp & ";"
                ElseIf InStr(strHostIp, ".") > 0 Then
                    mGetHostIpList &= strHostIp & ";"
                End If
            End With
        Next
    End Function

    Private Function mGetHostIp(IsIPv6 As Boolean, Optional IpHead As String = "") As String
        mGetHostIp = ""
        For Each oIPAddress As IPAddress In Dns.GetHostAddresses(Dns.GetHostName)
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
    Public Function GetRandString(StrLen As Integer, Optional MemberType As EnmGetRandString = EnmGetRandString.DisplayChar) As String
        Dim i As Integer
        Dim intChar As Integer
        Try
            GetRandString = ""
            For i = 1 To StrLen
                Select Case MemberType
                    Case EnmGetRandString.AllAsciiChar
                        intChar = GetRandNum(0, 255)
                    Case EnmGetRandString.DisplayChar    '!-~
                        intChar = GetRandNum(33, 126)
                    Case EnmGetRandString.NumberAndLetter
                        intChar = GetRandNum(1, 3)
                        Select Case intChar
                            Case 1  '0-9
                                intChar = GetRandNum(48, 57)
                            Case 2  'A-Z
                                intChar = GetRandNum(65, 90)
                            Case 3  'a-z
                                intChar = GetRandNum(97, 122)
                        End Select
                    Case EnmGetRandString.NumberOnly
                        intChar = GetRandNum(48, 57)
                End Select
                If MemberType = EnmGetRandString.AllAsciiChar Then
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
            strRndStr = GetRandString(64, EnmGetRandString.DisplayChar)
            strRndKey = GetRandString(intStrLen, EnmGetRandString.NumberOnly)
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

    Public Function GetPathPart(ByVal FileOrDirPath As String, Optional PathPart As EnmFathPart = EnmFathPart.FileOrDirTitle) As String
        Try
            Dim strPathPart As String = ""
            Select Case PathPart
                Case EnmFathPart.DriveLetter
                    strPathPart = Path.GetPathRoot(FileOrDirPath)
                    If InStr(strPathPart, ":") > 0 Then
                        strPathPart = UCase(Left(strPathPart, 1))
                    Else
                        strPathPart = ""
                    End If
                Case EnmFathPart.ExtName
                    strPathPart = Path.GetExtension(FileOrDirPath)
                    If InStr(strPathPart, ".") > 0 Then
                        strPathPart = Mid(strPathPart, 2)
                    End If
                Case EnmFathPart.FileOrDirTitle
                    strPathPart = Path.GetFileName(FileOrDirPath)
                Case EnmFathPart.ParentPath
                    strPathPart = Path.GetDirectoryName(FileOrDirPath)
                Case EnmFathPart.FileOrDirTitleWithoutExtName
                    strPathPart = Path.GetFileNameWithoutExtension(FileOrDirPath)
                Case Else
                    Throw New Exception("Invalid FilePart")
            End Select
            Return strPathPart
        Catch ex As Exception
            Me.SetSubErrInf("GetPathPart", ex)
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Get the part of the file|获取文件的组成部分
    ''' </summary>
    ''' <param name="FilePath">File absolute path|文件绝对路径</param>
    ''' <param name="FilePart">Part of the file|文件的部分</param>
    ''' <returns></returns>
    Public Function GetFilePart(ByVal FilePath As String, Optional FilePart As EnmFilePart = EnmFilePart.FileTitle) As String
        Dim strTemp As String, i As Long, lngLen As Long
        Dim strPath As String = "", strFileTitle As String = ""
        Dim strOsPathSep As String = ""
        Try
            If InStr(FilePath, "/") > 0 Then
                strOsPathSep = "/"
            Else
                strOsPathSep = "\"
            End If
            GetFilePart = ""
            Select Case FilePart
                Case EnmFilePart.DriveNo
                    GetFilePart = GetStr(FilePath, "", ":", False)
                    If GetFilePart = "" Then
                        GetFilePart = GetStr(FilePath, "", "$", False)
                        If GetFilePart <> "" Then GetFilePart = GetFilePart & "$"
                    End If
                Case EnmFilePart.ExtName
                    lngLen = Len(FilePath)
                    For i = lngLen To 1 Step -1
                        Select Case Mid(FilePath, i, 1)
                            Case "/", ":", "$", "\"
                                Exit For
                            Case "."
                                GetFilePart = Mid(FilePath, i + 1)
                                Exit For
                        End Select

                    Next
                Case EnmFilePart.FileTitle, EnmFilePart.Path
                    Do While True
                        strTemp = GetStr(FilePath, "", strOsPathSep, True)
                        If Len(strTemp) = 0 Then
                            If Right(strPath, 1) = strOsPathSep Then
                                If Right(strPath, 2) <> ":" & strOsPathSep Then
                                    strPath = Left(strPath, Len(strPath) - 1)
                                End If
                            ElseIf Left(FilePath, 1) = strOsPathSep Then
                                strPath = strOsPathSep
                                FilePath = Mid(FilePath, 2)
                            End If
                            If FilePath <> "" Then
                                strFileTitle = FilePath
                            Else
                                strFileTitle = "."
                            End If
                            Exit Do
                        End If
                        strPath = strPath & strTemp & strOsPathSep
                    Loop
                    If FilePart = EnmFilePart.FileTitle Then
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
        Try
            Dim lngBegin As Long
            Dim lngEnd As Long
            Dim lngBeginLen As Long
            Dim lngEndLen As Long
            lngBeginLen = Len(strBegin)
            lngBegin = InStr(SourceStr, strBegin, CompareMethod.Text)
            lngEndLen = Len(strEnd)
            If lngEndLen = 0 Then
                lngEnd = Len(SourceStr) + 1
            Else
                lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd, CompareMethod.Text)
                If lngBegin = 0 Then Return "" 'Throw New Exception("lngBegin=0")
            End If
            If lngEnd <= lngBegin Then Return "" ' Throw New Exception("lngEnd <= lngBegin")
            If lngBegin = 0 Then Return "" 'Throw New Exception("lngBegin=0[2]")
            GetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
            If IsCut = True Then
                SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
            End If
        Catch ex As Exception
            Return ""
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

        Catch ex As Exception
            Me.SetSubErrInf("UrlDecode", ex)
            Return Nothing
        End Try
    End Function

    Public Function IsRegexMatch(SrcStr As String, RegularExp As String) As Boolean
        Try
            If SrcStr = "" Then
                Return False
            Else
                Return System.Text.RegularExpressions.Regex.IsMatch(SrcStr, RegularExp)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsRegexMatch", ex)
            Return False
        End Try
    End Function

    Public Function GECBool(vData As String) As Boolean
        Try
            Return CBool(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECBool", ex)
            Return False
        End Try
    End Function

    Public Function GECBool(vData As String, IsEmptyTrue As Boolean) As Boolean
        Try
            If IsEmptyTrue = True And vData = "" Then
                Return True
            Else
                Return CBool(vData)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GECBool", ex)
            Return False
        End Try
    End Function

    Public Function GEInt(vData As String) As Integer
        Try
            Return CInt(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GEInt", ex)
            Return 0
        End Try
    End Function

    Public Function GECLng(vData As String) As Long
        Try
            Return CLng(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
        End Try
    End Function

    Public Function GECInt(vData As String) As Integer
        Try
            Return CInt(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECInt", ex)
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

    Public Function SQLCDate(vData As String) As DateTime
        Try
            Return Date.Parse(vData)
        Catch ex As Exception
            Me.SetSubErrInf("SQLCDate", ex)
            Return #1/1/1753#
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
            If SrcStr Is Nothing Then SrcStr = ""
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
            If CtlStr Is Nothing Then CtlStr = ""
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
            MainText &= strTabs & NewLine & Me.OsCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddMultiLineText", ex)
        End Try
    End Function

    Public Function AddMultiLineText(ByRef MainText As StringBuilder, NewLine As String, Optional LeftTabs As Integer = 0) As String
        Try
            If LeftTabs > 0 Then
                For i = 1 To LeftTabs
                    MainText.Append(vbTab)
                Next
            End If
            MainText.Append(NewLine & Me.OsCrLf)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddMultiLineText", ex)
        End Try
    End Function

    Public Function IsFileExists(FilePath As String) As Boolean
        Try
            Return IO.File.Exists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("IsFileExists", ex)
            Return Nothing
        End Try
    End Function

    Public Function IsFolderExists(FolderPath As String, Optional IsNotExistsCreate As Boolean = False) As Boolean
        Try
            If Directory.Exists(FolderPath) = True Then
                Return True
            ElseIf IsNotExistsCreate = True Then
                Dim strRet As String = Me.CreateFolder(FolderPath)
                If strRet <> "OK" Then Throw New Exception(strRet)
                Return Directory.Exists(FolderPath)
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsFolderExists", ex)
            Return False
        End Try
    End Function


    Public Function GetEnvVar(EnvVarName As String) As String
        Return GetEnvironmentVariable(EnvVarName)
    End Function

    Public Function GetUserName() As String
        Return Environment.UserName
    End Function
    Public Function CreateFolder(FolderPath As String) As String
        Try
            Directory.CreateDirectory(FolderPath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CreateFolder", ex)
        End Try
    End Function

    Public Function MoveFile(SrcFilePath As String, DestFilePath As String, Optional IsForceOverride As Boolean = False) As String
        Dim LOG As New PigStepLog("MoveFile")
        Try
            If IsForceOverride = True Then
                If File.Exists(DestFilePath) = True Then
                    LOG.StepName = "Delete(DestFilePath)"
                    File.Delete(DestFilePath)
                End If
            End If
            LOG.StepName = "Move(SrcFilePath, DestFilePath)"
            File.Move(SrcFilePath, DestFilePath)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(SrcFilePath)
            LOG.AddStepNameInf(DestFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function DeleteFile(FilePath As String) As String
        Try
            File.Delete(FilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFile", ex)
        End Try
    End Function

    Public Function DeleteFolder(FolderPath As String) As String
        Try
            Directory.Delete(FolderPath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFolder", ex)
        End Try
    End Function

    Public Function DeleteFolder(FolderPath As String, IsSubDir As Boolean) As String
        Try
            Directory.Delete(FolderPath, IsSubDir)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFolder", ex)
        End Try
    End Function

    Public Function GetComputerName() As String
        Return Dns.GetHostName()
    End Function

    Public Function IsOsWindows() As Boolean
        Return Me.IsWindows
    End Function

    Public Function MyOsCrLf() As String
        Return Me.OsCrLf
    End Function
    Public Function MyOsPathSep() As String
        Return Me.OsPathSep
    End Function

    Public Function GetFileVersion(FilePath As String, ByRef FileVersion As String) As String
        Dim LOG As New PigStepLog("GetFileVersion")
        Try
            LOG.StepName = "GetVersionInfo"
            Dim oGetVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(FilePath)
            FileVersion = oGetVersionInfo.FileVersion
            oGetVersionInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileVersion = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileMD5(FilePath As String, ByRef FileMD5 As String) As String
        Dim LOG As New PigStepLog("GetFileMD5")
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(FilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            FileMD5 = oPigFile.MD5
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            GetFileMD5 = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetBytesPigMD5(ByRef InBytes As Byte(), ByRef OutPigMD5 As String) As String
        Try
            Dim oPigMD5 As New PigMD5(InBytes)
            OutPigMD5 = oPigMD5.PigMD5
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            OutPigMD5 = ""
            Return Me.GetSubErrInf("GetBytesPigMD5", ex)
        End Try
    End Function

    Public Function GetTextPigMD5(SrcText As String, TextType As PigMD5.enmTextType, ByRef OutPigMD5 As String) As String
        Dim LOG As New PigStepLog("GetTextPigMD5")
        Try
            Dim oPigMD5 As New PigMD5(SrcText, TextType)
            OutPigMD5 = oPigMD5.PigMD5
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            OutPigMD5 = ""
            Return Me.GetSubErrInf("GetTextPigMD5", ex)
        End Try
    End Function

    Public Function GetFilePigMD5(FilePath As String, ByRef OutPigMD5 As String) As String
        Dim LOG As New PigStepLog("GetFilePigMD5")
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(FilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            OutPigMD5 = oPigFile.PigMD5
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            OutPigMD5 = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileCreateTime(FilePath As String, ByRef FileCreateTime As Date) As String
        Dim LOG As New PigStepLog("GetFileCreateTime")
        Try
            LOG.StepName = "New FileInfo"
            Dim oFileInfo As New FileInfo(FilePath)
            LOG.StepName = "CreationTime"
            FileCreateTime = oFileInfo.CreationTime
            oFileInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileCreateTime = #1/1/1900#
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileUpdateTime(FilePath As String, ByRef FileUpdateTime As Date) As String
        Dim LOG As New PigStepLog("GetFileUpdateTime")
        Try
            LOG.StepName = "New FileInfo"
            Dim oFileInfo As New FileInfo(FilePath)
            LOG.StepName = "LastWriteTime"
            FileUpdateTime = oFileInfo.LastWriteTime
            oFileInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileUpdateTime = #1/1/1900#
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileText(FilePath As String, ByRef FileText As String) As String
        Dim LOG As New PigStepLog("GetFileText")
        Try
            LOG.StepName = "New StreamReader"
            Dim srMain As New StreamReader(FilePath)
            LOG.StepName = "ReadToEnd"
            FileText = srMain.ReadToEnd()
            LOG.StepName = "Close"
            srMain.Close()
            Return "OK"
        Catch ex As Exception
            FileText = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ASyncSaveTextToFile(FilePath As String, SaveText As String, ByRef OutThreadID As Integer) As String
        Return Me.mASyncSaveTextToFile(FilePath, SaveText, OutThreadID)
    End Function


    Private Function mASyncSaveTextToFile(FilePath As String, SaveText As String, ByRef OutThreadID As Integer) As String
        Try
            Dim struMain As mStruSaveTextToFile
            With struMain
                .FilePath = FilePath
                .SaveText = SaveText
                .IsAsync = True
            End With
            Dim oThread As New Thread(AddressOf mSaveTextToFile)
            oThread.Start(struMain)
            OutThreadID = oThread.ManagedThreadId
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mASyncSaveTextToFile", ex)
        End Try
    End Function

    Public Function SaveTextToFile(FilePath As String, SaveText As String) As String
        Dim struMain As mStruSaveTextToFile
        With struMain
            .FilePath = FilePath
            .SaveText = SaveText
        End With
        Return Me.mSaveTextToFile(struMain)
    End Function


    Private Structure mStruSaveTextToFile
        Public FilePath As String
        Public SaveText As String
        Public IsAsync As Boolean
    End Structure

    Private Function mSaveTextToFile(StruMain As mStruSaveTextToFile) As String
        Dim LOG As New PigStepLog("mSaveTextToFile")
        Dim sarMain As StruASyncRet
        Try
            If StruMain.IsAsync = True Then
                sarMain.BeginTime = Now
            End If
            LOG.StepName = "New StreamWriter"
            Dim swMain As New StreamWriter(StruMain.FilePath, False)
            LOG.StepName = "Write"
            swMain.Write(StruMain.SaveText)
            LOG.StepName = "Close"
            swMain.Close()
            If StruMain.IsAsync = True Then
                sarMain.EndTime = Now
                sarMain.ThreadID = Thread.CurrentThread.ManagedThreadId
                sarMain.Ret = "OK"
                RaiseEvent ASyncRet_SaveTextToFile(sarMain)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(StruMain.FilePath)
            If StruMain.IsAsync = True Then
                sarMain.EndTime = Now
                sarMain.ThreadID = Thread.CurrentThread.ManagedThreadId
                sarMain.Ret = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
                RaiseEvent ASyncRet_SaveTextToFile(sarMain)
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 将一行PLSQL Developer 导出的CSV文本行转换成标准的SQL Server BCP格式行|Convert a line of CSV text exported by PL/SQL Developer into a standard SQL Server BCP format line
    ''' </summary>
    ''' <param name="CsvLine"></param>
    ''' <param name="BcpLine"></param>
    ''' <returns></returns>
    Public Function PLSqlCsv2Bcp(CsvLine As String, ByRef BcpLine As String) As String
        Try
            CsvLine &= ","
            Do While True
                If InStr(CsvLine, """"",") = 0 Then Exit Do
                CsvLine = Replace(CsvLine, """"",", """ "",")
            Loop
            BcpLine = ""
            Dim bolIsBegin As Boolean = False
            Do While True
                Dim strCol As String = Me.GetStr(CsvLine, """", """,")
                If strCol = "" Then Exit Do
                If strCol = " " Then strCol = ""
                If bolIsBegin = True Then
                    BcpLine &= vbTab
                Else
                    bolIsBegin = True
                End If
                BcpLine &= strCol
            Loop
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("PLSqlCsv2Bcp", ex)
        End Try
    End Function

    Public Function GetHumanSize(SrcSize As Decimal) As String
        Try
            Dim strSize As String = SrcSize.ToString
            If SrcSize < 0 Then Throw New Exception("Invalid size")
            If SrcSize < 1024 Then
                strSize = SrcSize.ToString & " bytes"
            Else
                SrcSize = SrcSize / 1024
                If SrcSize < 1024 Then
                    strSize = Math.Round(SrcSize, 2) & " KB"
                Else
                    SrcSize = SrcSize / 1024
                    If SrcSize < 1024 Then
                        strSize = Math.Round(SrcSize, 2) & " MB"
                    Else
                        SrcSize = SrcSize / 1024
                        If SrcSize < 1024 Then
                            strSize = Math.Round(SrcSize, 2) & " GB"
                        Else
                            SrcSize = SrcSize / 1024
                            If SrcSize < 1024 Then
                                strSize = Math.Round(SrcSize, 2) & " TB"
                            Else
                                SrcSize = SrcSize / 1024
                                If SrcSize < 1024 Then
                                    strSize = Math.Round(SrcSize, 2) & " PB"
                                Else
                                    SrcSize = SrcSize / 1024
                                    strSize = Math.Round(SrcSize, 2) & " EB"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Return strSize
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 将数字转换成中文显示|Convert numbers to Chinese display
    ''' </summary>
    ''' <param name="Digital">数字</param>
    ''' <param name="IsDecimal">是否小数，不是小数部分只支持千百十个位</param>
    ''' <param name="IsUCase">是否大写数字</param>
    ''' <returns></returns>
    Public Function DigitalToChnName(Digital As String, IsDecimal As Boolean, Optional IsUCase As Boolean = False) As String
        Try
            Dim i As Integer
            Dim lngDigital As Integer
            Dim lngDigitalLen As Integer
            Dim strDigital(0 To 9) As String
            Dim strUnit(0 To 3) As String
            If IsDecimal = False Then
                Select Case Len(Digital)
                    Case Is < 4
                        Digital = Right("0000" & Digital, 4)
                    Case Is > 4
                        Digital = Right(Digital, 4)
                End Select
            End If
            lngDigitalLen = Len(Digital)

            strDigital(0) = "零"
            If IsUCase = True Then
                strDigital(1) = "壹"
                strDigital(2) = "贰"
                strDigital(3) = "叁"
                strDigital(4) = "肆"
                strDigital(5) = "伍"
                strDigital(6) = "陆"
                strDigital(7) = "柒"
                strDigital(8) = "捌"
                strDigital(9) = "玖"

                strUnit(0) = "仟"
                strUnit(1) = "佰"
                strUnit(2) = "拾"
                strUnit(3) = ""
            Else
                strDigital(1) = "一"
                strDigital(2) = "二"
                strDigital(3) = "三"
                strDigital(4) = "四"
                strDigital(5) = "五"
                strDigital(6) = "六"
                strDigital(7) = "七"
                strDigital(8) = "八"
                strDigital(9) = "九"

                strUnit(0) = "千"
                strUnit(1) = "百"
                strUnit(2) = "十"
                strUnit(3) = ""
            End If
            DigitalToChnName = ""
            For i = 0 To lngDigitalLen - 1
                lngDigital = CLng(Mid(Digital, i + 1, 1))
                If IsDecimal = False Then
                    '这种情况lngDigitalLen为4
                    If lngDigital = 0 Then
                        DigitalToChnName &= strDigital(lngDigital)
                    Else
                        DigitalToChnName &= strDigital(lngDigital) & strUnit(i)
                    End If
                Else
                    DigitalToChnName &= strDigital(lngDigital)
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("DigitalToChnName", ex)
            Return ""
        End Try

    End Function

    ''' <summary>
    ''' 转换HTML标记
    ''' </summary>
    ''' <param name="HtmlStr">源HTML字符串|Source HTML string</param>
    ''' <param name="HowToConvHtml">如何转换|How to convert</param>
    ''' <returns></returns>
    Public Function ConvertHtmlStr(SrcHtml As String, HowToConvHtml As EmnHowToConvHtml) As String

        Try
            ConvertHtmlStr = SrcHtml
            Select Case HowToConvHtml
                Case EmnHowToConvHtml.DisableHTML
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "<", "&lt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, ">", "&gt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, " ", "&nbsp;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, vbCrLf, "<br>")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;br&gt;", "<br>")
                Case EmnHowToConvHtml.EnableHTML
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;", "<")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&gt;", ">")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&nbsp;&nbsp;&nbsp;&nbsp;", vbTab) '必须在空格转换前面
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&nbsp;", " ")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;br&gt;", "<br>")
                Case EmnHowToConvHtml.DisableHTMLOnlySymbol
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "<", "&lt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, ">", "&gt;")
                Case EmnHowToConvHtml.EnableHTMLOnlySymbol
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;", "<")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&gt;", ">")
                Case Else
                    Throw New Exception("Invalid HowToConvHtml " & HowToConvHtml.ToString)
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("ConvertHtmlStr", ex)
            Return ""
        End Try
    End Function

    Public Function GetWindowsProductId(Optional ByRef Res As String = "OK") As String
        Dim LOG As New PigStepLog("GetWindowsProductId")
        Try
            If Me.IsWindows = True Then
                LOG.StepName = "New PigReg"
                Dim oPigReg As New PigReg
                LOG.StepName = "GetRegValue"
                GetWindowsProductId = oPigReg.GetRegStrValue(PigReg.EmnRegRoot.LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId", "")
                If oPigReg.LastErr <> "" Then Throw New Exception(oPigReg.LastErr)
            Else
                Throw New Exception("Only run on Windows")
            End If
        Catch ex As Exception
            Res = ex.Message.ToString
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    Private Function mGetMachineGUID(ByRef OutGUID As String) As String
        Dim LOG As New PigStepLog("mGetMachineGUID")
        Try
            If Me.IsWindows = True Then
                LOG.StepName = "New PigReg"
                Dim oPigReg As New PigReg
                LOG.StepName = "GetRegValue"
                OutGUID = oPigReg.GetRegStrValue(PigReg.EmnRegRoot.LOCAL_MACHINE, "SOFTWARE\Microsoft\Cryptography", "MachineGuid", "")
                If oPigReg.LastErr <> "" Then Throw New Exception(oPigReg.LastErr)
                oPigReg = Nothing
            Else
                Throw New Exception("Only run on Windows")
            End If
            Return "OK"
        Catch ex As Exception
            OutGUID = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetMachineGUID() As String
        Try
            Dim strGUID As String = ""
            Dim strRet As String = Me.mGetMachineGUID(strGUID)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return strGUID
        Catch ex As Exception
            Me.SetSubErrInf("GetMachineGUID", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get the product ID of Linux operating system|获取Linux操作系统的产品标识
    ''' </summary>
    ''' <param name="RetInf">Returning OK indicates success, others indicate failure|返回OK表示成功，其他为失败</param>
    ''' <returns></returns>
    Public Function GetProductUuid(Optional ByRef RetInf As String = "") As String
        Dim LOG As New PigStepLog("GetProductUuid")
        Try
            Dim strProductUuid As String = ""
            Dim strFilePath As String = "/sys/class/dmi/id/product_uuid"
            If Me.IsFileExists(strFilePath) = True Then
                LOG.StepName = "GetFileText"
                LOG.Ret = Me.GetFileText(strFilePath, strProductUuid)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                strProductUuid = Replace(Trim(strProductUuid), vbCrLf, "")
                strProductUuid = Replace(strProductUuid, Me.OsCrLf, "")
            Else
                Throw New Exception(strFilePath & " not found.")
            End If
            RetInf = "OK"
            Return strProductUuid
        Catch ex As Exception
            RetInf = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get the unique boot ID of the Linux operating system|获取Linux操作系统的开机唯一标识
    ''' </summary>
    ''' <param name="RetInf">Returning OK indicates success, others indicate failure|返回OK表示成功，其他为失败</param>
    ''' <returns></returns>
    Public Function GetBootID(Optional ByRef RetInf As String = "") As String
        Dim LOG As New PigStepLog("GetBootID")
        Try
            Dim strBootID As String = ""
            Dim strFilePath As String = "/proc/sys/kernel/random/boot_id"
            If Me.IsFileExists(strFilePath) = True Then
                LOG.StepName = "GetFileText"
                LOG.Ret = Me.GetFileText(strFilePath, strBootID)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                strBootID = Replace(Trim(strBootID), vbCrLf, "")
                strBootID = Replace(strBootID, Me.OsCrLf, "")
            Else
                Throw New Exception(strFilePath & " not found.")
            End If
            RetInf = "OK"
            Return strBootID
        Catch ex As Exception
            RetInf = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    Private Function mGetUUID(ByRef OutUUID As String) As String
        Dim LOG As New PigStepLog("mGetUUID")
        Try
            Dim strFilePath As String = "/proc/sys/kernel/random/uuid"
            If Me.IsFileExists(strFilePath) = True Then
                LOG.StepName = "GetFileText"
                LOG.Ret = Me.GetFileText(strFilePath, OutUUID)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                OutUUID = Replace(Trim(OutUUID), vbCrLf, "")
                OutUUID = Replace(OutUUID, Me.OsCrLf, "")
            Else
                Throw New Exception(strFilePath & " not found.")
            End If
            Return "OK"
        Catch ex As Exception
            OutUUID = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function
    Public Function GetUUID() As String
        Try
            'If Me.IsWindows = True Then Throw New Exception("Only run on Linux")
            Dim strUUID As String = ""
            Dim strRet As String = Me.mGetUUID(strUUID)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return strUUID
        Catch ex As Exception
            Me.SetSubErrInf("GetUUID", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 获取对齐的字符串|Gets the aligned string
    ''' </summary>
    ''' <param name="SrcStr">源串|Source string</param>
    ''' <param name="Alignment">对齐方式|Alignment</param>
    ''' <param name="RowLen">行长度|Row length</param>
    ''' <returns></returns>
    Public Function GetAlignStr(SrcStr As String, Alignment As EnmAlignment, RowLen As Integer) As String
        Try
            Dim intSrcLen As Integer = Len(SrcStr)
            GetAlignStr = ""
            Select Case Alignment
                Case EnmAlignment.Left
                    If intSrcLen >= RowLen Then
                        GetAlignStr = Left(SrcStr, RowLen)
                    Else
                        GetAlignStr = SrcStr & Me.GetRepeatStr(RowLen - intSrcLen， " ")
                    End If
                Case EnmAlignment.Right
                    If intSrcLen >= RowLen Then
                        GetAlignStr = Right(SrcStr, RowLen)
                    Else
                        GetAlignStr = Me.GetRepeatStr(RowLen - intSrcLen， " ") & SrcStr
                    End If
                Case EnmAlignment.Center
                    Dim intBegin As Integer = (RowLen - intSrcLen) / 2
                    Dim intMidLen As Integer = intSrcLen
                    If intBegin < 1 Then intBegin = 1
                    Select Case intSrcLen
                        Case < RowLen
                            GetAlignStr = Me.GetRepeatStr(intBegin， " ") & SrcStr & Me.GetRepeatStr(RowLen - intBegin - intSrcLen, " ")
                        Case = RowLen
                            GetAlignStr = SrcStr
                        Case > RowLen
                            intBegin = (intSrcLen - RowLen) / 2
                            If intBegin < 1 Then intBegin = 1
                            intMidLen = RowLen
                            GetAlignStr = Mid(SrcStr, intBegin, intMidLen)
                    End Select
                Case Else
                    Throw New Exception("Invalid Alignment " & Alignment.ToString)
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("GetAlignStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 按 ASCII 习惯获取 Unicode 对齐的字符串|Obtain Unicode aligned strings according to ASCII conventions
    ''' </summary>
    ''' <param name="SrcStr">源串|Source string</param>
    ''' <param name="Alignment">对齐方式|Alignment</param>
    ''' <param name="RowLen">行长度|Row length</param>
    ''' <returns></returns>
    Public Function GetAlignStrA(SrcStr As String, Alignment As EnmAlignment, RowLen As Integer) As String
        Try
            Dim intSrcLen As Integer = LenA(SrcStr)
            GetAlignStrA = ""
            Select Case Alignment
                Case EnmAlignment.Left
                    If intSrcLen >= RowLen Then
                        GetAlignStrA = LeftA(SrcStr, RowLen)
                    Else
                        GetAlignStrA = SrcStr & Me.GetRepeatStr(RowLen - intSrcLen， " ")
                    End If
                Case EnmAlignment.Right
                    If intSrcLen >= RowLen Then
                        GetAlignStrA = RightA(SrcStr, RowLen)
                    Else
                        GetAlignStrA = Me.GetRepeatStr(RowLen - intSrcLen， " ") & SrcStr
                    End If
                Case EnmAlignment.Center
                    Dim intBegin As Integer = (RowLen - intSrcLen) / 2
                    Dim intMidLen As Integer = intSrcLen
                    If intBegin < 1 Then intBegin = 1
                    Select Case intSrcLen
                        Case < RowLen
                            GetAlignStrA = Me.GetRepeatStr(intBegin， " ") & SrcStr & Me.GetRepeatStr(RowLen - intBegin - intSrcLen, " ")
                        Case = RowLen
                            GetAlignStrA = SrcStr
                        Case > RowLen
                            intBegin = (intSrcLen - RowLen) / 2
                            If intBegin < 1 Then intBegin = 1
                            intMidLen = RowLen
                            GetAlignStrA = MidA(SrcStr, intBegin, intMidLen)
                    End Select
                Case Else
                    Throw New Exception("Invalid Alignment " & Alignment.ToString)
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("GetAlignStrA", ex)
            Return ""
        End Try
    End Function

    Public Function GetRepeatStr(Number As Integer, SrcStr As String) As String
        Try
            GetRepeatStr = ""
            For i = 1 To Number
                GetRepeatStr &= SrcStr
            Next
        Catch ex As Exception
            Me.SetSubErrInf("GetRepeatStr", ex)
            Return ""
        End Try
    End Function

    Public Function SaveShareMem(SMName As String, InBytes As Byte()) As String
        Dim LOG As New PigStepLog("SaveShareMem")
        Const SM_HEAD_LEN As Integer = 28
        Try
            If InBytes Is Nothing Then Throw New Exception("InBytes Is Nothing")
            LOG.StepName = "GetTextPigMD5"
            LOG.Ret = Me.GetTextPigMD5("~PigShareMem." & SMName & ">PigShareMem,", PigMD5.enmTextType.UTF8, SMName)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim intDataLen As Integer = InBytes.Length
            Dim dteCreate As Date = Now
            LOG.StepName = "New PigMD5"
            Dim pbHead As New PigBytes
            Dim oPigMD5 As New PigMD5(InBytes)
            If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
            With pbHead
                LOG.StepName = "pbHead.SetValue"
                .SetValue(intDataLen)
                .SetValue(dteCreate)
                .SetValue(oPigMD5.PigMD5Bytes)
                If .LastErr <> "" Then Throw New Exception(.LastErr)
            End With
            LOG.StepName = "New ShareMem(Head)"
            Dim smHead As New ShareMem
            With smHead
                LOG.StepName = "Init(Head)"
                LOG.Ret = .Init(SMName, SM_HEAD_LEN)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "Write(Head)"
                LOG.Ret = .Write(pbHead.Main)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            smHead = Nothing
            LOG.StepName = "New ShareMem(Body)"
            Dim smBody As New ShareMem
            With smBody
                LOG.StepName = "Init(Body)"
                LOG.Ret = .Init(oPigMD5.PigMD5, intDataLen)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "Write(Body)"
                LOG.Ret = .Write(InBytes)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            smBody = Nothing
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetPigMD5OrMD5(PigMD5OrMD5Bytes As Byte()) As String
        Try
            GetPigMD5OrMD5 = ""
            For i = 0 To 15
                GetPigMD5OrMD5 &= Right("00" & Hex(PigMD5OrMD5Bytes(i)).ToLower, 2)
            Next
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function GetShareMem(SMName As String, ByRef OutBytes As Byte(), ByRef CreateTime As Date) As String
        Dim LOG As New PigStepLog("GetShareMem")
        Const SM_HEAD_LEN As Integer = 28
        Try
            LOG.StepName = "GetTextPigMD5"
            LOG.Ret = Me.GetTextPigMD5("~PigShareMem." & SMName & ">PigShareMem,", PigMD5.enmTextType.UTF8, SMName)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "New ShareMem(Head)"
            Dim abHead(-1) As Byte
            Dim smHead As New ShareMem
            With smHead
                LOG.StepName = "Init(Head)"
                LOG.Ret = .Init(SMName, SM_HEAD_LEN)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                LOG.StepName = "Read(Head)"
                LOG.Ret = .Read(abHead)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            Dim intDataLen As Integer, abMD5(-1) As Byte
            Dim pbHead As New PigBytes(abHead)
            With pbHead
                LOG.StepName = "pbHead.GetValue"
                intDataLen = .GetInt32Value
                If intDataLen <= 0 Then Throw New Exception("No data")
                CreateTime = .GetDateTimeValue
                abMD5 = .GetBytesValue(16)
                If .LastErr <> "" Then Throw New Exception(.LastErr)
            End With
            Dim strDataMD5 As String = Me.GetPigMD5OrMD5(abMD5)
            If strDataMD5 = "" Then Throw New Exception("Invalid data")
            smHead = Nothing
            LOG.StepName = "New ShareMem(Body)"
            Dim smBody As New ShareMem
            With smBody
                LOG.StepName = "Init(Body)"
                LOG.Ret = .Init(strDataMD5, intDataLen)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                ReDim OutBytes(0)
                LOG.StepName = "Read(Body)"
                LOG.Ret = .Read(OutBytes)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End With
            smBody = Nothing
            Dim oPigMD5 As New PigMD5(OutBytes)
            If oPigMD5.LastErr <> "" Then Throw New Exception(oPigMD5.LastErr)
            If oPigMD5.PigMD5 <> strDataMD5 Then Throw New Exception("Data mismatch")
            oPigMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            ReDim OutBytes(0)
            CreateTime = Date.MinValue
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetMyExeName() As String
        Try
            Dim intPID As Integer = Me.fMyPID
            Dim oPigProc As New PigProc(Me.fMyPID)
            If oPigProc.LastErr <> "" Then Throw New Exception(oPigProc.LastErr)
            If Me.IsWindows = True Then
                GetMyExeName = oPigProc.ProcessName & ".exe"
            Else
                GetMyExeName = oPigProc.ProcessName
            End If
            oPigProc = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetMyExeName", ex)
            Return ""
        End Try
    End Function

    Public Function GetMyExePath() As String
        Try
            Dim intPID As Integer = Me.fMyPID
            Dim oPigProc As New PigProc(Me.fMyPID)
            If oPigProc.LastErr <> "" Then Throw New Exception(oPigProc.LastErr)
            GetMyExePath = oPigProc.FilePath
            oPigProc = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetMyExePath", ex)
            Return ""
        End Try
    End Function

    Public Function GetMyPigProc() As PigProc
        Try
            Dim intPID As Integer = Me.fMyPID
            GetMyPigProc = New PigProc(Me.fMyPID)
            If GetMyPigProc.LastErr <> "" Then Throw New Exception(GetMyPigProc.LastErr)
        Catch ex As Exception
            Me.SetSubErrInf("GetMyPigProc", ex)
            Return Nothing
        End Try
    End Function

    Public Function Is64Bit() As Boolean
        If System.Runtime.InteropServices.Marshal.SizeOf(IntPtr.Zero) * 8 = 64 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function IsFileDiff(FilePath1 As String, FilePath2 As String, Optional ByRef Res As String = "OK") As Boolean
        Try
            Dim strRet As String
            If FilePath1 = FilePath2 Then
                IsFileDiff = True
            Else
                Dim pfSrc As New PigFile(FilePath1)
                Dim lngSize1 As Long = pfSrc.Size
                If lngSize1 < 0 Then Throw New Exception(FilePath1 & " size invalid")
                Dim pfTar As New PigFile(FilePath2)
                Dim lngSize2 As Long = pfTar.Size
                If lngSize2 < 0 Then Throw New Exception(FilePath2 & " size invalid")
                If lngSize1 <> lngSize2 Then
                    IsFileDiff = False
                Else
                    Dim pmSrc As PigMD5 = Nothing, pmTar As PigMD5 = Nothing
                    strRet = pfSrc.GetFastPigMD5(pmSrc)
                    If strRet <> "OK" Then Throw New Exception(strRet)
                    strRet = pfSrc.GetFastPigMD5(pmTar)
                    If strRet <> "OK" Then Throw New Exception(strRet)
                    If pmSrc.PigMD5 <> pmTar.PigMD5 Then
                        IsFileDiff = False
                    Else
                        strRet = pfSrc.GetFullPigMD5(pmSrc)
                        If strRet <> "OK" Then Throw New Exception(strRet)
                        strRet = pfSrc.GetFullPigMD5(pmTar)
                        If strRet <> "OK" Then Throw New Exception(strRet)
                        If pmSrc.PigMD5 <> pmTar.PigMD5 Then
                            IsFileDiff = False
                        Else
                            IsFileDiff = True
                        End If
                    End If
                    pmSrc = Nothing
                    pmTar = Nothing
                End If
                pfTar = Nothing
                pfSrc = Nothing
            End If
        Catch ex As Exception
            IsFileDiff = False
            Res = Me.GetSubErrInf("IsFileDiff", ex)
        End Try
    End Function

    Public Function CheckFileDiff(SrcFile As String, TarFile As String, ByRef IsDiff As Boolean, Optional CheckDiffType As EnmCheckDiffType = EnmCheckDiffType.Size_Date_FastPigMD5) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "New SrcFile"
            Dim pfSrc As New PigFile(SrcFile)
            strStepName = "New SrcFile"
            Dim pfTar As New PigFile(TarFile)
            Dim lngSrcSize As Long = pfSrc.Size
            If lngSrcSize < 0 Then Throw New Exception("SrcFile size invalid")
            Dim lngTarSize As Long = pfTar.Size
            If lngTarSize < 0 Then Throw New Exception("TarFile size invalid")
            If lngSrcSize <> lngTarSize Then
                IsDiff = True
            Else
                Select Case CheckDiffType
                    Case EnmCheckDiffType.Size_Date, EnmCheckDiffType.Size_Date_FastPigMD5
                        If pfSrc.UpdateTime <> pfTar.UpdateTime Then
                            IsDiff = True
                        ElseIf CheckDiffType = EnmCheckDiffType.Size_Date Then
                            IsDiff = False
                        Else
                            Dim pmSrc As PigMD5 = Nothing, pmTar As PigMD5 = Nothing
                            strStepName = "GetFastPigMD5(pmSrc)"
                            strRet = pfSrc.GetFastPigMD5(pmSrc)
                            If strRet <> "OK" Then Throw New Exception(strRet)
                            strStepName = "GetFastPigMD5(pmTar)"
                            strRet = pfTar.GetFastPigMD5(pmTar)
                            If strRet <> "OK" Then Throw New Exception(strRet)
                            strStepName = "Check Diff"
                            If pmSrc.PigMD5 <> pmTar.PigMD5 Then
                                IsDiff = True
                            Else
                                IsDiff = False
                            End If
                            pmSrc = Nothing
                            pmTar = Nothing
                        End If
                    Case EnmCheckDiffType.Size_FullPigMD5
                        Dim pmSrc As PigMD5 = Nothing, pmTar As PigMD5 = Nothing
                        strStepName = "GetFullPigMD5(pmSrc)"
                        strRet = pfSrc.GetFullPigMD5(pmSrc)
                        If strRet <> "OK" Then Throw New Exception(strRet)
                        strStepName = "GetFullPigMD5(pmTar)"
                        strRet = pfTar.GetFullPigMD5(pmTar)
                        If strRet <> "OK" Then Throw New Exception(strRet)
                        strStepName = "Check Diff"
                        If pmSrc.PigMD5 <> pmTar.PigMD5 Then
                            IsDiff = True
                        Else
                            IsDiff = False
                        End If
                        pmSrc = Nothing
                        pmTar = Nothing
                    Case Else
                        Throw New Exception("Invalid CheckDiffType" & CheckDiffType.ToString)
                End Select
            End If
            Return "OK"
        Catch ex As Exception
            IsDiff = False
            Return Me.GetSubErrInf("IsFileDiff", strStepName, ex)
        End Try
    End Function

    Public Function IsMathDate(Date1 As Date, Date2 As Date, Optional DateFmt As String = "yyyy-MM-dd HH:mm.ss.f") As Boolean
        Try
            Dim strDate1 As String = Format(Date1, DateFmt)
            Dim strDate2 As String = Format(Date2, DateFmt)
            If strDate1 = strDate2 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsMathDate", ex)
            Return False
        End Try
    End Function

    Public Function IsMathDecimal(Dec1 As Decimal, Dec2 As Decimal, Optional Decimals As Integer = 4) As Boolean
        Try
            Dim strDec1 As String = Math.Round(Dec1, Decimals)
            Dim strDec2 As String = Math.Round(Dec2, Decimals)
            If strDec1 = strDec2 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsMathDecimal", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Get the longest line in a piece of text|获取一段文本中长度最大的行
    ''' </summary>
    ''' <param name="SrcStr">Source text|源文</param>
    ''' <param name="LineSeparator">Line Separator|行分隔符</param>
    ''' <returns></returns>
    Public Function GetMaxStr(SrcStr As String, Optional LineSeparator As String = vbCrLf) As String
        Try
            If LineSeparator = "" Then LineSeparator = Me.OsCrLf
            Dim oList As New List(Of String)
            Do While True
                Dim strLine As String = Me.GetStr(SrcStr, "", LineSeparator)
                If strLine = "" Then Exit Do
                oList.Add(strLine)
            Loop
            oList.Sort()
            GetMaxStr = oList.Item(oList.Count - 1)
            oList = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetMaxStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get the line with the smallest length in a piece of text|获取一段文本中长度最小的行
    ''' </summary>
    ''' <param name="SrcStr">Source text|源文</param>
    ''' <param name="LineSeparator">Line Separator|行分隔符</param>
    ''' <returns></returns>
    Public Function GetMinStr(SrcStr As String, Optional LineSeparator As String = vbCrLf) As String
        Try
            If LineSeparator = "" Then LineSeparator = Me.OsCrLf
            Dim oList As New List(Of String)
            Do While True
                Dim strLine As String = Me.GetStr(SrcStr, "", LineSeparator)
                If strLine = "" Then Exit Do
                oList.Add(strLine)
            Loop
            oList.Sort()
            GetMinStr = oList.Item(0)
            oList = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetMinStr", ex)
            Return ""
        End Try
    End Function

    Public Function GetEnmDispStr(EnmValue As Object, Optional IsCommaFirst As Boolean = True) As String
        Try
            If IsCommaFirst = True Then
                GetEnmDispStr = ","
            Else
                GetEnmDispStr = ""
            End If
            GetEnmDispStr &= CInt(EnmValue) & "-" & EnmValue.ToString
        Catch ex As Exception
            Me.SetSubErrInf("GetEnmDispStr", ex)
            Return ""
        End Try
    End Function

    Public Function IsStrongPassword(SrcPwd As String, Optional NeedLen As Integer = 8, Optional IsNeedSymbol As Boolean = True)
        Try
            Dim intPwdLen As Integer = Len(SrcPwd)
            If intPwdLen < 6 Or intPwdLen < NeedLen Then Throw New Exception("Password is too short")
            Const SYMBOL_LIST As String = "~!@#$%^&*()_+{}|:""<>?`-=[]\;',./"
            Dim bolUpperLetter As Boolean = False, bolLowerLetter As Boolean = False, bolNumber As Boolean = False, bolIsSymbol As Boolean = False
            For i = 0 To SrcPwd.Length - 1
                Dim strChar As String = SrcPwd.Substring(i, 1)
                If bolIsSymbol = False Then
                    If InStr(SYMBOL_LIST, strChar) > 0 Then
                        bolIsSymbol = True
                    End If
                End If
                If bolNumber = False Then
                    Select Case strChar
                        Case "0" To "9"
                            bolNumber = True
                    End Select
                End If
                If bolUpperLetter = False Then
                    Select Case strChar
                        Case "A" To "Z"
                            bolUpperLetter = True
                    End Select
                End If
                If bolLowerLetter = False Then
                    Select Case strChar
                        Case "a" To "z"
                            bolLowerLetter = True
                    End Select
                End If
            Next
            If IsNeedSymbol = True Then
                If bolIsSymbol And bolLowerLetter And bolNumber And bolLowerLetter Then
                    Return True
                Else
                    Return False
                End If
            ElseIf bolLowerLetter And bolNumber And bolLowerLetter Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsStrongPassword", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取一个格林威治时间最接近的分钟部分|Get the nearest minute part of Greenwich Mean Time
    ''' </summary>
    ''' <param name="SrcTime"></param>
    ''' <returns></returns>
    Public Function GetCompMinutePart(SrcTime As DateTime) As String
        Try
            Dim dteComp As DateTime = SrcTime.ToUniversalTime()
            If dteComp.Second > 30 Then dteComp.AddMinutes(1)
            GetCompMinutePart = Format(dteComp, "yyyyMMddHHmm")
        Catch ex As Exception
            Me.SetSubErrInf("", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 转义字符串
    ''' </summary>
    ''' <param name="SrcStr">源字符串</param>
    Public Function EscapeStr(SrcStr As String) As String
        Try
            If SrcStr Is Nothing Then
                EscapeStr = ""
            Else
                EscapeStr = SrcStr
            End If
            If InStr(EscapeStr, "&") > 0 Then EscapeStr = Replace(EscapeStr, "&", "&apos;")
            If InStr(EscapeStr, "<") > 0 Then EscapeStr = Replace(EscapeStr, "<", "&lt;")
            If InStr(EscapeStr, ">") > 0 Then EscapeStr = Replace(EscapeStr, ">", "&gt;")
        Catch ex As Exception
            Me.SetSubErrInf("EscapeStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 还原转义字符串
    ''' </summary>
    ''' <param name="EscapeStr">已转义字符串</param>
    Public Function UnEscapeStr(EscapeStr As String) As String
        Try
            If EscapeStr Is Nothing Then
                UnEscapeStr = ""
            Else
                UnEscapeStr = EscapeStr
            End If
            If InStr(UnEscapeStr, "&apos;") > 0 Then UnEscapeStr = Replace(UnEscapeStr, "&apos;", "&")
            If InStr(UnEscapeStr, "&lt;") > 0 Then UnEscapeStr = Replace(UnEscapeStr, "&lt;", "<")
            If InStr(UnEscapeStr, "&gt;") > 0 Then UnEscapeStr = Replace(UnEscapeStr, "&gt;", ">")
        Catch ex As Exception
            Me.SetSubErrInf("UnEscapeStr", ex)
            Return ""
        End Try
    End Function

    Public Function GetTextBase64(SrcText As String, TextType As PigText.enmTextType) As String
        Try
            Dim oPigText As New PigText(SrcText, TextType)
            GetTextBase64 = oPigText.Base64
            oPigText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetTextBase64", ex)
            Return ""
        End Try
    End Function

    Public Function GetTextSHA1(SrcText As String, TextType As PigText.enmTextType) As String
        Try
            Dim oPigText As New PigText(SrcText, TextType)
            Dim oSHA1 As New SHA1CryptoServiceProvider()
            GetTextSHA1 = BitConverter.ToString(oSHA1.ComputeHash(oPigText.TextBytes))
            GetTextSHA1 = Replace(GetTextSHA1, "-", "")
            oPigText = Nothing
        Catch ex As Exception
            Me.SetSubErrInf("GetTextSHA1", ex)
            Return ""
        End Try
    End Function

    Public Overloads Sub ClearErr()
        MyBase.ClearErr()
    End Sub

    ''' <summary>
    ''' Deviation from reference time|是否偏离参考时间
    ''' </summary>
    ''' <param name="ReferenceTime">Reference time|参考时间</param>
    ''' <param name="TimeCount">Time Count|时间计数</param>
    ''' <param name="DateInterval">Date Interval|日期间隔</param>
    ''' <returns></returns>
    Public Function IsDeviationTime(ReferenceTime As Date, TimeCount As Integer, Optional DateInterval As DateInterval = DateInterval.Minute) As Boolean
        Try
            If Math.Abs(DateDiff(DateInterval, ReferenceTime, Now)) > TimeCount Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsDeviationTime", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Timed out or not|是否超时
    ''' </summary>
    ''' <param name="BeginTime">Begin time|开始时间</param>
    ''' <param name="TimeCount">Time Count|时间计数</param>
    ''' <param name="DateInterval">Date Interval|日期间隔</param>
    ''' <returns></returns>
    Public Function IsTimeout(BeginTime As Date, TimeCount As Integer, Optional DateInterval As DateInterval = DateInterval.Minute) As Boolean
        Try
            If DateDiff(DateInterval, BeginTime, Now) > TimeCount Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsTimeout", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Detect if there is a new version|检测是否新版本
    ''' </summary>
    ''' <param name="OldVersion">Older version|旧版本</param>
    ''' <param name="LatestVersion">Latest version|新版本</param>
    ''' <returns></returns>
    Public Function IsNewVersion(OldVersion As String, LatestVersion As String, Optional ByRef Ret As String = "OK") As Boolean
        Try
            If InStr(OldVersion, ".") = 0 Or Right(OldVersion, 1) = "." Then Throw New Exception("Invalid version information(OldVersion)")
            If InStr(LatestVersion, ".") = 0 Or Right(LatestVersion, 1) = "." Then Throw New Exception("Invalid version information(LatestVersion)")
            Dim strOldVersion As String = "", strLatestVersion As String = ""
            OldVersion &= "."
            Do While True
                Dim strPart As String = Me.GetStr(OldVersion, "", ".")
                If strPart = "" Then Exit Do
                strPart = Right("0000000000" & strPart, 10)
                If strOldVersion = "" Then
                    strOldVersion = strPart
                Else
                    strOldVersion &= "." & strPart
                End If
            Loop
            OldVersion &= "."
            Do While True
                Dim strPart As String = Me.GetStr(OldVersion, "", ".")
                If strPart = "" Then Exit Do
                strPart = Right("0000000000" & strPart, 10)
                If strOldVersion = "" Then
                    strOldVersion = strPart
                Else
                    strOldVersion &= "." & strPart
                End If
            Loop
            LatestVersion &= "."
            Do While True
                Dim strPart As String = Me.GetStr(LatestVersion, "", ".")
                If strPart = "" Then Exit Do
                strPart = Right("0000000000" & strPart, 10)
                If strLatestVersion = "" Then
                    strLatestVersion = strPart
                Else
                    strLatestVersion &= "." & strPart
                End If
            Loop
            If strLatestVersion > strOldVersion Then
                Return True
            Else
                Return False
            End If
            Ret = "OK"
        Catch ex As Exception
            Ret = Me.GetSubErrInf("IsNewVersion", ex)
            Return False
        End Try
    End Function

    Public Function ClipboardGetText() As String
        Try
#If NETFRAMEWORK Then
            Return My.Computer.Clipboard.GetText()
#Else
            Throw New Exception("Only NETFRAMEWORK supported")
#End If
        Catch ex As Exception
            Me.SetSubErrInf("ClipboardGetText", ex)
            Return ""
        End Try
    End Function


    Public Function ClipboardSetText(Text As String) As String
        Dim LOG As New PigStepLog("ClipboardSetText")
        Try
#If NETFRAMEWORK Then
            My.Computer.Clipboard.Clear()
            My.Computer.Clipboard.SetText(Text)
#Else
            Throw New Exception("Only NETFRAMEWORK supported")
#End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function OpenUrl(Url As String) As String
        Try
            Process.Start(Url)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("OpenUrl", ex)
        End Try
    End Function

    Public Function GetDefaultBrowser() As String
        Dim LOG As New PigStepLog("GetDefaultBrowser")
        Dim strRegPath As String = "http\shell\open\command"
        Try
            If Me.IsWindows = False Then Throw New Exception("Only supports Windows")
            Dim oPigReg As New PigReg, strRegValue As String
            With oPigReg
                LOG.StepName = ""
                strRegValue = .GetRegStrValue(PigReg.EmnRegRoot.CLASSES_ROOT, strRegPath, "", "")
                If .LastErr <> "" Then Throw New Exception(.LastErr)
                If strRegValue = "" Then Throw New Exception("Unable to obtain registry value")
            End With
            oPigReg = Nothing
            If InStr(strRegValue, "%1") > 0 Then
                strRegValue = Replace(strRegValue, "%1", "")
            End If
            strRegValue = Trim(strRegValue)
            If InStr（strRegValue, """") > 0 Then
                strRegValue = Me.GetStr(strRegValue, """", """")
            End If
            Return strRegValue
        Catch ex As Exception
            LOG.AddStepNameInf("CLASSES_ROOT")
            LOG.AddStepNameInf(strRegPath)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Obtain the start and end times of the time period|获取时间段的开始和结束时间
    ''' </summary>
    ''' <param name="TimeSlot"></param>
    ''' <param name="BeginTime"></param>
    ''' <param name="EndTime"></param>
    ''' <returns></returns>
    Public Function GetTimeSlot(TimeSlot As EnmTimeSlot, ByRef BeginTime As Date, ByRef EndTime As Date) As String
        Dim LOG As New PigStepLog("GetTimeSlot")
        Try
            Dim dteTmp As Date
            LOG.StepName = TimeSlot.ToString
            Select Case TimeSlot
                Case EnmTimeSlot.CurrentWeek
                    Dim intDayOfWeek As Integer = Now.DayOfWeek
                    dteTmp = DateAdd(DateInterval.Day, 0 - intDayOfWeek, Now)
                    BeginTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 10))
                    dteTmp = DateAdd(DateInterval.Day, 7 - intDayOfWeek, Now)
                    EndTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 10))
                Case EnmTimeSlot.CurrentMonth
                    BeginTime = CDate(Left(Me.GetFmtDateTime(Now), 8) & "01")
                    dteTmp = DateAdd(DateInterval.Month, 1, Now)
                    EndTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 8) & "01")
                Case EnmTimeSlot.CurrentYear
                    BeginTime = CDate(Left(Me.GetFmtDateTime(Now), 5) & "01-01")
                    dteTmp = DateAdd(DateInterval.Year, 1, Now)
                    EndTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 5) & "01-01")
                Case EnmTimeSlot.LastWeek
                    Dim intDayOfWeek As Integer = Now.DayOfWeek
                    dteTmp = DateAdd(DateInterval.Day, 0 - intDayOfWeek - 7, Now)
                    BeginTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 10))
                    dteTmp = DateAdd(DateInterval.Day, 0 - intDayOfWeek, Now)
                    EndTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 10))
                Case EnmTimeSlot.LastMonth
                    dteTmp = DateAdd(DateInterval.Month, -1, Now)
                    BeginTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 8) & "01")
                    EndTime = CDate(Left(Me.GetFmtDateTime(Now), 8) & "01")
                Case EnmTimeSlot.LastYear
                    dteTmp = DateAdd(DateInterval.Year, -1, Now)
                    BeginTime = CDate(Left(Me.GetFmtDateTime(dteTmp), 5) & "01-01")
                    EndTime = CDate(Left(Me.GetFmtDateTime(Now), 5) & "01-01")
            End Select
            Return "OK"
        Catch ex As Exception
            BeginTime = Date.MinValue
            EndTime = Date.MinValue
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' Obtain the encoding of the text file|获取文本文件的编码
    ''' </summary>
    ''' <param name="FilePath">File path|文件路径</param>
    ''' <param name="EncCode">Encoding|编码</param>
    ''' <returns></returns>
    Public Function GetTextFileEncCode(FilePath As String, ByRef EncCode As String) As String
        Try
            Dim ecAny As Encoding
            Using srAny As New StreamReader(FilePath)
                srAny.ReadLine()
                ecAny = srAny.CurrentEncoding
            End Using
            EncCode = ecAny.EncodingName
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("GetTextFileEncCode", ex)
        End Try
    End Function

    ''' <summary>
    ''' Obtain the absolute path of the relative path of the basic path|获取基本路径的相对路径的绝对路径
    ''' </summary>
    ''' <param name="BasePath">Basic path|基本路径</param>
    ''' <param name="RelativePath">Relative path relative to the basic path|相对于基本路径的相对路径</param>
    ''' <returns></returns>
    Public Function CombinePath(BasePath As String, RelativePath As String) As String
        Try
            Dim strBasePath As String = System.IO.Path.GetFullPath(BasePath)
            Dim strCombinePath As String = System.IO.Path.Combine(strBasePath, RelativePath)
            Return System.IO.Path.GetFullPath(strCombinePath)
        Catch ex As Exception
            Me.SetSubErrInf("CombinePath", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Find the specified beginning and ending content in the source string|在源字符串中查找指定开头和结尾的内容
    ''' </summary>
    ''' <param name="SrcStr">Source string|源字符串</param>
    ''' <param name="BeginStr">Starting string|开头的字符串</param>
    ''' <param name="EndStr">A null-terminated string|结尾的字符串</param>
    ''' <param name="FindTimes">Search times|查找次数</param>
    ''' <returns></returns>
    Public Function GetStr(SrcStr As String, BeginStr As String, EndStr As String, FindTimes As Integer) As String
        Try
            Dim intBegin As Integer = 0, intEnd As Integer = 0, intCnt As Integer = 0, intBeginStrLen As Integer = Len(BeginStr), intEndStrLen As Integer = Len(EndStr)
            Dim intLastPos As Integer = 0
            If FindTimes < 1 Then FindTimes = 1

            For i As Integer = 1 To FindTimes
                If intBeginStrLen = 0 Then
                    intBegin = 0
                Else
                    intBegin = SrcStr.IndexOf(BeginStr, intLastPos)
                    If intBegin = -1 Then
                        Return ""
                    End If
                End If
                If intEndStrLen = 0 Then
                    intEnd = SrcStr.Length
                Else
                    intEnd = SrcStr.IndexOf(EndStr, intBegin + intBeginStrLen)
                    If intEnd = -1 Then
                        Return ""
                    End If
                End If
                intLastPos = intEnd + intEndStrLen
                intCnt += 1
            Next

            If intCnt = FindTimes Then
                Return SrcStr.Substring(intBegin + BeginStr.Length, intEnd - intBegin - BeginStr.Length)
            Else
                Return ""
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GetStr", ex)
            Return ""
        End Try
    End Function
    'Public Function GetStr(SrcStr As String, BeginStr As String, EndStr As String, FindTimes As Integer) As String
    '    Try
    '        Dim intBegin As Integer = 0, intEnd As Integer = 0, intCnt As Integer = 0, intBeginStrLen As Integer = Len(BeginStr), intEndStrLen As Integer = Len(EndStr)
    '        Dim intLastBegin As Integer = 0, intLastEnd As Integer = intLastBegin + intBeginStrLen, bolIsBeginEndSsame As Boolean
    '        If BeginStr = EndStr Then
    '            bolIsBeginEndSsame = True
    '            '                   intLastEnd = intLastBegin + intBeginStrLen + 1
    '            '                   intLastEnd = intLastBegin + intBeginStrLen
    '        Else
    '            bolIsBeginEndSsame = False
    '        End If
    '        If FindTimes < 1 Then FindTimes = 1

    '        For i As Integer = 1 To FindTimes
    '            If intBeginStrLen = 0 Then
    '                intBegin = 0
    '            Else
    '                intBegin = SrcStr.IndexOf(BeginStr, intLastBegin)
    '                If intBegin = -1 Then
    '                    Return ""
    '                End If
    '                intLastBegin = intBegin + intBeginStrLen
    '                If bolIsBeginEndSsame = True Then
    '                    intLastBegin = SrcStr.IndexOf(BeginStr, intLastBegin) + intBeginStrLen
    '                End If
    '            End If
    '            If intEndStrLen = 0 Then
    '                intEnd = SrcStr.Length
    '            Else
    '                intEnd = SrcStr.IndexOf(EndStr, intLastEnd)
    '                If intEnd = -1 Then
    '                    Return ""
    '                End If
    '                intLastEnd = intEnd + intEndStrLen
    '                If bolIsBeginEndSsame = True Then
    '                    intLastEnd = SrcStr.IndexOf(EndStr, intLastEnd) + intEndStrLen
    '                End If
    '            End If
    '            intCnt += 1
    '        Next

    '        If intCnt = FindTimes Then
    '            Return SrcStr.Substring(intBegin + BeginStr.Length, intEnd - intBegin - BeginStr.Length)
    '        Else
    '            Return ""
    '        End If
    '    Catch ex As Exception
    '        Me.SetSubErrInf("GetStr", ex)
    '        Return ""
    '    End Try
    'End Function
    ''' <summary>
    ''' Find the specified beginning and ending content in the source string and replace it with a new string|在源字符串中查找指定开头和结尾的内容并用新的字符串替换
    ''' </summary>
    ''' <param name="SrcStr">Source string|源字符串</param>
    ''' <param name="BeginStr">Starting string|开头的字符串</param>
    ''' <param name="EndStr">A null-terminated string|结尾的字符串</param>
    ''' <param name="FindTimes">Search times|查找次数</param>
    ''' <param name="ReplaceStr">The new string to replace|要替换的新字符串</param>
    ''' <returns></returns>
    Public Function GetStrAndReplace(SrcStr As String, BeginStr As String, EndStr As String, FindTimes As Integer, ReplaceStr As String) As String
        Try
            Dim intBegin As Integer = 0, intEnd As Integer = 0, intCnt As Integer = 0, intBeginStrLen As Integer = Len(BeginStr), intEndStrLen As Integer = Len(EndStr)
            Dim intLastPos As Integer = 0
            If FindTimes < 1 Then FindTimes = 1

            For i As Integer = 1 To FindTimes
                If intBeginStrLen = 0 Then
                    intBegin = 0
                Else
                    intBegin = SrcStr.IndexOf(BeginStr, intLastPos)
                    If intBegin = -1 Then
                        Return ""
                    End If
                End If
                If intEndStrLen = 0 Then
                    intEnd = SrcStr.Length
                Else
                    intEnd = SrcStr.IndexOf(EndStr, intBegin + intBeginStrLen)
                    If intEnd = -1 Then
                        Return ""
                    End If
                End If
                intLastPos = intEnd + intEndStrLen
                intCnt += 1
            Next

            If intCnt = FindTimes Then
                Return BeginStr & ReplaceStr & EndStr
            Else
                Return ""
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GetStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Is it an absolute path|是否绝对路径
    ''' </summary>
    ''' <param name="InPath">Input path|输入的路径</param>
    ''' <returns></returns>
    Public Function IsAbsolutePath(InPath As String) As Boolean
        Try
            Dim strInPath As String = System.IO.Path.GetFullPath(InPath)
            If strInPath = InPath Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.SetSubErrInf("IsAbsolutePath", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 按 ASCII 习惯获取 Uniocode 字符串的长度|Obtain the length of a Uniocode string according to ASCII conventions
    ''' </summary>
    ''' <param name="SrcStr">源字符串|Source string</param>
    ''' <returns></returns>
    ''' </summary>
    Public Function LenA(SrcStr As String) As Integer
        Try
            Dim abSrc As Byte() = Encoding.Unicode.GetBytes(SrcStr)
            Dim lngLen As Integer = 0
            For i = 0 To abSrc.Length - 1 Step 2
                If abSrc(i + 1) = 0 Then
                    lngLen += 1
                Else
                    lngLen += 2
                End If
            Next
            Return lngLen
        Catch ex As Exception
            Me.SetSubErrInf("LenA", ex)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' 按 ASCII 习惯获取 Uniocode 字符串的左边部分|Obtain the left part of the Uniocode string according to ASCII conventions
    ''' </summary>
    ''' <param name="SrcStr">源字符串|Source string</param>
    ''' <param name="LeftLen">左边部分的长度|The length of the left section</param>
    ''' <returns></returns>
    Public Function LeftA(SrcStr As String, LeftLen As Integer) As String
        Try
            If SrcStr = "" Then
                Return ""
            ElseIf LeftLen <= 0 Then
                Return ""
            Else
                Dim abSrc As Byte() = Encoding.Unicode.GetBytes(SrcStr)
                Dim abTar(-1) As Byte
                For i = 0 To abSrc.Length - 1 Step 2
                    If LeftLen <= 0 Then Exit For
                    Dim intPos As Integer = abTar.Length
                    ReDim Preserve abTar(intPos + 1)
                    abTar(intPos) = abSrc(i)
                    abTar(intPos + 1) = abSrc(i + 1)
                    If abSrc(i + 1) = 0 Then
                        LeftLen -= 1
                    Else
                        LeftLen -= 2
                    End If
                Next
                Return Encoding.Unicode.GetString(abTar)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("LeftA", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 按 ASCII 习惯获取 Uniocode 字符串的右边部分|Get the right part of the Uniocode string
    ''' </summary>
    ''' <param name="SrcStr">源字符串|Source string</param>
    ''' <param name="RightLen">右边部分的长度|The length of the right section</param>
    ''' <returns></returns>

    Public Function RightA(SrcStr As String, RightLen As Integer) As String
        Try
            If SrcStr = "" Then
                Return ""
            ElseIf RightLen <= 0 Then
                Return ""
            Else
                Dim abSrc As Byte() = Encoding.Unicode.GetBytes(SrcStr)
                Dim abTar(-1) As Byte
                Dim intPos As Integer = 0
                For i = abSrc.Length - 1 To 0 Step -2
                    If RightLen <= 0 Then Exit For
                    intPos = abTar.Length
                    ReDim Preserve abTar(intPos + 1)
                    abTar(intPos) = abSrc(i - 1)
                    abTar(intPos + 1) = abSrc(i)
                    If abSrc(i) = 0 Then
                        RightLen -= 1
                    Else
                        RightLen -= 2
                    End If
                Next
                Dim abTar2(abTar.Length - 1) As Byte
                intPos = 0
                For i = abTar.Length - 1 To 0 Step -2
                    abTar2(intPos) = abTar(i - 1)
                    abTar2(intPos + 1) = abTar(i)
                    intPos += 2
                Next
                Return Encoding.Unicode.GetString(abTar2)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("RightA", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 按 ASCII 习惯获取 Uniocode 字符串的中间部分|Get the middle part of the Uniocode string
    ''' </summary>
    ''' <param name="SrcStr">源字符串|Source string</param>
    ''' <param name="Start">开始位置|Start position</param>
    ''' <param name="MidLen">中间部分的长度|The length of the middle part</param>
    ''' <returns></returns>
    Public Function MidA(SrcStr As String, Start As Integer, Optional MidLen As Integer = 0) As String
        Try
            If MidLen <= 0 Then MidLen = Me.LenA(SrcStr)
            If SrcStr = "" Then
                Return ""
            Else
                Dim abSrc As Byte() = Encoding.Unicode.GetBytes(SrcStr)
                Dim abTar(-1) As Byte
                Dim intPos As Integer = 0
                For i = 0 To abSrc.Length - 1 Step 2
                    If Start <= 0 Then
                        intPos = abTar.Length
                        ReDim Preserve abTar(intPos + 1)
                        abTar(intPos) = abSrc(i)
                        abTar(intPos + 1) = abSrc(i + 1)
                        If abSrc(i + 1) = 0 Then
                            MidLen -= 1
                        Else
                            MidLen -= 2
                        End If
                        If MidLen <= 0 Then Exit For
                    Else
                        Start -= 1
                    End If
                Next
                Return Encoding.Unicode.GetString(abTar)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("MidA", ex)
            Return ""
        End Try
    End Function

    Public Function GetDateUnixTimestamp(SrcTime As Date) As Long
        Return DateDiff(DateInterval.Second, #1/1/1970#, SrcTime)
    End Function

    Public Function GetNowUnixTimestamp() As String
        Return DateDiff(DateInterval.Second, #1/1/1970#, Now)
    End Function

    Public Function DistinctString(InStrings As String(), ByRef OutStrings As String()) As String
        Try
            Dim oList As New List(Of String)
            For Each strItem As String In InStrings
                If oList.Contains(strItem) = False Then
                    oList.Add(strItem)
                End If
            Next
            OutStrings = oList.ToArray
            Return "OK"
        Catch ex As Exception
            ReDim OutStrings(-1)
            Return Me.GetSubErrInf("DistinctStr", ex)
        End Try
    End Function

    ''' <summary>
    ''' Replace multiple spaces in a string with two spaces|把字符串中多个空格替换为两个空格
    ''' </summary>
    ''' <param name="SrcStr">Source string|源字符串</param>
    ''' <param name="OutStr">Converted string|转换后的字符串</param>
    ''' <param name="IsTrimConvert">Do you want to convert double spaces to '<>'|是否把双空格转换成"<>"</param>
    ''' <returns></returns>
    ''' </summary>
    Public Function StrSpaceMulti2Double(SrcStr As String, ByRef OutStr As String, Optional IsTrimConvert As Boolean = True) As String
        Try
            Const SPACE_2 As String = "  "
            Const SPACE_3 As String = "   "
            OutStr = SrcStr
            Do While InStr(OutStr, SPACE_3) > 0
                OutStr = Replace(OutStr, SPACE_3, SPACE_2)
            Loop
            If IsTrimConvert = True Then
                OutStr = Trim(OutStr)
                OutStr = "<" & Replace(OutStr, SPACE_2, "><") & ">"
            End If
            Return "OK"
        Catch ex As Exception
            OutStr = ""
            Return Me.GetSubErrInf("StrSpaceMulti2Double", ex)
        End Try
    End Function

    ''' <summary>
    ''' Replace multiple spaces in a string with one space|把字符串中多个空格替换为一个空格
    ''' </summary>
    ''' <param name="SrcStr">Source string|源字符串</param>
    ''' <param name="OutStr">Converted string|转换后的字符串</param>
    ''' <param name="IsTrimConvert">Do you want to convert space to '<>'|是否把空格转换成"<>"</param>
    ''' <returns></returns>
    ''' </summary>
    Public Function StrSpaceMulti2One(SrcStr As String, ByRef OutStr As String, Optional IsTrimConvert As Boolean = True) As String
        Try
            Const SPACE_2 As String = "  "
            Const SPACE_1 As String = " "
            OutStr = SrcStr
            Do While InStr(OutStr, SPACE_2) > 0
                OutStr = Replace(OutStr, SPACE_2, SPACE_1)
            Loop
            If IsTrimConvert = True Then
                OutStr = Trim(OutStr)
                OutStr = "<" & Replace(OutStr, SPACE_1, "><") & ">"
            End If
            Return "OK"
        Catch ex As Exception
            OutStr = ""
            Return Me.GetSubErrInf("StrSpaceMulti2One", ex)
        End Try
    End Function

End Class
