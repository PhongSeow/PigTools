'**********************************
'* Name: PigFuncLite
'* Author: Seow Phong
'* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 经量的PigFunc，代码来自 PigFunc 的副本|A small part of PigFunc, the code comes from a copy of PigFunc
'* Home Url: https://en.seowphong.com
'* Version: 1.70
'* Create Time: 2/2/2021
'**********************************
Imports System.IO
Imports System.Text
Imports System.Environment
Imports System.Net
Imports System.Threading

Public Class PigFuncLite
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "70" & "." & "2"

    Public Event ASyncRet_SaveTextToFile(SyncRet As StruASyncRet)

    ''' <summary>文件的部分</summary>
    Public Enum EnmFilePart
        Path = 1         '路径
        FileTitle = 2    '文件名
        ExtName = 3      '扩展名
        DriveNo = 4      '驱动器名
    End Enum


    ''' <summary>获取随机字符串的方式</summary>
    Public Enum EnmGetRandString
        NumberOnly = 1      '只有数字
        NumberAndLetter = 2 '只有数字和字母(包括大小写)
        DisplayChar = 3     '全部可显示字符(ASCII 33-126)
        AllAsciiChar = 4    '全部ASCII码(返回结果以16进制方式显示)
    End Enum

    ''' <summary>对齐方式|Alignment</summary>
    Public Enum EnmAlignment
        Left = 1
        Right = 2
        Center = 3
    End Enum


    Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

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

    ''' <remarks>显示精确到毫秒的时间</remarks>
    Public Function GENow() As String
        Return Format(Now, "yyyy-MM-dd HH:mm:ss.fff")
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

    Public Function IsFileExists(FilePath As String) As Boolean
        Try
            Return IO.File.Exists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("IsFileExists", ex)
            Return Nothing
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

    Public Function CreateFolder(FolderPath As String) As String
        Try
            Directory.CreateDirectory(FolderPath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CreateFolder", ex)
        End Try
    End Function

    Public Function GetEnvVar(EnvVarName As String) As String
        Return GetEnvironmentVariable(EnvVarName)
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

    Public Function GetComputerName() As String
        Return Dns.GetHostName()
    End Function

    Public Function GEDec(vData As String) As Decimal
        Try
            Return CDec(vData)
        Catch ex As Exception
            Me.SetSubErrInf("CDec", ex)
            Return 0
        End Try
    End Function

    Public Function GEInt(vData As String) As Integer
        Try
            Return CInt(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
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

    ''' <remarks>同步等待</remarks>
    Public Sub Delay(DelayMs As Single)
        System.Threading.Thread.Sleep(DelayMs)
    End Sub

    Public Function GECLng(vData As String) As Long
        Try
            Return CLng(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
        End Try
    End Function

    Public Function GetHostIp() As String
        Return mGetHostIp(False)
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

    Public Function GetFmtDateTime(SrcTime As DateTime, Optional TimeFmt As String = "yyyy-MM-dd HH:mm:ss.fff") As String
        Return Format(SrcTime, TimeFmt)
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

    Public Function GetHostName() As String
        Return Dns.GetHostName()
    End Function

    Public Function ASyncSaveTextToFile(FilePath As String, SaveText As String, ByRef OutThreadID As Integer) As String
        Return Me.mASyncSaveTextToFile(FilePath, SaveText, OutThreadID)
    End Function

    Private Structure mStruSaveTextToFile
        Public FilePath As String
        Public SaveText As String
        Public IsAsync As Boolean
    End Structure

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

    Public Function GetUserName() As String
        Return Environment.UserName
    End Function

    Public Function SaveTextToFile(FilePath As String, SaveText As String) As String
        Dim struMain As mStruSaveTextToFile
        With struMain
            .FilePath = FilePath
            .SaveText = SaveText
        End With
        Return Me.mSaveTextToFile(struMain)
    End Function

    Public Function DeleteFile(FilePath As String) As String
        Try
            File.Delete(FilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFile", ex)
        End Try
    End Function

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


End Class
