''' <summary>
'''************************************
'''* Name: PigMLang
'''* Author: Seow Phong
'''* License: Copyright (c) 2020-2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'''* Describe: A lightweight multi language processing class, As long as you refer to this class, you can implement multilingual processing|一个轻量的多语言处理类，只要引用本类就可以实现多语言处理。 
'''* Home Url: https://en.seowphong.com
'''* Version: 1.7
'''* Create Time: 30/11/2020
'''* 1.0.2  1/12/2020   Modify GetAllLangInf, Add GetMLangText
'''* 1.0.3  1/12/2020   Modify mInitCultureSortList
'''* 1.0.4  7/12/2020   Modify mLoadMLangInf, add mIsWindows,MkMLangText
'''* 1.0.5  8/12/2020   Modify MkMLangText
'''* 1.0.6  12/12/2020  Modify mGetCultureInfo,mGetMLangText,mMLangFilePath,CurrLCID
'''* 1.0.7  13/12/2020  Modify SetCurrCulture,mSrc2CtlStr,mEscapeStr, Del mIsWindows, Add CanUseCultureList
'''* 1.0.8  14/12/2020  Remove Inherits PigBaseMini, Add mIsWindows,mOsCrLf,mOsPathSep,mEscapeStr,mUnEscapeStr,mConvStrWin2Linux
'''* 1.0.9  19/12/2020  Modify mLoadMLangInf
'''* 1.0.10 20/12/2020  Modify GetAllLangInf,mGetCultureInfoMD,mSetSubErrInf rename SetSubErrInf
'''* 1.0.11 5/1/2021    Modify mNew, auto copy lang file.
'''* 1.1 11/9/2022      Modify mNew, auto copy lang file.
'''* 1.2 17/10/2022     Add PigFunc and Rewrite Common Code
'''* 1.3 18/10/2022     Add GetCanUseCultureXml,SetCurrCulture,mInitCultureSortList
'''* 1.5 19/10/2022     Add MLangTextCnt
'''* 1.6 31/10/2022     Modify mGetMLangText
'''* 1.7 18/1/2022     Modify MkMLangText
'''************************************
''' </summary>
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.IO

''' <summary>
''' Multilingual processing class|多语言处理类
''' </summary>
Public Class PigMLang
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.7.2"

    Private ReadOnly Property mPigFunc As New PigFunc
    Private ReadOnly Property mFS As New mFileSystemObject

    ''' <summary>
    ''' 获取数据格式
    ''' </summary>
    Public Enum EnmGetInfFmt
        ''' <summary>
        ''' TAB键分隔|The separator is TAB key
        ''' </summary>
        TabSeparator = 0
        ''' <summary>
        ''' Markdown
        ''' </summary>
        Markdown = 10
    End Enum

    ''' <summary>
    ''' 区域集合
    ''' </summary>
    Private mslCultureInfo As New SortedList

    ''' <summary>
    ''' 当前可用的区域数组
    ''' </summary>
    Private mciaCanUseCultureList As CultureInfo()
    Public ReadOnly Property CanUseCultureList() As CultureInfo()
        Get
            Return mciaCanUseCultureList
        End Get
    End Property

    ''' <summary>
    ''' 多语言文本集合
    ''' </summary>
    Private mslMLangText As New SortedList
    ''' <summary>
    ''' 当前LCID
    ''' </summary>
    Private mciCurrent As CultureInfo
    Public ReadOnly Property CurrLCID() As Integer
        Get
            Return mciCurrent.LCID
        End Get
    End Property

    Public ReadOnly Property MLangTextCnt() As Integer
        Get
            If mslMLangText Is Nothing Then
                Return 0
            Else
                Return mslMLangText.Count
            End If
        End Get
    End Property

    ''' <summary>
    ''' 当前PigMLang标题
    ''' </summary>
    Private mstrCurrMLangTitle As String
    Public ReadOnly Property CurrMLangTitle() As String
        Get
            Return mstrCurrMLangTitle
        End Get
    End Property

    ''' <summary>
    ''' 当前PigMLang目录
    ''' </summary>
    Private mstrCurrMLangDir As String
    Public ReadOnly Property CurrMLangDir() As String
        Get
            Return mstrCurrMLangDir
        End Get
    End Property


    ''' <summary>
    ''' 当前PigMLang文件
    ''' </summary>
    Private mstrCurrMLangFile As String
    Public ReadOnly Property CurrMLangFile() As String
        Get
            Return mstrCurrMLangFile
        End Get
    End Property


    ''' <summary>
    ''' 当前区域名称
    ''' </summary>
    Public ReadOnly Property CurrCultureName() As String
        Get
            Return mciCurrent.Name
        End Get
    End Property

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.mNew("", "")
        Me.SetCurrCulture()
    End Sub

    Public Sub New(MLangTitle As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(MLangTitle, "")
        Me.SetCurrCulture()
    End Sub

    Public Sub New(MLangTitle As String, LCID As Integer)
        MyBase.New(CLS_VERSION)
        Me.mNew(MLangTitle, "")
        Me.SetCurrCulture(LCID)
    End Sub

    Public Sub New(MLangTitle As String, MLangDir As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(MLangTitle, MLangDir)
        Me.SetCurrCulture()
    End Sub

    Public Sub New(MLangTitle As String, MLangDir As String, LCID As Integer)
        MyBase.New(CLS_VERSION)
        Me.mNew(MLangTitle, MLangDir)
        Me.SetCurrCulture(LCID)
    End Sub

    Public Sub New(MLangTitle As String, MLangDir As String, CultureName As String)
        MyBase.New(CLS_VERSION)
        Me.mNew(MLangTitle, MLangDir)
        Me.SetCurrCulture(CultureName)
    End Sub

    ''' <summary>
    ''' 设置当前区域
    ''' </summary>
    ''' <param name="LCID">LCID</param>
    Public Function SetCurrCulture(LCID As Integer) As String
        Try
            Dim oCultureInfo As CultureInfo = Nothing
            Dim strRet As String = Me.mNewCultureInfo("", oCultureInfo, LCID)
            If strRet <> "OK" Then Throw New Exception(strRet)
            mciCurrent = oCultureInfo
            Return "OK"
        Catch ex As Exception
            mciCurrent = New CultureInfo("en-US")
            Return Me.GetSubErrInf("SetCurrCulture", ex)
        End Try
    End Function

    Private Function mMLangFilePath(LCID As Integer) As String
        Return mstrCurrMLangDir & Me.OsPathSep & mstrCurrMLangTitle & "." & LCID.ToString
    End Function

    Private Function mMLangFilePath(CultureName As String) As String
        Return mstrCurrMLangDir & Me.OsPathSep & mstrCurrMLangTitle & "." & CultureName
    End Function


    Private Function mGetLCIDByCultureName(CultureName As String) As Integer
        Try
            Dim oCultureInfo As CultureInfo, strRet As String = ""
            oCultureInfo = Me.mGetCultureInfo(CultureName)
            If oCultureInfo Is Nothing Then
                strRet = Me.mNewCultureInfo(CultureName, oCultureInfo)
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            If oCultureInfo Is Nothing Then
                Return 0
            Else
                mGetLCIDByCultureName = oCultureInfo.LCID
                oCultureInfo = Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetLCIDByCultureName", ex)
            Return 0
        End Try
    End Function

    Private Function mGetCultureNameByLCID(LCID As Integer) As String
        Try
            Dim oCultureInfo As CultureInfo, strRet As String = ""
            oCultureInfo = Me.mGetCultureInfo(LCID)
            If oCultureInfo Is Nothing Then
                strRet = Me.mNewCultureInfo("", oCultureInfo, LCID)
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            If oCultureInfo Is Nothing Then
                mGetCultureNameByLCID = ""
            Else
                mGetCultureNameByLCID = oCultureInfo.Name
                oCultureInfo = Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetCultureNameByLCID", ex)
            Return ""
        End Try
    End Function

    Public Function IsMLangExists(LCID As Integer) As Boolean
        Dim strMLangFile As String = Me.mGetMLangFile(LCID, False)
        If strMLangFile = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function IsMLangExists(LCID As Integer, IsFindAlike As Boolean) As Boolean
        Dim strMLangFile As String = Me.mGetMLangFile(LCID, IsFindAlike)
        If strMLangFile = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function IsMLangExists(LCID As Integer, IsFindAlike As Boolean, IsUseEnglish As Boolean) As Boolean
        Dim strMLangFile As String = Me.mGetMLangFile(LCID, IsFindAlike)
        If strMLangFile = "" Then
            If IsUseEnglish = True Then
                strMLangFile = Me.mGetMLangFile(1033, IsFindAlike)  'en-US
                If strMLangFile = "" Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Private Function mGetMLangFile(LCID As Integer, Optional IsFindAlike As Boolean = True) As String
        Try
            mGetMLangFile = Me.mMLangFilePath(LCID)
            If File.Exists(mGetMLangFile) = False Then
                Dim strCultureName As String = Me.mGetCultureNameByLCID(LCID)
                If strCultureName = "" Then
                    mGetMLangFile = ""
                Else
                    mGetMLangFile = Me.mMLangFilePath(strCultureName)
                    If File.Exists(mGetMLangFile) = False Then
                        If IsFindAlike = True Then
                            Me.mInitCultureSortList()
                            Dim strLangName As String = Me.mPigFunc.GetStr(strCultureName, "", "-", False) & "-"
                            Dim intLangLen As Integer = Len(strLangName)
                            Dim oCultureInfo As CultureInfo
                            Dim bolIsFind As Boolean = False
                            For Each obj In mslCultureInfo
                                oCultureInfo = obj.value
                                If Left(oCultureInfo.Name, intLangLen) = strLangName Then
                                    mGetMLangFile = Me.mMLangFilePath(oCultureInfo.LCID)
                                    If File.Exists(mGetMLangFile) = True Then
                                        bolIsFind = True
                                        Exit For
                                    End If
                                    mGetMLangFile = Me.mMLangFilePath(oCultureInfo.Name)
                                    If File.Exists(mGetMLangFile) = True Then
                                        bolIsFind = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If bolIsFind = False Then mGetMLangFile = ""
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetMLangFile", ex)
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' 设置当前区域
    ''' </summary>
    ''' <param name="CultureName">区域名称</param>
    Public Function SetCurrCulture(CultureName As String) As String
        Dim LOG As New PigStepLog("SetCurrCulture")
        Try
            Dim oCultureInfo As CultureInfo = Nothing
            Dim strRet As String = Me.mNewCultureInfo(CultureName, oCultureInfo)
            If strRet <> "OK" Then Throw New Exception(strRet)
            mciCurrent = oCultureInfo
            Return "OK"
        Catch ex As Exception
            mciCurrent = New CultureInfo("en-US")
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 设置当前区域
    ''' </summary>
    Public Function SetCurrCulture() As String
        Try
            mciCurrent = CultureInfo.CurrentCulture
            If mciCurrent.LCID = 127 Then
                Throw New Exception("Unknow Culture (LCID=127),Switch to en-US")
            End If
            Return "OK"
        Catch ex As Exception
            mciCurrent = New CultureInfo("en-US")
            Return Me.GetSubErrInf("SetCurrCulture", ex)
        End Try
    End Function

    Private Sub mMkDir(DirPath As String)
        Try
            Directory.CreateDirectory(DirPath)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mMkDir", ex)
        End Try
    End Sub

    Private Sub mNew(MLangTitle As String, MLangDir As String)
        Try
            If MLangTitle = "" Then MLangTitle = Me.AppTitle
            mstrCurrMLangTitle = MLangTitle
            Dim bolIsFindDir As Boolean = False
            If MLangDir = "" Then
                bolIsFindDir = True
            ElseIf Directory.Exists(MLangDir) = False Then
                bolIsFindDir = True
            End If

            If bolIsFindDir = True Then
                MLangDir = Me.AppPath
                If Right(MLangDir, 1) <> Me.OsPathSep Then MLangDir &= Me.OsPathSep
                MLangDir = Me.AppPath & "Lang"
                If Directory.Exists(MLangDir) = False Then
                    Me.mMkDir(MLangDir)
                    If Directory.Exists(MLangDir) = False Then
                        MLangDir = Me.AppPath
                    End If
                    Dim strUpLangDir As String, strUpDir As String
                    If Me.IsWindows = True Then
                        strUpDir = "..\..\..\"
                    Else
                        strUpDir = "../../../"
                    End If
                    strUpLangDir = strUpDir & "Lang"
                    If Directory.Exists(strUpLangDir) = True Then
                        Dim strSrcDir As String = Directory.GetParent(strUpDir).FullName & Me.OsPathSep & "Lang"
                        Dim diSrc As New DirectoryInfo(strSrcDir)
                        For Each oFile In diSrc.GetFiles
                            If InStr(oFile.Extension, "-") > 0 Then
                                Dim strTarFile As String = MLangDir & Me.OsPathSep & oFile.Name
                                If File.Exists(strTarFile) = False Then
                                    File.Copy(oFile.FullName, strTarFile)
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            mstrCurrMLangDir = MLangDir
            Me.mInitCultureSortList()
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub

    Private Function mGetCultureInfoMD(IsColName As Boolean, Optional LCID As Integer = -1) As String
        Dim strStepName As String = ""
        Try
            mGetCultureInfoMD = ""
            If IsColName = True Then
                mGetCultureInfoMD &= "|" & "Name"
                mGetCultureInfoMD &= "|" & "LCID"
                mGetCultureInfoMD &= "|" & "EnglishName"
                mGetCultureInfoMD &= "|" & "DisplayName"
                mGetCultureInfoMD &= "|" & "NativeName"
                mGetCultureInfoMD &= "|" & Me.OsCrLf & "| ---- | ---- | ---- | ---- | ---- |"
            Else
                strStepName = "New CultureInfo(" & LCID & ")"
                Dim oCultureInfo As CultureInfo = New CultureInfo(LCID)
                With oCultureInfo
                    mGetCultureInfoMD &= "|" & .Name
                    mGetCultureInfoMD &= "|" & .LCID
                    mGetCultureInfoMD &= "|" & .EnglishName
                    mGetCultureInfoMD &= "|" & .DisplayName
                    mGetCultureInfoMD &= "|" & .NativeName
                    mGetCultureInfoMD &= "|"
                End With
            End If
            Return mGetCultureInfoMD
        Catch ex As Exception
            Return ""
            Me.SetSubErrInf("mGetCultureInfoMD", ex)
        End Try
    End Function

    Private Function mGetCultureInfoTab(IsColName As Boolean, Optional LCID As Integer = -1) As String
        Dim strStepName As String = ""
        Try
            mGetCultureInfoTab = ""
            If IsColName = True Then
                mGetCultureInfoTab &= "Name"
                mGetCultureInfoTab &= vbTab & "LCID"
                mGetCultureInfoTab &= vbTab & "EnglishName"
                mGetCultureInfoTab &= vbTab & "DisplayName"
                mGetCultureInfoTab &= vbTab & "NativeName"
            Else
                strStepName = "New CultureInfo(" & LCID & ")"
                Dim oCultureInfo As CultureInfo = New CultureInfo(LCID)
                With oCultureInfo
                    mGetCultureInfoTab &= .Name
                    mGetCultureInfoTab &= vbTab & .LCID
                    mGetCultureInfoTab &= vbTab & .EnglishName
                    mGetCultureInfoTab &= vbTab & .DisplayName
                    mGetCultureInfoTab &= vbTab & .NativeName
                End With
            End If
            Return mGetCultureInfoTab
        Catch ex As Exception
            Return ""
            Me.SetSubErrInf("mGetCultureInfoTab", ex)
        End Try
    End Function

    Private Sub mInitCultureSortList()
        Dim strStepName As String = ""
        Try
            Dim bolIsInit As Boolean = False
            If mslCultureInfo Is Nothing Then
                bolIsInit = True
            ElseIf mslCultureInfo.Count = 0 Then
                bolIsInit = True
            End If

            If bolIsInit = True Then
                strStepName = "New SortedList And Add"
                Dim mslLCID As New SortedList
                With mslLCID
                    .Add(1078, 1078)
                    .Add(1052, 1052)
                    .Add(1118, 1118)
                    .Add(5121, 5121)
                    .Add(15361, 15361)
                    .Add(3073, 3073)
                    .Add(2049, 2049)
                    .Add(11265, 11265)
                    .Add(13313, 13313)
                    .Add(12289, 12289)
                    .Add(4097, 4097)
                    .Add(6145, 6145)
                    .Add(8193, 8193)
                    .Add(16385, 16385)
                    .Add(1025, 1025)
                    .Add(10241, 10241)
                    .Add(7169, 7169)
                    .Add(14337, 14337)
                    .Add(9217, 9217)
                    .Add(1067, 1067)
                    .Add(1101, 1101)
                    .Add(2092, 2092)
                    .Add(1068, 1068)
                    .Add(1069, 1069)
                    .Add(1059, 1059)
                    .Add(2117, 2117)
                    .Add(1093, 1093)
                    .Add(5146, 5146)
                    .Add(1026, 1026)
                    .Add(1109, 1109)
                    .Add(1027, 1027)
                    .Add(2052, 2052)
                    .Add(3076, 3076)
                    .Add(5124, 5124)
                    .Add(4100, 4100)
                    .Add(1028, 1028)
                    .Add(1050, 1050)
                    .Add(1029, 1029)
                    .Add(1030, 1030)
                    .Add(2067, 2067)
                    .Add(1043, 1043)
                    .Add(1126, 1126)
                    .Add(3081, 3081)
                    .Add(10249, 10249)
                    .Add(4105, 4105)
                    .Add(9225, 9225)
                    .Add(2057, 2057)
                    .Add(16393, 16393)
                    .Add(6153, 6153)
                    .Add(8201, 8201)
                    .Add(5129, 5129)
                    .Add(13321, 13321)
                    .Add(7177, 7177)
                    .Add(11273, 11273)
                    .Add(1033, 1033)
                    .Add(12297, 12297)
                    .Add(1061, 1061)
                    .Add(1071, 1071)
                    .Add(1080, 1080)
                    .Add(1065, 1065)
                    .Add(1124, 1124)
                    .Add(1035, 1035)
                    .Add(2060, 2060)
                    .Add(11276, 11276)
                    .Add(3084, 3084)
                    .Add(9228, 9228)
                    .Add(12300, 12300)
                    .Add(1036, 1036)
                    .Add(5132, 5132)
                    .Add(13324, 13324)
                    .Add(6156, 6156)
                    .Add(14348, 14348)
                    .Add(10252, 10252)
                    .Add(4108, 4108)
                    .Add(7180, 7180)
                    .Add(1122, 1122)
                    .Add(2108, 2108)
                    .Add(1110, 1110)
                    .Add(1079, 1079)
                    .Add(3079, 3079)
                    .Add(1031, 1031)
                    .Add(5127, 5127)
                    .Add(4103, 4103)
                    .Add(2055, 2055)
                    .Add(1032, 1032)
                    .Add(1140, 1140)
                    .Add(1095, 1095)
                    .Add(1037, 1037)
                    .Add(1081, 1081)
                    .Add(1038, 1038)
                    .Add(1039, 1039)
                    .Add(1136, 1136)
                    .Add(1057, 1057)
                    .Add(1040, 1040)
                    .Add(2064, 2064)
                    .Add(1041, 1041)
                    .Add(1099, 1099)
                    .Add(1120, 1120)
                    .Add(1087, 1087)
                    .Add(1107, 1107)
                    .Add(1111, 1111)
                    .Add(1042, 1042)
                    .Add(1088, 1088)
                    .Add(1108, 1108)
                    .Add(1142, 1142)
                    .Add(1062, 1062)
                    .Add(1063, 1063)
                    .Add(2110, 2110)
                    .Add(1086, 1086)
                    .Add(1100, 1100)
                    .Add(1082, 1082)
                    .Add(1112, 1112)
                    .Add(1153, 1153)
                    .Add(1102, 1102)
                    .Add(2128, 2128)
                    .Add(1104, 1104)
                    .Add(1121, 1121)
                    .Add(1044, 1044)
                    .Add(2068, 2068)
                    .Add(1096, 1096)
                    .Add(1045, 1045)
                    .Add(1046, 1046)
                    .Add(2070, 2070)
                    .Add(1094, 1094)
                    .Add(1047, 1047)
                    .Add(2072, 2072)
                    .Add(1048, 1048)
                    .Add(1049, 1049)
                    .Add(2073, 2073)
                    .Add(1083, 1083)
                    .Add(1103, 1103)
                    .Add(3098, 3098)
                    .Add(2074, 2074)
                    .Add(1072, 1072)
                    .Add(1074, 1074)
                    .Add(1113, 1113)
                    .Add(1051, 1051)
                    .Add(1060, 1060)
                    .Add(1143, 1143)
                    .Add(1070, 1070)
                    .Add(11274, 11274)
                    .Add(16394, 16394)
                    .Add(13322, 13322)
                    .Add(9226, 9226)
                    .Add(5130, 5130)
                    .Add(7178, 7178)
                    .Add(12298, 12298)
                    .Add(17418, 17418)
                    .Add(4106, 4106)
                    .Add(18442, 18442)
                    .Add(2058, 2058)
                    .Add(19466, 19466)
                    .Add(6154, 6154)
                    .Add(15370, 15370)
                    .Add(10250, 10250)
                    .Add(20490, 20490)
                    .Add(1034, 1034)
                    .Add(14346, 14346)
                    .Add(8202, 8202)
                    .Add(1089, 1089)
                    .Add(2077, 2077)
                    .Add(1053, 1053)
                    .Add(1114, 1114)
                    .Add(1064, 1064)
                    .Add(1097, 1097)
                    .Add(1092, 1092)
                    .Add(1098, 1098)
                    .Add(1054, 1054)
                    .Add(1105, 1105)
                    .Add(1073, 1073)
                    .Add(1055, 1055)
                    .Add(1090, 1090)
                    .Add(1058, 1058)
                    '.Add(0, 0)
                    .Add(1056, 1056)
                    .Add(2115, 2115)
                    .Add(1091, 1091)
                    .Add(1075, 1075)
                    .Add(1066, 1066)
                    .Add(1106, 1106)
                    .Add(1076, 1076)
                    .Add(1085, 1085)
                    .Add(1077, 1077)
                End With
                mslCultureInfo = New SortedList
                Dim i As Integer, oCultureInfo As CultureInfo, strErr As String = "", intLCID As Integer
                For i = 0 To mslLCID.Count - 1
                    intLCID = mslLCID.GetByIndex(i)
                    strStepName = "New CultureInfo(" & intLCID.ToString & ")"
                    oCultureInfo = Nothing
                    Dim strRet As String = Me.mNewCultureInfo("", oCultureInfo, intLCID)
                    If strRet = "OK" Then
                        mslCultureInfo.Add(oCultureInfo.LCID, oCultureInfo)
                    Else
                        strErr &= strRet & ";"
                    End If
                Next
                If strErr <> "" Then
                    strStepName = "mNewCultureInfo Error"
                    Throw New Exception(strErr)
                End If
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mInitCultureSortList", ex)
        End Try
    End Sub


    Private Function mNewCultureInfo(CultureName As String, ByRef OutCultureInfo As CultureInfo, Optional LCID As Integer = 2052) As String
        Try
            If CultureName <> "" Then
                OutCultureInfo = New CultureInfo(CultureName)
            Else
                OutCultureInfo = New CultureInfo(LCID)
            End If
            Return "OK"
        Catch ex As Exception
            OutCultureInfo = Nothing
            Return Me.GetSubErrInf("mNewCultureInfo", ex)
        End Try
    End Function


    Private Function mGetCultureInfo(LCID As Integer) As CultureInfo
        Try
            If mslCultureInfo.IndexOfKey(LCID) >= 0 Then
                Return mslCultureInfo.Item(LCID)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetCultureInfo", ex)
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' 获取多语言文本
    ''' </summary>
    ''' <param name="GlobalKey">全局键值</param>
    ''' <param name="DefaultText">默认值，若找不到则使用之</param>
    Public Function GetMLangText(GlobalKey As String, DefaultText As String) As String
        Return Me.mGetMLangText("Global", GlobalKey, DefaultText)
    End Function

    ''' <summary>
    ''' 获取多语言文本
    ''' </summary>
    ''' <param name="ObjName">对象名称，如窗体名称</param>
    ''' <param name="Key">键值</param>
    ''' <param name="DefaultText">默认值，若找不到则使用之</param>
    Public Function GetMLangText(ObjName As String, Key As String, DefaultText As String) As String
        Return Me.mGetMLangText(ObjName, Key, DefaultText)
    End Function

    Private Function mGetMLangText(ObjName As String, Key As String, DefaultText As String) As String
        Try
            Dim strKey As String = ObjName & "." & Key
            If mslMLangText.IndexOfKey(strKey) >= 0 Then
                Return mslMLangText.Item(strKey)
            ElseIf DefaultText = "" Then
                Return strKey
            Else
                Return DefaultText
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mGetMLangText", ex, True)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 获取与本LCID相似的语言区域对象
    ''' </summary>
    ''' <param name="LCID">LCID</param>
    ''' <returns></returns>
    Public Function GetAlikeCulture(LCID As Integer) As CultureInfo
        Dim strStepName As String = ""
        Try
            Dim strKey As String = LCID.ToString
            GetAlikeCulture = Me.mGetCultureInfo(LCID)
            If GetAlikeCulture Is Nothing Then
                strStepName = "mInitCultureSortList"
                Me.mInitCultureSortList()
                If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GetAlikeCulture", ex)
            Return Nothing
        End Try
    End Function


    Public Function GetAllLangInf(GetInfFmt As EnmGetInfFmt) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "mInitCultureSortList"
            Me.mInitCultureSortList()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            GetAllLangInf = ""
            Select Case GetInfFmt
                Case EnmGetInfFmt.TabSeparator
                    GetAllLangInf = Me.mGetCultureInfoTab(True) & Me.OsCrLf
                Case EnmGetInfFmt.Markdown
                    GetAllLangInf = Me.mGetCultureInfoMD(True) & Me.OsCrLf
            End Select
            Dim oCultureInfo As CultureInfo, strRow As String = ""
            strStepName = "Get Rows"
            For Each obj In mslCultureInfo
                oCultureInfo = obj.value
                Select Case GetInfFmt
                    Case EnmGetInfFmt.TabSeparator
                        strRow = Me.mGetCultureInfoTab(False, oCultureInfo.LCID)
                    Case EnmGetInfFmt.Markdown
                        strRow = Me.mGetCultureInfoMD(False, oCultureInfo.LCID)
                End Select
                If Me.LastErr = "" Then
                    GetAllLangInf &= strRow & Me.OsCrLf
                Else
                    GetAllLangInf &= Me.LastErr & Me.OsCrLf
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("GetAllLangInf", ex)
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' 导入多语言信息
    ''' </summary>
    ''' <param name="IsAuto">是否自动，是则当前语言区域文件找不到会自动寻找近似的</param>
    ''' <returns></returns>
    Public Function LoadMLangInf(IsAuto As Boolean) As String
        Return Me.mLoadMLangInf(IsAuto)
    End Function

    Public Function LoadMLangInf() As String
        Return Me.mLoadMLangInf(True)
    End Function

    Public Function GetCanUseCultureXml() As String
        Dim LOG As New PigStepLog("GetCanUseCultureXml")
        Try
            GetCanUseCultureXml = ""
            LOG.StepName = "RefCanUseCultureList"
            Me.RefCanUseCultureList()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            If Me.CanUseCultureList IsNot Nothing Then
                Dim oPigXml As New PigXml(False)
                oPigXml.AddEleLeftSign("CanUseCulture")
                For Each oCultureInfo As CultureInfo In Me.CanUseCultureList
                    With oCultureInfo
                        oPigXml.AddEle("Name", .Name)
                        oPigXml.AddEle("LCID", .LCID)
                        oPigXml.AddEle("MLangFileName", Me.CurrMLangTitle & "." & .Name)
                        oPigXml.AddEle("DisplayName", .DisplayName)
                        oPigXml.AddEle("NativeName", .NativeName)
                    End With
                Next
                oPigXml.AddEleRightSign("CanUseCulture")
                GetCanUseCultureXml = oPigXml.MainXmlStr
                oPigXml = Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    Public Sub RefCanUseCultureList()
        Try
            Dim oDirectoryInfo As DirectoryInfo = New DirectoryInfo(Me.CurrMLangDir)
            ReDim mciaCanUseCultureList(-1)
            For Each oFile As FileInfo In oDirectoryInfo.GetFiles(Me.CurrMLangTitle & ".*")
                Dim strExtName As String = oFile.Extension
                If Me.mIsMLangFileExt(strExtName) = True Then
                    strExtName = Mid(strExtName, 2)
#If NETFRAMEWORK Then
                    Dim intCnt As Integer = mciaCanUseCultureList.Length
#Else
                    Dim intCnt As Integer = mciaCanUseCultureList.Count
#End If
                    ReDim Preserve mciaCanUseCultureList(intCnt)
                    If IsNumeric(strExtName) Then
                        mciaCanUseCultureList(intCnt) = New CultureInfo(CInt(strExtName))
                    Else
                        mciaCanUseCultureList(intCnt) = New CultureInfo(strExtName)
                    End If
                End If
            Next
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("RefCanUseCultureList", ex)
        End Try
    End Sub

    Private Function mIsMLangFileExt(ExtName As String) As Boolean
        Dim strStepName As String = ""
        Try
            strStepName = "mInitCultureSortList"
            Me.mInitCultureSortList()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            mIsMLangFileExt = False
            For Each obj In mslCultureInfo
                Dim oCultureInfo As CultureInfo = obj.value
                Dim strMatName As String = "." & oCultureInfo.LCID.ToString
                If ExtName = strMatName Then
                    mIsMLangFileExt = True
                    Exit For
                End If
                strMatName = "." & oCultureInfo.Name
                If Me.IsWindows = True Then
                    ExtName = UCase(ExtName)
                    strMatName = UCase(strMatName)
                End If
                If ExtName = strMatName Then
                    mIsMLangFileExt = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("mIsMLangFile", strStepName, ex)
            Return False
        End Try
    End Function

    Private Function mLoadMLangInf(IsAuto As Boolean) As String
        Dim LOG As New PigStepLog("mLoadMLangInf")
        Try
            Dim strMLangFile As String
            LOG.StepName = "mGetMLangFile"
            If IsAuto = True Then
                strMLangFile = Me.mGetMLangFile(mciCurrent.LCID, True)
                If strMLangFile = "" Then strMLangFile = Me.mGetMLangFile(1033, True) 'en-us
            Else
                strMLangFile = Me.mGetMLangFile(mciCurrent.LCID)
            End If
            If strMLangFile = "" Then
                mstrCurrMLangFile = ""
                Throw New Exception("No available MLangFile(" & Me.CurrMLangTitle & ") on " & Me.CurrMLangDir)
            Else
                mstrCurrMLangFile = strMLangFile
            End If
            If Me.mPigFunc.IsFileExists(Me.CurrMLangFile) = False Then Throw New Exception("MLangFile" & Me.CurrMLangFile & " not found")
            mslMLangText = New SortedList
            Dim tsMain As mTextStream
            LOG.StepName = "OpenTextFile"
            tsMain = Me.mFS.OpenTextFile(Me.CurrMLangFile, mFileSystemObject.IOMode.ForReading)
            If Me.mFS.LastErr <> "" Then Throw New Exception(Me.mFS.LastErr)
            Dim strObjName As String = "", strNewObjName As String, strKey As String, strMLangText As String
            Do While Not tsMain.AtEndOfStream
                Dim strLine As String = tsMain.ReadLine
                If strLine.IndexOf("{") >= 0 And strLine.IndexOf("}") > 0 Then
                    strNewObjName = Me.mPigFunc.GetStr(strLine, "{", "}")
                    If strObjName <> strNewObjName Then
                        strObjName = strNewObjName
                    End If
                Else
                    strKey = Me.mPigFunc.GetStr(strLine, "[", "]=")
                    If strKey <> "" Then
                        strMLangText = strLine
                        Me.mAddMLangText(strObjName, strKey, strMLangText)
                    End If
                End If
            Loop
            tsMain.Close()
            'Dim srMLang As New StreamReader(strMLangFile)
            'Dim strObjName As String = "", strNewObjName As String, strKey As String, strMLangText As String
            'Do Until srMLang.EndOfStream
            '    Dim strLine As String = srMLang.ReadLine
            '    If strLine.IndexOf("{") >= 0 And strLine.IndexOf("}") > 0 Then
            '        strNewObjName = Me.mPigFunc.GetStr(strLine, "{", "}")
            '        If strObjName <> strNewObjName Then
            '            strObjName = strNewObjName
            '        End If
            '    Else
            '        strKey = Me.mPigFunc.GetStr(strLine, "[", "]=")
            '        If strKey <> "" Then
            '            strMLangText = strLine
            '            Me.mAddMLangText(strObjName, strKey, strMLangText)
            '        End If
            '    End If
            'Loop
            'srMLang.Close()
            'srMLang = Nothing
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(Me.CurrMLangFile)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try

    End Function

    Public Sub AddMLangText(GlobalKey As String, MLangText As String)
        Me.mAddMLangText("Global", GlobalKey, MLangText)
    End Sub

    Public Sub AddMLangText(ObjName As String, Key As String, MLangText As String)
        Me.mAddMLangText(ObjName, Key, MLangText)
    End Sub


    Private Sub mAddMLangText(ObjName As String, Key As String, MLangText As String)
        Try
            Dim strKey As String = ObjName & "." & Key
            If mslMLangText.IndexOfKey(strKey) = -1 Then
                Me.mPigFunc.EscapeStr(MLangText)
                mslMLangText.Add(strKey, MLangText)
            Else
                Throw New Exception("Key value already exists")
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mAddMLangText", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 生成一个多语言文本|Generate a multilingual text
    ''' </summary>
    ''' <param name="GlobalKey">键值|Key value</param>
    ''' <param name="MLangText">多语言文本|Multilingual text</param>
    Public Function MkMLangText(GlobalKey As String, MLangText As String) As String
        Dim strRet As String = ""
        Try
            MkMLangText = ""
            strRet = Me.mMkMLangText("Global", GlobalKey, MLangText, MkMLangText)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            MkMLangText = ""
            Return Me.GetSubErrInf("MkMLangText", ex)
        End Try
    End Function

    ''' <summary>
    ''' 生成一个多语言文本|Generate a multilingual text
    ''' </summary>
    ''' <param name="ObjName">对象名称|Object Name</param>
    ''' <param name="Key">键值|Key value</param>
    ''' <param name="MLangText">多语言文本|Multilingual text</param>
    Public Function MkMLangText(ObjName As String, Key As String, MLangText As String) As String
        Dim strRet As String = ""
        Try
            MkMLangText = ""
            strRet = Me.mMkMLangText(ObjName, Key, MLangText, MkMLangText)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            MkMLangText = ""
            Return Me.GetSubErrInf("MkMLangText", ex)
        End Try
    End Function

    Private Function mMkMLangText(ObjName As String, Key As String, MLangText As String, ByRef OutMLangText As String) As String
        Try
            Me.mPigFunc.EscapeStr(MLangText)
            OutMLangText = "{" & ObjName & "}" & Me.OsCrLf
            OutMLangText &= "[" & Key & "]=" & MLangText & Me.OsCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mMkMLangText", ex)
        End Try
    End Function

End Class
