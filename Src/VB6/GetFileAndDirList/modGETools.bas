Attribute VB_Name = "modGETools"
Option Explicit

Public Enum xpFilePart          '文件的部分
    xpFilePart_Path = 1         '路径
    xpFilePart_FileTitle = 2    '文件名
    xpFilePart_ExtName = 3      '扩展名
    xpFilePart_DriveNo = 4      '驱动器名
End Enum

Public Enum xpXMLAddWhere   '增加一个XML标记的部分
    xpXMLAddWhere_Left = 0  '左部，如<Version>
    xpXMLAddWhere_Right = 1 '右部，如</Version>
    xpXMLAddWhere_Both = 2    '两者，如<Version></Version>
End Enum

Public Enum xpEncStr_Enc            '加解密操作方式，带键头，即加密后密文会增加一个随机字符，不带键头，即原文和密文的长度相同
    xpEncStr_EncWithKeyHead = 0     '加密带键头
    xpEncStr_UnEncWithKeyHead = 1   '解密带键头
    xpEncStr_EncNoKeyHead = 2       '加密不带键头
    xpEncStr_UnEncNoKeyHead = 3     '解密不带键头
End Enum
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Public goFS As FileSystemObject



Public Function GEMkDir(ByVal Path As String) As String
Dim strDir As String, strDrive As String, strTmp As String
    
    On Error GoTo ErrOcc:
    GEMkDir = "OK"
    strDrive = GetFilePart(Path, xpFilePart_DriveNo)
    If Len(strDrive) = 0 Then
        strDrive = GetFilePart(CurDir(), xpFilePart_DriveNo) & ":\"
        If Left(Path, 1) = "\" Then
            Path = strDrive & Path
        Else
            Path = CurDir() & "\" & Path
        End If
    Else
        If InStr(strDrive, "$") = 0 Then strDrive = strDrive & ":\"
    End If
    
    '从根目录向上逐级建立目录
    If Right(Path, 1) <> "\" Then
        Path = Path & "\" '后面处理需要
    End If
    strDir = GetStr(Path, strDrive, "\", False)
    strDir = strDrive & strDir & "\"
    On Error Resume Next
    If goFS Is Nothing Then Set goFS = New FileSystemObject
    Do While True
        DoEvents
        goFS.CreateFolder strDir
        Err.Clear
        strTmp = GetStr(Path, strDir, "\", False)
        If Len(strTmp) = 0 Then
            Exit Do
        End If
        strDir = strDir & strTmp & "\"
    Loop
    
    On Error GoTo 0
    Exit Function
ErrOcc:
    If Err.Number <> 0 Then GEMkDir = Err.Description
    On Error GoTo 0
End Function


Public Function GetFilePart(ByVal FileName As String, Optional FilePart As xpFilePart = xpFilePart_FileTitle) As String
Dim strTemp As String, i As Long, lngLen As Long
Dim strPath As String, strFileTitle As String

    Select Case FilePart
    Case xpFilePart_DriveNo
        GetFilePart = GetStr(FileName, "", ":", False)
        If GetFilePart = "" Then
            GetFilePart = GetStr(FileName, "", "$", False)
            If GetFilePart <> "" Then GetFilePart = GetFilePart & "$"
        End If
    Case xpFilePart_ExtName
        'GetFilePart = GetStr(FileName, ".", "", False)
        lngLen = Len(FileName)
        For i = lngLen To 1 Step -1
            Select Case Mid(FileName, i, 1)
            Case "/", ":", "$"
                Exit For
            Case "."
                GetFilePart = Mid(FileName, i + 1)
                Exit For
            End Select
            
        Next
    Case xpFilePart_FileTitle, xpFilePart_Path
        Do While True
            DoEvents
            strTemp = GetStr(FileName, "", "\", True)
            If Len(strTemp) = 0 Then
                If Right(strPath, 1) = "\" Then
                    If Right(strPath, 2) <> ":\" Then
                        strPath = Left(strPath, Len(strPath) - 1)
                    End If
                ElseIf Left(FileName, 1) = "\" Then
                    strPath = "\"
                    FileName = Mid(FileName, 2)
                End If
                If FileName <> "" Then
                    strFileTitle = FileName
                Else
                    strFileTitle = "."
                End If
                Exit Do
            End If
            strPath = strPath & strTemp & "\"
        Loop
        If FilePart = xpFilePart_FileTitle Then
            GetFilePart = strFileTitle
        Else
            GetFilePart = strPath
        End If
    Case Else
        GetFilePart = "Error"
    End Select
End Function

Public Function GetStr(SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
Dim lngBegin As Long
Dim lngEnd As Long
Dim lngBeginLen As Long
Dim lngEndLen As Long
    
    On Error GoTo ErrOcc:
    lngBeginLen = Len(strBegin)
    lngBegin = InStr(SourceStr, strBegin)
    lngEndLen = Len(strEnd)
    If lngEndLen = 0 Then
        lngEnd = Len(SourceStr) + 1
    Else
        lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd): If lngBegin = 0 Then GoTo ErrOcc:
    End If
    If lngEnd <= lngBegin Then GoTo ErrOcc:
    If lngBegin = 0 Then GoTo ErrOcc:
    
    GetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
    If IsCut = True Then
        SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
    End If
    On Error GoTo 0
    Exit Function
    
ErrOcc:
    GetStr = ""
    On Error GoTo 0
End Function

Public Sub OptLogInf(ByVal OptString As String, Optional ByVal LogFileName As String, Optional IsShowTime As Boolean = True)
'作者：萧鹏
'创建时间：1999.06
'记下操作错误，如无法连接数据库等
Dim intFreeFile As Integer
    On Error Resume Next
    intFreeFile = FreeFile()
    Open LogFileName For Append Shared As #intFreeFile
    If IsShowTime = True Then
        Print #intFreeFile, Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
    Print #intFreeFile, OptString
    Close #intFreeFile
    On Error GoTo 0
End Sub


Public Function IsFileExists(ByVal FilePath As String) As Boolean
Dim oFile As File
    On Error Resume Next
    If goFS Is Nothing Then Set goFS = New FileSystemObject
    Err.Clear
    Set oFile = goFS.GetFile(FilePath)
    If Err.Number = 0 Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
    Set oFile = Nothing
    On Error GoTo 0
End Function

Public Function IsFolderExists(ByVal FolderPath As String) As Boolean
Dim oFolder As Folder
    On Error Resume Next
    Err.Clear
    If goFS Is Nothing Then Set goFS = New FileSystemObject
    Set oFolder = goFS.GetFolder(FolderPath)
    If Err.Number = 0 Then
        IsFolderExists = True
    Else
        IsFolderExists = False
    End If
    Set oFolder = Nothing
    On Error GoTo 0
End Function


Public Function XmlGetStr(SourceStr As String, XMLSign As String, Optional IsCrlf As Boolean = False) As String
Dim strBegin As String, strEnd As String
    strBegin = "<" & XMLSign & ">"
    strEnd = "</" & XMLSign & ">"
    If IsCrlf = True Then
        strBegin = strBegin & vbCrLf
        strEnd = strEnd & vbCrLf
    End If
    XmlGetStr = GetStr(SourceStr, strBegin, strEnd, True)
End Function

Public Sub XMLAddStr(MainStr As String, XMLSign As String, XMLValue As String, _
                     Optional ByVal AddWhere As xpXMLAddWhere = xpXMLAddWhere_Both, _
                     Optional ByVal IsCrlf As Boolean = False, _
                     Optional ByVal LeftTab As Integer = 0)
Dim strTmp As String
    
    If LeftTab > 0 Then
        MainStr = MainStr & String(LeftTab, vbTab)
    End If
    Select Case AddWhere
    Case xpXMLAddWhere_Left
        strTmp = "<" & XMLSign & ">"
    Case xpXMLAddWhere_Right
        strTmp = "</" & XMLSign & ">"
    Case xpXMLAddWhere_Both
        strTmp = "<" & XMLSign & ">" & XMLValue & "</" & XMLSign & ">"
    Case Else
        strTmp = "<" & XMLSign & ">"
    End Select
    MainStr = MainStr & strTmp
    
    If IsCrlf = True Then
        MainStr = MainStr & vbCrLf
    End If
    On Error GoTo 0
End Sub


Public Sub GEOptLogInf(ByVal OptString, ByVal LogFileName As String, Optional ByVal IsShowTime As Boolean = True)
    OptString = "[PID=" & GetCurrentProcessId() & "]" & OptString
    If IsShowTime = True Then
        OptString = "[" & GENow() & "]" & OptString
    End If
    OptLogInf OptString, LogFileName, False
End Sub

Public Function EncUnicodeStr(ByVal strSourse As String, ByVal strEncKey As String, ByVal lngOpt As xpEncStr_Enc) As String
'Unicode字符串加解密函数，可以处理汉字
'<返回值>'加密或解密后的字符串，返回空串表示加密或解密失败，解密后的按16进制形式显示，如：92D1CDB405
'<strSourse>'原文或密文，密文按16进制形式显示
'<strEncKey>'密钥
'<lngOpt>'操作方式，具体参见xpEncStr_Enc枚举常量的说明
'作者：萧鹏
'创建时间:1999.11
'修改时间:2001.4
'专门用于对Unitcode字符串进行加解密
'密文以十六进制形式显示
'如:
'lngOpt :
'   xpEncStr_EncWithKeyHead = 0－加密带键头
'   xpEncStr_UnEncWithKeyHead = 1－解密带键头
'   xpEncStr_EncNoKeyHead = 2－加密不带键头
'   xpEncStr_unEncNoKeyHead = 3－解密不带键头
'返回值为空表示加解密失败

Const MAXUNICODE As Integer = 255 '~
Const MINUNICODE As Integer = 0 '空格
Dim i As Long
Dim lngKeyLength As Integer
Dim lngSourse As Integer
Dim lngSourseLen As Integer
Dim lngTarget As Integer
Dim lngKey As Integer
Dim strKeyHead As String
Dim bolHasKeyHead As Boolean
Dim bolIsEnc As Boolean
Dim lngKeyMod As Integer
Dim lngMaxMinRange As Integer

    On Error GoTo ErrOcc:
    
    Select Case lngOpt
    Case xpEncStr_EncWithKeyHead '加密带键头
        bolIsEnc = True
        bolHasKeyHead = True
    Case xpEncStr_UnEncWithKeyHead '解密带键头
        bolIsEnc = False
        bolHasKeyHead = True
    Case xpEncStr_EncNoKeyHead '加密不带键头
        bolIsEnc = True
        bolHasKeyHead = False
    Case xpEncStr_UnEncNoKeyHead '解密不带键头
        bolIsEnc = False
        bolHasKeyHead = False
    End Select
    
    If Len(strEncKey) = 0 Then strEncKey = "萧鹏"
    
    '原文及密钥转换成ASCII码串,密文由十六进制形式还原成ASCII字符串
    If bolIsEnc = True Then
        strSourse = StrConv(strSourse, vbFromUnicode)
    Else
        strSourse = HexStrToUniCodeStr(strSourse, False)
    End If
    strEncKey = StrConv(strEncKey, vbFromUnicode)
    
    Randomize   '初始化随机数生成器
    If bolHasKeyHead = True Then
        If bolIsEnc = True Then
            strKeyHead = ChrB(Int((MAXUNICODE - MINUNICODE + 1) * Rnd + MINUNICODE)) '产生ASCII码MINUNICODE至MAXUNICODE随机数
'            strKeyHead = ChrB(gmGetRandNum(MINUNICODE, MAXUNICODE)) '产生ASCII码MINUNICODE至MAXUNICODE随机数
            EncUnicodeStr = strKeyHead
        Else
            strKeyHead = LeftB(strSourse, 1)
            strSourse = MidB(strSourse, 2)
        End If
    Else
        strKeyHead = ChrB(120)
    End If
    
    lngSourse = LenB(strSourse)
    lngKeyLength = LenB(strEncKey)
    
    lngMaxMinRange = MAXUNICODE - MINUNICODE
    For i = 1 To lngKeyLength
        lngKeyMod = lngKeyMod + AscB(MidB(strEncKey, i, 1))
        lngKeyMod = lngKeyMod Mod lngMaxMinRange
    Next
    
    For i = 1 To lngSourse
        lngKey = AscB(MidB(strEncKey, (i Mod lngKeyLength) + 1, 1))
        lngKey = lngKey + AscB(strKeyHead) + lngKeyMod
        'Debug.Print "lngKey=" & lngKey
        lngSourse = AscB(MidB(strSourse, i, 1))
        If lngSourse < MINUNICODE Or lngSourse > MAXUNICODE Then
            GoTo ErrOcc:
        End If
        If bolIsEnc = True Then '加密
           'Debug.Print "加密前lngSourse(" & i & "):" & CStr(lngSourse)
            lngSourse = lngSourse + lngKey
            If lngSourse > MAXUNICODE Then
                Do While lngSourse > MAXUNICODE
                    lngSourse = lngSourse - (MAXUNICODE - MINUNICODE)
                Loop
            End If
            'Debug.Print "加密后lngSourse(" & i & "):" & CStr(lngSourse)
        Else '解密
            lngSourse = lngSourse - lngKey
            'Debug.Print "解密前lngSourse(" & i & "):" & CStr(lngSourse)
            If lngSourse < MINUNICODE Then
                Do While lngSourse < MINUNICODE
                    lngSourse = lngSourse + (MAXUNICODE - MINUNICODE)
                Loop
            End If
            'Debug.Print "解密后lngSourse(" & i & "):" & CStr(lngSourse)
        End If
        EncUnicodeStr = EncUnicodeStr & ChrB(lngSourse)
    Next
    
    If bolIsEnc = True Then
        EncUnicodeStr = UnicodeStrToHexStr(EncUnicodeStr, False)
    Else
        EncUnicodeStr = StrConv(EncUnicodeStr, vbUnicode)
    End If
        
    On Error GoTo 0
    Exit Function
    
ErrOcc:
    EncUnicodeStr = ""
    On Error GoTo 0
End Function

Public Function GENow() As String
'显示精确到毫秒的时间
    GENow = CStr(Round(Timer - DateDiff("s", Date, Now), 3))
    If InStr(GENow, ".") = 0 Then
        GENow = Now & ".000"
    Else
        If Len(GENow) < 4 Then GENow = GENow & String(4 - Len(GENow), "0")
        GENow = Format(Now, "yyyy-mm-dd hh:mm:ss") & GENow
    End If
End Function

Public Function HexStrToUniCodeStr(ByVal strInput As String, Optional IsConvToUnicode As Boolean = True) As String
'将十六进制形式的字符串转换成UnitCode字符串
'例如："中1"->D6D031
'转换失败返回空串
    
    On Error GoTo ErrOcc:
    HexStrToUniCodeStr = ""
    Do While Len(strInput) > 0
        HexStrToUniCodeStr = HexStrToUniCodeStr & ChrB(CLng("&h" & Left(strInput, 2)))
        strInput = Mid(strInput, 3)
    Loop
    If IsConvToUnicode = True Then
        HexStrToUniCodeStr = StrConv(HexStrToUniCodeStr, vbUnicode)
    End If
    On Error GoTo 0
    Exit Function

ErrOcc:
    HexStrToUniCodeStr = ""
    On Error GoTo 0
End Function

Public Function UnicodeStrToHexStr(ByVal strInput As String, Optional IsConvToFromUnicode As Boolean = True) As String
'将ASCII字符串转换成十六进制形式
'例如："中1"->D6D031
Dim lngStrLen As Integer
Dim i, lngAscCode As Integer
    
    If IsConvToFromUnicode = True Then
        strInput = StrConv(strInput, vbFromUnicode)
    End If
    lngStrLen = LenB(strInput)
    UnicodeStrToHexStr = ""
    For i = 1 To lngStrLen
        lngAscCode = AscB(MidB(strInput, i, 1))
        UnicodeStrToHexStr = UnicodeStrToHexStr & Right("0" & Hex(lngAscCode), 2)
    Next
    UnicodeStrToHexStr = UnicodeStrToHexStr
End Function

