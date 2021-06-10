Attribute VB_Name = "modGETools"
Option Explicit

Public Enum xpFilePart          '�ļ��Ĳ���
    xpFilePart_Path = 1         '·��
    xpFilePart_FileTitle = 2    '�ļ���
    xpFilePart_ExtName = 3      '��չ��
    xpFilePart_DriveNo = 4      '��������
End Enum

Public Enum xpXMLAddWhere   '����һ��XML��ǵĲ���
    xpXMLAddWhere_Left = 0  '�󲿣���<Version>
    xpXMLAddWhere_Right = 1 '�Ҳ�����</Version>
    xpXMLAddWhere_Both = 2    '���ߣ���<Version></Version>
End Enum

Public Enum xpEncStr_Enc            '�ӽ��ܲ�����ʽ������ͷ�������ܺ����Ļ�����һ������ַ���������ͷ����ԭ�ĺ����ĵĳ�����ͬ
    xpEncStr_EncWithKeyHead = 0     '���ܴ���ͷ
    xpEncStr_UnEncWithKeyHead = 1   '���ܴ���ͷ
    xpEncStr_EncNoKeyHead = 2       '���ܲ�����ͷ
    xpEncStr_UnEncNoKeyHead = 3     '���ܲ�����ͷ
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
    
    '�Ӹ�Ŀ¼�����𼶽���Ŀ¼
    If Right(Path, 1) <> "\" Then
        Path = Path & "\" '���洦����Ҫ
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
'���ߣ�����
'����ʱ�䣺1999.06
'���²����������޷��������ݿ��
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
'Unicode�ַ����ӽ��ܺ��������Դ�����
'<����ֵ>'���ܻ���ܺ���ַ��������ؿմ���ʾ���ܻ����ʧ�ܣ����ܺ�İ�16������ʽ��ʾ���磺92D1CDB405
'<strSourse>'ԭ�Ļ����ģ����İ�16������ʽ��ʾ
'<strEncKey>'��Կ
'<lngOpt>'������ʽ������μ�xpEncStr_Encö�ٳ�����˵��
'���ߣ�����
'����ʱ��:1999.11
'�޸�ʱ��:2001.4
'ר�����ڶ�Unitcode�ַ������мӽ���
'������ʮ��������ʽ��ʾ
'��:
'lngOpt :
'   xpEncStr_EncWithKeyHead = 0�����ܴ���ͷ
'   xpEncStr_UnEncWithKeyHead = 1�����ܴ���ͷ
'   xpEncStr_EncNoKeyHead = 2�����ܲ�����ͷ
'   xpEncStr_unEncNoKeyHead = 3�����ܲ�����ͷ
'����ֵΪ�ձ�ʾ�ӽ���ʧ��

Const MAXUNICODE As Integer = 255 '~
Const MINUNICODE As Integer = 0 '�ո�
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
    Case xpEncStr_EncWithKeyHead '���ܴ���ͷ
        bolIsEnc = True
        bolHasKeyHead = True
    Case xpEncStr_UnEncWithKeyHead '���ܴ���ͷ
        bolIsEnc = False
        bolHasKeyHead = True
    Case xpEncStr_EncNoKeyHead '���ܲ�����ͷ
        bolIsEnc = True
        bolHasKeyHead = False
    Case xpEncStr_UnEncNoKeyHead '���ܲ�����ͷ
        bolIsEnc = False
        bolHasKeyHead = False
    End Select
    
    If Len(strEncKey) = 0 Then strEncKey = "����"
    
    'ԭ�ļ���Կת����ASCII�봮,������ʮ��������ʽ��ԭ��ASCII�ַ���
    If bolIsEnc = True Then
        strSourse = StrConv(strSourse, vbFromUnicode)
    Else
        strSourse = HexStrToUniCodeStr(strSourse, False)
    End If
    strEncKey = StrConv(strEncKey, vbFromUnicode)
    
    Randomize   '��ʼ�������������
    If bolHasKeyHead = True Then
        If bolIsEnc = True Then
            strKeyHead = ChrB(Int((MAXUNICODE - MINUNICODE + 1) * Rnd + MINUNICODE)) '����ASCII��MINUNICODE��MAXUNICODE�����
'            strKeyHead = ChrB(gmGetRandNum(MINUNICODE, MAXUNICODE)) '����ASCII��MINUNICODE��MAXUNICODE�����
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
        If bolIsEnc = True Then '����
           'Debug.Print "����ǰlngSourse(" & i & "):" & CStr(lngSourse)
            lngSourse = lngSourse + lngKey
            If lngSourse > MAXUNICODE Then
                Do While lngSourse > MAXUNICODE
                    lngSourse = lngSourse - (MAXUNICODE - MINUNICODE)
                Loop
            End If
            'Debug.Print "���ܺ�lngSourse(" & i & "):" & CStr(lngSourse)
        Else '����
            lngSourse = lngSourse - lngKey
            'Debug.Print "����ǰlngSourse(" & i & "):" & CStr(lngSourse)
            If lngSourse < MINUNICODE Then
                Do While lngSourse < MINUNICODE
                    lngSourse = lngSourse + (MAXUNICODE - MINUNICODE)
                Loop
            End If
            'Debug.Print "���ܺ�lngSourse(" & i & "):" & CStr(lngSourse)
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
'��ʾ��ȷ�������ʱ��
    GENow = CStr(Round(Timer - DateDiff("s", Date, Now), 3))
    If InStr(GENow, ".") = 0 Then
        GENow = Now & ".000"
    Else
        If Len(GENow) < 4 Then GENow = GENow & String(4 - Len(GENow), "0")
        GENow = Format(Now, "yyyy-mm-dd hh:mm:ss") & GENow
    End If
End Function

Public Function HexStrToUniCodeStr(ByVal strInput As String, Optional IsConvToUnicode As Boolean = True) As String
'��ʮ��������ʽ���ַ���ת����UnitCode�ַ���
'���磺"��1"->D6D031
'ת��ʧ�ܷ��ؿմ�
    
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
'��ASCII�ַ���ת����ʮ��������ʽ
'���磺"��1"->D6D031
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

