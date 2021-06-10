Attribute VB_Name = "modGetFileAndDirList"
Option Explicit
Const VSSVER_SCC_FILE As String = "vssver.scc"
Dim strLogFile As String
Dim oFS As FileSystemObject
Dim strNoScanDirListPath As String
Dim strNoScanDirList As String
Dim oNoScanDirItems As NoScanDirItems
Public Sub Main()
Dim strError As String, tsDir As TextStream, tsFile As TextStream, oFolder As Folder, oFile As File
Dim i As Long, j As Long, strDirListPath As String, strFileListPath As String, strRootDir As String, strLine As String
Dim strDirList As String, strOptRes As String, strFilePath As String, strDirPath As String, lngSubLen As Long
Dim strCmd As String, bolIsAbsDir As Boolean
    On Error GoTo ErrOcc:
    strLogFile = App.Path & "\" & App.Title & ".log"
    Set oFS = New FileSystemObject
    strCmd = LCase(Trim(Command()))
    If InStr(strCmd, "/abs") <> 0 Then
        bolIsAbsDir = True
    Else
        bolIsAbsDir = False
    End If
    strRootDir = App.Path
    '--------------------
    strNoScanDirListPath = strRootDir & "\NoScanDir.txt"
    strError = "打开" & strNoScanDirListPath
    Set tsDir = oFS.OpenTextFile(strNoScanDirListPath, ForReading, True)
    Set oNoScanDirItems = New NoScanDirItems
    If tsDir.AtEndOfStream = False Then
        strNoScanDirList = tsDir.ReadAll
        tsDir.Close
        If Right(strNoScanDirList, 2) <> vbCrLf Then strNoScanDirList = strNoScanDirList & vbCrLf
        Do While True
            strLine = GetStr(strNoScanDirList, "", vbCrLf)
            If strLine = "" Then Exit Do
            oNoScanDirItems.Add strLine
            DoEvents
        Loop
    End If
    '--------------------
    strDirListPath = strRootDir & "\DirList.txt"
    strError = "打开" & strDirListPath
    Set tsDir = oFS.OpenTextFile(strDirListPath, ForWriting, True)
    '--------------------
    strFileListPath = strRootDir & "\FileList.txt"
    strError = "打开" & strFileListPath
    Set tsFile = oFS.OpenTextFile(strFileListPath, ForWriting, True)
    '--------------------
    strError = "获取根目录" & strRootDir
    Set oFolder = oFS.GetFolder(strRootDir)
    strDirList = GetDirList(strRootDir)
    Do While True
        strLine = GetStr(strDirList, "", vbCrLf)
        If strLine = "" Then Exit Do
        If Left(strLine, 1) = "#" Then
            OptLogInf strLine & "为出错目录，不处理", strLogFile
        ElseIf IsNoScanDir(strLine) = True Then
            OptLogInf strLine & "为免扫描目录，不处理", strLogFile
        ElseIf IsFolderExists(strLine) = False Then
            OptLogInf strLine & "已经不存在，不处理", strLogFile
        Else
            strError = "获取当前目录" & strLine
            strDirPath = strLine
            If bolIsAbsDir = False Then
                lngSubLen = Len(strDirPath) - Len(strRootDir) - 1
                If lngSubLen <= 0 Then
                    strDirPath = ".\"
                Else
                    strDirPath = ".\" & Right(strDirPath, lngSubLen)
                End If
            End If
            tsDir.WriteLine strDirPath
            Set oFolder = oFS.GetFolder(strLine)
            strError = "遍历当前目录" & strLine & "的文件"
            For Each oFile In oFolder.Files
                If LCase(GetFilePart(oFile.Name)) <> VSSVER_SCC_FILE Then
                    strFilePath = oFile.Path
                    If bolIsAbsDir = False Then
                        strFilePath = ".\" & Right(strFilePath, Len(strFilePath) - Len(strRootDir) - 1)
                    End If
                    tsFile.WriteLine strFilePath & vbTab & oFile.Size & vbTab & Format(oFile.DateLastModified, "yyyy-mm-dd hh:mm:ss")
                End If
                DoEvents
            Next
        End If
        DoEvents
    Loop
    tsDir.Close
    tsFile.Close
    On Error GoTo 0
    Set oFolder = Nothing
    Set oFile = Nothing
    Set tsDir = Nothing
    Set tsFile = Nothing
    Set oFS = Nothing
    Exit Sub
ErrOcc:
    If Err.Number <> 0 Then strError = strError & "错误：" & Err.Description
    OptLogInf strError, strLogFile
    Set oFolder = Nothing
    Set oFile = Nothing
    Set tsDir = Nothing
    Set tsFile = Nothing
    Set oFS = Nothing
    On Error GoTo 0
End Sub

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

Public Sub gmGetDirList(oFolder As Folder, DirList As String, OptRes As String)
'递归获取子目录列表
Dim oSubFolder As Folder
    On Error GoTo ErrOcc:
    If DirList = "" Then
        DirList = oFolder.Path & vbCrLf
    Else
        DirList = DirList & oFolder.Path & vbCrLf
    End If
    If oFolder.SubFolders.Count = 0 Then GoTo ExitSub:
    For Each oSubFolder In oFolder.SubFolders
        OptRes = ""
        gmGetDirList oSubFolder, DirList, OptRes
        If OptRes <> "OK" Then
            DirList = DirList & "#获取" & oSubFolder.Path & "的子目录出错：" & OptRes & vbCrLf
        End If
        DoEvents
    Next
ExitSub:
    OptRes = "OK"
    Set oSubFolder = Nothing
    Exit Sub
ErrOcc:
    If Err.Number <> 0 Then OptRes = OptRes & Err.Description
    Set oSubFolder = Nothing
End Sub


Public Function GetDirList(ByVal BaseDir As String) As String
'出错的以#开头
Dim oFolder As Folder, strRet As String, strOptRes As String
    Set oFolder = oFS.GetFolder(BaseDir)
    gmGetDirList oFolder, strRet, strOptRes
    Set oFolder = Nothing
    GetDirList = strRet
End Function

Public Function IsNoScanDir(ByVal ChkDirPath As String) As Boolean
Dim i As Long
    For i = 1 To oNoScanDirItems.Count
        With oNoScanDirItems(i)
            IsNoScanDir = .IsMeNoScan(ChkDirPath)
            If IsNoScanDir = True Then Exit Function
        End With
    Next
End Function
