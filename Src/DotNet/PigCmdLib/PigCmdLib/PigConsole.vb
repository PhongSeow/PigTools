'**********************************
'* Name: PigConsole
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 增加控制台的功能|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.11
'* Create Time: 15/1/2022
'*1.1 23/1/2022    Add GetKeyType1, modify GetPwdStr
'*1.2 3/2/2022     Add GetLine
'*1.3 4/2/2022     Add mGetKeyTypeForLine,ClearLine,mGetLine
'*1.4 5/2/2022     Modify GetLine.
'*1.5 6/2/2022     Modify mGetKeyTypeForLine,mGetKeyTypeForPwd,mGetLine
'*1.6 19/3/2022    Modify mGetLine,GetLine,GetPwdStr, add mGetPwdStr, 
'*1.7 16/4/2022    Add IsYesOrNo and SimpleMenu.
'*1.8 17/4/2022    Modify SimpleMenu
'*1.9 29/4/2022    Add DisplayPause,mDisplayPause
'*1.10 26/7/2022   Modify Imports
'*1.11 29/7/2022   Modify Imports,mGetLine
'**********************************
Imports PigToolsLiteLib

Public Class PigConsole
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.11.3"
    Private moPigFunc As New PigFunc

    Public Enum EnmSimpleMenuExitType
        Null = 0
        QtoExit = 1
        QtoUp = 2
    End Enum

    Friend Enum EnmLineEdit
        Insert = 0
        Delete = 1
        Backspace = 2
    End Enum

    Friend Enum EnmKeyTypeForLine
        ''' <summary>
        ''' 识别输入文本的字符类型|Identifies the character type of the input text
        ''' </summary>
        TextChar = 0
        Enter = 1
        ''' <summary>
        ''' Other characters
        ''' </summary>
        LeftArrow = 2
        RightArrow = 3
        Home = 4
        [End] = 5
        Backspace = 6
        Delete = 7
        OtherChar = 8
        Escape = 9
    End Enum

    Friend Enum EnmKeyTypeForPwd
        ''' <summary>
        ''' 识别输入密码的字符类型|Identify the character type of the input password
        ''' </summary>
        PasswordChar = 0
        Enter = 1
        ''' <summary>
        ''' Other characters
        ''' </summary>
        OtherChar = 2
        Backspace = 3
    End Enum

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Private Function mGetKeyTypeForLine(ByRef InConsoleKeyInfo As ConsoleKeyInfo) As EnmKeyTypeForLine
        mGetKeyTypeForLine = EnmKeyTypeForLine.OtherChar
        Select Case InConsoleKeyInfo.Key
            Case ConsoleKey.Backspace
            Case ConsoleKey.Tab
            Case ConsoleKey.Clear
            Case ConsoleKey.Enter
                mGetKeyTypeForLine = EnmKeyTypeForLine.Enter
            Case ConsoleKey.Pause
            Case ConsoleKey.Escape
                mGetKeyTypeForLine = EnmKeyTypeForLine.Escape
            Case ConsoleKey.Spacebar
                mGetKeyTypeForLine = EnmKeyTypeForLine.TextChar
            Case ConsoleKey.PageUp
            Case ConsoleKey.PageDown
            Case ConsoleKey.[End]
                mGetKeyTypeForLine = EnmKeyTypeForLine.End
            Case ConsoleKey.Home
                mGetKeyTypeForLine = EnmKeyTypeForLine.Home
            Case ConsoleKey.LeftArrow
                mGetKeyTypeForLine = EnmKeyTypeForLine.LeftArrow
            Case ConsoleKey.UpArrow
            Case ConsoleKey.RightArrow
                mGetKeyTypeForLine = EnmKeyTypeForLine.RightArrow
            Case ConsoleKey.DownArrow
            Case ConsoleKey.[Select]
            Case ConsoleKey.Print
            Case ConsoleKey.Execute
            Case ConsoleKey.PrintScreen
            Case ConsoleKey.Insert
            Case ConsoleKey.Delete
                mGetKeyTypeForLine = EnmKeyTypeForLine.Delete
            Case ConsoleKey.Help
            Case ConsoleKey.D0 To ConsoleKey.D9
                mGetKeyTypeForLine = EnmKeyTypeForLine.TextChar
            'Case ConsoleKey.D1
            'Case ConsoleKey.D2
            'Case ConsoleKey.D3
            'Case ConsoleKey.D4
            'Case ConsoleKey.D5
            'Case ConsoleKey.D6
            'Case ConsoleKey.D7
            'Case ConsoleKey.D8
            'Case ConsoleKey.D9
            Case ConsoleKey.A To ConsoleKey.Z
                mGetKeyTypeForLine = EnmKeyTypeForLine.TextChar
            'Case ConsoleKey.B
            'Case ConsoleKey.C
            'Case ConsoleKey.D
            'Case ConsoleKey.E
            'Case ConsoleKey.F
            'Case ConsoleKey.G
            'Case ConsoleKey.H
            'Case ConsoleKey.I
            'Case ConsoleKey.J
            'Case ConsoleKey.K
            'Case ConsoleKey.L
            'Case ConsoleKey.M
            'Case ConsoleKey.N
            'Case ConsoleKey.O
            'Case ConsoleKey.P
            'Case ConsoleKey.Q
            'Case ConsoleKey.R
            'Case ConsoleKey.S
            'Case ConsoleKey.T
            'Case ConsoleKey.U
            'Case ConsoleKey.V
            'Case ConsoleKey.W
            'Case ConsoleKey.X
            'Case ConsoleKey.Y
            'Case ConsoleKey.Z
            Case ConsoleKey.LeftWindows
            Case ConsoleKey.RightWindows
            Case ConsoleKey.Applications
            Case ConsoleKey.Sleep
            Case ConsoleKey.NumPad0 To ConsoleKey.NumPad9
                mGetKeyTypeForLine = EnmKeyTypeForLine.TextChar
            'Case ConsoleKey.NumPad1
            'Case ConsoleKey.NumPad2
            'Case ConsoleKey.NumPad3
            'Case ConsoleKey.NumPad4
            'Case ConsoleKey.NumPad5
            'Case ConsoleKey.NumPad6
            'Case ConsoleKey.NumPad7
            'Case ConsoleKey.NumPad8
            'Case ConsoleKey.NumPad9
            Case ConsoleKey.Multiply
            Case ConsoleKey.Add
            Case ConsoleKey.Separator
            Case ConsoleKey.Subtract
            Case ConsoleKey.[Decimal]
            Case ConsoleKey.Divide
            Case ConsoleKey.F1
            Case ConsoleKey.F2
            Case ConsoleKey.F3
            Case ConsoleKey.F4
            Case ConsoleKey.F5
            Case ConsoleKey.F6
            Case ConsoleKey.F7
            Case ConsoleKey.F8
            Case ConsoleKey.F9
            Case ConsoleKey.F10
            Case ConsoleKey.F11
            Case ConsoleKey.F12
            Case ConsoleKey.F13
            Case ConsoleKey.F14
            Case ConsoleKey.F15
            Case ConsoleKey.F16
            Case ConsoleKey.F17
            Case ConsoleKey.F18
            Case ConsoleKey.F19
            Case ConsoleKey.F20
            Case ConsoleKey.F21
            Case ConsoleKey.F22
            Case ConsoleKey.F23
            Case ConsoleKey.F24
            Case ConsoleKey.BrowserBack
            Case ConsoleKey.BrowserForward
            Case ConsoleKey.BrowserRefresh
            Case ConsoleKey.BrowserStop
            Case ConsoleKey.BrowserSearch
            Case ConsoleKey.BrowserFavorites
            Case ConsoleKey.BrowserHome
            Case ConsoleKey.VolumeMute
            Case ConsoleKey.VolumeDown
            Case ConsoleKey.VolumeUp
            Case ConsoleKey.MediaNext
            Case ConsoleKey.MediaPrevious
            Case ConsoleKey.MediaStop
            Case ConsoleKey.MediaPlay
            Case ConsoleKey.LaunchMail
            Case ConsoleKey.LaunchMediaSelect
            Case ConsoleKey.LaunchApp1
            Case ConsoleKey.LaunchApp2
            Case ConsoleKey.Oem1
            Case ConsoleKey.OemPlus
            Case ConsoleKey.OemComma
            Case ConsoleKey.OemMinus
            Case ConsoleKey.OemPeriod
            Case ConsoleKey.Oem2
            Case ConsoleKey.Oem3
            Case ConsoleKey.Oem4
            Case ConsoleKey.Oem5
            Case ConsoleKey.Oem6
            Case ConsoleKey.Oem7
            Case ConsoleKey.Oem8
            Case ConsoleKey.Oem102
            Case ConsoleKey.Process
            Case ConsoleKey.Packet
            Case ConsoleKey.Attention
            Case ConsoleKey.CrSel
            Case ConsoleKey.ExSel
            Case ConsoleKey.EraseEndOfFile
            Case ConsoleKey.Play
            Case ConsoleKey.Zoom
            Case ConsoleKey.NoName
            Case ConsoleKey.Pa1
            Case ConsoleKey.OemClear
            Case Else
                mGetKeyTypeForLine = EnmKeyTypeForLine.TextChar
        End Select
    End Function

    Private Function mGetKeyTypeForPwd(Key As Char) As EnmKeyTypeForPwd
        Select Case Key
            Case "a" To "z"
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "A" To "Z"
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "0" To "9"
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "`", "-", "=", "[", "]", "\", ";", "'", ".", "/", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "{", "}", "|", ":", "<", ">", " ", "?", """", ","
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case Chr(ConsoleKey.Backspace)
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.Backspace
            Case Chr(ConsoleKey.Enter)
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.Enter
            Case Else
                mGetKeyTypeForPwd = EnmKeyTypeForPwd.OtherChar
        End Select
    End Function

    ''' <summary>
    ''' 从控制台当前行读取密码|Read the password from the current line of the console
    ''' </summary>
    ''' <param name="PromptInf">提示信息|Prompt information</param>
    ''' <returns></returns>
    Public Function GetPwdStr(PromptInf As String) As String
        Return Me.mGetPwdStr(PromptInf)
    End Function

    ''' <summary>
    ''' 从控制台当前行读取密码|Read the password from the current line of the console
    ''' </summary>
    ''' <returns></returns>
    Public Function GetPwdStr() As String
        Return Me.mGetPwdStr("")
    End Function

    ''' <summary>
    ''' 从控制台当前行读取密码|Read the password from the current line of the console
    ''' </summary>
    ''' <returns></returns>
    Private Function mGetPwdStr(PromptInf As String) As String
        Try
            Dim strPwd As String = ""
            '            Dim bolCurrCursorVisible As Boolean
            'If Me.IsWindows = True Then
            '    bolCurrCursorVisible = Console.CursorVisible
            '    If bolCurrCursorVisible = False Then Console.CursorVisible = True
            'End If
            If PromptInf <> "" Then Console.Write(PromptInf & ":")
            Do While True
                Dim oConsoleKeyInfo As ConsoleKeyInfo = Console.ReadKey(True)
                Select Case Me.mGetKeyTypeForPwd(oConsoleKeyInfo.KeyChar)
                    Case EnmKeyTypeForPwd.PasswordChar
                        strPwd &= oConsoleKeyInfo.KeyChar
                        Console.Write("*")
                    Case EnmKeyTypeForPwd.Backspace
                        If strPwd <> "" Then
                            Console.Write(vbBack & " " & vbBack)
                            strPwd = Left(strPwd, Len(strPwd) - 1)
                        End If
                    Case EnmKeyTypeForPwd.Enter
                        Console.WriteLine()
                        Exit Do
                End Select
            Loop
            'If Me.IsWindows = True Then
            '    If bolCurrCursorVisible = Console.CursorVisible Then Console.CursorVisible = bolCurrCursorVisible
            'End If
            Return strPwd
        Catch ex As Exception
            Me.SetSubErrInf("mGetPwdStr", ex)
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' 从控制台当前行读取文本|Read text from the current line of the console
    ''' </summary>
    ''' <param name="PromptInf">提示信息|Prompt information</param>
    ''' <param name="OutLine">输出的文本|Output text</param>
    ''' <param name="IsShowCurrLine">是否显示当前文本|Is show current text</param>
    ''' <returns></returns>
    Public Function GetLine(PromptInf As String, ByRef OutLine As String, IsShowCurrLine As Boolean) As String
        Return Me.mGetLine(PromptInf, OutLine, IsShowCurrLine)
    End Function

    ''' <summary>
    ''' 从控制台当前行读取文本|Read text from the current line of the console
    ''' </summary>
    ''' <param name="PromptInf">提示信息|Prompt information</param>
    ''' <param name="OutLine">输出的文本，也作为默认文本|The output text is also used as the default text</param>
    ''' <returns></returns>
    Public Function GetLine(PromptInf As String, ByRef OutLine As String) As String
        Return Me.mGetLine(PromptInf, OutLine)
    End Function

    ''' <summary>
    ''' 从控制台当前行读取文本|Read text from the current line of the console
    ''' </summary>
    ''' <param name="OutLine">输出的文本，也作为默认文本|The output text is also used as the default text</param>
    ''' <returns></returns>
    Public Function GetLine(ByRef OutLine As String) As String
        Return Me.mGetLine("", OutLine)
    End Function


    ''' <summary>
    ''' 显示信息并暂停|Display message and pause
    ''' </summary>
    ''' <param name="DisplayInf">display information</param>
    ''' <returns></returns>
    Public Function DisplayPause(DisplayInf As String) As String
        Return Me.mDisplayPause(DisplayInf)
    End Function

    Public Function DisplayPause() As String
        Return Me.mDisplayPause("")
    End Function

    Private Function mDisplayPause(DisplayInf As String) As String
        Try
            Console.Write(DisplayInf & "(Press any key to continue)")
            Dim oConsoleKey As ConsoleKey = Console.ReadKey(True).Key
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mDisplayPause", ex)
        End Try
    End Function

    ''' <summary>
    ''' 选择是或否|Select Yes or no
    ''' </summary>
    ''' <param name="PromptInf">Prompt information</param>
    ''' <returns></returns>
    Public Function IsYesOrNo(PromptInf As String) As Boolean
        Try
            IsYesOrNo = Nothing
            Console.Write(PromptInf & ":(Press Y to Yes, N to No)")
            Do While True
                Select Case Console.ReadKey(True).Key
                    Case ConsoleKey.Y
                        IsYesOrNo = True
                        Exit Do
                    Case ConsoleKey.N
                        IsYesOrNo = False
                        Exit Do
                End Select
            Loop
        Catch ex As Exception
            Me.SetSubErrInf("IsYesOrNo", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 简单菜单|Simple Menu
    ''' </summary>
    ''' <param name="MenuTitle">菜单标题|Menu title</param>
    ''' <param name="MenuDefinition">菜单定义，格式：键值1#名称1|键值2#名称2|...|Menu definition, format:Key1#Name1|Key2#Name2|...</param>
    ''' <param name="OutMenuKey">选择的键值|Selected key value</param>
    ''' <param name="EnmSimpleMenuExitType">退出方式|Exit mode</param>
    ''' <returns></returns>
    Public Function SimpleMenu(MenuTitle As String, MenuDefinition As String, ByRef OutMenuKey As String, Optional MenuExitType As EnmSimpleMenuExitType = EnmSimpleMenuExitType.QtoExit) As String
        Dim LOG As New PigStepLog("SimpleMenu")
        Try
            Dim abMenuName As String(), abMenuKey As String(), abLetter As String()
            ReDim abMenuName(0)
            ReDim abMenuKey(0)
            ReDim abLetter(0)
            abMenuName(0) = MenuTitle
            Dim intMaxLen As Integer = Len(MenuTitle), intItems As Integer = 0
            LOG.StepName = "Handle MenuDefinition"
            If Right(MenuDefinition, 1) <> "|" Then MenuDefinition &= "|"
            Do While True
                Dim strLine As String = moPigFunc.GetStr(MenuDefinition, "", "|")
                If strLine = "" Then Exit Do
                intItems += 1
                ReDim Preserve abMenuKey(intItems)
                ReDim Preserve abMenuName(intItems)
                ReDim Preserve abLetter(intItems)
                abMenuKey(intItems) = moPigFunc.GetStr(strLine, "", "#")
                abMenuName(intItems) = strLine
                If Len(abMenuName(intItems)) > intMaxLen Then intMaxLen = Len(abMenuName(intItems))
                abLetter(intItems) = Chr(64 + intItems)
                If MenuExitType <> EnmSimpleMenuExitType.Null And abLetter(intItems) >= "Q" Then
                    abLetter(intItems) += Chr(64 + intItems + 1)
                End If
            Loop
            If intItems <= 0 Then
                Throw New Exception("No menu item defined")
            End If
            intMaxLen += 8
            LOG.StepName = "Print Menu"
            Dim strStarLine As String = moPigFunc.GetRepeatStr(intMaxLen, "*")
            Console.WriteLine(strStarLine)
            Console.WriteLine("* " & MenuTitle)
            Console.WriteLine(strStarLine)
            Select Case MenuExitType
                Case EnmSimpleMenuExitType.QtoExit
                    Console.WriteLine("* Q - Exit")
                Case EnmSimpleMenuExitType.QtoUp
                    Console.WriteLine("* Q - Up")
            End Select
            For i = 1 To intItems
                Dim strLetter As String = Chr(64 + i)
                If strLetter = "Q" Then strLetter = Chr(65 + i)
                Console.WriteLine("* " & strLetter & " - " & abMenuName(i))
            Next
            Console.WriteLine(strStarLine)
            LOG.StepName = "Select Menu"
            Do While True
                Dim oConsoleKey As ConsoleKey = Console.ReadKey(True).Key
                Select Case oConsoleKey
                    Case Asc(abLetter(1)) To Asc(abLetter(intItems)), ConsoleKey.Q
                        If oConsoleKey = ConsoleKey.Q And MenuExitType <> EnmSimpleMenuExitType.Null Then
                            OutMenuKey = ""
                        Else
                            Dim intKey As Integer = CInt(oConsoleKey) - 64
                            If MenuExitType <> EnmSimpleMenuExitType.Null And intKey > ConsoleKey.Q Then intKey -= 1
                            OutMenuKey = abMenuKey(intKey)
                        End If
                        Exit Do
                End Select
            Loop
            Return "OK"
        Catch ex As Exception
            OutMenuKey = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Function mGetLine(PromptInf As String, ByRef OutLine As String, Optional IsShowCurrLine As Boolean = True) As String
        Dim LOG As New PigStepLog("mGetLine")
        Try
            Dim strOldLine As String = OutLine
            '            Dim bolCurrCursorVisible As Boolean
            'If Me.IsWindows = True Then
            '    LOG.StepName = "Save CursorVisible"
            '    bolCurrCursorVisible = Console.CursorVisible
            '    If bolCurrCursorVisible = False Then Console.CursorVisible = True
            'End If
            Dim intBeginLeft As Integer = Console.CursorLeft, intBeginTop As Integer = Console.CursorTop
            If OutLine <> "" And IsShowCurrLine = True Then
                Console.WriteLine("(Press ENTER to set the current value to the following text)")
                Console.WriteLine(OutLine)
            End If
            If PromptInf <> "" Then Console.Write(PromptInf & ":")
            Dim strLine As String = Console.ReadLine
            If strLine <> "" Then
                OutLine = strLine
            End If
            'If Me.IsWindows = True Then
            '    If bolCurrCursorVisible = Console.CursorVisible Then Console.CursorVisible = bolCurrCursorVisible
            'End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ClearLine(LineLen As Integer, BeginLeft As Integer, BeginTop As Integer) As String
        Try
            Console.SetCursorPosition(BeginLeft, BeginTop)
            For i = 1 To LineLen
                Console.Write(" ")
            Next
            Console.SetCursorPosition(BeginLeft, BeginTop)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ClearLine", ex)
        End Try
    End Function

    Private Function mGetLinePos(BeginTop As Integer, NowLeft As Integer, NowTop As Integer, LineLen As Integer, ByRef OutLinePos As Integer, ByRef OutNowTop As Integer) As String
        Try
            OutLinePos = (NowTop - BeginTop) * Console.BufferWidth + NowLeft
            If OutLinePos <= 0 Then
                OutLinePos = 0
            End If
            If OutLinePos > LineLen Then
                OutLinePos = LineLen - 1
                OutNowTop = NowTop - 1
            End If
            Return "OK"
        Catch ex As Exception
            OutLinePos = 0
            Return Me.GetSubErrInf("mGetLinePos", ex)
        End Try
    End Function

    Private Function mEditLine(ByRef LineStr As String, BeginTop As Integer, LineEdit As EnmLineEdit, Optional InsChar As Char = "") As String
        Dim LOG As New PigStepLog("mEditLine")
        Try
            Dim intLinePos As Integer, intNowLeft As Integer = Console.CursorLeft, intNowTop As Integer = Console.CursorTop, intLineLen As Integer = Len(LineStr)
            Dim intOutNowTop As Integer = intNowTop
            If LineStr Is Nothing Then LineStr = ""
            LOG.StepName = "mGetLinePos"
            LOG.Ret = Me.mGetLinePos(BeginTop, intNowLeft, intNowTop, intLineLen, intLinePos, intOutNowTop)
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
            If intNowTop <> intNowTop Then
                intNowTop = intOutNowTop
                Console.SetCursorPosition(Console.CursorLeft, intNowTop)
            End If
            LOG.StepName = "Left and Right Str"
            Dim strLeft As String = "", strRight As String = ""
            Select Case LineEdit
                Case EnmLineEdit.Backspace
                    If intLinePos >= 1 Then strLeft = Left(LineStr, intLinePos - 1)
                Case Else
                    If intLinePos >= 0 Then strLeft = Left(LineStr, intLinePos)
            End Select
            Select Case LineEdit
                Case EnmLineEdit.Delete
                    If intLinePos < intLineLen And intLineLen > 0 Then
                        strRight = Mid(LineStr, intLineLen - intLinePos - 1)
                    End If
                Case Else
                    If intLinePos < intLineLen And intLineLen > 0 Then
                        strRight = Right(LineStr, intLineLen - intLinePos)
                    End If
            End Select
            Select Case LineEdit
                Case EnmLineEdit.Backspace
                    If intLinePos > 0 Then
                        intLinePos -= 1
                    Else

                    End If
                Case EnmLineEdit.Delete

                Case EnmLineEdit.Insert
                    Console.Write(InsChar)
                    If strRight <> "" Then Console.Write(strRight)
                    intLineLen += 1
                    intLinePos += 1
                    LOG.StepName = "mSetLinePos"
                    LOG.Ret = Me.mSetLinePos(intLinePos, intLineLen, BeginTop)
                    If LOG.Ret <> "OK" Then
                        Throw New Exception(LOG.Ret)
                    End If
                    LineStr = strLeft & InsChar & strRight
            End Select
            Select Case intLinePos
                Case < 0
                    intLinePos = 0
                Case 0 To intLineLen
                Case > intLineLen
                    intLinePos = intLineLen
            End Select

            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mSetLinePos(LinePos As Integer, LineLen As Integer, BeginTop As Integer) As String
        Try
            If LineLen < 0 Then Throw New Exception("LineLen < 0")
            If LinePos > LineLen Then Throw New Exception("LinePos > LineLen")
            If LinePos < 0 Then Throw New Exception("LinePos < 0")
            Dim intTopAdd As Integer = CInt(LineLen / Console.BufferWidth)
            Dim intLinePos As Integer = LinePos Mod Console.BufferWidth
            Console.SetCursorPosition(intLinePos, Console.CursorTop + intTopAdd)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSetLinePos", ex)
        End Try
    End Function

End Class
