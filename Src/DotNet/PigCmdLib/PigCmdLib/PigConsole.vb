﻿'**********************************
'* Name: PigConsole
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 增加控制台的功能|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 15/1/2022
'*1.1 23/1/2022    Add GetKeyType1, modify GetPwdStr
'*1.2 3/2/2022     Add GetLine
'*1.3 4/2/2022     Add mGetKeyTypeForLine,ClearLine,mGetLine
'*1.4 5/2/2022     Modify GetLine.
'*1.5 6/2/2022     Modify mGetKeyTypeForLine,mGetKeyTypeForPwd,mGetLine
'**********************************
Imports PigToolsLiteLib

Public Class PigConsole
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.8"

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
    ''' <returns></returns>
    Public Function GetPwdStr() As String
        Try
            GetPwdStr = ""
            Dim bolCurrCursorVisible As Boolean
            If Me.IsWindows = True Then
                bolCurrCursorVisible = Console.CursorVisible
                If bolCurrCursorVisible = False Then Console.CursorVisible = True
            End If
            Do While True
                Dim oConsoleKeyInfo As ConsoleKeyInfo = Console.ReadKey(True)
                Select Case Me.mGetKeyTypeForPwd(oConsoleKeyInfo.KeyChar)
                    Case EnmKeyTypeForPwd.PasswordChar
                        GetPwdStr &= oConsoleKeyInfo.KeyChar
                        Console.Write("*")
                    Case EnmKeyTypeForPwd.Backspace
                        If GetPwdStr <> "" Then
                            Console.Write(vbBack & " " & vbBack)
                            GetPwdStr = Left(GetPwdStr, Len(GetPwdStr) - 1)
                        End If
                    Case EnmKeyTypeForPwd.Enter
                        Console.WriteLine()
                        Exit Do
                End Select
            Loop
            If Me.IsWindows = True Then
                If bolCurrCursorVisible = Console.CursorVisible Then Console.CursorVisible = bolCurrCursorVisible
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GetPwdStr", ex)
            Return ""
        End Try
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

    Private Function mGetLine(PromptInf As String, ByRef OutLine As String) As String
        Dim LOG As New PigStepLog("mGetLine")
        Try
            Dim strOldLine As String = OutLine
            Dim bolCurrCursorVisible As Boolean
            If Me.IsWindows = True Then
                LOG.StepName = "Save CursorVisible"
                bolCurrCursorVisible = Console.CursorVisible
                If bolCurrCursorVisible = False Then Console.CursorVisible = True
            End If
            If PromptInf <> "" Then Console.WriteLine(PromptInf)
            Dim intBeginLeft As Integer = Console.CursorLeft, intBeginTop As Integer = Console.CursorTop
            If OutLine <> "" Then Console.Write(OutLine)
            Do While True
                Dim intLineLen As Integer = Len(OutLine)
                Dim oConsoleKeyInfo As ConsoleKeyInfo = Console.ReadKey(True)
                Dim strLeft As String = "", strRight As String = ""
                Select Case Me.mGetKeyTypeForLine(oConsoleKeyInfo)
                    Case EnmKeyTypeForLine.TextChar
                        Me.mEditLine(OutLine, intBeginTop, EnmLineEdit.Insert, oConsoleKeyInfo.KeyChar)
                    Case EnmKeyTypeForLine.Backspace
                    Case EnmKeyTypeForLine.Home
                        Me.mSetLinePos(0, intLineLen, intBeginTop)
                    Case EnmKeyTypeForLine.End
                        Me.mSetLinePos(intLineLen, intLineLen, intBeginTop)
                    Case EnmKeyTypeForLine.Delete
                    Case EnmKeyTypeForLine.LeftArrow
                        If Console.CursorLeft > 0 Then Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop)
                    Case EnmKeyTypeForLine.RightArrow
                        If Console.CursorLeft < OutLine.Length - 1 Then Console.SetCursorPosition(Console.CursorLeft + 1, Console.CursorTop)
                    Case EnmKeyTypeForLine.Escape
                        OutLine = strOldLine
                        LOG.StepName = "ClearLine"
                        LOG.Ret = Me.ClearLine(OutLine.Length, 0, intBeginTop)
                        If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                        Console.Write(OutLine)
                    Case EnmKeyTypeForLine.Enter
                        Console.WriteLine()
                        Exit Do
                End Select
            Loop
            If Me.IsWindows = True Then
                If bolCurrCursorVisible = Console.CursorVisible Then Console.CursorVisible = bolCurrCursorVisible
            End If
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