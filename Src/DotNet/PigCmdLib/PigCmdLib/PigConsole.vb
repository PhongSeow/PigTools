'**********************************
'* Name: PigConsole
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 增加控制台的功能|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 15/1/2022
'*1.1 23/1/2022     Add GetKeyType1, modify GetPwdStr
'**********************************
Public Class PigConsole
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.9"

    Public Enum EnmKeyTypeForPwd
        ''' <summary>
        ''' Characters that can be used as passwords
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

    Public Function GetKeyTypeForPwd(Key As Char) As EnmKeyTypeForPwd
        Select Case Key
            Case "a" To "z"
                GetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "A" To "Z"
                GetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "0" To "9"
                GetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case "`", "-", "=", "[", "]", "\", ";", "'", ".", "/", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "{", "}", "|", ":", "<", ">", " ", "?", """", ","
                GetKeyTypeForPwd = EnmKeyTypeForPwd.PasswordChar
            Case Chr(ConsoleKey.Backspace)
                GetKeyTypeForPwd = EnmKeyTypeForPwd.Backspace
            Case Chr(ConsoleKey.Enter)
                GetKeyTypeForPwd = EnmKeyTypeForPwd.Enter
            Case Else
                GetKeyTypeForPwd = EnmKeyTypeForPwd.OtherChar
        End Select
    End Function

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
                Select Case Me.GetKeyTypeForPwd(oConsoleKeyInfo.KeyChar)
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
End Class
