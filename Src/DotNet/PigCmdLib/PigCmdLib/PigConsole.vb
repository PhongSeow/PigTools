'**********************************
'* Name: PigConsole
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 增加控制台的功能|Application of calling operating system commands
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 15/1/2022
'**********************************
Public Class PigConsole
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function GetPwdStr() As String
        Try
            GetPwdStr = ""
            Dim bolCurrCursorVisible As Boolean = Console.CursorVisible
            If bolCurrCursorVisible = False Then Console.CursorVisible = True
            Do While True
                Dim oConsoleKeyInfo As ConsoleKeyInfo = Console.ReadKey(True)
                If oConsoleKeyInfo.Key = ConsoleKey.Enter Then
                    Console.WriteLine()
                    Exit Do
                End If
                GetPwdStr &= oConsoleKeyInfo.Key.ToString
                Console.Write("*")
            Loop
            If bolCurrCursorVisible = Console.CursorVisible Then Console.CursorVisible = bolCurrCursorVisible
        Catch ex As Exception
            Me.SetSubErrInf("GetPwdStr", ex)
            Return ""
        End Try
    End Function
End Class
