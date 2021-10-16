'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 16/10/2021
'************************************
Imports PigToolsLiteLib

Public Class ConsoleDemo
    Public ShareMem As ShareMem
    Public SMName As String
    Public SMSize As Integer = 1024
    Public MainText As String
    Public Ret As String
    Public Base64EncKey As String
    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to PigCompressor")
            Console.WriteLine("Press B to PigText")
            Console.WriteLine("Press C to PigFile")
            Console.WriteLine("Press D to PigFunc")
            Console.WriteLine("Press E to PigWebReq")
            Console.WriteLine("Press F to ShareMem")
            Console.WriteLine("Press G to PigBytes")
            Console.WriteLine("Press H to PigAes")
            Console.WriteLine("Press I to PigRsa")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Me.PigCompressorDemo()
                Case ConsoleKey.B
                    Me.PigTextDemo()
                Case ConsoleKey.C
                    Dim strFilePath As String
                    Console.WriteLine("Input filepath:")
                    strFilePath = Console.ReadLine
                    Me.PigFileDemo(strFilePath)
                Case ConsoleKey.D
                    Me.PigFuncDemo()
                Case ConsoleKey.E
                    Console.WriteLine("Input Url:")
                    Dim strUrl As String = Console.ReadLine
                    Me.PigWebReqDemo(strUrl)
                Case ConsoleKey.F
                    Console.WriteLine("*******************")
                    Console.WriteLine("ShareMem")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Press Q to Up")
                    Console.WriteLine("Press A to Write")
                    Console.WriteLine("Press B to Read")
                    Console.WriteLine("*******************")
                    Do While True
                        Dim intConsoleKey As ConsoleKey = Console.ReadKey().Key
                        Select Case intConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine("Input ShareMem:")
                                Me.SMName = Console.ReadLine
                                Console.WriteLine("Input MainText:")
                                Me.MainText = Console.ReadLine
                                Dim oPigText As New PigText(Me.MainText, PigText.enmTextType.UTF8)
                                Me.ShareMem = New ShareMem
                                With Me.ShareMem
                                    Console.WriteLine(" " & Me.SMName & ",Len=" & oPigText.TextBytes.Length)
                                    .Init(Me.SMName, oPigText.TextBytes.Length)
                                    If .LastErr <> "" Then
                                        Console.WriteLine(.LastErr)
                                    Else
                                        Console.WriteLine("OK")
                                    End If
                                End With
                                If Me.ShareMem.IsInit = False Then
                                    Console.WriteLine("ShareMem Not Init")
                                Else
                                    With Me.ShareMem
                                        Console.WriteLine("Write")
                                        Me.Ret = .Write(oPigText.TextBytes)
                                        Console.WriteLine(Me.Ret)
                                    End With
                                End If
                                Exit Do
                            Case ConsoleKey.B
                                If Me.ShareMem Is Nothing Then Me.ShareMem = New ShareMem
                                If Me.ShareMem.IsInit = False Then
                                    Console.WriteLine("ShareMem Not Init")
                                    Console.WriteLine("Input ShareMem:")
                                    Me.SMName = Console.ReadLine
                                    Console.WriteLine("Input TextLen:")
                                    Dim intLen As Integer = CInt(Console.ReadLine)
                                    With Me.ShareMem
                                        Console.WriteLine("Init " & Me.SMName & ",Len=" & intLen.ToString)
                                        .Init(Me.SMName, intLen)
                                        If .LastErr <> "" Then
                                            Console.WriteLine(.LastErr)
                                        Else
                                            Console.WriteLine("OK")
                                        End If
                                    End With
                                End If
                                With Me.ShareMem
                                    Dim abAny(0) As Byte
                                    Console.WriteLine("Read")
                                    Me.Ret = .Read(abAny)
                                    If Me.Ret <> "OK" Then
                                        Console.WriteLine(Me.Ret)
                                    Else
                                        Console.WriteLine("OK")
                                        Console.WriteLine("New PigText")
                                        Dim oPigText As New PigText(abAny)
                                        If .LastErr <> "" Then
                                            Console.WriteLine(.LastErr)
                                        Else
                                            Console.WriteLine("OK")
                                            Console.WriteLine(oPigText.Text)
                                        End If
                                    End If
                                End With
                                Exit Do
                        End Select
                    Loop
                Case ConsoleKey.G
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigBytes")
                    Console.WriteLine("*******************")
                    Dim oPigBytes As New PigBytes
                    With oPigBytes
                        Console.WriteLine(".SetValue(CLng(1))")
                        .SetValue(CLng(1))
                        Console.WriteLine(".SetValue(CDate(Now))")
                        .SetValue(CDate(Now))
                        Console.WriteLine(".SetValue(CInt(2)")
                        .SetValue(CInt(2))
                        Console.WriteLine("Dim abAny(15) As Byte")
                        Console.WriteLine(".SetValue(abAny)")
                        Dim abAny(15) As Byte
                        .SetValue(abAny)
                        Console.WriteLine(".Main.Length=" & .Main.Length)
                    End With
                    With oPigBytes
                        Console.WriteLine(".SetValue(True)")
                        .SetValue(True)
                        Console.WriteLine(".SetValue(CInt(1))")
                        .SetValue(CInt(1))
                        Console.WriteLine(".SetValue(CDate(Now))")
                        .SetValue(CDate(Now))
                        Console.WriteLine(".SetValue(CDec(1.1))")
                        .SetValue(CDec(1.1))
                        Console.WriteLine("Dim abAny(9) As Byte")
                        Console.WriteLine(".SetValue(abAny)")
                        Dim abAny(9) As Byte
                        .SetValue(abAny)
                        Console.WriteLine(".Main.Length=" & .Main.Length)
                    End With
                    Dim oPigBytes2 As New PigBytes(oPigBytes.Main)
                    With oPigBytes2
                        Console.WriteLine(".Main.Length=" & .Main.Length)
                        Console.WriteLine("GetBooleanValue=" & .GetBooleanValue())
                        Console.WriteLine("GetInt32Value=" & .GetInt32Value())
                        Console.WriteLine("GetDateTimeValue=" & .GetDateTimeValue())
                        Console.WriteLine("GetDoubleValue=" & .GetDoubleValue())
                        Dim abAny() As Byte = .GetBytesValue(9)
                        Console.WriteLine("GetBytesValue=" & abAny.Length)
                        Console.WriteLine(".IsMatchBytes(oPigBytes.Main)=" & .IsMatchBytes(oPigBytes.Main))
                    End With
                Case ConsoleKey.H
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigAes")
                    Console.WriteLine("*******************")
                    Dim oPigAes As New PigAes
                    Dim strBase64EncKey As String = ""
                    Dim strSrc As String = ""
                    Dim strBase64EncStr As String = ""
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to MkEncKey")
                        Console.WriteLine("Press B to LoadEncKey")
                        Console.WriteLine("Press C to Encrypt")
                        Console.WriteLine("Press D to Decrypt")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Me.Ret = oPigAes.MkEncKey(strBase64EncKey)
                                Me.Base64EncKey = strBase64EncKey
                                Console.WriteLine("MkEncKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Console.WriteLine("Inupt Base64EncKey:" & vbCrLf & Me.Base64EncKey)
                                strBase64EncKey = Console.ReadLine
                                If strBase64EncKey = "" Then strBase64EncKey = Me.Base64EncKey
                                Me.Ret = oPigAes.LoadEncKey(strBase64EncKey)
                                Console.WriteLine("LoadEncKey=" & Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:")
                                strSrc = Console.ReadLine
                                Dim oPigText As New PigText(strSrc, PigText.enmTextType.UTF8)
                                Me.Ret = oPigAes.Encrypt(oPigText.TextBytes, strBase64EncStr)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & strBase64EncStr)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:")
                                strBase64EncStr = Console.ReadLine
                                strSrc = ""
                                Me.Ret = oPigAes.Decrypt(strBase64EncStr, strSrc, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & strSrc)
                        End Select
                    Loop
                Case ConsoleKey.I
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigRsa")
                    Console.WriteLine("*******************")
                    Dim oPigRsa As New PigRsa
                    Dim strBase64EncKey As String = ""
                    Dim strSrc As String = ""
                    Dim strBase64EncStr As String = ""
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to MkPubKey")
                        Console.WriteLine("Press B to LoadPubKey")
                        Console.WriteLine("Press C to Encrypt")
                        Console.WriteLine("Press D to Decrypt")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Me.Ret = oPigRsa.MkPubKey(True, strBase64EncKey)
                                Me.Base64EncKey = strBase64EncKey
                                Console.WriteLine("MkPubKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Console.WriteLine("Inupt Base64EncKey:" & vbCrLf & Me.Base64EncKey)
                                strBase64EncKey = Console.ReadLine
                                If strBase64EncKey = "" Then strBase64EncKey = Me.Base64EncKey
                                Me.Ret = oPigRsa.LoadPubKey(strBase64EncKey)
                                Console.WriteLine("LoadPubKey=" & Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:")
                                strSrc = Console.ReadLine
                                Dim oPigText As New PigText(strSrc, PigText.enmTextType.UTF8)
                                Me.Ret = oPigRsa.Encrypt(oPigText.TextBytes, strBase64EncStr)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & strBase64EncStr)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:")
                                strBase64EncStr = Console.ReadLine
                                strSrc = ""
                                Me.Ret = oPigRsa.Decrypt(strBase64EncStr, strSrc, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & strSrc)
                        End Select
                    Loop
                Case Else
                    Console.WriteLine("Coming soon...")
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub PigWebReqDemo(Url As String)
        Dim strDisplay As String = ""
        strDisplay &= vbCrLf & "***PigWebReqDemo Sample code***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        strDisplay &= "Dim oPigWebReq As New PigWebReq(Url)" & vbCrLf
        strDisplay &= "With oPigWebReq" & vbCrLf
        strDisplay &= vbTab & ".GetText()" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".LastErr="",.LastErr)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".ResString="",.ResString)" & vbCrLf
        strDisplay &= "End With" & vbCrLf
        strDisplay &= "```" & vbCrLf

        Dim oPigWebReq As New PigWebReq(Url)
        strDisplay &= vbCrLf & "***Return results***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        With oPigWebReq
            strDisplay &= ".GetText()" & vbCrLf
            .GetText()
            strDisplay &= ".LastErr" & .LastErr & vbCrLf
            strDisplay &= ".ResString" & .ResString & vbCrLf
        End With
        strDisplay &= "```" & vbCrLf
        Console.WriteLine(strDisplay)
    End Sub

    Public Sub PigFuncDemo()
        Dim strDisplay As String = ""
        strDisplay &= vbCrLf & "***PigFuncDemo Sample code***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        strDisplay &= "Dim oPigFunc As New PigFunc" & vbCrLf
        strDisplay &= "With oPigFunc" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GENow = "" & .GENow)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GetFilePart(""c:\temp\aaa"", PigFunc.enmFilePart.FileTitle)= & .GetFilePart(""c:\temp\aaa"", PigFunc.enmFilePart.FileTitle))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GetProcThreadID = "" & .GetProcThreadID)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GetRandNum(1, 100) = "" & .GetRandNum(1, 100))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter) = "" & GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".GetRateDesc(16.88)="" & .GetRateDesc(16.88))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".GetStr(""<hi>"", ""<"", "">"")="" & .GetStr(""<hi>"", ""<"", "">""))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".IsRegexMatch(""B2b"", ""^[A-Za-z0-9]+$"")="" & .IsRegexMatch(""B2b"", ""^[A-Za-z0-9]+$""))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""GetProcThreadID = "" & .GetProcThreadID)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".UrlEncode(""https://www.seowphong.com/oss/PigTools"")="" & .UrlEncode(""https://www.seowphong.com/oss/PigTools""))" & vbCrLf
        strDisplay &= vbTab & "Debug.Print("".UrlDecode(""https%3A%2F%2Fwww.seowphong.com%2Foss%2FPigTools"")="" & .UrlDecode(""https%3A%2F%2Fwww.seowphong.com%2Foss%2FPigTools""))" & vbCrLf
        strDisplay &= "End With" & vbCrLf
        strDisplay &= "```" & vbCrLf

        Dim strSrc As String = "PigTextDemo Sample code"
        Dim oPigFunc As New PigFunc
        strDisplay &= vbCrLf & "***Return results***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        With oPigFunc
            strDisplay &= "GENow=" & .GENow & vbCrLf
            strDisplay &= "GetFilePart(""c:\temp\aaa"", PigFunc.enmFilePart.FileTitle)=" & .GetFilePart("c:\temp\aaa", PigFunc.enmFilePart.FileTitle) & vbCrLf
            strDisplay &= "GetProcThreadID=" & .GetProcThreadID & vbCrLf
            strDisplay &= "GetRandNum(1, 100)=" & .GetRandNum(1, 100) & vbCrLf
            strDisplay &= "GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter)=" & .GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter) & vbCrLf
            strDisplay &= ".GetRateDesc(16.88)=" & .GetRateDesc(16.88) & vbCrLf
            strDisplay &= ".GetStr(""<hi>"", ""<"", "">"")=" & .GetStr("<hi>", "<", ">") & vbCrLf
            strDisplay &= ".IsRegexMatch(""B2b"", ""^[A-Za-z0-9]+$"")=" & .IsRegexMatch("B2b", "^[A-Za-z0-9]+$") & vbCrLf
            strDisplay &= ".GetProcThreadID=" & .GetProcThreadID & vbCrLf
            strDisplay &= ".UrlEncode=" & .UrlEncode("https://www.seowphong.com/oss/PigTools") & vbCrLf
            strDisplay &= ".UrlDecode=" & .UrlDecode("https%3A%2F%2Fwww.seowphong.com%2Foss%2FPigTools") & vbCrLf
        End With
        strDisplay &= "```" & vbCrLf
        Console.WriteLine(strDisplay)
    End Sub

    Public Sub PigFileDemo(FilePath As String)
        Dim strDisplay As String = ""
        strDisplay &= vbCrLf & "***PigFileDemo Sample code***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        strDisplay &= "Dim oPigFile As New PigFile(FilePath)" & vbCrLf
        strDisplay &= "With oPigText" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""CreationTime = "" & .CreationTime)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""UpdateTime = "" & .UpdateTime)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""FileTitle = "" & .FileTitle)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""ExtName = "" & .ExtName)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""FilePath = "" & .FilePath)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""DirPath = "" & .DirPath)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""Size = "" & .Size)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""MD5 = "" & .MD5)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""PigMD5 = "" & .PigMD5)" & vbCrLf
        strDisplay &= "End With" & vbCrLf
        strDisplay &= "```" & vbCrLf

        Dim strSrc As String = "PigTextDemo Sample code"
        Dim oPigFile As New PigFile(FilePath)
        strDisplay &= vbCrLf & "***Return results***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        With oPigFile
            If .LastErr = "" Then
                strDisplay &= "CreationTime=" & .CreationTime & vbCrLf
                strDisplay &= "UpdateTime=" & .UpdateTime & vbCrLf
                strDisplay &= "FileTitle=" & .FileTitle & vbCrLf
                strDisplay &= "ExtName=" & .ExtName & vbCrLf
                strDisplay &= "FilePath=" & .FilePath & vbCrLf
                strDisplay &= "DirPath=" & .DirPath & vbCrLf
                strDisplay &= "Size=" & .Size & vbCrLf
                strDisplay &= "MD5=" & .MD5 & vbCrLf
                strDisplay &= "PigMD5=" & .PigMD5 & vbCrLf
            Else
                strDisplay &= "LastErr=" & .LastErr & vbCrLf
            End If
        End With
        strDisplay &= "```" & vbCrLf
        Console.WriteLine(strDisplay)
    End Sub

    Public Sub PigTextDemo()
        Dim strDisplay As String = ""
        strDisplay &= vbCrLf & "***PigTextDemo Sample code***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        strDisplay &= "oPigText = New PigText(strSrc, PigText.enmTextType.UTF8)" & vbCrLf
        strDisplay &= "With oPigText" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""Base64 = "" & .Base64)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""HexStr = "" & .HexStr)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""MD5 = "" & .MD5)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""PigMD5 = "" & .PigMD5)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""SHA1 = "" & .SHA1)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""TextBytes.Length = "" & .TextBytes.Length)" & vbCrLf
        strDisplay &= "End With" & vbCrLf
        strDisplay &= "```" & vbCrLf

        Dim strSrc As String = "PigTextDemo Sample code"
        Dim oPigText As PigText
        oPigText = New PigText(strSrc, PigText.enmTextType.UTF8)
        strDisplay &= vbCrLf & "***Return results***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        With oPigText
            strDisplay &= "Text=" & .Text & vbCrLf
            strDisplay &= "Base64=" & .Base64 & vbCrLf
            strDisplay &= "HexStr=" & .HexStr & vbCrLf
            strDisplay &= "MD5=" & .MD5 & vbCrLf
            strDisplay &= "PigMD5=" & .PigMD5 & vbCrLf
            strDisplay &= "SHA1=" & .SHA1 & vbCrLf
            strDisplay &= "TextBytes.Length=" & .TextBytes.Length & vbCrLf
        End With
        strDisplay &= "```" & vbCrLf
        Console.WriteLine(strDisplay)
    End Sub

    Public Sub PigCompressorDemo()
        Dim strDisplay As String = ""
        strDisplay &= vbCrLf & "***PigCompressor Sample code***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        strDisplay &= "Dim oPigCompressor As New PigCompressor" & vbCrLf
        strDisplay &= "With oPigCompressor" & vbCrLf
        strDisplay &= vbTab & "abCompressor = .Compress(strSrc)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""strSrc.Length="" & strSrc.Length.ToString)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""abCompressor.Length="" & abCompressor.Length.ToString)" & vbCrLf
        strDisplay &= vbTab & "abSrc = .Depress(abCompressor)" & vbCrLf
        strDisplay &= vbTab & "Debug.Print(""abSrc.Length="" & abSrc.Length.ToString)" & vbCrLf
        strDisplay &= "End With" & vbCrLf
        strDisplay &= "```" & vbCrLf

        Dim oPigCompressor As New PigCompressor
        Dim strSrc As String = "11111111111111113333333333444444444444446666666666666ggggggggggdddddddddddddddddeeeeeeeeeeeeeeeeedsfffffffffffffeeeeeegggggggggghhhhhhhhh"
        Dim abSrc(-1) As Byte
        Dim abCompressor(-1) As Byte

        strDisplay &= vbCrLf & "***Return results***" & vbCrLf
        strDisplay &= "```" & vbCrLf
        With oPigCompressor
            abCompressor = .Compress(strSrc)
            If .LastErr <> "" Then
                strDisplay &= "Compress=" & .LastErr & vbCrLf
            Else
                strDisplay &= "strSrc.Length=" & strSrc.Length.ToString & vbCrLf
                strDisplay &= "abCompressor.Length=" & abCompressor.Length.ToString & vbCrLf
                abSrc = .Depress(abCompressor)
                If .LastErr <> "" Then
                    strDisplay &= "Depress=" & .LastErr & vbCrLf
                Else
                    strDisplay &= "abSrc.Length=" & abSrc.Length.ToString & vbCrLf
                End If
            End If
        End With
        strDisplay &= "```" & vbCrLf
        Console.WriteLine(strDisplay)
    End Sub

End Class
