Imports PigToolsLib

Public Class ConsoleDemo

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
            'Console.WriteLine("Press F to PigFile")
            'Console.WriteLine("Press G to PigMD5")
            'Console.WriteLine("Press H to PigXml")
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
