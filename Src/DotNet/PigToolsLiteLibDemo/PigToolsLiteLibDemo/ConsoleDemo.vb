'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6.2
'* Create Time: 16/10/2021
'* 1.1    21/12/2021   Add PigConfig
'* 1.2    22/12/2021   Modify PigConfig
'* 1.3    26/12/2021   Modify PigConfig
'* 1.4    2/1/2022   Modify PigConfig
'* 1.5    3/2/2022   Modify PigConfig, Main, PigFunc demo,PigCompressorDemo
'* 1.6    22/2/2022   Modify PigConfig
'************************************
Imports PigToolsLiteLib

Public Class ConsoleDemo
    Public ShareMem As ShareMem
    Public SMName As String
    Public SMSize As Integer = 1024
    Public MainText As String
    Public Ret As String
    Public Base64EncKey As String
    Public PigConfigApp As PigConfigApp
    Public PigConfigSession As PigConfigSession
    Public PigConfig As PigConfig
    Public ConfName As String
    Public ConfValue As String
    Public SessionName As String
    Public SessionDesc As String
    Public ConfDesc As String
    Public ConfData As String
    Public FilePath As String
    Public FilePart As PigFunc.enmFilePart
    Public Line As String
    Public SaveType As PigConfigApp.EnmSaveType
    Public PigFunc As New PigFunc
    Public Url As String
    Public SrcStr As String
    Public Base64EncStr As String
    Public LeftStr As String
    Public RightStr As String
    Public EnvVar As String
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
            Console.WriteLine("Press J to PigConfig")
            Console.WriteLine("Press K to PigXml")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Me.PigCompressorDemo()
                Case ConsoleKey.B
                    Me.PigTextDemo()
                Case ConsoleKey.C
                    Console.CursorVisible = True
                    Console.WriteLine("Input FilePath:" & Me.FilePath)
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.FilePath = Me.Line
                    End If
                    Me.PigFileDemo(Me.FilePath)
                Case ConsoleKey.D
                    Me.PigFuncDemo()
                Case ConsoleKey.E
                    Console.CursorVisible = True
                    Console.WriteLine("Input Url:" & Me.Url)
                    Me.Line = Console.ReadLine
                    If Me.Line <> "" Then
                        Me.Url = Me.Line
                    End If
                    Me.PigWebReqDemo(Me.Url)
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
                                Console.CursorVisible = True
                                Console.WriteLine("Input ShareMem:" & Me.SMName)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.SMName = Me.Line
                                End If
                                Console.WriteLine("Input MainText:" & Me.MainText)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.MainText = Me.Line
                                End If
                                Dim oPigText As New PigText(Me.MainText, PigText.enmTextType.UTF8)
                                Me.ShareMem = New ShareMem
                                With Me.ShareMem
                                    Console.WriteLine(" " & Me.SMName & ",Len=" & oPigText.TextBytes.Length)
                                    Me.Ret = .Init(Me.SMName, oPigText.TextBytes.Length)
                                    Console.WriteLine(Me.Ret)
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
                    Console.CursorVisible = False
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to MkEncKey")
                        Console.WriteLine("Press B to LoadEncKey")
                        Console.WriteLine("Press C to Encrypt")
                        Console.WriteLine("Press D to Decrypt")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Me.Ret = oPigAes.MkEncKey(Me.Base64EncKey)
                                Console.WriteLine("MkEncKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Console.WriteLine("Inupt Base64EncKey:" & vbCrLf & Me.Base64EncKey)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncKey = Me.Line
                                End If
                                Me.Ret = oPigAes.LoadEncKey(Me.Base64EncKey)
                                Console.WriteLine("LoadEncKey=" & Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:" & Me.SrcStr)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.SrcStr = Me.Line
                                End If
                                Dim oPigText As New PigText(Me.SrcStr, PigText.enmTextType.UTF8)
                                Me.Ret = oPigAes.Encrypt(oPigText.TextBytes, Me.Base64EncStr)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & Me.Base64EncStr)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:")
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncStr = Me.Line
                                End If
                                Me.SrcStr = ""
                                Me.Ret = oPigAes.Decrypt(Me.Base64EncStr, Me.SrcStr, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & Me.SrcStr)
                        End Select
                    Loop
                Case ConsoleKey.I
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigRsa")
                    Console.WriteLine("*******************")
                    Dim oPigRsa As New PigRsa
                    Console.CursorVisible = False
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
                                Me.Ret = oPigRsa.MkPubKey(True, Me.Base64EncKey)
                                Me.Base64EncKey = Me.Base64EncKey
                                Console.WriteLine("MkPubKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Console.CursorVisible = True
                                Console.WriteLine("Inupt Base64EncKey:" & vbCrLf & Me.Base64EncKey)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncKey = Me.Line
                                End If
                                Me.Ret = oPigRsa.LoadPubKey(Me.Base64EncKey)
                                Console.WriteLine("LoadPubKey=" & Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:")
                                Me.SrcStr = Console.ReadLine
                                Dim oPigText As New PigText(Me.SrcStr, PigText.enmTextType.UTF8)
                                Me.Ret = oPigRsa.Encrypt(oPigText.TextBytes, Me.Base64EncStr)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & Me.Base64EncStr)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:" & Me.Base64EncKey)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncKey = Me.Line
                                End If
                                Me.SrcStr = ""
                                Me.Ret = oPigRsa.Decrypt(Me.Base64EncStr, Me.SrcStr, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & Me.SrcStr)
                        End Select
                    Loop
                Case ConsoleKey.J
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigConfig")
                    Console.WriteLine("*******************")
                    Dim oPigRsa As New PigRsa
                    Dim strBase64EncKey As String = ""
                    Dim strSrc As String = ""
                    Dim strBase64EncStr As String = ""
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to New")
                        Console.WriteLine("Press B to MkEncKey")
                        Console.WriteLine("Press C to GetEncStr")
                        Console.WriteLine("Press D to AddOrGet ConfigSession")
                        Console.WriteLine("Press E to AddOrGet ConfigItem")
                        Console.WriteLine("Press F to SaveConfig and SaveConfigFile")
                        Console.WriteLine("Press G to LoadConfig")
                        Console.WriteLine("Press H to LoadConfigFile")
                        Console.WriteLine("Press I to ShowConfig")
                        Console.WriteLine("Press J to Edit ConfigItem")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.CursorVisible = True
                                Console.WriteLine("EncKey=")
                                Me.Line = Console.ReadLine
                                Me.PigConfigApp = New PigConfigApp(Me.Line)
                                Console.WriteLine("New PigConfigApp=")
                                If Me.PigConfigApp.LastErr = "" Then
                                    Console.WriteLine("OK")
                                    Me.PigConfigApp.OpenDebug()
                                Else
                                    Console.WriteLine(Me.PigConfigApp.LastErr)
                                End If
                            Case ConsoleKey.B
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("PigConfigApp.MkEncKey=")
                                    Me.Ret = Me.PigConfigApp.MkEncKey(Me.Base64EncKey)
                                    Console.WriteLine(Me.Ret)
                                    If Me.Ret = "OK" Then
                                        Console.WriteLine(Me.Base64EncKey)
                                    End If
                                End If
                            Case ConsoleKey.C
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("Input source string=")
                                    Me.MainText = Console.ReadLine
                                    Console.WriteLine("GetEncStr=")
                                    Me.Ret = Me.PigConfigApp.GetEncStr(Me.MainText)
                                    Console.WriteLine(Me.Ret)
                                End If
                            Case ConsoleKey.D
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("SessionName=")
                                    Me.SessionName = Console.ReadLine
                                    Console.WriteLine("SessionDesc=")
                                    Me.SessionDesc = Console.ReadLine
                                    Me.PigConfigSession = Me.PigConfigApp.PigConfigSessions.AddOrGet(Me.SessionName, Me.SessionDesc)
                                    If Me.PigConfigApp.PigConfigSessions.LastErr <> "" Then
                                        Console.WriteLine(Me.PigConfigApp.PigConfigSessions.LastErr)
                                    ElseIf Me.PigConfigSession Is Nothing Then
                                        Console.WriteLine("PigConfigSession Is Nothing")
                                    Else
                                        Console.WriteLine("OK")
                                    End If
                                End If
                            Case ConsoleKey.E
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("SessionName=" & Me.SessionName)
                                    Me.Line = Console.ReadLine
                                    If Me.Line <> "" Then Me.SessionName = Me.Line
                                    Me.PigConfigSession = Me.PigConfigApp.GetPigConfigSession(Me.SessionName)
                                    If Me.PigConfigApp.LastErr <> "" Then
                                        Console.WriteLine(Me.PigConfigApp.LastErr)
                                    ElseIf Me.PigConfigSession Is Nothing Then
                                        Console.WriteLine("PigConfigSession Is Nothing")
                                    Else
                                        Console.WriteLine("OK")
                                        Console.WriteLine("ConfName=" & Me.ConfName)
                                        Me.ConfName = Console.ReadLine
                                        Console.WriteLine("ConfValue=" & Me.ConfValue)
                                        Me.ConfValue = Console.ReadLine
                                        Console.WriteLine("ConfDesc=" & Me.ConfDesc)
                                        Me.ConfDesc = Console.ReadLine
                                        Me.PigConfig = Me.PigConfigSession.PigConfigs.AddOrGet(Me.ConfName, Me.ConfValue, Me.ConfDesc)
                                        If Me.PigConfigSession.PigConfigs.LastErr <> "" Then
                                            Console.WriteLine(Me.PigConfigSession.PigConfigs.LastErr)
                                        ElseIf Me.PigConfig Is Nothing Then
                                            Console.WriteLine("PigConfig Is Nothing")
                                        Else
                                            Console.WriteLine("OK")
                                        End If
                                    End If
                                End If
                            Case ConsoleKey.F
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("PigConfigApp.SaveConfig")
                                    Console.WriteLine("SaveType(" & PigConfigApp.EnmSaveType.Xml.ToString & "=" & PigConfigApp.EnmSaveType.Xml & "," & PigConfigApp.EnmSaveType.Ini.ToString & "=" & PigConfigApp.EnmSaveType.Ini & ")=" & Me.SaveType.ToString)
                                    Me.SaveType = Console.ReadLine
                                    Me.Ret = Me.PigConfigApp.SaveConfig(Me.ConfData, Me.SaveType)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("ConfData=")
                                    Console.WriteLine(Me.ConfData)
                                    Console.WriteLine("PigConfigApp.SaveConfigFile")
                                    Console.WriteLine("FilePath=" & Me.FilePath)
                                    Me.FilePath = Console.ReadLine
                                    Me.Ret = Me.PigConfigApp.SaveConfigFile(Me.FilePath, Me.SaveType)
                                    Console.WriteLine(Me.Ret)
                                End If
                            Case ConsoleKey.G
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("SaveType(" & PigConfigApp.EnmSaveType.Xml.ToString & "=" & PigConfigApp.EnmSaveType.Xml & "," & PigConfigApp.EnmSaveType.Ini.ToString & "=" & PigConfigApp.EnmSaveType.Ini & ")=" & Me.SaveType.ToString)
                                    Me.SaveType = Console.ReadLine
                                    Console.WriteLine("ConfData=" & Me.ConfData)
                                    Me.Line = Console.ReadLine
                                    If Me.Line <> "" Then Me.FilePath = Me.Line
                                    Console.WriteLine("PigConfigApp.LoadConfig")
                                    Me.Ret = Me.PigConfigApp.LoadConfig(Me.ConfData, Me.SaveType)
                                    Console.WriteLine(Me.Ret)
                                End If
                            Case ConsoleKey.H
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("SaveType(" & PigConfigApp.EnmSaveType.Xml.ToString & "=" & PigConfigApp.EnmSaveType.Xml & "," & PigConfigApp.EnmSaveType.Ini.ToString & "=" & PigConfigApp.EnmSaveType.Ini & ")=" & Me.SaveType.ToString)
                                    Me.SaveType = Console.ReadLine
                                    Console.WriteLine("FilePath=" & Me.FilePath)
                                    Me.Line = Console.ReadLine
                                    If Me.Line <> "" Then Me.FilePath = Me.Line
                                    Console.WriteLine("PigConfigApp.LoadConfigFile")
                                    Me.Ret = Me.PigConfigApp.LoadConfigFile(Me.FilePath, Me.SaveType)
                                    Console.WriteLine(Me.Ret)
                                End If
                            Case ConsoleKey.I
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("PigConfigApp.IsChange=" & Me.PigConfigApp.IsChange)
                                    For Each oPigConfigSession As PigConfigSession In Me.PigConfigApp.PigConfigSessions
                                        Console.WriteLine("SessionName=" & oPigConfigSession.SessionName)
                                        Console.WriteLine("SessionDesc=" & oPigConfigSession.SessionDesc)
                                        For Each oPigConfig As PigConfig In oPigConfigSession.PigConfigs
                                            With oPigConfig
                                                Console.WriteLine("ConfName=" & .ConfName)
                                                Console.WriteLine("ConfValue=" & .ConfValue)
                                                Console.WriteLine("ConfDesc=" & .ConfDesc)
                                                Console.WriteLine("ContentType=" & .ContentType.ToString)
                                            End With
                                        Next
                                    Next
                                End If
                            Case ConsoleKey.J
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("SessionName=" & Me.SessionName)
                                    Me.Line = Console.ReadLine
                                    If Me.Line <> "" Then Me.SessionName = Me.Line
                                    Me.PigConfigSession = Me.PigConfigApp.GetPigConfigSession(Me.SessionName)
                                    If Me.PigConfigApp.LastErr <> "" Then
                                        Console.WriteLine(Me.PigConfigApp.LastErr)
                                    ElseIf Me.PigConfigSession Is Nothing Then
                                        Console.WriteLine("PigConfigSession Is Nothing")
                                    Else
                                        Console.WriteLine("OK")
                                        Console.WriteLine("ConfName=" & Me.ConfName)
                                        Me.ConfName = Console.ReadLine
                                        Console.WriteLine("ConfValue=" & Me.ConfValue)
                                        Me.ConfValue = Console.ReadLine
                                        Console.WriteLine("ConfDesc=" & Me.ConfDesc)
                                        Me.ConfDesc = Console.ReadLine
                                        With Me.PigConfigSession.PigConfigs.Item(ConfName)
                                            .ConfValue = Me.ConfValue
                                            .ConfDesc = Me.ConfDesc
                                        End With
                                    End If
                                End If
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
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("PigFunc")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to Show Function Demo")
            Console.WriteLine("Press B to GetFilePart")
            Console.WriteLine("Press C to GetFileText")
            Console.WriteLine("Press D to SaveTextToFile")
            Console.WriteLine("Press E to GetEnvVar")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    With Me.PigFunc
                        Console.WriteLine("GENow=" & .GENow)
                        Console.WriteLine("GetHostName=" & .GetHostName)
                        Console.WriteLine("GetHostIpList=" & .GetHostIpList)
                        Console.WriteLine("GetHostIpList(True)=" & .GetHostIpList(True))
                        Console.WriteLine("GetHostIp=" & .GetHostIp)
                        Console.WriteLine("GetHostIp(True)=" & .GetHostIp(True))
                        Console.WriteLine("GetHostIp(False, ""169.254.79."")=" & .GetHostIp(False, "169.254.79."))
                        Console.WriteLine("GetUserName=" & .GetUserName)
                        Console.WriteLine("GetEnvVar(""Path"")=" & .GetEnvVar("Path"))
                        'Console.WriteLine("GetFilePart(""c: \temp\aaa"", PigFunc.enmFilePart.FileTitle)=" & .GetFilePart("c:\temp\aaa", PigFunc.enmFilePart.FileTitle))
                        Console.WriteLine("GetProcThreadID=" & .GetProcThreadID)
                        Console.WriteLine("GetRandNum(1, 100)=" & .GetRandNum(1, 100))
                        Console.WriteLine("GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter)=" & .GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter))
                        Console.WriteLine(".GetRateDesc(16.88)=" & .GetRateDesc(16.88))
                        Console.WriteLine(".GetStr(""<hi>"", ""<"", "">"")=" & .GetStr("<hi>", "<", ">"))
                        Console.WriteLine(".IsRegexMatch(""B2b"", ""^[A-Za-z0-9]+$"")=" & .IsRegexMatch("B2b", "^[A-Za-z0-9]+$"))
                        Console.WriteLine(".GetProcThreadID=" & .GetProcThreadID)
                        Console.WriteLine(".UrlEncode=" & .UrlEncode("https://www.seowphong.com/oss/PigTools"))
                        Console.WriteLine(".UrlDecode=" & .UrlDecode("https%3A%2F%2Fwww.seowphong.com%2Foss%2FPigTools"))
                    End With
                Case ConsoleKey.B
                    Console.CursorVisible = True
                    Me.GetLine("FilePath", Me.FilePath)
                    Console.WriteLine("FilePart=DriveNo-" & PigFunc.enmFilePart.DriveNo & ",ExtName-" & PigFunc.enmFilePart.ExtName & ",FileTitle-" & PigFunc.enmFilePart.FileTitle & ",Path-" & PigFunc.enmFilePart.Path)
                    Me.Line = Console.ReadLine
                    Select Case Me.Line
                        Case PigFunc.enmFilePart.DriveNo, PigFunc.enmFilePart.ExtName, PigFunc.enmFilePart.FileTitle, PigFunc.enmFilePart.Path
                            Me.FilePart = CInt(Me.Line)
                            Console.WriteLine("GetFilePart(" & Me.FilePath & "," & Me.FilePart & ")=" & Me.PigFunc.GetFilePart(Me.FilePath, Me.FilePart))
                        Case Else
                            Console.WriteLine("Invalid FilePart")
                    End Select
                Case ConsoleKey.C
                    Console.CursorVisible = True
                    Me.GetLine("FilePath", Me.FilePath)
                    Me.Ret = Me.PigFunc.GetFileText(Me.FilePath, Me.SrcStr)
                    Console.WriteLine("GetFileText(" & Me.FilePath & ")=" & Me.Ret)
                    Console.WriteLine("FileText=" & Me.SrcStr)
                Case ConsoleKey.D
                    Console.CursorVisible = True
                    Me.GetLine("FilePath", Me.FilePath)
                    Me.GetLine("SaveText", Me.SrcStr)
                    Me.Ret = Me.PigFunc.SaveTextToFile(Me.FilePath, Me.SrcStr)
                    Console.WriteLine("SaveTextToFile(" & Me.FilePath & ")=" & Me.Ret)
                Case ConsoleKey.E
                    Console.CursorVisible = True
                    Me.GetLine("EnvVarName", Me.EnvVar)
                    Me.SrcStr = Me.PigFunc.GetEnvVar(Me.EnvVar)
                    Console.WriteLine("GetEnvVar(" & Me.EnvVar & ")=" & Me.SrcStr)
            End Select
        Loop
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
        oPigFile.LoadFile()
        Dim o As New PigText(oPigFile.GbMain.Main, PigText.enmTextType.Ascii)
        Console.WriteLine(o.Text)
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

    Public Sub GetLine(ShowInf As String, ByRef OutLine As String)
        Console.WriteLine("Input " & ShowInf & " : " & OutLine)
        Dim strLine As String = Console.ReadLine
        If strLine <> "" Then
            OutLine = strLine
        End If
    End Sub
    Public Sub PigCompressorDemo()
        Console.CursorVisible = True
        Me.GetLine("SrcStr", Me.SrcStr)
        Console.WriteLine("SrcStr.Length=" & Me.SrcStr.Length)
        Dim oPigCompressor As New PigCompressor
        Dim abSrc(0) As Byte
        Dim abCompressor(0) As Byte
        With oPigCompressor
            Console.WriteLine("abCompressor = .Compress(Me.SrcStr)")
            abCompressor = .Compress(Me.SrcStr, PigText.enmTextType.UTF8)
            If .LastErr = "" Then
                Console.WriteLine("OK")
                Console.WriteLine("abCompressor.Length=" & abCompressor.Length)
                Console.WriteLine("abSrc = .Depress(abCompressor)")
                abSrc = .Depress(abCompressor)
                If .LastErr = "" Then
                    Console.WriteLine("OK")
                    Console.WriteLine("abSrc.Length=" & abSrc.Length)
                    Console.WriteLine("Dim oPigText As New PigText(abSrc, PigText.enmTextType.UTF8)")
                    Dim oPigText As New PigText(abSrc, PigText.enmTextType.UTF8)
                    If oPigText.LastErr = "" Then
                        Console.WriteLine("OK")
                        Console.WriteLine("oPigText.Text=" & oPigText.Text)
                        If oPigText.Text = Me.SrcStr Then
                            Console.WriteLine("oPigText.Text = Me.SrcStr")
                        Else
                            Console.WriteLine("oPigText.Text <> Me.SrcStr")
                        End If
                    Else
                        Console.WriteLine(oPigText.LastErr)
                    End If
                Else
                    Console.WriteLine(.LastErr)
                End If
            Else
                Console.WriteLine(.LastErr)
            End If
        End With
    End Sub

End Class
