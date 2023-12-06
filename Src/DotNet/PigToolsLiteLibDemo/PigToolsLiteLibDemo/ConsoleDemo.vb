'**********************************
'* Name: ConsoleDemo
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.32.2
'* Create Time: 16/10/2021
'* 1.1    21/12/2021   Add PigConfig
'* 1.2    22/12/2021   Modify PigConfig
'* 1.3    26/12/2021   Modify PigConfig
'* 1.4    2/1/2022   Modify PigConfig
'* 1.5    3/2/2022   Modify PigConfig, Main, PigFunc demo,PigCompressorDemo
'* 1.6    22/2/2022   Modify PigConfig
'* 1.7    23/2/2022   Add PLSqlCsv2Bcp
'* 1.8    20/3/2022   Add PigCmdLib.PigConsole
'* 1.9    26/3/2022   Add PigProcApp
'* 1.10   2/4/2022   Modify PigProcDemo, add GetHumanSize
'* 1.11   9/4/2022   Modify PigFunc demo
'* 1.12   10/5/2022   Modify PigConfig demo
'* 1.13   15/5/2022   Modify PigFuncDemo
'* 1.14   29/5/2022   Add PigXmlDemo
'* 1.15   30/5/2022   Modify PigXmlDemo
'* 1.16   3/6/2022   Modify PigXmlDemo
'* 1.17   16/6/2022   Add  PigVBCode Demo
'* 1.18   2/9/2022   Modify PigFunc demo
'* 1.19   9/10/2022   Modify PigConfig
'* 1.20   16/10/2022  Add PigMLangDemo
'* 1.21   27/10/2022  Add PigWebReqDemo
'* 1.22   8/11/2022  Add PigSendDemo
'* 1.23   15/11/2022 Modify PigFuncDemo,PigRsa
'* 1.25   18/1/2023 Modify PigMLangDemo
'* 1.26   2/4/2023  Remove reference to PigCmdLib
'* 1.27   28/4/2023  Modify Aes code
'* 1.28   11/6/2023  Add PigFSDemo
'* 1.29   15/6/2023  Modify PigFSDemo
'* 1.30   28/9/2023  Modify MkStr2Func
'* 1.31   17/10/2023  Modify MkStr2Func
'* 1.32   4/12/2023  Modify PigFileDemo
'************************************
Imports PigToolsLiteLib
Imports System.Xml
Imports System.Globalization
Imports Microsoft.VisualBasic

Public Class ConsoleDemo

    Public ShareMem As ShareMem
    Public SMName As String
    Public SMSize As Integer = 1024
    Public MainText As String
    Public Ret As String
    Public Base64Str As String
    Public Base64EncKey As String
    Public BytesEncKey As Byte()
    Public PigConfigApp As PigConfigApp
    Public PigConfigSession As PigConfigSession
    Public PigConfig As PigConfig
    Public ConfName As String
    Public ConfValue As String
    Public SessionName As String
    Public SessionDesc As String
    Public ConfDesc As String
    Public ConfData As String
    Public FolderPath As String
    Public PigFolder As PigFolder
    Public FilePath As String
    Public FilePath2 As String
    Public FilePart As PigFunc.EnmFilePart
    Public Line As String
    Public SaveType As PigConfigApp.EnmSaveType
    Public WithEvents PigFunc As New PigFunc
    Public Url As String
    Public Para As String
    Public UserAgent As String
    Public SrcStr As String
    Public UnEncStr As String
    Public SignBase64 As String
    Public Base64EncStr As String
    Public LeftStr As String
    Public RightStr As String
    Public EnvVar As String
    Public CsvLine As String
    Public BcpLine As String
    Public PigConsole As New PigConsole
    Public PigProcApp As PigProcApp
    Public PID As Integer
    Public ProcName As String
    Public SrcSize As String
    Public MenuKey As String
    Public MenuKey2 As String
    Public MenuDefinition As String
    Public MenuDefinition2 As String
    Public EncKey As String
    Public ThreadID As Integer
    Public PigXml As New PigXml(True)
    Public XmlKey As String
    Public TextValue As String
    Public XmlKey2 As String
    Public XmlStr As String = "<a><b id='1'>b1</b><b id='2'>b2</b><b id='3'>b3</b></a>"
    Public SkipTimes As String = "0"
    Public XmlNode As XmlNode
    Public PigVBCode As New PigVBCode
    Public ClassName As String
    Public KeyName As String
    Public WithEvents PigFile As PigFile
    Public PigFile2 As PigFile
    Public SrcFile As String
    Public TarFile As String
    Public PigSort As PigSort
    Public SortWhat As PigSort.EnmSortWhat
    Public ComprssType As SeowEnc.EmnComprssType
    Public CompressRate As Decimal
    Public UseTime As New UseTime
    Public PigMLang As PigMLang
    Public MLangTitle As String = "MLangTest"
    Public MLangDir As String = "C:\temp"
    Public LCID As Integer = 2052
    Public CultureName As String = "en-US"
    Public ObjName As String
    Public Key As String
    Public MLangText As String
    Public CurrCultureName As String
    Public PigWebReq As PigWebReq
    Public PigSend As PigSend
    Public InitStr As String
    Public EncKeyFilePath As String
    Public PigFS As New PigFileSystem
    Public IsOverwrite As Boolean
    Public OldVersion As String
    Public LatestVersion As String
    Public FuncName As String
    Public ToFuncStr As String
    Public FuncStr As String
    Public TextStream As TextStream
    Public AnyText As String


    Public Sub Main()

        Dim o As New PigXml(False)
        Do While True
            Console.Clear()
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
            Console.WriteLine("Press L to PigProc")
            Console.WriteLine("Press M to PigVBCode")
            Console.WriteLine("Press N to PigTripleDES")
            Console.WriteLine("Press O to PigSort")
            Console.WriteLine("Press P to SeowEnc")
            Console.WriteLine("Press R to PigMLang")
            Console.WriteLine("Press S to PigSend")
            Console.WriteLine("Press T to PigFileSystem")
            Console.WriteLine("Press U to TextStream")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.U
                    Me.TextStreamDemo()
                Case ConsoleKey.T
                    Me.PigFSDemo()
                Case ConsoleKey.O
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigSort")
                    Console.WriteLine("*******************")
                    Me.Line = "Select SortWhat=("
                    Me.Line &= PigSort.EnmSortWhat.SortByte.ToString & "=" & PigSort.EnmSortWhat.SortByte
                    Me.Line &= "," & PigSort.EnmSortWhat.SortDate.ToString & "=" & PigSort.EnmSortWhat.SortDate
                    Me.Line &= "," & PigSort.EnmSortWhat.SortDecimal.ToString & "=" & PigSort.EnmSortWhat.SortDecimal
                    Me.Line &= "," & PigSort.EnmSortWhat.SortLong.ToString & "=" & PigSort.EnmSortWhat.SortLong
                    Me.Line &= "," & PigSort.EnmSortWhat.SortString.ToString & "=" & PigSort.EnmSortWhat.SortString
                    Me.Line &= ")"
                    Me.PigConsole.GetLine(Me.Line, Me.SortWhat）
                    Me.PigSort = New PigSort(Me.SortWhat)
                    Select Case Me.SortWhat
                        Case PigSort.EnmSortWhat.SortByte
                            Me.PigSort.AddByteValue(168)
                            Me.PigSort.AddByteValue(18)
                            Me.PigSort.AddByteValue(218)
                        Case PigSort.EnmSortWhat.SortDate
                            Me.PigSort.AddValue(#2023-12-1#)
                            Me.PigSort.AddValue(Now)
                            Me.PigSort.AddValue(#2022-1-1#)
                        Case PigSort.EnmSortWhat.SortDecimal
                            Me.PigSort.AddDecValue(168.88)
                            Me.PigSort.AddDecValue(18.88)
                            Me.PigSort.AddDecValue(888.888)
                        Case PigSort.EnmSortWhat.SortLong
                            Me.PigSort.AddValue(168)
                            Me.PigSort.AddValue(1888)
                            Me.PigSort.AddValue(8888)
                        Case PigSort.EnmSortWhat.SortString
                            Me.PigSort.AddValue("aaaaaa")
                            Me.PigSort.AddValue("cccccc")
                            Me.PigSort.AddValue("dddddd")
                    End Select
                    Me.PigSort.Sort()
                    Console.WriteLine("GetMaxValue=" & Me.PigSort.GetMaxValue)
                    Console.WriteLine("GetMinValue=" & Me.PigSort.GetMinValue)
                    For i = 0 To Me.PigSort.MaxSortNo
                        Console.WriteLine(i & "=" & Me.PigSort.GetValue(i))
                    Next
                    Me.PigConsole.DisplayPause()
                Case ConsoleKey.M
                    Me.VBCodeDemo()
                Case ConsoleKey.A
                    Me.PigCompressorDemo()
                Case ConsoleKey.B
                    Me.PigTextDemo()
                Case ConsoleKey.C
                    Console.CursorVisible = True
                    Me.PigConsole.GetLine("Input FilePath", Me.FilePath)
                    Console.WriteLine()
                    If Me.PigConsole.IsYesOrNo("Is test FastPigMD5") = True Then
                        Me.PigFile = New PigFile(Me.FilePath)
                        Console.WriteLine()
                        Dim strPigMD5 As String = ""
                        Console.WriteLine("GetFastPigMD5")
                        Dim oUseTime As New UseTime
                        oUseTime.GoBegin()
                        Me.Ret = Me.PigFile.GetFastPigMD5(strPigMD5)
                        oUseTime.ToEnd()
                        Console.WriteLine(Me.Ret)
                        Console.WriteLine("PigMD5=" & strPigMD5)
                        Console.WriteLine("UseTime=" & oUseTime.AllDiffSeconds)
                        If Me.PigConsole.IsYesOrNo("Is test SegLoadFile") = True Then
                            Dim intSegSize As Integer = 2048
                            Me.PigConsole.GetLine("Input Segment size", Me.Line)
                            intSegSize = CInt(Me.Line)
                            Me.PigFile = New PigFile(Me.FilePath)
                            Console.WriteLine("SegLoadFile")
                            Me.Ret = Me.PigFile.SegLoadFile(intSegSize)
                            Console.WriteLine(Me.Ret)
                        End If
                        Console.WriteLine("GetTextRows=" & Me.PigFile.GetTextRows)
                        Dim intRows As Integer = 10
                        Me.PigConsole.GetLine("Enter the number of rows to display", intRows)
                        Console.WriteLine("GetTailText(" & intRows & ")")
                        Console.WriteLine(Me.PigFile.GetTailText(intRows))
                        Console.WriteLine("GetTopText(" & intRows & ")")
                        Console.WriteLine(Me.PigFile.GetTopText(intRows))
                    Else
                        Me.PigFileDemo(Me.FilePath)
                    End If
                    Me.PigConsole.DisplayPause()
                Case ConsoleKey.D
                    Me.PigFuncDemo()
                Case ConsoleKey.E
                    Me.PigWebReqDemo()
                Case ConsoleKey.F
                    Console.WriteLine("*******************")
                    Console.WriteLine("ShareMem")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Press Q to Up")
                    Console.WriteLine("Press A to Write")
                    Console.WriteLine("Press B to Read")
                    Console.WriteLine("*******************")
                    Do While True
                        Dim intConsoleKey As ConsoleKey = Console.ReadKey(True).Key
                        Select Case intConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.CursorVisible = True
                                Me.PigConsole.GetLine("Input ShareMem", Me.SMName)
                                Me.PigConsole.GetLine("Input MainText", Me.MainText)
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
                                    Me.PigConsole.GetLine("Input ShareMem", Me.SMName)
                                    Me.PigConsole.GetLine("Input TextLen", Me.Line)
                                    Dim intLen As Integer = CInt(Me.Line)
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
                        Console.WriteLine("Press C to Encrypt(Text)")
                        Console.WriteLine("Press D to Decrypt(Text)")
                        Console.WriteLine("Press E to Encrypt(Bytes)")
                        Console.WriteLine("Press F to Decrypt(Bytes)")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim bolIsInitKey As Boolean = Me.PigConsole.IsYesOrNo("Is there an InitKey")
                                Dim bolIsBase64 As Boolean = Me.PigConsole.IsYesOrNo("Whether to use Base64")
                                If bolIsInitKey = True Then
                                    Dim strIniKey As String = ""
                                    Me.PigConsole.GetLine("Input InitKey", strIniKey)
                                    Dim ptInitKey As New PigText(strIniKey, PigText.enmTextType.UTF8)
                                    If bolIsBase64 = True Then
                                        Console.WriteLine("MkEncKey(Base64InitKey As String, ByRef Base64EncKey As String)")
                                        Me.Ret = oPigAes.MkEncKey(ptInitKey.Base64, Me.Base64EncKey)
                                    Else
                                        Console.WriteLine("MkEncKey(InitKey As Byte(), ByRef EncKey As Byte())")
                                        Me.Ret = oPigAes.MkEncKey(ptInitKey.TextBytes, Me.BytesEncKey)
                                    End If
                                ElseIf bolIsBase64 = True Then
                                    Console.WriteLine("MkEncKey(ByRef Base64EncKey As String)")
                                    Me.Ret = oPigAes.MkEncKey(Me.Base64EncKey)
                                Else
                                    Console.WriteLine("MkEncKey(ByRef EncKey As Byte())")
                                    Me.Ret = oPigAes.MkEncKey(Me.BytesEncKey)
                                End If
                                Console.WriteLine(Me.Ret)
                                Me.PigConsole.GetLine("Input EncKey file path", Me.EncKeyFilePath)
                                If bolIsBase64 = True Then
                                    Console.WriteLine("SaveTextToFile")
                                    Me.Ret = Me.PigFunc.SaveTextToFile(Me.EncKeyFilePath, Me.Base64EncKey)
                                Else
                                    Dim oPigFile As New PigFile(Me.EncKeyFilePath)
                                    Console.WriteLine("SetData")
                                    Me.Ret = oPigFile.SetData(Me.BytesEncKey)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("SaveFile")
                                    Me.Ret = oPigFile.SaveFile()
                                End If
                                Console.WriteLine(Me.Ret)
                            Case ConsoleKey.B
                                Dim bolIsBase64 As Boolean = Me.PigConsole.IsYesOrNo("Whether to use Base64")
                                Dim strEncKeyBase64 As String = "", abEncKey(0) As Byte
                                Me.PigConsole.GetLine("Input EncKey file path", Me.EncKeyFilePath)
                                If bolIsBase64 = True Then
                                    Console.WriteLine("GetFileText(Base64EncKey)")
                                    Me.PigFunc.GetFileText(Me.EncKeyFilePath, Me.Base64EncKey)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("LoadEncKey")
                                    Me.Ret = oPigAes.LoadEncKey(Me.Base64EncKey)
                                Else
                                    Dim oPigFile As New PigFile(Me.EncKeyFilePath)
                                    Console.WriteLine("LoadFile")
                                    Me.Ret = oPigFile.LoadFile
                                    Console.WriteLine(Me.Ret)
                                    If Me.Ret = "OK" Then
                                        Console.WriteLine("LoadEncKey")
                                        Me.Ret = oPigAes.LoadEncKey(oPigFile.GbMain.Main)
                                    End If
                                End If
                                Console.WriteLine(Me.Ret)
                            Case ConsoleKey.C
                                Me.PigConsole.GetLine("Enter the string to encrypt", Me.SrcStr)
                                Console.WriteLine("Encrypt(SrcString As String, ByRef EncBase64Str As String, Optional SrcTextType As PigText.enmTextType = PigText.enmTextType.UTF8)")
                                Me.Ret = oPigAes.Encrypt(Me.SrcStr, Me.Base64EncStr, PigText.enmTextType.UTF8)
                                Console.WriteLine(Me.Ret)
                                Console.WriteLine("Base64EncStr=" & Base64EncStr)
                            Case ConsoleKey.D
                                Me.PigConsole.GetLine("Enter the Base64EncStr", Me.Base64EncStr)
                                Console.WriteLine("Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType)")
                                Me.Ret = oPigAes.Decrypt(Me.Base64EncStr, Me.SrcStr, PigText.enmTextType.UTF8)
                                Console.WriteLine(Me.Ret)
                                Console.WriteLine("UnEncStr")
                                Console.WriteLine(SrcStr)
                            Case ConsoleKey.E
                                Me.PigConsole.GetLine("Enter the source file path", Me.FilePath)
                                Dim oPigFile As New PigFile(Me.FilePath)
                                Console.WriteLine("LoadFile")
                                Me.Ret = oPigFile.LoadFile
                                Console.WriteLine(Me.Ret)
                                Dim bolIsBase64 As Boolean = Me.PigConsole.IsYesOrNo("Whether to use Base64")
                                If bolIsBase64 = True Then
                                    Console.WriteLine("Encrypt(SrcBytes As Byte(), ByRef EncBase64Str As String)")
                                    Me.Ret = oPigAes.Encrypt(oPigFile.GbMain.Main, Me.Base64EncStr)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("Base64EncStr")
                                    Console.WriteLine(Me.Base64EncStr)
                                Else
                                    Dim abEnc(0) As Byte
                                    Console.WriteLine("Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte())")
                                    Me.Ret = oPigAes.Encrypt(oPigFile.GbMain.Main, abEnc)
                                    Console.WriteLine(Me.Ret)
                                    If Me.Ret = "OK" Then
                                        Me.PigConsole.GetLine("Enter the Encrypt file path", Me.FilePath2)
                                        Dim oPigFile2 As New PigFile(Me.FilePath2)
                                        Console.WriteLine("SetData")
                                        Me.Ret = oPigFile2.SetData(abEnc)
                                        Console.WriteLine(Me.Ret)
                                        Console.WriteLine("SaveFile")
                                        Me.Ret = oPigFile2.SaveFile()
                                        Console.WriteLine(Me.Ret)
                                    End If
                                End If
                            Case ConsoleKey.F
                                Me.PigConsole.GetLine("Enter the encrypted file path", Me.FilePath2)
                                Dim oPigFile2 As New PigFile(Me.FilePath2)
                                Console.WriteLine("LoadFile")
                                Me.Ret = oPigFile2.LoadFile
                                Console.WriteLine(Me.Ret)
                                Dim abUnEnc(0) As Byte
                                Console.WriteLine("Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte())")
                                Me.Ret = oPigAes.Decrypt(oPigFile2.GbMain.Main, abUnEnc)
                                Console.WriteLine(Me.Ret)
                                Me.PigConsole.GetLine("Enter the decrypt file path", Me.FilePath)
                                Dim oPigFile As New PigFile(Me.FilePath)
                                Console.WriteLine("SetData")
                                Me.Ret = oPigFile.SetData(abUnEnc)
                                Console.WriteLine(Me.Ret)
                                Console.WriteLine("SaveFile")
                                Me.Ret = oPigFile.SaveFile()
                                Console.WriteLine(Me.Ret)
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
                        Console.WriteLine("Press A to MkEncKey")
                        Console.WriteLine("Press B to LoadEncKey")
                        Console.WriteLine("Press C to Encrypt")
                        Console.WriteLine("Press D to Decrypt")
                        Console.WriteLine("Press E to SignData")
                        Console.WriteLine("Press F to VerifyData")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim bolIsSepKey As Boolean
                                bolIsSepKey = Me.PigConsole.IsYesOrNo("Do you want to separate the public key from the private key?")
                                If bolIsSepKey = True Then
                                    Me.Ret = oPigRsa.MkEncKey(Me.XmlKey, Me.XmlKey2)
                                Else
                                    Me.Ret = oPigRsa.MkEncKey(Me.XmlKey)
                                End If
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("PrivateKeyXml=" & Me.PigFunc.MyOsCrLf & Me.XmlKey & Me.PigFunc.MyOsCrLf)
                                    If bolIsSepKey = True Then
                                        Console.WriteLine("PublicKeyXml=" & Me.PigFunc.MyOsCrLf & Me.XmlKey2 & Me.PigFunc.MyOsCrLf)
                                    End If
                                End If
                            Case ConsoleKey.B
                                Console.CursorVisible = True
                                If Me.PigConsole.IsYesOrNo("Is load EncKey from file") = True Then
                                    Me.PigConsole.GetLine("Input file path", Me.FilePath)
                                    Me.Ret = Me.PigFunc.GetFileText(Me.FilePath, Me.XmlKey)
                                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                                Else
                                    Me.PigConsole.GetLine("Inupt EncKey Xml", Me.XmlKey)
                                End If
                                Me.Ret = oPigRsa.LoadEncKey(Me.XmlKey)
                                Console.WriteLine(Me.Ret)
                            Case ConsoleKey.C
                                Me.PigConsole.GetLine("Enter the string to encrypt", Me.SrcStr)
                                Dim oPigText As New PigText(Me.SrcStr, PigText.enmTextType.UTF8)
                                Me.Ret = oPigRsa.Encrypt(oPigText.TextBytes, Me.Base64EncStr)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("EncBase64=" & Me.Base64EncStr)
                                End If
                            Case ConsoleKey.D
                                Me.PigConsole.GetLine("Enter the EncBase64 string", Me.Base64EncKey)
                                Me.Ret = oPigRsa.Decrypt(Me.Base64EncStr, Me.UnEncStr, PigText.enmTextType.UTF8)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("UnEncStr=" & Me.UnEncStr)
                                End If
                            Case ConsoleKey.E
                                Me.PigConsole.GetLine("Enter the string to SignData", Me.SrcStr)
                                Me.Ret = oPigRsa.SignData(Me.SrcStr, PigText.enmTextType.UTF8, Me.SignBase64)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("SignBase64=" & Me.SignBase64)
                                End If
                            Case ConsoleKey.F
                                Me.PigConsole.GetLine("Enter Source String", Me.SrcStr)
                                Me.PigConsole.GetLine("Enter SignBase64", Me.SignBase64)
                                Dim bolIsVerify As Boolean
                                Me.Ret = oPigRsa.VerifyData(Me.SrcStr, PigText.enmTextType.UTF8, Me.SignBase64, bolIsVerify)
                                If Me.Ret <> "OK" Then
                                    Console.WriteLine(Me.Ret)
                                Else
                                    Console.WriteLine("IsVerify=" & bolIsVerify)
                                End If
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
                        Console.Clear()
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
                                Dim intEncType As PigConfigApp.EnmEncType = PigConfigApp.EnmEncType.RSA_AES
                                Me.PigConsole.GetLine("EncType(0-RSA_AES,1-RSA_3DES)", Me.Line)
                                Select Case Me.Line
                                    Case "0", "1"
                                        intEncType = CInt(Me.Line)
                                        Me.PigConsole.GetLine("Input EncKey file path", Me.FilePath)
                                        Console.WriteLine("GetFileText...")
                                        Me.Ret = Me.PigFunc.GetFileText(Me.FilePath, Me.EncKey)
                                        Console.WriteLine(Me.Ret)
                                        If Me.Ret = "OK" Then
                                            Me.PigConfigApp = New PigConfigApp(Me.EncKey, intEncType)
                                            Console.WriteLine("New PigConfigApp=")
                                            If Me.PigConfigApp.LastErr = "" Then
                                                Console.WriteLine("OK")
                                                Me.PigConfigApp.OpenDebug()
                                                Console.WriteLine("EncType=" & Me.PigConfigApp.EncType.ToString)
                                            Else
                                                Console.WriteLine(Me.PigConfigApp.LastErr)
                                            End If
                                        End If
                                    Case Else
                                        Console.WriteLine("invalid EncType")
                                End Select
                            Case ConsoleKey.B
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Console.WriteLine("MkEncKey=")
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
                                    Me.PigConsole.GetLine("Input source string", Me.MainText)
                                    Console.WriteLine("GetEncStr=")
                                    Me.Ret = Me.PigConfigApp.GetEncStr(Me.MainText)
                                    Console.WriteLine(Me.Ret)
                                End If
                            Case ConsoleKey.D
                                Console.CursorVisible = True
                                If Me.PigConfigApp Is Nothing Then
                                    Console.WriteLine("PigConfigApp Is Nothing")
                                Else
                                    Me.PigConsole.GetLine("Input SessionName", Me.SessionName)
                                    Me.PigConsole.GetLine("Input SessionDesc", Me.SessionDesc)
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
                                    Me.PigConsole.GetLine("Input SessionName", Me.SessionName)
                                    Me.PigConfigSession = Me.PigConfigApp.GetPigConfigSession(Me.SessionName)
                                    If Me.PigConfigApp.LastErr <> "" Then
                                        Console.WriteLine(Me.PigConfigApp.LastErr)
                                    ElseIf Me.PigConfigSession Is Nothing Then
                                        Console.WriteLine("PigConfigSession Is Nothing")
                                    Else
                                        Me.PigConsole.GetLine("Input ConfName", Me.ConfName)
                                        Me.PigConsole.GetLine("Input ConfValue", Me.ConfValue)
                                        Me.PigConsole.GetLine("Input ConfDesc", Me.ConfDesc)
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
                                    Me.MenuDefinition2 = "Xml#Xml|Ini#Ini"
                                    Me.PigConsole.SimpleMenu("Select SaveType", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.Null)
                                    Select Case Me.MenuKey2
                                        Case "Xml"
                                            Me.SaveType = PigConfigApp.EnmSaveType.Xml
                                        Case "Ini"
                                            Me.SaveType = PigConfigApp.EnmSaveType.Ini
                                    End Select
                                    Console.WriteLine("Me.SaveType=" & Me.SaveType.ToString)
                                    Console.WriteLine("PigConfigApp.SaveConfig")
                                    Me.Ret = Me.PigConfigApp.SaveConfig(Me.ConfData, Me.SaveType)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("ConfData=")
                                    Console.WriteLine(Me.ConfData)
                                    Me.PigConsole.GetLine("Input FilePath", Me.FilePath)
                                    Console.WriteLine("PigConfigApp.SaveConfigFile")
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
                                    Me.MenuDefinition2 = "Xml#Xml|Ini#Ini"
                                    Me.PigConsole.SimpleMenu("Select SaveType", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.Null)
                                    Select Case Me.MenuKey2
                                        Case "Xml"
                                            Me.SaveType = PigConfigApp.EnmSaveType.Xml
                                        Case "Ini"
                                            Me.SaveType = PigConfigApp.EnmSaveType.Ini
                                    End Select
                                    Console.WriteLine("Me.SaveType=" & Me.SaveType.ToString)
                                    Me.PigConsole.GetLine("Input FilePath", Me.FilePath)
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
                                    Me.PigConsole.GetLine("Input SessionName", Me.SessionName)
                                    Me.PigConfigSession = Me.PigConfigApp.GetPigConfigSession(Me.SessionName)
                                    If Me.PigConfigApp.LastErr <> "" Then
                                        Console.WriteLine(Me.PigConfigApp.LastErr)
                                    ElseIf Me.PigConfigSession Is Nothing Then
                                        Console.WriteLine("PigConfigSession Is Nothing")
                                    Else
                                        Me.PigConsole.GetLine("Input ConfName", Me.ConfName)
                                        Me.PigConsole.GetLine("Input ConfValue", Me.ConfValue)
                                        Me.PigConsole.GetLine("Input ConfDesc", Me.ConfDesc)
                                        If Me.PigConfigSession.PigConfigs.IsItemExists(ConfName) = True Then
                                            With Me.PigConfigSession.PigConfigs.Item(ConfName)
                                                .ConfValue = Me.ConfValue
                                                .ConfDesc = Me.ConfDesc
                                            End With
                                        Else
                                            Console.WriteLine(ConfName & " not exists.")
                                        End If
                                    End If
                                End If
                        End Select
                        Me.PigConsole.DisplayPause()
                    Loop
                Case ConsoleKey.L
                    Me.PigProcDemo()
                Case ConsoleKey.K
                    Me.PigXmlDemo()
                Case ConsoleKey.N
                    Console.WriteLine("*******************")
                    Console.WriteLine("PigTripleDES")
                    Console.WriteLine("*******************")
                    Dim oPigTripleDES As New PigTripleDES
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
                                Me.Ret = oPigTripleDES.MkEncKey(Me.Base64EncKey)
                                Console.WriteLine("MkEncKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Console.WriteLine("Inupt Base64EncKey:" & vbCrLf & Me.Base64EncKey)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncKey = Me.Line
                                End If
                                Me.Ret = oPigTripleDES.LoadEncKey(Me.Base64EncKey)
                                Console.WriteLine("LoadEncKey=" & Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:" & Me.SrcStr)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.SrcStr = Me.Line
                                End If
                                Dim oPigText As New PigText(Me.SrcStr, PigText.enmTextType.UTF8)
                                Me.Ret = oPigTripleDES.Encrypt(oPigText.TextBytes, Me.Base64EncStr)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & Me.Base64EncStr)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:")
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncStr = Me.Line
                                End If
                                Me.SrcStr = ""
                                Me.Ret = oPigTripleDES.Decrypt(Me.Base64EncStr, Me.SrcStr, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & Me.SrcStr)
                        End Select
                    Loop
                Case ConsoleKey.P
                    Console.WriteLine("*******************")
                    Console.WriteLine("SeowEnc")
                    Console.WriteLine("*******************")
                    Me.Line = Me.PigFunc.GetEnmDispStr(SeowEnc.EmnComprssType.NoComprss, False)
                    Me.Line &= Me.PigFunc.GetEnmDispStr(SeowEnc.EmnComprssType.AutoComprss)
                    Me.Line &= Me.PigFunc.GetEnmDispStr(SeowEnc.EmnComprssType.Over328ToComprss)
                    Me.Line &= Me.PigFunc.GetEnmDispStr(SeowEnc.EmnComprssType.PigCompressor)
                    Me.PigConsole.GetLine("ComprssType=(" & Me.Line & ")", Me.ComprssType)
                    Dim oSeowEnc As New SeowEnc(Me.ComprssType)
                    Console.CursorVisible = False
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to MkEncKey")
                        Console.WriteLine("Press B to LoadEncKey")
                        Console.WriteLine("Press C to Encrypt")
                        Console.WriteLine("Press D to Decrypt")
                        Console.WriteLine("Press E to Encrypt file")
                        Console.WriteLine("Press F to Decrypt file")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Me.PigConsole.GetLine("Input InitStr", Me.InitStr)
                                Dim ptInit As New PigText(Me.InitStr, PigText.enmTextType.UTF8)
                                Me.Ret = oSeowEnc.MkEncKey(256, ptInit.Base64, Me.Base64EncKey)
                                Console.WriteLine("MkEncKey=" & Me.Ret)
                                Console.WriteLine("Base64EncKey=" & Me.Base64EncKey)
                            Case ConsoleKey.B
                                Dim bolIsBase64 As Boolean = Me.PigConsole.IsYesOrNo("Whether to use Base64")
                                Dim strEncKeyBase64 As String = "", abEncKey(0) As Byte
                                Me.PigConsole.GetLine("Input EncKey file path", Me.EncKeyFilePath)
                                If bolIsBase64 = True Then
                                    Console.WriteLine("GetFileText(Base64EncKey)")
                                    Me.PigFunc.GetFileText(Me.EncKeyFilePath, Me.Base64EncKey)
                                    Console.WriteLine(Me.Ret)
                                    Console.WriteLine("LoadEncKey")
                                    Me.Ret = oSeowEnc.LoadEncKey(Me.Base64EncKey)
                                Else
                                    Dim oPigFile As New PigFile(Me.EncKeyFilePath)
                                    Console.WriteLine("LoadFile")
                                    Me.Ret = oPigFile.LoadFile
                                    Console.WriteLine(Me.Ret)
                                    If Me.Ret = "OK" Then
                                        Console.WriteLine("LoadEncKey")
                                        Me.Ret = oSeowEnc.LoadEncKey(oPigFile.GbMain.Main)
                                    End If
                                End If
                                Console.WriteLine(Me.Ret)
                            Case ConsoleKey.C
                                Console.WriteLine("Enter the string to encrypt:" & Me.SrcStr)
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.SrcStr = Me.Line
                                End If
                                Dim oPigText As New PigText(Me.SrcStr, PigText.enmTextType.UTF8)
                                Dim decRate As Decimal = 0
                                Me.Ret = oSeowEnc.Encrypt(oPigText.TextBytes, Me.Base64EncStr, decRate)
                                Console.WriteLine("Encrypt=" & Me.Ret)
                                Console.WriteLine("Base64EncStr=" & vbCrLf & Me.Base64EncStr)
                                Console.WriteLine("CompressRate=" & vbCrLf & decRate)
                            Case ConsoleKey.D
                                Console.WriteLine("Enter the Base64EncStr:")
                                Me.Line = Console.ReadLine
                                If Me.Line <> "" Then
                                    Me.Base64EncStr = Me.Line
                                End If
                                Me.SrcStr = ""
                                Me.Ret = oSeowEnc.Decrypt(Me.Base64EncStr, Me.SrcStr, PigText.enmTextType.UTF8)
                                Console.WriteLine("Decrypt=" & Me.Ret)
                                Console.WriteLine("Decrypt string=" & vbCrLf & Me.SrcStr)
                            Case ConsoleKey.E
                                Me.PigConsole.GetLine("Input srouce file", Me.FilePath)
                                Me.FilePath2 = Me.FilePath & ".se"
                                Me.PigConsole.GetLine("Input SeowEnc file", Me.FilePath2)
                                Dim pfSrc As New PigFile(Me.FilePath)
                                pfSrc.LoadFile()
                                Me.UseTime.GoBegin()
                                Dim pfEnc As New PigFile(Me.FilePath2)
                                Console.WriteLine("")
                                Me.Ret = oSeowEnc.Encrypt(pfSrc.GbMain.Main, pfEnc.GbMain.Main, Me.CompressRate)
                                Console.WriteLine(Me.Ret)
                                Console.WriteLine("CompressRate=" & Me.CompressRate)
                                pfEnc.SaveFile()
                                Me.UseTime.ToEnd()
                                Console.WriteLine("UseTime=" & Me.UseTime.AllDiffSeconds)
                            Case ConsoleKey.F
                                Me.PigConsole.GetLine("Input SeowEnc file", Me.FilePath2)
                                Me.PigConsole.GetLine("Input srouce file", Me.FilePath)
                                Dim pfEnc As New PigFile(Me.FilePath2)
                                pfEnc.LoadFile()
                                Dim pfSrc As New PigFile(Me.FilePath)
                                Me.UseTime.GoBegin()
                                Console.WriteLine("")
                                Me.Ret = oSeowEnc.Decrypt(pfEnc.GbMain.Main, pfSrc.GbMain.Main)
                                Console.WriteLine(Me.Ret)
                                pfSrc.SaveFile()
                                Me.UseTime.ToEnd()
                                Console.WriteLine("UseTime=" & Me.UseTime.AllDiffSeconds)
                        End Select
                    Loop
                Case ConsoleKey.R
                    Me.PigMLangDemo()
                Case ConsoleKey.S
                    Me.PigSendDemo()
                Case Else
                    Console.WriteLine("Coming soon...")
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub ShowPigProc(ByRef oPigProc As PigProc)
        With oPigProc
            Console.WriteLine("ProcessID=" & .ProcessID)
            Console.WriteLine("ProcessName=" & .ProcessName)
            Console.WriteLine("ModuleName=" & .ModuleName)
            Console.WriteLine("FilePath=" & .FilePath)
            Console.WriteLine("MemoryUse=" & CDec(.MemoryUse) / 1024 / 1024)
            Console.WriteLine("TotalProcessorTime=" & .TotalProcessorTime.ToString)
            Console.WriteLine("UserProcessorTime=" & .UserProcessorTime.ToString)
        End With
    End Sub

    Public Sub PigXmlDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = "NewPigXml#New PigXml|"
            Me.MenuDefinition2 = "LoadXmlDocumentByFile#Load XmlDocument by file|"
            Me.MenuDefinition2 &= "LoadXmlDocumentByXml#Load XmlDocument by xml string|"
            Me.MenuDefinition2 &= "GetXmlDocText#GetXmlDocText|"
            Me.MenuDefinition2 &= "SetXmlDocValue#SetXmlDocValue|"
            Me.MenuDefinition2 &= "ShowXmdDocMainStr#Show XmdDocMainStr|"
            Me.MenuDefinition2 &= "GetXmlDocAttribute#GetXmlDocAttribute|"
            Me.MenuDefinition2 &= "GetXmlDocNode#GetXmlDocNode|"
            Me.MenuDefinition2 &= "GetXmlDocTextFromNode#GetXmlDocText(FromNode)|"
            Me.MenuDefinition2 &= "GetXmlDocAttributeFromNode#GetXmlDocAttribute(FromNode)|"
            Me.MenuDefinition2 &= "XmlAddTest#XmlAdd test|"
            Me.PigConsole.SimpleMenu("PigXml", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "ShowXmdDocMainStr"
                    Console.WriteLine("XmdDocMainStr=" & Me.PigXml.XmdDocMainStr)
                Case "NewPigXml"
                    Dim bolIsCrLf As Boolean = Me.PigConsole.IsYesOrNo("Is need CrLf")
                    Me.PigXml = New PigXml(bolIsCrLf)
                Case "XmlAddTest"
                    With Me.PigXml
                        .Clear()
                        .AddEleLeftSign("Root", True)
                        .AddEleLeftAttribute("ID", "123")
                        .AddEleLeftAttribute("Ref", "abc")
                        .AddEleLeftSignEnd()
                        .AddEle("a", "aaa")
                        .AddEle("b", "bbb")
                        .AddEleLeftSign("c")
                        .AddEle("c1", "ccc1")
                        .AddEle("c2", "ccc1")
                        .AddEleRightSign("c")
                        .AddEleRightSign("Root")
                        Console.WriteLine("MainXmlStr=" & .MainXmlStr)
                        .InitXmlDocument()
                        Console.WriteLine("Root.a=" & .GetXmlDocText("Root.a"))
                        Console.WriteLine("Root.b=" & .GetXmlDocText("Root.b"))
                        Console.WriteLine("Root.c.c1=" & .GetXmlDocText("Root.c.c1"))
                        Console.WriteLine("Root.c.c2=" & .GetXmlDocText("Root.c.c2"))
                        Console.WriteLine("Root.ID=" & .GetXmlDocAttribute("Root.ID"))
                        Console.WriteLine("Root.Ref=" & .GetXmlDocAttribute("Root.Ref"))
                    End With
                Case "LoadXmlDocumentByFile"
                    Me.PigConsole.GetLine("Input xml FilePath", Me.FilePath)
                    Console.WriteLine("InitXmlDocument")
                    Console.WriteLine(Me.FilePath)
                    Me.Ret = Me.PigXml.InitXmlDocument(Me.FilePath)
                    Console.WriteLine(Me.Ret)
                Case "LoadXmlDocumentByXml"
                    Me.PigConsole.GetLine("Input xml string", Me.XmlStr)
                    Me.PigXml.SetMainXml(Me.XmlStr)
                    Console.WriteLine("InitXmlDocument")
                    Me.Ret = Me.PigXml.InitXmlDocument()
                    Console.WriteLine(Me.Ret)
                Case "GetXmlDocText"
                    Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                    Me.PigConsole.GetLine("Input SkipTimes", Me.SkipTimes)
                    If IsNumeric(Me.SkipTimes) = False Or Me.SkipTimes <= 0 Then
                        Console.WriteLine("IsXmlNodeExists(" & Me.XmlKey & ")=" & Me.PigXml.IsXmlNodeExists(Me.XmlKey))
                        Console.WriteLine(Me.XmlKey & "=" & Me.PigXml.GetXmlDocText(Me.XmlKey))
                    Else
                        Console.WriteLine("IsXmlNodeExists(" & Me.XmlKey & ")=" & Me.PigXml.IsXmlNodeExists(Me.XmlKey, CInt(Me.SkipTimes)))
                        Console.WriteLine(Me.XmlKey & "=" & Me.PigXml.GetXmlDocText(Me.XmlKey, CInt(Me.SkipTimes)))
                    End If
                Case "SetXmlDocValue"
                    Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                    Me.PigConsole.GetLine("Input SkipTimes", Me.SkipTimes)
                    Me.PigConsole.GetLine("Input Value", Me.TextValue)
                    If IsNumeric(Me.SkipTimes) = False Or Me.SkipTimes <= 0 Then
                        Console.WriteLine("Set String value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.TextValue))
                        Console.WriteLine("Set Integer value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GEInt(Me.TextValue)))
                        Console.WriteLine("Set Long value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECLng(Me.TextValue)))
                        Console.WriteLine("Set Date value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECBool(Me.TextValue)))
                        Console.WriteLine("Set CBool value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECDate(Me.TextValue)))
                        Console.WriteLine("Set Decimal value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GEDec(Me.TextValue)))
                    Else
                        Console.WriteLine("Set String value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.TextValue, CInt(Me.SkipTimes)))
                        Console.WriteLine("Set Integer value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GEInt(Me.TextValue), CInt(Me.SkipTimes)))
                        Console.WriteLine("Set Long value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECLng(Me.TextValue), CInt(Me.SkipTimes)))
                        Console.WriteLine("Set Date value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECBool(Me.TextValue), CInt(Me.SkipTimes)))
                        Console.WriteLine("Set CBool value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GECDate(Me.TextValue), CInt(Me.SkipTimes)))
                        Console.WriteLine("Set Decimal value:" & Me.PigXml.SetXmlDocValue(Me.XmlKey, Me.PigFunc.GEDec(Me.TextValue), CInt(Me.SkipTimes)))
                    End If
                Case "GetXmlDocAttribute"
                    Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                    Me.PigConsole.GetLine("Input SkipTimes", Me.SkipTimes)
                    If IsNumeric(Me.SkipTimes) = False Or Me.SkipTimes <= 0 Then
                        Console.WriteLine(Me.XmlKey & "=" & Me.PigXml.GetXmlDocAttribute(Me.XmlKey))
                    Else
                        Console.WriteLine(Me.XmlKey & "=" & Me.PigXml.GetXmlDocAttribute(Me.XmlKey, CInt(Me.SkipTimes)))
                    End If
                Case "GetXmlDocTextFromNode"
                    If Me.XmlNode Is Nothing Then
                        Console.WriteLine("XmlNode Is Nothing")
                    Else
                        Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                        Console.WriteLine("GetXmlDocText")
                        Console.WriteLine(Me.PigXml.GetXmlDocText(Me.XmlKey, Me.XmlNode))
                        If Me.PigXml.LastErr <> "" Then
                            Console.WriteLine(Me.PigXml.LastErr)
                        End If
                    End If
                Case "GetXmlDocAttributeFromNode"
                    If Me.XmlNode Is Nothing Then
                        Console.WriteLine("XmlNode Is Nothing")
                    Else
                        Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                        Console.WriteLine("GetXmlDocAttribute")
                        Console.WriteLine(Me.PigXml.GetXmlDocAttribute(Me.XmlKey, Me.XmlNode))
                        If Me.PigXml.LastErr <> "" Then
                            Console.WriteLine(Me.PigXml.LastErr)
                        End If
                    End If
                Case "GetXmlDocNode"
                    Me.PigConsole.GetLine("Input xml key", Me.XmlKey)
                    Me.PigConsole.GetLine("Input SkipTimes", Me.SkipTimes)
                    If IsNumeric(Me.SkipTimes) = False Or Me.SkipTimes <= 0 Then
                        Me.XmlNode = Me.PigXml.GetXmlDocNode(Me.XmlKey)
                    Else
                        Me.XmlNode = Me.PigXml.GetXmlDocNode(Me.XmlKey, CInt(Me.SkipTimes))
                    End If
                    If Me.PigXml.LastErr <> "" Then
                        Console.WriteLine(Me.PigXml.LastErr)
                    Else
                        If Me.XmlNode IsNot Nothing Then Console.WriteLine(Me.XmlNode.OuterXml)
                    End If
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigProcDemo()
        If Me.PigProcApp Is Nothing Then Me.PigProcApp = New PigProcApp
        Console.WriteLine("*******************")
        Console.WriteLine("PigProcApp")
        Console.WriteLine("*******************")
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Up")
            Console.WriteLine("Press A to GetPigProc")
            Console.WriteLine("Press B to GetPigProcs")
            Console.WriteLine("Press C to KillPigProc")
            Console.WriteLine("Press D to KillPigProcs")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Me.PigConsole.GetLine("Input PID", Me.PID)
                    Dim oPigProc As PigProc = Me.PigProcApp.GetPigProc(Me.PID)
                    If Me.PigProcApp.LastErr <> "" Then
                        Console.WriteLine("PigProcApp.GetPigProc=" & PigProcApp.LastErr)
                    Else
                        Me.ShowPigProc(oPigProc)
                    End If
                Case ConsoleKey.B
                    Me.PigConsole.GetLine("Input ProcName", Me.ProcName)
                    Dim oPigProcs As PigProcs = Me.PigProcApp.GetPigProcs(Me.ProcName)
                    If Me.PigProcApp.LastErr <> "" Then
                        Console.WriteLine("PigProcApp.GetPigProcs=" & PigProcApp.LastErr)
                    Else
                        For Each oPigProc As PigProc In oPigProcs
                            Me.ShowPigProc(oPigProc)
                        Next
                    End If
                Case ConsoleKey.C
                    Me.PigConsole.GetLine("Input PID", Me.PID)
                    Dim oPigProc As PigProc = Me.PigProcApp.GetPigProc(Me.PID)
                    If Me.PigProcApp.LastErr <> "" Then
                        Console.WriteLine("PigProcApp.GetPigProc=" & PigProcApp.LastErr)
                    Else
                        Me.ShowPigProc(oPigProc)
                    End If
                Case ConsoleKey.D
                    Me.PigConsole.GetLine("Input ProcName", Me.ProcName)
                    Dim oPigProcs As PigProcs = Me.PigProcApp.GetPigProcs(Me.ProcName)
                    If Me.PigProcApp.LastErr <> "" Then
                        Console.WriteLine("PigProcApp.GetPigProcs=" & PigProcApp.LastErr)
                    Else
                        For Each oPigProc As PigProc In oPigProcs
                            Me.ShowPigProc(oPigProc)
                        Next
                    End If
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub VBCodeDemo()
        Console.WriteLine("*******************")
        Console.WriteLine("PigVBCodeDemo")
        Console.WriteLine("*******************")
        Me.PigConsole.GetLine("Input save file path", Me.FilePath）
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = "MkCollectionClass#MkCollectionClass|"
            Me.MenuDefinition2 &= "MkBytes2Func#MkBytes2Func|"
            Me.MenuDefinition2 &= "MkStr2Func#MkStr2Func|"
            Me.PigConsole.SimpleMenu("PigVBCodeDemo", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Dim strData As String = ""
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "MkCollectionClass"
                    Me.PigConsole.GetLine("Input Member Class Name", Me.ClassName）
                    Me.PigConsole.GetLine("Input Member Class KeyName", Me.KeyName）
                    Me.Ret = Me.PigVBCode.MkCollectionClass(strData, Me.ClassName, Me.KeyName)
                Case "MkBytes2Func"
                    Me.PigConsole.GetLine("Enter function name", Me.FuncName)
                    Me.PigConsole.GetLine("Enter a string to convert to a function", Me.ToFuncStr)
                    Dim ptAny As New PigText(Me.ToFuncStr, PigText.enmTextType.UTF8)
                    Console.WriteLine("MkBytes2Func")
                    Me.Ret = Me.PigVBCode.MkBytes2Func(ptAny.TextBytes, Me.FuncName, Me.FuncStr)
                    Console.WriteLine(Me.Ret)
                    Console.WriteLine("FuncStr")
                    Console.WriteLine(Me.FuncStr)
                    Me.PigFunc.ClipboardSetText(Me.FuncStr)
                Case "MkStr2Func"
                    Me.PigConsole.GetLine("Enter function name", Me.FuncName)
                    Me.PigConsole.GetLine("Enter a string to convert to a function", Me.ToFuncStr)
                    Console.WriteLine("MkStr2Func")
                    Me.Ret = Me.PigVBCode.MkStr2Func(Me.ToFuncStr, PigText.enmTextType.UTF8, Me.FuncName, Me.FuncStr)
                    Console.WriteLine(Me.Ret)
                    Console.WriteLine("FuncStr")
                    Console.WriteLine(Me.FuncStr)
                    Me.PigFunc.ClipboardSetText(Me.FuncStr)
                Case "DownloadFile"
            End Select
            If Me.Ret <> "OK" Then
                Console.WriteLine(Me.Ret)
            Else
                Me.PigFunc.SaveTextToFile(Me.FilePath, strData)
            End If
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigSendDemo()
        Console.WriteLine("*******************")
        Console.WriteLine("PigSendDemo")
        Console.WriteLine("*******************")
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = ""
            Me.MenuDefinition2 &= "NewOriginal#New and EncType is Original|"
            Me.MenuDefinition2 &= "NewSeowEnc#New and EncType is SeowEnc|"
            Me.MenuDefinition2 &= "NewSeowEncAndPigAes#New and EncType is SeowEncAndPigAes|"
            Me.MenuDefinition2 &= "SendBytesData#SendBytesData|"
            Me.MenuDefinition2 &= "SendStrData#SendStrData|"
            Me.MenuDefinition2 &= "ReceiveBytesData#ReceiveBytesData|"
            Me.MenuDefinition2 &= "ReceiveStrData#ReceiveStrData|"
            Me.PigConsole.SimpleMenu("PigVBCodeDemo", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Dim strData As String = ""
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "NewOriginal"
                    Console.WriteLine("New PigSend(Original)")
                    Me.PigSend = New PigSend(PigSend.EnmEncType.Original)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                Case "NewSeowEnc"
                    Console.WriteLine("New PigSend(SeowEnc)")
                    Me.PigSend = New PigSend(PigSend.EnmEncType.SeowEnc)
                    Dim oSeowEnc As New SeowEnc(SeowEnc.EmnComprssType.AutoComprss)
                    Dim strEncKey As String = ""
                    Console.WriteLine("SeowEnc.MkEncKey")
                    Me.Ret = oSeowEnc.MkEncKey(256, "", strEncKey)
                    If Me.Ret <> "OK" Then
                        Console.WriteLine(Me.Ret)
                    Else
                        Console.WriteLine("SeowEnc.LoadEncKey")
                        Me.Ret = oSeowEnc.LoadEncKey(strEncKey)
                        If Me.Ret <> "OK" Then
                            Console.WriteLine(Me.Ret)
                        Else
                            Console.WriteLine("InitEnc(SeowEnc)")
                            Me.Ret = Me.PigSend.InitEnc(oSeowEnc)
                            If Me.Ret <> "OK" Then
                                Console.WriteLine(Me.Ret)
                            End If
                        End If
                    End If
                Case "SendBytesData"
                    Me.PigConsole.GetLine("Input send base64 string", Me.Base64Str)
                    Dim oPigBytes As New PigBytes(Me.Base64Str), strTarBase64 As String = ""
                    Me.Ret = Me.PigSend.SendBytesData(oPigBytes.Main, strTarBase64)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Console.WriteLine("Send TarBase64=" & strTarBase64)
                Case "SendStrData"
                    Me.PigConsole.GetLine("Input send string", Me.Line)
                    Dim strTarBase64 As String = ""
                    Me.Ret = Me.PigSend.SendStrData(Me.Line, PigText.enmTextType.UTF8, strTarBase64)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Console.WriteLine("Send TarBase64=" & strTarBase64)
                Case "ReceiveBytesData"
                    Me.PigConsole.GetLine("Input recvie base64 string", Me.Base64Str)
                    Dim oPigBytes As New PigBytes(Me.Base64Str), strTarBase64 As String = ""
                    Me.Ret = Me.PigSend.ReceiveBytesData(oPigBytes.Main, strTarBase64)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Console.WriteLine("Receive TarBase64=" & strTarBase64)
                Case "ReceiveStrData"
                    Me.PigConsole.GetLine("Input recvie base64 string", Me.Base64Str)
                    Dim strTarStr As String = ""
                    Me.Ret = Me.PigSend.ReceiveStrData(Me.Base64Str, PigText.enmTextType.UTF8, strTarStr)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Console.WriteLine("Receive TarStr=" & strTarStr)
            End Select
            If Me.Ret <> "OK" Then
                Console.WriteLine(Me.Ret)
            Else
                Me.PigFunc.SaveTextToFile(Me.FilePath, strData)
            End If
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigWebReqDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = "NewPigWebReq#New PigWebReq|"
            Me.MenuDefinition2 &= "DownloadFile#DownloadFile|"
            Me.MenuDefinition2 &= "GetText#GetText|"
            Me.MenuDefinition2 &= "PostText#PostText|"
            Me.PigConsole.SimpleMenu("PigWebReqDemo", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "NewPigWebReq"
                    Me.PigConsole.GetLine("Input Url", Me.Url)
                    If Me.Url = "" Then
                        Console.WriteLine("Url cannot be empty")
                    Else
                        Dim bolIsNeedPara As Boolean, bolIsNeedUserAgent As Boolean
                        bolIsNeedPara = Me.PigConsole.IsYesOrNo("Is need Para")
                        If bolIsNeedPara = True Then
                            Me.PigConsole.GetLine("Input Para", Me.Para)
                        End If
                        bolIsNeedUserAgent = Me.PigConsole.IsYesOrNo("Is need UserAgent")
                        If bolIsNeedUserAgent = True Then
                            Me.PigConsole.GetLine("Input UserAgent", Me.UserAgent)
                        End If
                        If Me.UserAgent <> "" And Me.Para <> "" Then
                            Me.PigWebReq = New PigWebReq(Me.Url, Me.Para, Me.UserAgent)
                        ElseIf Me.Para <> "" Then
                            Me.PigWebReq = New PigWebReq(Me.Url, Me.Para)
                        Else
                            Me.PigWebReq = New PigWebReq(Me.Url)
                        End If
                    End If
                Case "DownloadFile"
                    Me.PigConsole.GetLine("Enter the download file path to save", Me.FilePath)
                    Console.WriteLine("DownloadFile...")
                    Me.Ret = Me.PigWebReq.DownloadFile(Me.FilePath)
                    Console.WriteLine(Me.Ret)
                Case "GetText"
                    Me.PigWebReq.GetText()
                Case "PostText"
                    'Me.PigWebReq.PostText()
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigFuncDemo()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("PigFunc")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q To Exit")
            Console.WriteLine("Press A To Show Function Demo")
            Console.WriteLine("Press B To GetFilePart")
            Console.WriteLine("Press C To GetFileText")
            Console.WriteLine("Press D To SaveTextToFile")
            Console.WriteLine("Press E To GetEnvVar")
            Console.WriteLine("Press F To PLSqlCsv2Bcp")
            Console.WriteLine("Press G To GetHumanSize")
            Console.WriteLine("Press H To OptLogInf")
            Console.WriteLine("Press I To SaveShareMem")
            Console.WriteLine("Press J To GetShareMem")
            Console.WriteLine("Press K To CheckFileDiff")
            Console.WriteLine("Press L To IsFileDiff")
            Console.WriteLine("Press M To IsNewVersion")
            Console.WriteLine("Press O To OpenUrl")
            Console.WriteLine("Press P To GetTextFileEncCode")
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
                        Console.WriteLine("GetHostIp(""192."")=" & .GetHostIp("192."))
                        Console.WriteLine("GetUserName=" & .GetUserName)
                        Console.WriteLine("GetEnvVar(""Path"")=" & .GetEnvVar("Path"))
                        'Console.WriteLine("GetFilePart(""c: \temp\aaa"", PigFunc.enmFilePart.FileTitle)=" & .GetFilePart("c:\temp\aaa", PigFunc.enmFilePart.FileTitle))
                        Console.WriteLine("GetProcThreadID=" & .GetProcThreadID)
                        Console.WriteLine("GetRandNum(1, 100)=" & .GetRandNum(1, 100))
                        Console.WriteLine("GetRandString(16, PigFunc.enmGetRandString.NumberAndLetter)=" & .GetRandString(16, PigFunc.EnmGetRandString.NumberAndLetter))
                        Console.WriteLine(".GetRateDesc(16.88)=" & .GetRateDesc(16.88))
                        Console.WriteLine(".GetStr(""<hi>"", ""<"", "">"")=" & .GetStr("<hi>", "<", ">"))
                        Console.WriteLine(".IsRegexMatch(""B2b"", ""^[A-Za-z0-9]+$"")=" & .IsRegexMatch("B2b", "^[A-Za-z0-9]+$"))
                        Console.WriteLine(".GetProcThreadID=" & .GetProcThreadID)
                        Console.WriteLine(".UrlEncode=" & .UrlEncode("https://www.seowphong.com/oss/PigTools"))
                        Console.WriteLine(".UrlDecode=" & .UrlDecode("https%3A%2F%2Fwww.seowphong.com%2Foss%2FPigTools"))
                        .ClearErr()
                        Console.WriteLine(".GetUUID=" & .GetUUID())
                        If .LastErr <> "" Then Console.WriteLine(.LastErr)
                        .ClearErr()
                        Console.WriteLine(".GetMachineGUID=" & .GetMachineGUID())
                        If .LastErr <> "" Then Console.WriteLine(.LastErr)
                        .ClearErr()
                        Console.WriteLine(".GetWindowsProductId=" & .GetWindowsProductId())
                        If .LastErr <> "" Then Console.WriteLine(.LastErr)
                        Console.WriteLine(".GetBootID=" & .GetBootID())
                        Console.WriteLine(".DigitalToChnName(1234567890)=" & .DigitalToChnName("1234567890", True)）
                        Console.WriteLine(".DigitalToChnName(3389)=" & .DigitalToChnName("3389", False)）
                        Console.WriteLine(".DigitalToChnName(大写1234567890)=" & .DigitalToChnName("1234567890", True, True)）
                        Console.WriteLine(".ConvertHtmlStr=" & .ConvertHtmlStr("<html><title>Title</titel><body>Body</body></html>", PigFunc.EmnHowToConvHtml.DisableHTML)）
                        Console.WriteLine(".GetAlignStr(Left Alignment)=" & vbCrLf & "*" & .GetAlignStr("Left Alignment", PigFunc.EnmAlignment.Left, 80) & "*")
                        Console.WriteLine(".GetAlignStr(Right Alignment)=" & vbCrLf & "*" & .GetAlignStr("Right Alignment", PigFunc.EnmAlignment.Right, 80) & "*")
                        Console.WriteLine(".GetAlignStr(Center Alignment)=" & vbCrLf & "*" & .GetAlignStr("Center Alignment", PigFunc.EnmAlignment.Center, 80) & "*")
                        Console.WriteLine(".GetMyExeName()=" & .GetMyExeName)
                        Console.WriteLine(".GetEnmDispStr(PigText.enmTextType.Unicode)=" & .GetEnmDispStr(PigText.enmTextType.Unicode))
                        Console.WriteLine(".GetCompMinutePart(Now)=" & .GetCompMinutePart(Now))
                        Console.WriteLine(".GetTextSHA1(Hi,UTF8)=" & .GetTextSHA1("Hi", PigText.enmTextType.UTF8))
                        Console.WriteLine(".GetTextBase64(Hi,UTF8)=" & .GetTextBase64("Hi", PigText.enmTextType.UTF8))
                        Dim strPwd As String = ""
                        Me.PigConsole.GetLine("Input the password", strPwd)
                        Console.WriteLine(".IsStrongPassword(" & strPwd & ")=" & .IsStrongPassword(strPwd))
                        Dim strValue As String = .EscapeStr("<1><2>")
                        Console.WriteLine(".EscapeStr(<1><2>)=" & strValue)
                        Console.WriteLine(".UnEscapeStr(" & strValue & ")=" & .UnEscapeStr(strValue))
                        Console.WriteLine("DefaultBrowser=" & .GetDefaultBrowser())
                        Dim dteBeginTime As Date, dteEndTime As Date
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.CurrentMonth, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(CurrentMonth)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.CurrentWeek, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(CurrentWeek)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.CurrentYear, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(CurrentYear)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.LastMonth, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(LastMonth)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.LastWeek, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(LastWeek)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                        Me.Ret = .GetTimeSlot(PigFunc.EnmTimeSlot.LastYear, dteBeginTime, dteEndTime)
                        Console.WriteLine("GetTimeSlot(LastYear)=" & Me.Ret & "，BeginTime=" & .GetFmtDateTime(dteBeginTime) & ",EndTime=" & .GetFmtDateTime(dteEndTime))
                    End With
                Case ConsoleKey.B
                    Console.CursorVisible = True
                    Me.GetLine("FilePath", Me.FilePath)
                    Console.WriteLine("FilePart=DriveNo-" & PigFunc.EnmFilePart.DriveNo & ",ExtName-" & PigFunc.EnmFilePart.ExtName & ",FileTitle-" & PigFunc.EnmFilePart.FileTitle & ",Path-" & PigFunc.EnmFilePart.Path)
                    Me.Line = Console.ReadLine
                    Select Case Me.Line
                        Case PigFunc.EnmFilePart.DriveNo, PigFunc.EnmFilePart.ExtName, PigFunc.EnmFilePart.FileTitle, PigFunc.EnmFilePart.Path
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
                    If Me.PigConsole.IsYesOrNo("Is asynchronous processing") = True Then
                        Me.Ret = Me.PigFunc.ASyncSaveTextToFile(Me.FilePath, Me.SrcStr, Me.ThreadID)
                        Console.WriteLine("ASyncSaveTextToFile(" & Me.FilePath & ")=" & Me.Ret)
                        Console.WriteLine("ThreadID=" & Me.ThreadID)
                    Else
                        Me.Ret = Me.PigFunc.SaveTextToFile(Me.FilePath, Me.SrcStr)
                        Console.WriteLine("SaveTextToFile(" & Me.FilePath & ")=" & Me.Ret)
                    End If
                Case ConsoleKey.E
                    Console.CursorVisible = True
                    Me.GetLine("EnvVarName", Me.EnvVar)
                    Me.SrcStr = Me.PigFunc.GetEnvVar(Me.EnvVar)
                    Console.WriteLine("GetEnvVar(" & Me.EnvVar & ")=" & Me.SrcStr)
                Case ConsoleKey.F
                    Console.CursorVisible = True
                    Me.GetLine("CsvLine", Me.CsvLine)
                    Console.WriteLine("CsvLine is ")
                    Console.WriteLine(Me.CsvLine)
                    Console.WriteLine("PigFunc.PLSqlCsv2Bcp=")
                    Me.Ret = Me.PigFunc.PLSqlCsv2Bcp(Me.CsvLine, Me.BcpLine)
                    Console.WriteLine(Me.Ret)
                    Console.WriteLine("BcpLine is ")
                    Console.WriteLine(Me.BcpLine)
                Case ConsoleKey.G
                    Console.CursorVisible = True
                    Me.GetLine("Input SrcSize", Me.SrcSize)
                    Console.WriteLine("SrcSize is " & Me.SrcSize)
                    Console.WriteLine("GetHumanSize is " & Me.PigFunc.GetHumanSize(CDec(Me.SrcSize)))
                Case ConsoleKey.H
                    Console.CursorVisible = True
                    Me.GetLine("Input SrcStr", Me.SrcStr)
                    Me.GetLine("Input FilePath", Me.FilePath)
                    If Me.PigConsole.IsYesOrNo("Is asynchronous processing") = True Then
                        Me.PigFunc.ASyncOptLogInf(Me.SrcStr, Me.FilePath)
                    Else
                        Me.PigFunc.OptLogInf(Me.SrcStr, Me.FilePath)
                    End If
                Case ConsoleKey.I
                    Me.PigConsole.GetLine("Input ShareMem name", Me.KeyName)
                    Me.PigConsole.GetLine("Input Text", Me.Line)
                    Dim oPigText As New PigText(Me.Line, PigText.enmTextType.UTF8)
                    Me.Ret = Me.PigFunc.SaveShareMem(Me.KeyName, oPigText.TextBytes)
                    Console.WriteLine(Me.Ret)
                Case ConsoleKey.J
                    Me.PigConsole.GetLine("Input ShareMem name", Me.KeyName)
                    Dim abOut(0) As Byte, dteCreate As Date
                    Me.Ret = Me.PigFunc.GetShareMem(Me.KeyName, abOut, dteCreate)
                    If Me.Ret = "OK" Then
                        Dim gtAny As New PigText(abOut)
                        Console.WriteLine(gtAny.Text)
                        Console.WriteLine(dteCreate)
                    Else
                        Console.WriteLine(Me.Ret)
                    End If
                Case ConsoleKey.K
                    Me.PigConsole.GetLine("Input source file", Me.SrcFile)
                    Me.PigConsole.GetLine("Input targe file", Me.TarFile)
                    Dim strDisp As String = "CheckDiffType=("
                    strDisp &= PigFunc.EnmCheckDiffType.Size_Date & "=Size_Date,"
                    strDisp &= PigFunc.EnmCheckDiffType.Size_Date_FastPigMD5 & "=Size_Date_FastPigMD5,"
                    strDisp &= PigFunc.EnmCheckDiffType.Size_FullPigMD5 & "=Size_FullPigMD5)"
                    Me.PigConsole.GetLine(strDisp, Me.Line)
                    Dim bolIsDiff As Boolean
                    Dim intCheckDiffType As PigFunc.EnmCheckDiffType = CInt(Me.Line)
                    Dim oUseTime As New UseTime
                    oUseTime.GoBegin()
                    Me.Ret = Me.PigFunc.CheckFileDiff(Me.SrcFile, Me.TarFile, bolIsDiff, intCheckDiffType)
                    oUseTime.ToEnd()
                    Console.WriteLine("CheckFileDiff=" & Me.Ret)
                    Console.WriteLine("IsDiff=" & bolIsDiff)
                    Console.WriteLine("UseTime=" & oUseTime.AllDiffSeconds)
                Case ConsoleKey.L
                    Me.PigConsole.GetLine("Input file1 path", Me.SrcFile)
                    Me.PigConsole.GetLine("Input file2 path", Me.TarFile)
                    Console.WriteLine(CStr(Me.PigFunc.IsFileDiff(Me.SrcFile, Me.TarFile, Me.Ret)))
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                Case ConsoleKey.M
                    Me.PigConsole.GetLine("Input OldVersion", Me.OldVersion)
                    Me.PigConsole.GetLine("Input LatestVersion", Me.LatestVersion)
                    Console.WriteLine(CStr(Me.PigFunc.IsNewVersion(Me.OldVersion, Me.LatestVersion, Me.Ret)))
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                Case ConsoleKey.O
                    Me.PigConsole.GetLine("Input Url", Me.Url)
                    Console.WriteLine("OpenUrl")
                    Me.Ret = Me.PigFunc.OpenUrl(Me.Url)
                    Console.WriteLine(Me.Ret)
            End Select
        Loop
    End Sub

    Public Sub PigFileDemo(FilePath As String)
        Dim strSrc As String = "PigTextDemo Sample code"
        Dim oPigFile As New PigFile(FilePath)
        With oPigFile
            If .LastErr = "" Then
                Console.WriteLine("CreationTime=" & .CreationTime)
                Console.WriteLine("UpdateTime=" & .UpdateTime)
                Console.WriteLine("FileTitle=" & .FileTitle)
                Console.WriteLine("ExtName=" & .ExtName)
                Console.WriteLine("FilePath=" & .FilePath)
                Console.WriteLine("DirPath=" & .DirPath)
                Console.WriteLine("Size=" & .Size)
                Console.WriteLine("MD5=" & .MD5)
                Console.WriteLine("PigMD5=" & .PigMD5)
                Dim strPigMD5 As String = ""
                .GetFastPigMD5(strPigMD5)
                Console.WriteLine("GetFastPigMD5=" & strPigMD5)
                .GetFullPigMD5(strPigMD5)
                Console.WriteLine("GetFullPigMD5=" & strPigMD5)
            End If
        End With
        If Me.PigConsole.IsYesOrNo("Is load file?") = True Then
            oPigFile.LoadFile()
            Dim o As New PigText(oPigFile.GbMain.Main, PigText.enmTextType.Ascii)
            Console.WriteLine(o.Text)
        End If
        If Me.PigConsole.IsYesOrNo("Is test GetFullText") = True Then
            Me.MenuDefinition2 = ""
            Dim intTextType As PigText.enmTextType
            intTextType = PigText.enmTextType.Ascii
            Me.MenuDefinition2 &= intTextType & "#" & intTextType.ToString & "|"
            intTextType = PigText.enmTextType.Unicode
            Me.MenuDefinition2 &= intTextType & "#" & intTextType.ToString & "|"
            intTextType = PigText.enmTextType.UTF8
            Me.MenuDefinition2 &= intTextType & "#" & intTextType.ToString & "|"
            intTextType = PigText.enmTextType.GB2312
            Me.MenuDefinition2 &= intTextType & "#" & intTextType.ToString & "|"
            Me.PigConsole.SimpleMenu("Select TextType", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.Null)
            intTextType = CInt(Me.MenuKey2)
            Dim ptAny As PigText = Nothing
            Console.WriteLine("GetFullText")
            Me.Ret = oPigFile.GetFullText(ptAny, intTextType)
            Console.WriteLine(Me.Ret)
            If Me.Ret = "OK" Then
                Console.WriteLine(ptAny.Text)
            End If
        End If
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

    Private Sub PigFunc_ASyncRet_SaveTextToFile(SyncRet As PigToolsLiteLib.PigBaseMini.StruASyncRet) Handles PigFunc.ASyncRet_SaveTextToFile
        Console.WriteLine("PigFunc_ASyncRet_SaveTextToFile")
        With SyncRet
            Console.WriteLine("BeginTime=" & .BeginTime)
            Console.WriteLine("EndTime=" & .EndTime)
            Console.WriteLine("Ret=" & .Ret)
            Console.WriteLine("ThreadID=" & .ThreadID)
        End With
    End Sub

    Private Sub PigFile_EnvSegLoadFile(SegNo As Integer, SegBytes() As Byte, IsEnd As Boolean) Handles PigFile.EnvSegLoadFile
        Console.WriteLine("SegNo=" & SegNo)
        If SegBytes Is Nothing Then
            Console.WriteLine("SegBytes Is Nothing")
        Else
            Console.WriteLine("SegBytes.Length=" & SegBytes.Length)
        End If
        If SegNo = 0 Then
            Me.PigFile2 = New PigFile(Me.FilePath & ".diff")
            Me.PigFile2.GbMain = New PigBytes
        End If
        Me.PigFile2.GbMain.SetValue(SegBytes)
        If IsEnd = True Then
            Console.WriteLine("To End")
            Me.PigFile2.SaveFile()
        End If
    End Sub

    Public Sub PigMLangDemo()
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = "NewPigMLang#New PigMLang|"
            Me.MenuDefinition2 &= "ShowProperty#Show Properties|"
            Me.MenuDefinition2 &= "LoadMLangInf#LoadMLangInf|"
            Me.MenuDefinition2 &= "GetAllLangInf#GetAllLangInf|"
            Me.MenuDefinition2 &= "AddMLangText#AddMLangText|"
            Me.MenuDefinition2 &= "GetMLangText#GetMLangText|"
            Me.MenuDefinition2 &= "MkMLangText#MkMLangText|"
            Me.MenuDefinition2 &= "GetCanUseCultureXml#GetCanUseCultureXml|"
            Me.PigConsole.SimpleMenu("PigMLang", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "NewPigMLang"
                    Me.PigConsole.GetLine("Input MLangTitle", Me.MLangTitle)
                    Me.PigConsole.GetLine("Input MLangDir", Me.MLangDir)
                    If Me.MLangDir = "" Then
                        Me.PigMLang = New PigMLang(Me.MLangTitle)
                    Else
                        Me.PigMLang = New PigMLang(Me.MLangTitle, Me.MLangDir)
                    End If
                    If Me.PigMLang.LastErr <> "" Then Console.WriteLine(Me.PigMLang.LastErr)
                Case "GetAllLangInf"
                    Console.WriteLine(Me.PigMLang.GetAllLangInf(PigMLang.EnmGetInfFmt.Markdown))
                Case "LoadMLangInf"
                    If Me.PigConsole.IsYesOrNo("Is Auto?") = True Then
                        Console.WriteLine("LoadMLangInf(True)")
                        Me.Ret = Me.PigMLang.LoadMLangInf(True)
                        Console.WriteLine(Me.Ret)
                    Else
                        Me.PigConsole.GetLine("Input CurrCultureName", Me.CurrCultureName)
                        Console.WriteLine("SetCurrCulture")
                        Me.Ret = Me.PigMLang.SetCurrCulture(Me.CurrCultureName)
                        Console.WriteLine(Me.Ret)
                        Console.WriteLine("LoadMLangInf(False)")
                        Me.Ret = Me.PigMLang.LoadMLangInf(False)
                        Console.WriteLine(Me.Ret)
                    End If
                    Console.WriteLine(Me.Ret)
                Case "ShowProperty"
                    Dim strDisp As String = "", strOsCrLf As String = Me.PigFunc.MyOsCrLf
                    With Me.PigMLang
                        strDisp &= "CurrCultureName=" & .CurrCultureName & strOsCrLf
                        strDisp &= "CurrLCID=" & .CurrLCID & strOsCrLf
                        strDisp &= "CurrMLangDir=" & .CurrMLangDir & strOsCrLf
                        strDisp &= "CurrMLangFile=" & .CurrMLangFile & strOsCrLf
                        strDisp &= "CurrMLangTitle=" & .CurrMLangTitle & strOsCrLf
                        strDisp &= "MLangTextCnt=" & .MLangTextCnt & strOsCrLf
                        strDisp &= "CanUseCultureList=" & strOsCrLf
                        If .CanUseCultureList IsNot Nothing Then
                            For Each oCultureInfo As CultureInfo In .CanUseCultureList
                                With oCultureInfo
                                    strDisp &= "[" & .LCID & "]" & .Name & strOsCrLf
                                End With
                            Next
                        End If
                    End With
                    Console.WriteLine(strDisp)
                Case "AddMLangText"
                    With Me.PigMLang
                        If Me.PigConsole.IsYesOrNo("Is GlobalKey") = True Then
                            Me.PigConsole.GetLine("Input GlobalKey", Me.Key)
                            Me.PigConsole.GetLine("Input MLangText", Me.MLangText)
                            .AddMLangText(Me.Key, Me.MLangText)
                        Else
                            Me.PigConsole.GetLine("Input ObjName", Me.ObjName)
                            Me.PigConsole.GetLine("Input Key", Me.Key)
                            Me.PigConsole.GetLine("Input MLangText", Me.MLangText)
                            .AddMLangText(Me.ObjName, Me.Key, Me.MLangText)
                        End If
                    End With
                Case "GetMLangText"
                    With Me.PigMLang
                        If Me.PigConsole.IsYesOrNo("Is GlobalKey") = True Then
                            Me.PigConsole.GetLine("Input GlobalKey", Me.Key)
                            Me.PigConsole.GetLine("Input DefaultText", Me.MLangText)
                            Console.WriteLine(.GetMLangText(Me.Key, Me.MLangText))
                        Else
                            Me.PigConsole.GetLine("Input ObjName", Me.ObjName)
                            Me.PigConsole.GetLine("Input Key", Me.Key)
                            Me.PigConsole.GetLine("Input DefaultText", Me.MLangText)
                            Console.WriteLine(.GetMLangText(Me.ObjName, Me.Key, Me.MLangText))
                        End If
                    End With
                Case "MkMLangText"
                    With Me.PigMLang
                        If Me.PigConsole.IsYesOrNo("Is GlobalKey") = True Then
                            Me.PigConsole.GetLine("Input GlobalKey", Me.Key)
                            Me.PigConsole.GetLine("Input DefaultText", Me.MLangText)
                            'Console.WriteLine(.MkMLangText(Me.Key, Me.MLangText))
                        Else
                            Me.PigConsole.GetLine("Input ObjName", Me.ObjName)
                            Me.PigConsole.GetLine("Input Key", Me.Key)
                            Me.PigConsole.GetLine("Input DefaultText", Me.MLangText)
                            'Console.WriteLine(.MkMLangText(Me.ObjName, Me.Key, Me.MLangText))
                        End If
                    End With
                Case "GetCanUseCultureXml"
                    With Me.PigMLang
                        Console.WriteLine("GetCanUseCultureXml...")
                        Console.WriteLine(.GetCanUseCultureXml)
                    End With
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub


    Private Function GetMenuDefinition_TextType() As String
        Try
            Dim strMenuDefinition As String = ""
            '            Dim intTextType As PigText.enmTextType
            For Each name As String In [Enum].GetNames(GetType(PigText.enmTextType))
                Dim value As PigText.enmTextType = [Enum].Parse(GetType(PigText.enmTextType), name)
                Console.WriteLine("{0} = {1}", name, CInt(value))
            Next

            'intTextType = PigText.enmTextType.Ascii : strMenuDefinition &= intTextType & "#" & intTextType.ToString & "|"
            'intTextType = PigText.enmTextType.Unicode : strMenuDefinition &= intTextType & "#" & intTextType.ToString & "|"
            'intTextType = PigText.enmTextType.UTF8 : strMenuDefinition &= intTextType & "#" & intTextType.ToString & "|"
            'intTextType = PigText.enmTextType.GB2312 : strMenuDefinition &= intTextType & "#" & intTextType.ToString & "|"
            Return strMenuDefinition
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function


    Public Sub TextStreamDemo()

        Console.WriteLine("*******************")
        Console.WriteLine("TextStream Demo")
        Console.WriteLine("*******************")
        Do While True
            Console.Clear()
            Me.MenuDefinition = ""
            With Me.PigConsole
                .AddMenuDefinition(Me.MenuDefinition, "OpenTextFile", "Open Text File")
                .AddMenuDefinition(Me.MenuDefinition, "ReadLine")
                .AddMenuDefinition(Me.MenuDefinition, "ReadAll")
                .AddMenuDefinition(Me.MenuDefinition, "WriteBlankLines")
                .AddMenuDefinition(Me.MenuDefinition, "Write")
                .AddMenuDefinition(Me.MenuDefinition, "WriteLine")
                .AddMenuDefinition(Me.MenuDefinition, "Close")
            End With
            Me.PigConsole.SimpleMenu("TextStream Demo", Me.MenuDefinition, Me.MenuKey, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Dim strData As String = ""
            Select Case Me.MenuKey
                Case ""
                    Exit Do
                Case "Write", "WriteLine"
                    If Me.TextStream Is Nothing Then
                        Console.WriteLine("TextStream Is Nothing")
                    Else
                        Me.PigConsole.GetLine("Input Text", Me.AnyText)
                        Select Case Me.MenuKey
                            Case "Write"
                                Me.TextStream.Write(Me.AnyText)
                            Case "WriteLine"
                                Me.TextStream.WriteLine(Me.AnyText)
                        End Select
                        If Me.TextStream.LastErr <> "" Then Console.WriteLine(Me.TextStream.LastErr)
                    End If
                Case "WriteBlankLines"
                    If Me.TextStream Is Nothing Then
                        Console.WriteLine("TextStream Is Nothing")
                    Else
                        Me.TextStream.WriteBlankLines(1)
                        If Me.TextStream.LastErr <> "" Then Console.WriteLine(Me.TextStream.LastErr)
                    End If
                Case "Close"
                    If Me.TextStream Is Nothing Then
                        Console.WriteLine("TextStream Is Nothing")
                    Else
                        Me.TextStream.Close()
                        If Me.TextStream.LastErr <> "" Then Console.WriteLine(Me.TextStream.LastErr)
                    End If
                Case "ReadAll"
                    If Me.TextStream Is Nothing Then
                        Console.WriteLine("TextStream Is Nothing")
                    Else
                        Console.WriteLine(Me.TextStream.ReadAll)
                    End If
                Case "ReadLine"
                    If Me.TextStream Is Nothing Then
                        Console.WriteLine("TextStream Is Nothing")
                    Else
                        Console.WriteLine(Me.TextStream.ReadLine)
                    End If
                Case "OpenTextFile"
                    Me.PigConsole.GetLine("Enter text file path", Me.SrcFile)
                    Dim intIOMode As PigFileSystem.IOMode = Me.PigFunc.GECLng(Me.PigConsole.SelectMenuOfEnumeration(PigConsole.EnmWhatTypeOfMenuDefinition.PigFileSystem_IOMode))
                    Dim intTextType As PigText.enmTextType = Me.PigFunc.GECLng(Me.PigConsole.SelectMenuOfEnumeration(PigConsole.EnmWhatTypeOfMenuDefinition.PigText_EnmTextType))
                    Me.TextStream = Me.PigFS.OpenTextFile(Me.SrcFile, intIOMode, intTextType, True)
            End Select
            Me.PigConsole.DisplayPause()
        Loop
    End Sub

    Public Sub PigFSDemo()

        Console.WriteLine("*******************")
        Console.WriteLine("PigFileSystem Demo")
        Console.WriteLine("*******************")
        Do While True
            Console.Clear()
            Me.MenuDefinition2 = ""
            Me.MenuDefinition2 &= "GetPigFolder#GetPigFolder|"
            Me.MenuDefinition2 &= "GetPigFile#GetPigFile|"
            Me.MenuDefinition2 &= "CopyFile#CopyFile|"
            Me.MenuDefinition2 &= "MoveFile#MoveFile|"
            Me.MenuDefinition2 &= "PigFolder_RefSubPigFolders#PigFolder.RefSubPigFolders|"
            Me.MenuDefinition2 &= "PigFolder_RefPigFiles#PigFolder.RefPigFiles|"
            Me.MenuDefinition2 &= "PigFolder_FindSubFolders#PigFolder.FindSubFolders|"
            Me.PigConsole.SimpleMenu("PigFileSystem Demo", Me.MenuDefinition2, Me.MenuKey2, PigConsole.EnmSimpleMenuExitType.QtoUp)
            Dim strData As String = ""
            Select Case Me.MenuKey2
                Case ""
                    Exit Do
                Case "MoveFile"
                    Me.PigConsole.GetLine("SourceFile=", Me.SrcFile)
                    Me.PigConsole.GetLine("TargetFile=", Me.TarFile)
                    Me.IsOverwrite = Me.PigConsole.IsYesOrNo("IsOverwrite=")
                    Console.WriteLine("MoveFile")
                    Me.Ret = Me.PigFS.MoveFile(Me.SrcFile, Me.TarFile, Me.IsOverwrite)
                    Console.WriteLine(Me.Ret)
                Case "CopyFile"
                    Me.PigConsole.GetLine("SourceFile=", Me.SrcFile)
                    Me.PigConsole.GetLine("TargetFile=", Me.TarFile)
                    Me.IsOverwrite = Me.PigConsole.IsYesOrNo("IsOverwrite=")
                    Console.WriteLine("CopyFile")
                    Me.Ret = Me.PigFS.CopyFile(Me.SrcFile, Me.TarFile, Me.IsOverwrite)
                    Console.WriteLine(Me.Ret)
                Case "PigFolder_FindSubFolders"
                    If Me.PigFolder Is Nothing Then
                        Console.WriteLine("Me.PigFolder Is Nothing")
                    Else
                        Dim bolIsDeep As Boolean, oPigFolders As PigFolders = Nothing
                        bolIsDeep = Me.PigConsole.IsYesOrNo("IsDeep")
                        Console.WriteLine("FindSubFolders")
                        Me.Ret = Me.PigFolder.FindSubFolders(bolIsDeep, oPigFolders)
                        Console.WriteLine(Me.Ret)
                        Console.WriteLine("AllFilesSize=" & oPigFolders.AllFilesSize)
                        Console.WriteLine("AllFiles=" & oPigFolders.AllFiles)
                        Dim strPigMD5 As String = ""
                        oPigFolders.GetAllFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileSize_Files)
                        Console.WriteLine("GetAllFastPigMD5(FileSize_Files)=" & strPigMD5)
                        oPigFolders.GetAllFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileSize_Files_UpdateTime)
                        Console.WriteLine("GetAllFastPigMD5(FileSize_Files_UpdateTime)=" & strPigMD5)
                        oPigFolders.GetAllFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileFastPigMD5)
                        Console.WriteLine("GetAllFastPigMD5(FileFastPigMD5)=" & strPigMD5)
                        If oPigFolders IsNot Nothing Then
                            If Me.PigConsole.IsYesOrNo("Is show PigFolders") = True Then
                                For Each oPigFolder In oPigFolders
                                    With oPigFolder
                                        Console.WriteLine("----------------------")
                                        Console.WriteLine("FolderPath=" & .FolderPath)
                                    End With
                                Next
                            End If
                        End If
                    End If
                Case "PigFolder_RefPigFiles"
                    If Me.PigFolder Is Nothing Then
                        Console.WriteLine("Me.PigFolder Is Nothing")
                    Else
                        Console.WriteLine("RefPigFiles")
                        Me.Ret = Me.PigFolder.RefPigFiles
                        Console.WriteLine(Me.Ret)
                        For Each oPigFile In Me.PigFolder.PigFiles
                            With oPigFile
                                Console.WriteLine("----------------------")
                                Console.WriteLine("FileTitle=" & .FileTitle)
                                Console.WriteLine("FilePath=" & .FilePath)
                                Console.WriteLine("CreationTime=" & .CreationTime)
                                Console.WriteLine("UpdateTime=" & .UpdateTime)
                            End With
                        Next
                    End If
                Case "PigFolder_RefSubPigFolders"
                    If Me.PigFolder Is Nothing Then
                        Console.WriteLine("Me.PigFolder Is Nothing")
                    Else
                        Console.WriteLine("RefSubPigFolders")
                        Me.Ret = Me.PigFolder.RefSubPigFolders
                        Console.WriteLine(Me.Ret)
                        For Each oPigFolder In Me.PigFolder.SubPigFolders
                            With oPigFolder
                                Console.WriteLine("----------------------")
                                Console.WriteLine("FolderName=" & .FolderName)
                                Console.WriteLine("FolderPath=" & .FolderPath)
                                Console.WriteLine("CreationTime=" & .CreationTime)
                                Console.WriteLine("UpdateTime=" & .UpdateTime)
                                Console.WriteLine("IsRootFolder=" & .IsRootFolder)
                            End With
                        Next
                    End If
                Case "GetPigFolder"
                    Me.PigConsole.GetLine("Input folder path", Me.FolderPath)
                    Me.PigFolder = Me.PigFS.GetPigFolder(Me.FolderPath)
                    With Me.PigFolder
                        Console.WriteLine("FolderName=" & .FolderName)
                        Console.WriteLine("FolderPath=" & .FolderPath)
                        Console.WriteLine("CreationTime=" & .CreationTime)
                        Console.WriteLine("UpdateTime=" & .UpdateTime)
                        Console.WriteLine("IsRootFolder=" & .IsRootFolder)
                        Console.WriteLine("FilesSize=" & .FilesSize)
                        Console.WriteLine("PigFiles.Count=" & .PigFiles.Count)
                        Dim strPigMD5 As String = ""
                        .GetFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileSize_Files)
                        Console.WriteLine("GetFastPigMD5(FileSize_Files)=" & strPigMD5)
                        .GetFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileSize_Files_UpdateTime)
                        Console.WriteLine("GetFastPigMD5(FileSize_Files_UpdateTime)=" & strPigMD5)
                        .GetFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.FileFastPigMD5)
                        Console.WriteLine("GetFastPigMD5(FileFastPigMD5)=" & strPigMD5)
                        .GetFastPigMD5(strPigMD5, PigFolder.EnmGetFastPigMD5Type.CurrDirInfo)
                        Console.WriteLine("GetFastPigMD5(CurrDirInfo)=" & strPigMD5)
                    End With
                Case "GetPigFile"
                    Me.PigConsole.GetLine("Input file path", Me.FilePath)
                    Me.PigFile = Me.PigFS.GetPigFile(Me.FilePath)
                    With Me.PigFile
                        Console.WriteLine("FileTitle=" & .FileTitle)
                        Console.WriteLine("FilePath=" & .FilePath)
                        Console.WriteLine("CreationTime=" & .CreationTime)
                        Console.WriteLine("UpdateTime=" & .UpdateTime)
                    End With
            End Select
            'If Me.Ret <> "OK" Then
            '    Console.WriteLine(Me.Ret)
            'Else
            '    Me.PigFunc.SaveTextToFile(Me.FilePath, strData)
            'End If
            Me.PigConsole.DisplayPause()
        Loop
    End Sub



End Class
