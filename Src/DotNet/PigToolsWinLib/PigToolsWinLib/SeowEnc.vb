'**********************************
'* Name: SeowEnc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 这是一个简单的压缩和移位加密算法，可以减少密文的长度，以及每次加密结果的密文都不相同，不要单独使用本加密算法，因为很容易破解，建议与其他加密算法一起使用，如AES、RSA和3DES等。|This is a simple compression and shift encryption algorithm, which can reduce the length of the ciphertext, and the ciphertext of each encryption result is different. Do not use this encryption algorithm alone, because it is easy to crack. It is recommended to use it together with other encryption algorithms, such as AES, RSA and 3DES.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 11/9/2022
'* 1.1  12/9/2022   Modify EmnComprssType,Encrypt
'* 1.2  24/9/2022   Add IsRandAdd
'* 1.3  8/11/2022   Modify mDecrypt,mEncrypt
'************************************
''' <summary>
''' Seow encryption processing class|萧氏加密处理类
''' </summary>
Public Class SeowEnc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.2"

    Private mabEncKey As Byte()
    Public Sub New(ComprssType As EmnComprssType)
        MyBase.New(CLS_VERSION)
        Me.ComprssType = ComprssType
    End Sub

    Public Enum EmnComprssType
        NoComprss = 0
        AutoComprss = 1
        Over328ToComprss = 2
        PigCompressor = 3
    End Enum

    ''' <summary>
    ''' 是否随机增加位移，是否每次产生的密文长度不同|Whether the displacement is increased randomly, and whether the length of ciphertext generated each time is different
    ''' </summary>
    ''' <returns></returns>
    Public Property IsRandAdd As Boolean = True

    Public ReadOnly Property ComprssType As EmnComprssType

    Private mIsLoadEncKey As Boolean
    Public ReadOnly Property IsLoadEncKey As Boolean
        Get
            IsLoadEncKey = mIsLoadEncKey
        End Get
    End Property

    Private Function mLoadEncKey(EncKey As Byte()) As String
        Dim LOG As New PigStepLog("mLoadEncKey")
        Try
            Select Case EncKey.Length
                Case 1 To 256
                    mabEncKey = EncKey
                    Me.mIsLoadEncKey = True
                Case Else
                    Throw New Exception("Invalid EncKey")
            End Select
            Return "OK"
        Catch ex As Exception
            Me.mIsLoadEncKey = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function LoadEncKey(EncKey As Byte()) As String
        Return Me.mLoadEncKey(EncKey)
    End Function

    Public Function LoadEncKey(Base64EncKey As String) As String
        Dim LOG As New PigStepLog("LoadEncKey")
        Try
            LOG.StepName = "New PigBytes"
            Dim pbEncKey As New PigBytes(Base64EncKey)
            If pbEncKey.LastErr <> "" Then Throw New Exception(pbEncKey.LastErr)
            LOG.StepName = "mLoadEncKey"
            LOG.Ret = Me.mLoadEncKey(pbEncKey.Main)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            pbEncKey = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function MkEncKey(EncKeyLen As Integer, InitKey As Byte(), ByRef EncKey As Byte()) As String
        Return Me.mMkEncKey(EncKeyLen, InitKey, EncKey)
    End Function

    Public Function MkEncKey(EncKeyLen As Integer, Base64InitKey As String, ByRef Base64EncKey As String) As String
        Dim LOG As New PigStepLog("MkEncKey")
        Try
            LOG.StepName = "New PigBytes(Base64InitKey)"
            Dim pbInitKey As New PigBytes(Base64InitKey)
            If pbInitKey.LastErr <> "" Then Throw New Exception(pbInitKey.LastErr)
            Dim abEncKey(0) As Byte
            LOG.StepName = "mMkEncKey"
            LOG.Ret = Me.mMkEncKey(EncKeyLen, pbInitKey.Main, abEncKey)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Dim pbEncKey As New PigBytes(mabEncKey)
            Base64EncKey = pbEncKey.Base64Str
            pbEncKey = Nothing
            pbInitKey = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mMkEncKey(EncKeyLen As Integer, InitKey As Byte(), ByRef EncKey As Byte()) As String
        Try
            Select Case EncKeyLen
                Case 1 To 256
                Case Else
                    Throw New Exception("Key length cannot be greater than " & EncKeyLen.ToString)
            End Select
            If InitKey IsNot Nothing Then
                If InitKey.Length > EncKeyLen Then ReDim InitKey(EncKeyLen - 1)
                ReDim EncKey(EncKeyLen - 1)
                For i = 0 To EncKeyLen - 1
                    If i < InitKey.Length Then
                        EncKey(i) = InitKey(i)
                    Else
                        EncKey(i) = Me.mGetRandNum(0, 255)
                    End If
                Next
            End If
            mabEncKey = EncKey
            Me.mIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mMkEncKey", ex)
        End Try
    End Function

    Private Function mGetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function

    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte(), Optional ByRef CompressRate As Decimal = 1) As String
        Return Me.mEncrypt(SrcBytes, EncBytes, CompressRate)
    End Function

    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte(), ByRef CompressRate As Decimal) As String
        Dim strStepName As String = "", intPos As Integer = 0
        Try
            If Me.mIsLoadEncKey = False Then Throw New Exception("Key not imported")
            Dim intSrcLen As Integer = SrcBytes.Length
            Dim intComprssType As EmnComprssType
            Select Case Me.ComprssType
                Case EmnComprssType.AutoComprss, EmnComprssType.Over328ToComprss
                    strStepName = "Check SrcBytes Len"
                    If intSrcLen > 328 Then
                        intComprssType = EmnComprssType.PigCompressor
                    Else
                        intComprssType = EmnComprssType.NoComprss
                    End If
                    If Me.ComprssType = EmnComprssType.AutoComprss Then
                        If intSrcLen > 3280 Then
                            Dim intChkTimes As Integer = CInt(intSrcLen / 3280), decCompressRate As Decimal = 0
                            If intChkTimes < 1 Then
                                intChkTimes = 1
                            ElseIf intChkTimes > 10 Then
                                intChkTimes = 10
                            End If
                            Dim pcTest As New PigCompressor
                            For i = 1 To intChkTimes
                                Dim intBegin As Integer = Me.mGetRandNum(0, intSrcLen - 328 + 1)
                                Dim abTest(327) As Byte
                                For j = 0 To 327
                                    abTest(j) = SrcBytes(intBegin + j)
                                Next
                                Dim abTestEnc As Byte() = pcTest.Compress(abTest)
                                If pcTest.LastErr = "" Then
                                    If decCompressRate = 0 Then
                                        decCompressRate = CDec(abTestEnc.Length) / 328
                                    Else
                                        decCompressRate = (CDec(abTestEnc.Length) / 328 + decCompressRate) / 2
                                    End If
                                End If
                            Next
                            If decCompressRate < 0.85 Then
                                intComprssType = EmnComprssType.PigCompressor
                            Else
                                intComprssType = EmnComprssType.NoComprss
                            End If
                        End If
                    End If
                Case EmnComprssType.NoComprss, EmnComprssType.PigCompressor
                    intComprssType = Me.ComprssType
                Case Else
                    Throw New Exception("Invalid ComprssType is " & Me.ComprssType.ToString)
            End Select
            Dim strRet As String = "", bytAdd As Byte
            If Me.IsRandAdd = True Then
                bytAdd = Me.mGetRandNum(0, 255)
            Else
                bytAdd = 163
            End If
            Select Case intComprssType
                Case EmnComprssType.PigCompressor
                    Dim pcSrc As New PigCompressor
                    strStepName = "Compress"
                    EncBytes = pcSrc.Compress(SrcBytes)
                    If pcSrc.LastErr <> "" Then Throw New Exception(pcSrc.LastErr)
                    pcSrc = Nothing
                Case EmnComprssType.NoComprss
                    EncBytes = SrcBytes
            End Select
            strStepName = "Set Len"
            Dim intDataLen As Integer = EncBytes.Length, bytEncKeyLen As Byte = CByte(Me.mabEncKey.Length - 1)
            strStepName = "For Next"
            For intPos = 0 To intDataLen - 1
                Dim intItem As Integer = EncBytes(intPos), bytMod As Byte = intPos Mod bytEncKeyLen
                intItem += Me.mabEncKey(bytMod)
                If intItem > 255 Then intItem -= 256
                intItem += bytAdd
                If intItem > 255 Then intItem -= 256
                EncBytes(intPos) = intItem
            Next
            strStepName = "ReDim"
            ReDim Preserve EncBytes(intDataLen + 1)
            EncBytes(intDataLen) = intComprssType
            EncBytes(intDataLen + 1) = bytAdd
            CompressRate = CDec(EncBytes.Length) / intSrcLen
            Return "OK"
        Catch ex As Exception
            EncBytes = Nothing
            CompressRate = 0
            strStepName &= ",Pos=" & intPos
            Return Me.GetSubErrInf("mEncrypt", strStepName, ex)
        End Try
    End Function

    Public Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Return Me.mDecrypt(EncBytes, UnEncBytes)
    End Function

    Public Function Decrypt(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As PigText
            Dim abUnEncBytes As Byte()
            ReDim abUnEncBytes(0)
            strStepName = "mDecrypt"
            strRet = Me.mDecrypt(EncBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oPigText.Text
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", strStepName, ex)
        End Try
    End Function

    Private Function mDecrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Dim strStepName As String = "", intPos As Integer = 0
        Try
            If Me.mIsLoadEncKey = False Then Throw New Exception("Key not imported")
            Dim intDataLen As Integer = EncBytes.Length - 2
            strStepName = "Set EmnComprssType"
            Dim intComprssType As EmnComprssType = EncBytes(intDataLen)
            strStepName = "Set Add"
            Dim bytAdd As Byte = EncBytes(intDataLen + 1)
            strStepName = "Set EncKeyLen"
            Dim bytEncKeyLen As Byte = CByte(Me.mabEncKey.Length - 1)
            strStepName = "For Next"
            For intPos = 0 To intDataLen - 1
                Dim intItem As Integer = EncBytes(intPos), bytMod As Byte = intPos Mod bytEncKeyLen
                intItem -= Me.mabEncKey(bytMod)
                If intItem < 0 Then intItem += 256
                intItem -= bytAdd
                If intItem < 0 Then intItem += 256
                EncBytes(intPos) = intItem
            Next
            strStepName = "ReDim"
            ReDim Preserve EncBytes(intDataLen - 1)
            Select Case intComprssType
                Case EmnComprssType.NoComprss
                    UnEncBytes = EncBytes
                Case EmnComprssType.PigCompressor
                    Dim pcEnc As New PigCompressor
                    strStepName = "Depress"
                    UnEncBytes = pcEnc.Depress(EncBytes)
                    If pcEnc.LastErr <> "" Then Throw New Exception(pcEnc.LastErr)
                    pcEnc = Nothing
            End Select
            Return "OK"
        Catch ex As Exception
            UnEncBytes = Nothing
            strStepName &= ",Pos=" & intPos
            Return Me.GetSubErrInf("mDecrypt", strStepName, ex)
        End Try
    End Function

    Public Function Encrypt(SrcStr As String, TextType As PigText.enmTextType, ByRef Base64EncStr As String, Optional ByRef CompressRate As Decimal = 1) As String
        Dim strStepName As String = ""
        Try
            Dim strRet As String
            Dim abEncBytes(0) As Byte
            strStepName = "New PigText"
            Dim ptSrc As New PigText(SrcStr, TextType)
            If ptSrc.LastErr <> "" Then Throw New Exception(ptSrc.LastErr)
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(ptSrc.TextBytes, abEncBytes, CompressRate)
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "ToBase64String"
            Base64EncStr = Convert.ToBase64String(abEncBytes)
            ptSrc = Nothing
            Return "OK"
        Catch ex As Exception
            Base64EncStr = ""
            Return Me.GetSubErrInf("Encrypt", ex)
        End Try
    End Function

    Public Function Encrypt(SrcBytes As Byte(), ByRef Base64EncStr As String, Optional ByRef CompressRate As Decimal = 1) As String
        Dim strStepName As String = ""
        Try
            Dim strRet As String
            Dim abEncBytes(0) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(SrcBytes, abEncBytes, CompressRate)
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "ToBase64String"
            Base64EncStr = Convert.ToBase64String(abEncBytes)
            Return "OK"
        Catch ex As Exception
            Base64EncStr = ""
            Return Me.GetSubErrInf("Encrypt", ex)
        End Try
    End Function

    Public Function Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As PigText
            oPigText = New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            Dim abUnEncBytes(0) As Byte
            strStepName = "mDecrypt"
            strRet = Me.mDecrypt(oPigText.TextBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oPigText.Text
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", strStepName, ex)
        End Try
    End Function

End Class
