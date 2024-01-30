'**********************************
'* Name: PigRsa
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: RSA Processing Class|RSA处理类
'* Home Url: https://en.seowphong.com
'* Version: 1.5
'* Create Time: 17/10/2019
'1.0.2  2019-10-18 Decrypt和Encrypt重载函数
'1.0.3  2019-11-17 增加 密钥 bytes 数据处理
'1.1  16/10/2021 Modify LoadPubKey,mDecrypt,mEncrypt
'1.2  14/11/2021 Add MkEncKey,LoadEncKey, modify Decrypt
'1.3  15/11/2021 Add SignData,mSignData,mVerifyData,VerifyData
'1.5  17/1/2024 Add Decrypt,Encrypt
'**********************************

Imports System.Security.Cryptography
''' <summary>
''' RSA Processing Class|RSA处理类
''' </summary>
Public Class PigRsa
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "5" & "." & "8"

    ''' <summary>密钥组成部分结构长度</summary>
    Private Structure struEncKeyPartLen
        Dim Modulus As Short
        Dim Exponent As Short
        ''' <remarks>以下的可能为0，只有选择了私钥才有</remarks>
        Dim P As Short
        Dim Q As Short
        Dim DP As Short
        Dim DQ As Short
        Dim InverseQ As Short
        Dim D As Short
    End Structure

    Private mrsaMain As RSACryptoServiceProvider = New RSACryptoServiceProvider
    Private mbolIsLoadEncKey As Boolean = False

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <remarks>是否已导入密钥</remarks>
    Public ReadOnly Property IsLoadEncKey As Boolean
        Get
            IsLoadEncKey = mbolIsLoadEncKey
        End Get
    End Property

    ''' <summary>
    ''' 解密|Decrypt
    ''' </summary>
    ''' <param name="EncBase64Str">加密后的字节数组的Base64|Base64 of encrypted byte array</param>
    ''' <param name="UnEncBytes">解密后的源字节数组|Decrypted source byte array</param>
    ''' <param name="TextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Function Decrypt(EncBase64Str As String, ByRef UnEncBytes As Byte(), TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            strStepName = "mDecrypt"
            strRet = Me.mDecrypt(oPigText.TextBytes, UnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", strStepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' Decryption, generally public key encryption and private key decryption|解密，一般是公钥加密，私钥解密
    ''' </summary>
    ''' <param name="EncBase64Str">Base64 of ciphertext|密文的Base64</param>
    ''' <param name="UnEncStr">Decrypted string|解密后的字符串</param>
    ''' <param name="TextType">Text encoding type|文本编码类型</param>
    ''' <returns></returns>
    Public Function Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oGEText As PigText
            oGEText = New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            Dim abUnEncBytes(-1) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mDecrypt(oGEText.TextBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oGEText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oGEText.Text
            oGEText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", ex, True)
        End Try
    End Function

    ''' <summary>
    ''' Decryption, generally public key encryption and private key decryption|解密，一般是公钥加密，私钥解密
    ''' </summary>
    ''' <param name="EncBytes">Byte array of ciphertext|密文的字节数组</param>
    ''' <param name="UnEncStr">Decrypted string|解密后的字符串</param>
    ''' <param name="TextType">Text encoding type|文本编码类型</param>
    ''' <returns></returns>
    Public Function Decrypt(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oGEText As PigText
            Dim abUnEncBytes(-1) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mDecrypt(EncBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oGEText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oGEText.Text
            oGEText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", ex, True)
        End Try
    End Function

    ''' <summary>
    ''' Decryption, generally public key encryption and private key decryption|解密，一般是公钥加密，私钥解密
    ''' </summary>
    ''' <param name="EncBytes">Byte array of ciphertext|密文的字节数组</param>
    ''' <param name="UnEncBytes">Byte array after decryption|解密后的字节数组</param>
    ''' <returns></returns>
    Public Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Decrypt = Me.mDecrypt(EncBytes, UnEncBytes)
    End Function


    ''' <remarks>解密数据（内部）</remarks>
    Private Function mDecrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Dim strStepName As String = ""
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            UnEncBytes = mrsaMain.Decrypt(EncBytes, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mDecrypt", ex, False)
        End Try
    End Function

    ''' <summary>
    ''' Encryption, generally public key encryption and private key decryption|加密，一般是公钥加密，私钥解密
    ''' Keep this interface for compatibility with previous versions|保留这个接口是为了兼容以前的版本 
    ''' </summary>
    ''' <param name="SrcStr">Original string|原文字符串</param>
    ''' <param name="TextType">Text encoding type|文本编码类型</param>
    ''' <param name="EncBase64Str">Encrypted Base64|加密后的Base64</param>
    ''' <returns></returns>
    Public Function Encrypt(SrcStr As String, TextType As PigText.enmTextType, ByRef EncBase64Str As String) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As New PigText(SrcStr, TextType)
            Dim abEncBytes As Byte()
            ReDim abEncBytes(-1)
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(oPigText.TextBytes, abEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            EncBase64Str = Convert.ToBase64String(abEncBytes)
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Encrypt", ex, False)
        End Try
    End Function

    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBase64Str As String) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim abEncBytes As Byte()
            ReDim abEncBytes(0)
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(SrcBytes, abEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            EncBase64Str = Convert.ToBase64String(abEncBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Encrypt", ex, False)
        End Try
    End Function


    ''' <summary>
    ''' Encryption, generally public key encryption and private key decryption|加密，一般是公钥加密，私钥解密
    ''' </summary>
    ''' <param name="SrcBytes">Original byte array|原文的字节数组</param>
    ''' <param name="EncBytes">Encrypted byte array|加密后的字节数组</param>
    ''' <returns></returns>
    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Return Me.mEncrypt(SrcBytes, EncBytes)
    End Function


    ''' <remarks>加密数据（内部）</remarks>
    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Dim strStepName As String = ""
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            strStepName = "mrsaMain.Encrypt"
            EncBytes = mrsaMain.Encrypt(SrcBytes, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mEncrypt", ex, False)
        End Try
    End Function


    ''' <summary>
    ''' Generate Key|生成密钥
    ''' </summary>
    ''' <param name="OutEncKeyXml">Output key XML, including public key and private key|输出的密钥XML，包括公钥和私钥</param>
    ''' <returns></returns>
    Public Function MkEncKey(ByRef OutEncKeyXml As String) As String
        Try
            OutEncKeyXml = mrsaMain.ToXmlString(True)
            Return "OK"
        Catch ex As Exception
            OutEncKeyXml = ""
            Return Me.GetSubErrInf("MkEncKey", ex)
        End Try
    End Function

    ''' <summary>
    ''' Generate Key|生成密钥
    ''' </summary>
    ''' <param name="OutPubKeyXml">Output public key XML|输出的公钥XML</param>
    ''' <param name="OutPrivateKeyXml">Output private key XML|输出的私钥XML</param>
    ''' <returns></returns>
    Public Function MkEncKey(ByRef OutPrivateKeyXml As String, ByRef OutPublicKeyXml As String) As String
        Dim LOG As New PigStepLog("MkEncKey")
        Try
            LOG.StepName = "Mk OutEncKeyXml"
            OutPrivateKeyXml = mrsaMain.ToXmlString(True)
            Dim pxMain As New PigXml(False)
            LOG.StepName = "Main.SetMainXml"
            LOG.Ret = pxMain.SetMainXml(OutPrivateKeyXml)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Main.InitXmlDocument"
            LOG.Ret = pxMain.InitXmlDocument()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "Mk DecryptOnlyKey"
            Dim pxDecryp As New PigXml(False)
            With pxDecryp
                .AddEleLeftSign("RSAKeyValue")
                .AddEle("Modulus", pxMain.GetXmlDocText("RSAKeyValue.Modulus"))
                .AddEle("Exponent", pxMain.GetXmlDocText("RSAKeyValue.Exponent"))
                .AddEleRightSign("RSAKeyValue")
                If .LastErr <> "" Then Throw New Exception(.LastErr)
                OutPublicKeyXml = .MainXmlStr
                If OutPublicKeyXml = "" Then Throw New Exception("Generate failed.")
            End With
            pxMain = Nothing
            pxDecryp = Nothing
            Return "OK"
        Catch ex As Exception
            OutPrivateKeyXml = ""
            OutPublicKeyXml = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>The old key generation interface, MkEncKey is recommended|旧的生成密钥接口，建议使用MkEncKey</summary>
    ''' <param name="IsIncPriKey">是否生成私钥</param>
    ''' <param name="OutXmlKey">输出为XML</param>
    Public Function MkPubKey(IsIncPriKey As Boolean, ByRef OutXmlKey As String) As String
        Try
            OutXmlKey = mrsaMain.ToXmlString(IsIncPriKey)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkPubKey", ex, False)
        End Try
    End Function

    ''' <summary>The old key generation interface, MkEncKey is recommended|旧的生成密钥接口，建议使用MkEncKey</summary>
    ''' <param name="IsIncPriKey">是否生成私钥</param>
    ''' <param name="OutBytes">输出为字节数组</param>
    Public Function MkPubKey(IsIncPriKey As Boolean, ByRef OutBytes As Byte()) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strXmlKey As String = mrsaMain.ToXmlString(IsIncPriKey)
            Dim ekplAny As struEncKeyPartLen
            Dim oGEXml As New PigXml(False), oGEText As PigText, strItem As String
            strXmlKey = oGEXml.XmlGetStr(strXmlKey, "RSAKeyValue")
            Dim oGEBytes As New PigBytes
            oGEBytes.RestPos()
            With ekplAny
                strItem = oGEXml.XmlGetStr(strXmlKey, "Modulus")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .Modulus = oGEText.TextBytes.Length
                strStepName = ".SetValue(Modulus)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "Exponent")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8)
                .Exponent = oGEText.TextBytes.Length
                strStepName = ".SetValue(Exponent)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "P")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .P = oGEText.TextBytes.Length
                strStepName = ".SetValue(P)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "Q")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .Q = oGEText.TextBytes.Length
                strStepName = ".SetValue(Q)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "DP")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .DP = oGEText.TextBytes.Length
                strStepName = ".SetValue(DP)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "DQ")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .DQ = oGEText.TextBytes.Length
                strStepName = ".SetValue(DQ)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "InverseQ")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .InverseQ = oGEText.TextBytes.Length
                strStepName = ".SetValue(InverseQ)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "D")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .D = oGEText.TextBytes.Length
                strStepName = ".SetValue(D)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
            End With
            Dim oStru2Bytes As New Stru2Bytes
            strStepName = ".Stru2Bytes(ekplAny, OutBytes)"
            strRet = oStru2Bytes.Stru2Bytes(ekplAny, OutBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Dim gbOutBytes As New PigBytes
            strStepName = ".SetValue(CInt(OutBytes.Length))"
            strRet = gbOutBytes.SetValue(CInt(OutBytes.Length))
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = ".SetValue(OutBytes)"
            strRet = gbOutBytes.SetValue(OutBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = ".SetValue(oGEBytes.Main)"
            strRet = gbOutBytes.SetValue(oGEBytes.Main)
            If strRet <> "OK" Then Throw New Exception(strRet)
            OutBytes = gbOutBytes.Main
            oStru2Bytes = Nothing
            oGEXml = Nothing
            oGEBytes = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkPubKey", ex, False)
        End Try
    End Function

    ''' <summary>导入密钥</summary>
    ''' <param name="KeyBytes">密钥数组</param>
    Public Overloads Function LoadPubKey(KeyBytes As Byte()) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            strStepName = "New PigBytes"
            Dim oGEBytes As New PigBytes(KeyBytes)
            If oGEBytes.LastErr <> "" Then Throw New Exception(oGEBytes.LastErr)
            Dim oStru2Bytes As New Stru2Bytes
            Dim ekplAny As struEncKeyPartLen
            Dim intStruLen As Integer = oGEBytes.GetInt32Value
            strStepName = "Bytes2Stru"
            strRet = oStru2Bytes.Bytes2Stru(oGEBytes.GetBytesValue(intStruLen), ekplAny.GetType, ekplAny)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Dim oGEXml As New PigXml(False), oGEText As PigText
            strStepName = "Mk Xml"
            oGEXml.AddEleLeftSign("RSAKeyValue")
            strStepName = "New PigText(ekplAny.Modulus)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Modulus))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Modulus", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.Exponent)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Exponent), PigText.enmTextType.UTF8)
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Exponent", oGEText.Text)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.P)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.P))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("P", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.Q)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Q))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Q", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.DP)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.DP))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("DP", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.DQ)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.DQ))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("DQ", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.InverseQ)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.InverseQ))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("InverseQ", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.D)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.D))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("D", oGEText.Base64)
            oGEText = Nothing
            oGEXml.AddEleRightSign("RSAKeyValue")
            strStepName = "FromXmlString"
            mrsaMain.FromXmlString(oGEXml.MainXmlStr)
            mbolIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadPubKey", ex)
        End Try
    End Function


    ''' <summary>
    ''' Import the key, which is generated by MkEncKey|导入密钥，密钥通过MkEncKey生成
    ''' </summary>
    ''' <param name="XmlKey">The XML of the key can be either the key that can be encrypted or decrypted, or the key that can only be decrypted|密钥的XML，可以是即可加密也可解密的密钥，也可以是只能解密的密钥</param>
    ''' <returns></returns>
    Public Function LoadEncKey(XmlKey As String) As String
        Try
            mrsaMain.FromXmlString(XmlKey)
            mbolIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadEncKey", ex)
        End Try
    End Function


    ''' <summary>The old import key interface, LoadEncKey is recommended|旧的导入密钥接口，建议使用LoadEncKey</summary>
    ''' <param name="XmlKey">XML密钥</param>
    Public Overloads Function LoadPubKey(XmlKey As String) As String
        Dim strStepName As String = ""
        Try
            strStepName = "FromXmlString"
            mrsaMain.FromXmlString(XmlKey)
            mbolIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadPubKey", strStepName, ex)
        End Try
    End Function

    Protected Overrides Sub Finalize()
        mrsaMain = Nothing
        MyBase.Finalize()
    End Sub


    Private Function mSignData(SrcBytes As Byte(), ByRef SignBase64 As String) As String
        Try
            If mbolIsLoadEncKey = False Then Throw New Exception("Key not imported")
            Dim abSign As Byte() = mrsaMain.SignData(SrcBytes, "SHA1")
            SignBase64 = Convert.ToBase64String(abSign)
            Return "OK"
        Catch ex As Exception
            SignBase64 = ""
            Return Me.GetSubErrInf("mSignData", ex)
        End Try
    End Function

    ''' <summary>
    ''' Sign the input data, sign the private key, and verify the public key|对输入数据签名，为私钥签名，公钥验证
    ''' </summary>
    ''' <param name="SrcBytes">Input byte array|输入字节数组</param>
    ''' <param name="SignBase64">Output signature result SHA1 base64|输出签名结果SHA1的Base64</param>
    ''' <returns></returns>
    Public Function SignData(SrcBytes As Byte(), ByRef SignBase64 As String) As String
        Return Me.mSignData(SrcBytes, SignBase64)
    End Function

    ''' <summary>
    ''' Sign the input data, sign the private key, and verify the public key|对输入数据签名，为私钥签名，公钥验证
    ''' </summary>
    ''' <param name="SrcStr">Original string|原文字符串</param>
    ''' <param name="TextType">Text encoding type|文本编码类型</param>
    ''' <param name="SignBase64">Output signature result SHA1 base64|输出签名结果SHA1的Base64</param>
    ''' <returns></returns>
    Public Function SignData(SrcStr As String, TextType As PigText.enmTextType, ByRef SignBase64 As String) As String
        Try
            Dim oPigText As New PigText(SrcStr, TextType)
            Dim strRet As String = Me.mSignData(oPigText.TextBytes, SignBase64)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SignData", ex)
        End Try
    End Function

    ''' <summary>
    ''' Verify the signature result of input data, private key signature, public key verification|验证输入数据的签名结果，私钥签名，公钥验证
    ''' </summary>
    ''' <param name="SrcStr">Original string|原文字符串</param>
    ''' <param name="TextType">Text encoding type|文本编码类型</param>
    ''' <param name="SignBase64">Signed SHA1 base64|签名的SHA1 base64</param>
    ''' <param name="IsVerify">Output check signature results|输出检查签名结果</param>
    ''' <returns></returns>
    Public Function VerifyData(SrcStr As String, TextType As PigText.enmTextType, SignBase64 As String, ByRef IsVerify As Boolean) As String
        Try
            Dim oPigText As New PigText(SrcStr, TextType)
            Dim strRet As String = Me.mVerifyData(oPigText.TextBytes, SignBase64, IsVerify)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("VerifyData", ex)
        End Try
    End Function

    ''' <summary>
    ''' Verify the signature result of input data, private key signature, public key verification|验证输入数据的签名结果，私钥签名，公钥验证
    ''' </summary>
    ''' <param name="SrcBytes">Input byte array|输入字节数组</param>
    ''' <param name="SignBase64">Signed SHA1 base64|签名的SHA1 base64</param>
    ''' <param name="IsVerify">Output check signature results|输出检查签名结果</param>
    ''' <returns></returns>
    Public Function VerifyData(SrcBytes As Byte(), SignBase64 As String, ByRef IsVerify As Boolean) As String
        Return Me.mVerifyData(SrcBytes, SignBase64, IsVerify)
    End Function

    Private Function mVerifyData(SrcBytes As Byte(), SignBase64 As String, ByRef IsVerify As Boolean) As String
        Try
            If mbolIsLoadEncKey = False Then Throw New Exception("Key not imported")
            Dim abSign As Byte() = Convert.FromBase64String(SignBase64)
            IsVerify = mrsaMain.VerifyData(SrcBytes, "SHA1", abSign)
            Return "OK"
        Catch ex As Exception
            IsVerify = False
            Return Me.GetSubErrInf("mVerifyData", ex)
        End Try
    End Function


    ''' <summary>
    ''' 加密|encryption
    ''' </summary>
    ''' <param name="SrcString">源字符串</param>
    ''' <param name="EncBytes">加密后的字节数组|Encrypted byte array</param>
    ''' <param name="SrcTextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Function Encrypt(SrcString As String, ByRef EncBytes As Byte(), Optional SrcTextType As PigText.enmTextType = PigText.enmTextType.UTF8) As String
        Dim LOG As New PigStepLog("Encrypt")
        Try
            Dim ptSrc As New PigText(SrcString, SrcTextType)
            LOG.StepName = "mEncrypt"
            LOG.Ret = Me.mEncrypt(ptSrc.TextBytes, EncBytes)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            ptSrc = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 加密|encryption
    ''' </summary>
    ''' <param name="SrcString">源字符串</param>
    ''' <param name="EncBase64Str">加密后的字节数组的Base64|Base64 of encrypted byte array</param>
    ''' <param name="SrcTextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Function Encrypt(SrcString As String, ByRef EncBase64Str As String, Optional SrcTextType As PigText.enmTextType = PigText.enmTextType.UTF8) As String
        Dim LOG As New PigStepLog("Encrypt")
        Try
            Dim ptSrc As New PigText(SrcString, SrcTextType)
            Dim abEnc(-1) As Byte
            LOG.StepName = "mEncrypt"
            LOG.Ret = Me.mEncrypt(ptSrc.TextBytes, abEnc)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            EncBase64Str = Convert.ToBase64String(abEnc)
            ptSrc = Nothing
            Return "OK"
        Catch ex As Exception
            EncBase64Str = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    '''' <summary>
    '''' 私钥签名 
    '''' </summary>
    '''' <returns></returns>
    'Public Shared Function RSA_Sign(prikey As String, KEY As String) As String
    '    Dim key_byt As Byte() = Encoding.UTF8.GetBytes(KEY)
    '    Dim RSA_SP As RSACryptoServiceProvider = New RSACryptoServiceProvider()
    '    RSA_SP.FromXmlString(Encoding.Default.GetString(Convert.FromBase64String(prikey)))
    '    Dim res As Byte() = RSA_SP.SignData(key_byt, "SHA1")
    '    Dim sign As String = Convert.ToBase64String(res)
    '    Return sign
    'End Function

    '私钥签名，公钥验证

    '''' <summary>
    '''' 公钥验证 
    '''' </summary>
    '''' <returns></returns>
    'Public Shared Function RSA_Verify(pubkey As String, KEY As String, sign As String) As Boolean
    '    Dim key_byt As Byte() = Encoding.UTF8.GetBytes(KEY)
    '    Dim RSA_SP As RSACryptoServiceProvider = New RSACryptoServiceProvider()
    '    RSA_SP.FromXmlString(Encoding.Default.GetString(Convert.FromBase64String(pubkey)))
    '    Dim sign_byt As Byte() = Convert.FromBase64String(sign)
    '    Dim Verify_bool As Boolean = RSA_SP.VerifyData(key_byt, "SHA1", sign_byt)
    '    Return Verify_bool
    'End Function


End Class
