'**********************************
'* Name: PigAes
'* Author: Seow Phong
'* License: Copyright (c) 2020-2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: AES Processing Class|AES处理类
'* Home Url: https://en.seowphong.com
'* Version: 1.6
'* Create Time: 2019-10-27
'1.0.2  2019-10-29
'1.0.3  2019-10-31  稳定版本，去掉 EncriptStr 和 DecriptStr
'1.0.5  2019-12-8   修改LoadEncKey
'1.0.6  24/8/2021   Modify mDecrypt,mEncrypt,mMkEncKey
'1.1  16/10/2021   Modify mDecrypt,mEncrypt,LoadEncKey,LoadEncKey,Decrypt
'1.2  11/9/2021   Modify mMkEncKey
'1.3  17/2/2023   Modify LoadEncKey,mEncrypt,mDecrypt
'1.5  28/4/2023   Add Decrypt,Encrypt
'1.6  29/1/2024   Add LoadEncKeySub,DecryptSub
'************************************
Imports System.Security.Cryptography
Imports System.Text
''' <summary>
''' AES Processing Class|AES处理类
''' </summary>
Public Class PigAes
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "6" & "." & "38"
    Private mabEncKey As Byte()

#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Private Const C_A As String = "A"
    Private Const C_E As String = "E"
    Private Const C_S As String = "S"
    Private maesMain As Aes = Aes.Create(C_A & C_E & C_S)
#End If

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub


    ''' <remarks>是否已导入密钥</remarks>
    Private mIsLoadEncKey As Boolean
    Public ReadOnly Property IsLoadEncKey As Boolean
        Get
            IsLoadEncKey = mIsLoadEncKey
        End Get
    End Property


    ''' <summary>
    ''' 解密|Decrypt
    ''' </summary>
    ''' <param name="EncBase64Str">加密后的字节数组的Base64|Base64 of encrypted byte array</param>
    ''' <param name="UnEncStr">解密后的源字符串|Decrypted source string</param>
    ''' <param name="TextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Overloads Function Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As PigText
            oPigText = New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            Dim abUnEncBytes(-1) As Byte
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

    ''' <summary>
    ''' 解密|Decrypt
    ''' </summary>
    ''' <param name="EncBase64Str">加密后的字节数组的Base64|Base64 of encrypted byte array</param>
    ''' <param name="UnEncBytes">解密后的源字节数组|Decrypted source byte array</param>
    ''' <param name="TextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Overloads Function Decrypt(EncBase64Str As String, ByRef UnEncBytes As Byte(), TextType As PigText.enmTextType) As String
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
    ''' 解密|Decrypt
    ''' </summary>
    ''' <param name="EncBytes">加密后的字节数组|Encrypted byte array</param>
    ''' <param name="UnEncStr">解密后的源字符串|Decrypted source string</param>
    ''' <param name="TextType">源字符串文本类型|Source String Text Type</param>
    ''' <returns></returns>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
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

    ''' <summary>
    ''' 解密|Decrypt
    ''' </summary>
    ''' <param name="EncBytes">加密后的字节数组|Encrypted byte array</param>
    ''' <param name="UnEncBytes">解密后的源字节数组|Decrypted source byte array</param>
    ''' <returns></returns>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Return Me.mDecrypt(EncBytes, UnEncBytes)
    End Function


    ''' <remarks>解密数据（内部）</remarks>
    Private Function mDecrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Dim LOG As New PigStepLog("mDecrypt")
        Try
            If Me.mIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            LOG.StepName = "maesMain.CreateDecryptor"
            Dim ictAny As ICryptoTransform = maesMain.CreateDecryptor()
            LOG.StepName = "TransformFinalBlock"
            UnEncBytes = ictAny.TransformFinalBlock(EncBytes, 0, EncBytes.Length)
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
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
    ''' <param name="SrcBytes">源字节数组|Source Byte Array</param>
    ''' <param name="EncBase64Str">加密后的字节数组的Base64|Base64 of encrypted byte array</param>
    ''' <returns></returns>
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
            Return Me.GetSubErrInf("Encrypt", ex)
        End Try
    End Function


    ''' <summary>
    ''' 加密|encryption
    ''' </summary>
    ''' <param name="SrcBytes">源字节数组|Source Byte Array</param>
    ''' <param name="EncBytes">加密后的字节数组|Encrypted byte array</param>
    ''' <returns></returns>
    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Return Me.mEncrypt(SrcBytes, EncBytes)
    End Function


    ''' <remarks>加密数据（内部）</remarks>
    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Dim LOG As New PigStepLog("mEncrypt")
        Try
            If Me.mIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            LOG.StepName = "maesMain.CreateEncryptor"
            Dim ictAny As ICryptoTransform = maesMain.CreateEncryptor()
            LOG.StepName = "TransformFinalBlock"
            EncBytes = ictAny.TransformFinalBlock(SrcBytes, 0, SrcBytes.Length)
            ictAny = Nothing
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>获取随机数</summary>
    ''' <param name="BeginNum">起始数</param>
    ''' <param name="EndNum">结束数</param>
    Private Function mGetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()   '初始化随机数生成器
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function


    Public Function MkEncKey(ByRef EncKey As Byte()) As String
        Return Me.mMkEncKey(Nothing, 256, EncKey)
    End Function

    Public Function MkEncKey(InitKey As Byte(), ByRef EncKey As Byte()) As String
        Return Me.mMkEncKey(InitKey, 256, EncKey)
    End Function

    Public Function MkEncKey(Base64InitKey As String, ByRef Base64EncKey As String) As String
        Dim LOG As New PigStepLog("MkEncKey")
        Try
            LOG.StepName = "Base64InitKey to bytes"
            Dim oPigBytes As New PigBytes(Base64InitKey)
            If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
            Dim abEncKey(-1) As Byte
            LOG.StepName = "mMkEncKey"
            LOG.Ret = Me.mMkEncKey(oPigBytes.Main, 256, abEncKey)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            Base64EncKey = Convert.ToBase64String(mabEncKey)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkEncKey", ex)
        End Try
    End Function

    Public Function MkEncKey(ByRef Base64EncKey As String) As String
        Try
            Dim abEncKey(-1) As Byte
            Dim strRet As String = Me.mMkEncKey(Nothing, 256, abEncKey)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Base64EncKey = Convert.ToBase64String(mabEncKey)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkEncKey", ex)
        End Try
    End Function

    Private Function mMkEncKey(InitKey As Byte(), EncKeyLen As Integer, ByRef EncKey As Byte()) As String
        Try
            Select Case EncKeyLen
                Case 128, 192, 256
                Case Else
                    Throw New Exception("Key length cannot be greater than " & EncKeyLen.ToString)
            End Select
            Dim i As Integer
            ReDim EncKey(EncKeyLen - 1)
            If InitKey IsNot Nothing Then
                For i = 0 To InitKey.Length - 1
                    EncKey(i) = InitKey(i)
                Next
                For i = InitKey.Length To EncKeyLen - 1
                    EncKey(i) = Me.mGetRandNum(0, 255)
                Next
            Else
                For i = 0 To EncKeyLen - 1
                    EncKey(i) = Me.mGetRandNum(0, 255)
                Next
            End If
            If Me.IsLoadEncKey = False Then
                mabEncKey = EncKey
                Me.mIsLoadEncKey = True
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mMkEncKey", ex)
        End Try
    End Function

    '''' <remarks>生成密钥</remarks>
    'Public Function MkEncKey(ByRef Base64EncKey As String) As String
    '    Try
    '        Return Me.mMkEncKey(256, Base64EncKey)
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf("MkEncKey", ex, False)
    '    End Try
    'End Function

    'Private Function mMkEncKey(EncKeyLen As Integer, ByRef Base64EncKey As String) As String
    '    Try
    '        Select Case EncKeyLen
    '            Case 128, 192, 256
    '            Case Else
    '                Throw New Exception("Key length cannot be greater than " & EncKeyLen.ToString)
    '        End Select
    '        Dim i As Integer
    '        ReDim mabEncKey(EncKeyLen - 1)
    '        For i = 0 To EncKeyLen - 1
    '            mabEncKey(i) = Me.mGetRandNum(0, 255)
    '        Next
    '        Dim oPigBytes As New PigBytes(mabEncKey)
    '        Base64EncKey = oPigBytes.Base64Str
    '        oPigBytes = Nothing
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf("MkEncKey", ex)
    '    End Try
    'End Function

    ''' <remarks>导入密钥</remarks>
    Public Overloads Function LoadEncKey(Base64EncKey As String) As String
        Dim LOG As New PigStepLog("LoadEncKey.Base64EncKey")
        Try
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            LOG.StepName = "New PigText"
            Dim oPigText As New PigText(Base64EncKey, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
            LOG.StepName = "New MD5CryptoServiceProvider"
            mabEncKey = (New MD5CryptoServiceProvider).ComputeHash(oPigText.TextBytes)
            LOG.StepName = "Set maesMain"
            With maesMain
                .BlockSize = mabEncKey.Length * 8
                .Key = mabEncKey
                .IV = mabEncKey
                .Mode = CipherMode.CBC
                .Padding = PaddingMode.PKCS7
            End With
            Me.mIsLoadEncKey = True
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            Me.mIsLoadEncKey = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <remarks>导入密钥</remarks>
    Public Overloads Function LoadEncKey(EncKey As Byte()) As String
        Dim LOG As New PigStepLog("LoadEncKey.EncKey")
        Try
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            LOG.StepName = "New MD5CryptoServiceProvider"
            mabEncKey = (New MD5CryptoServiceProvider).ComputeHash(EncKey)
            LOG.StepName = "Set maesMain"
            With maesMain
                .BlockSize = mabEncKey.Length * 8
                .Key = mabEncKey
                .IV = mabEncKey
                .Mode = CipherMode.CBC
                .Padding = PaddingMode.PKCS7
            End With
            Me.mIsLoadEncKey = True
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            Me.mIsLoadEncKey = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub LoadEncKeySub(EncKey As Byte())
        Const SUB_NAME As String = "LoadEncKeySub"
        Const NEED_NET4_OR_HI As String = "Need to run in .net 4.0 or higher framework"
        Try
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            mabEncKey = (New MD5CryptoServiceProvider).ComputeHash(EncKey)
            With maesMain
                .BlockSize = mabEncKey.Length * 8
                .Key = mabEncKey
                .IV = mabEncKey
                .Mode = CipherMode.CBC
                .Padding = PaddingMode.PKCS7
            End With
            Me.mIsLoadEncKey = True
#Else
            Throw New Exception(NEED_NET4_OR_HI)
#End If
            Me.ClearErr()
        Catch ex As Exception
            Me.mIsLoadEncKey = False
            Me.SetSubErrInf(SUB_NAME, ex)
        End Try
    End Sub

    Public Sub DecryptSub(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType)
        Const SUB_NAME As String = "DecryptSub"
        Const NEED_NET4_OR_HI As String = "Need to run in .net 4.0 or higher framework"
        Const KEY_NOT_IMPORTED As String = "Key not imported"
        Try
            If Me.mIsLoadEncKey = False Then
                Throw New Exception(KEY_NOT_IMPORTED)
            End If
            Dim oPigText As PigText
            Dim abUnEncBytes As Byte()
            ReDim abUnEncBytes(0)
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            Dim ictAny As ICryptoTransform = maesMain.CreateDecryptor()
            abUnEncBytes = ictAny.TransformFinalBlock(EncBytes, 0, EncBytes.Length)
#Else
            Throw New Exception(NEED_NET4_OR_HI)
#End If
            oPigText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oPigText.Text
            oPigText = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, ex)
        End Try
    End Sub

End Class
