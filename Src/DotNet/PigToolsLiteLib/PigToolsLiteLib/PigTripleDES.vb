'**********************************
'* Name: PigTripleDES
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Processing TripleDES encryption algorithm
'* Home Url: https://en.seowphong.com
'* Version: 1.3
'* Create Time: 20/8/2022
'1.1  23/8/2002 Modify New, add IsLoadEncKey,LoadEncKey
'1.2  24/8/2002 Add MkKey
'1.3  25/8/2002 Modify MkEncKey,mMkEncKey, add MkEncKey
'************************************
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
''' <summary>
''' 3DES processing class|3DES处理类
''' </summary>
Public Class PigTripleDES
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.3.8"
    Private mabEncKey As Byte()

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Try
            Me.mTripleDES = New TripleDESCryptoServiceProvider
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub
    Private ReadOnly Property mTripleDES As TripleDESCryptoServiceProvider

    Private mIsLoadEncKey As Boolean
    Public Property IsLoadEncKey As Boolean
        Get
            IsLoadEncKey = mIsLoadEncKey
        End Get
        Friend Set(value As Boolean)
            mIsLoadEncKey = value
        End Set
    End Property

    'Public Function EncryptDes(ByVal SourceStr As String, ByVal myKey As String, ByVal myIV As String) As String '使用的DES对称加密
    '    Dim des As New System.Security.Cryptography.DESCryptoServiceProvider 'DES算法
    '    'Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider'TripleDES算法
    '    Dim inputByteArray As Byte()
    '    inputByteArray = System.Text.Encoding.Default.GetBytes(SourceStr)
    '    des.Key = System.Text.Encoding.UTF8.GetBytes(myKey) 'myKey DES用8个字符，TripleDES要24个字符
    '    des.IV = System.Text.Encoding.UTF8.GetBytes(myIV) 'myIV DES用8个字符，TripleDES要24个字符
    '    Dim ms As New System.IO.MemoryStream
    '    Dim cs As New System.Security.Cryptography.CryptoStream(ms, des.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
    '    Dim sw As New System.IO.StreamWriter(cs)
    '    sw.Write(SourceStr)
    '    sw.Flush()
    '    cs.FlushFinalBlock()
    '    ms.Flush()
    '    EncryptDes = Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    'End Function
    'Public Function DecryptDes(ByVal SourceStr As String, ByVal myKey As String, ByVal myIV As String) As String    '使用标准DES对称解密
    '    Dim des As New System.Security.Cryptography.DESCryptoServiceProvider 'DES算法
    '    'Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider'TripleDES算法
    '    des.Key = System.Text.Encoding.UTF8.GetBytes(myKey) 'myKey DES用8个字符，TripleDES要24个字符
    '    des.IV = System.Text.Encoding.UTF8.GetBytes(myIV) 'myIV DES用8个字符，TripleDES要24个字符
    '    Dim buffer As Byte() = Convert.FromBase64String(SourceStr)
    '    Dim ms As New System.IO.MemoryStream(buffer)
    '    Dim cs As New System.Security.Cryptography.CryptoStream(ms, des.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read)
    '    Dim sr As New System.IO.StreamReader(cs)
    '    DecryptDes = sr.ReadToEnd()
    'End Function

    Public Overloads Function LoadEncKey(Base64EncKey As String) As String
        Try
            Dim pbEncKey As New PigBytes(Base64EncKey)
            If pbEncKey.LastErr <> "" Then Throw New Exception(pbEncKey.LastErr)
            Dim strRet As String = Me.LoadEncKey(pbEncKey.Main)
            pbEncKey = Nothing
            Return strRet
        Catch ex As Exception
            Me.IsLoadEncKey = False
            Return Me.GetSubErrInf("LoadEncKey", ex, False)
        End Try
    End Function

    ''' <remarks>导入密钥</remarks>
    Public Overloads Function LoadEncKey(EncKey As Byte()) As String
        Dim LOG As New PigStepLog("LoadEncKey")
        Try
            If EncKey Is Nothing Then Throw New Exception("EncKey is Nothing")
            If EncKey.Length <> 32 Then Throw New Exception("The length of EncKey must be 32")
            LOG.StepName = "Set abIV"
            Dim abIV(7) As Byte
            For i = 0 To 7
                abIV(i) = EncKey(i + 24)
            Next
            ReDim Preserve EncKey(23)
            LOG.StepName = "Set Key"
            Me.mTripleDES.Key = EncKey
            LOG.StepName = "Set IV"
            Me.mTripleDES.IV = abIV
            Me.IsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            Me.IsLoadEncKey = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Dim LOG As New PigStepLog("mEncrypt")
        Try
            If Me.IsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            LOG.StepName = "New MemoryStream"
            Dim msMain As New MemoryStream
            LOG.StepName = "New CryptoStream"
            Dim csMain As New CryptoStream(msMain, Me.mTripleDES.CreateEncryptor(), CryptoStreamMode.Write)
            LOG.StepName = "New StreamWriter"
            Dim swMain As New StreamWriter(csMain)
            LOG.StepName = "Write"
            swMain.Write(Convert.ToBase64String(SrcBytes))
            swMain.Flush()
            csMain.FlushFinalBlock()
            msMain.Flush()
            LOG.StepName = "GetBuffer"
            EncBytes = msMain.GetBuffer()
            ReDim Preserve EncBytes(msMain.Length - 1)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Return Me.mEncrypt(SrcBytes, EncBytes)
    End Function

    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBase64Str As String) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim abEncBytes(-1) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(SrcBytes, abEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            EncBase64Str = Convert.ToBase64String(abEncBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Encrypt", ex)
        End Try
    End Function

    Public Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Return Me.mDecrypt(EncBytes, UnEncBytes)
    End Function

    Public Function Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
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
        Dim LOG As New PigStepLog("mEncrypt")
        Try
            If Me.IsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            LOG.StepName = "New MemoryStream"
            Dim msMain As New MemoryStream(EncBytes)
            LOG.StepName = "New CryptoStream"
            Dim csMain As New CryptoStream(msMain, Me.mTripleDES.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read)
            LOG.StepName = "New StreamReader"
            Dim srMain As New StreamReader(csMain)
            LOG.StepName = "ReadToEnd"
            Dim strUnEncBase64 As String = srMain.ReadToEnd()
            LOG.StepName = "Set Main"
            Dim oPigBytes As New PigBytes(strUnEncBase64)
            UnEncBytes = oPigBytes.Main
            If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
            oPigBytes = Nothing
            Return "OK"
        Catch ex As Exception
            UnEncBytes = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
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
            Randomize()   '初始化随机数生成器
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function

    Public Function MkEncKey(ByRef Base64EncKey As String, Optional InitKeyStr As String = "", Optional InitKeyTextType As PigText.enmTextType = PigText.enmTextType.UTF8) As String
        Dim LOG As New PigStepLog("MkKey")
        Try
            Dim abKey(-1) As Byte
            LOG.StepName = "New PigText"
            Dim ptInitKey As New PigText(InitKeyStr, InitKeyTextType)
            If ptInitKey.LastErr <> "" Then Throw New Exception(ptInitKey.LastErr)
            LOG.StepName = "mMkEncKey"
            LOG.Ret = Me.mMkEncKey(ptInitKey.TextBytes, abKey)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            LOG.StepName = "ToBase64String"
            Base64EncKey = Convert.ToBase64String(abKey)
            ptInitKey = Nothing
            Return "OK"
        Catch ex As Exception
            Base64EncKey = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function MkEncKey(InitKey As Byte(), ByRef EncKey As Byte()) As String
        Return Me.mMkEncKey(InitKey, EncKey)
    End Function

    Private Function mMkEncKey(InitKey As Byte(), ByRef EncKey As Byte()) As String
        Dim LOG As New PigStepLog("mMkEncKey")
        Try
            If InitKey Is Nothing Then Throw New Exception("InitKey Is Nothing")
            Dim intKenLen As Integer = InitKey.Length
            Select Case intKenLen
                Case <= 0
                    LOG.StepName = "GenerateKey"
                    Me.mTripleDES.GenerateKey()
                Case 1 To 23
                    ReDim Preserve InitKey(23)
                    For i = intKenLen - 1 To 23
                        InitKey(i) = Me.mGetRandNum(0, 255)
                        'InitKey(i) = Me.mGetRandNum(4, 122)
                    Next
                    Me.mTripleDES.Key = InitKey
                Case 24
                    LOG.StepName = "Set Key"
                    Me.mTripleDES.Key = InitKey
                Case > 24
                    ReDim Preserve InitKey(23)
                    Me.mTripleDES.Key = InitKey
            End Select
            LOG.StepName = "GenerateIV"
            Me.mTripleDES.GenerateIV()
            Dim oPigBytes As New PigBytes
            LOG.StepName = "SetValue(Key)"
            oPigBytes.SetValue(Me.mTripleDES.Key)
            LOG.StepName = "SetValue(IV)"
            oPigBytes.SetValue(Me.mTripleDES.IV)
            EncKey = oPigBytes.Main
            Return "OK"
        Catch ex As Exception
            EncKey = Nothing
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    'Sub Main()
    '    Dim strSrc As String = "nyfort的BLOG"
    '    Dim abIV(-1) As Byte
    '    '加密
    '    '注意KEY和IV只能用英文和数字,des是8个字符
    '    Dim strEnc As String = EncryptDes(strSrc, "Qz\p{s+萧u(R0e$^j/=*c[\", abIV)
    '    Console.WriteLine("源文=" & strSrc)
    '    Console.WriteLine("密文=" & strEnc)
    '    '解密
    '    Dim strUnEnc As String = ""
    '    strUnEnc = DecryptDes(strEnc, "Qz\p{s+萧u(R0e$^j/=*c[\", abIV)
    '    Console.WriteLine("解密文=" & strUnEnc)
    '    Console.WriteLine("OK")
    'End Sub

    'Public Function DecryptDes(ByVal SourceStr As String, ByVal myKey As String, InIV As Byte()) As String    '使用标准DES对称解密
    '    'Dim des As New System.Security.Cryptography.DESCryptoServiceProvider 'DES算法
    '    Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider 'TripleDES算法
    '    DES.Key = System.Text.Encoding.UTF8.GetBytes(myKey) 'myKey DES用8个字符，TripleDES要24个字符
    '    'DES.IV = System.Text.Encoding.UTF8.GetBytes(myIV) 'myIV DES用8个字符，TripleDES要24个字符
    '    DES.IV = InIV
    '    Dim buffer As Byte() = Convert.FromBase64String(SourceStr)
    '    Dim ms As New System.IO.MemoryStream(buffer)
    '    Dim cs As New System.Security.Cryptography.CryptoStream(ms, DES.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read)
    '    Dim sr As New System.IO.StreamReader(cs)
    '    DecryptDes = sr.ReadToEnd()
    'End Function

    'Public Function EncryptDes(ByVal SourceStr As String, ByVal myKey As String, ByRef OutIV As Byte()) As String '使用的DES对称加密
    '    'Dim des As New System.Security.Cryptography.DESCryptoServiceProvider 'DES算法
    '    Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider 'TripleDES算法
    '    Dim inputByteArray As Byte()
    '    inputByteArray = System.Text.Encoding.Default.GetBytes(SourceStr)
    '    DES.Key = System.Text.Encoding.UTF8.GetBytes(myKey) 'myKey DES用8个字符，TripleDES要24个字符
    '    DES.GenerateIV()
    '    OutIV = DES.IV
    '    'des.IV = System.Text.Encoding.UTF8.GetBytes(myIV) 'myIV DES用8个字符，TripleDES要24个字符
    '    Dim ms As New System.IO.MemoryStream
    '    Dim cs As New System.Security.Cryptography.CryptoStream(ms, DES.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
    '    Dim sw As New System.IO.StreamWriter(cs)
    '    sw.Write(SourceStr)
    '    sw.Flush()
    '    cs.FlushFinalBlock()
    '    ms.Flush()
    '    EncryptDes = Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)
    'End Function

End Class
