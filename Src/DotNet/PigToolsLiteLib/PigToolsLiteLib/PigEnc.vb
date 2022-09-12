'**********************************
'* Name: PigLocalEnc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: A localized encryption and decryption class, ciphertext cannot be decrypted to other machines.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 20/8/2022
'1.1  6/9/2022  Add EnmEncKeySaveType,EnmEncType
'**********************************
Friend Class PigEnc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"

    ''' <summary>
    ''' 密钥保存方式|Key storage method
    ''' </summary>
    Public Enum EnmEncKeySaveType
        Rsa_TripleDES = 0
        Rsa_AES = 1
        UserDefined1 = 10
        UserDefined2 = 20
        UserDefined3 = 30
    End Enum

    ''' <summary>
    ''' 加密方式|Encryption method
    ''' </summary>
    Public Enum EnmEncType
        TripleDES = 0
        Aes = 1
        Rsa = 2
        CompressTripleDES = 3
        CompressAes = 3
        CompressRsa = 4
        Compress2TripleDES = 5
        Compress2Aes = 6
        Compress2Rsa = 7
        UserDefined1 = 10
        UserDefined2 = 20
        UserDefined3 = 30
    End Enum

    Private Property mPigRsa As PigRsa
    Private Property mPigAes As PigAes
    Private Property mPigTripleDES As PigTripleDES

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function MkEncKey() As String
        Try
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkEncKey", ex)
        End Try
    End Function

End Class
