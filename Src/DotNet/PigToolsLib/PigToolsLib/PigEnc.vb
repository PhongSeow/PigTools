'**********************************
'* Name: PigLocalEnc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: A localized encryption and decryption class, ciphertext cannot be decrypted to other machines.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 20/8/2022
'1.1  6/9/2022   Add EnmEncKeySaveType,EnmEncType
'1.2  25/9/2022  Modify EnmDataEncType
'**********************************
Friend Class PigEnc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.2"

    ''' <summary>
    ''' 密钥使用方式|How to use the key
    ''' </summary>
    Public Enum EnmEncKeyUseType
        ''' <summary>
        ''' 可以在不同机器使用
        ''' </summary>
        Multifocal = 0
        ''' <summary>
        ''' 只能在本机使用
        ''' </summary>
        LocalHost = 1
        ''' <summary>
        ''' 只能在本机用户使用
        ''' </summary>
        LocalUser = 2
        ''' <summary>
        ''' 密钥需要导入到本机后才能使用
        ''' </summary>
        LocalHostLoad = 3
        ''' <summary>
        ''' 密钥需要导入到本机用户后才能使用
        ''' </summary>
        LocalUserLoad = 4
        ''' <summary>
        ''' 密钥需要在客户端向服务器端注册后才能使用，
        ''' </summary>
        Register = 5
        ''' <summary>
        ''' 密钥需要在客户端向服务器端注册，每次通讯时需要由客户端向服务器商登录获取动态的短期密钥。
        ''' </summary>
        RegisterDynamic = 6
    End Enum

    ''' <summary>
    ''' 密钥加密方式|Data encryption mode
    ''' </summary>
    Public Enum EnmDataEncType
        TripleDES = 0
        Aes = 1
        Rsa = 2
    End Enum


    Private Property mPigRsa As PigRsa
    Private Property mPigAes As PigAes
    Private Property mPigTripleDES As PigTripleDES

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Private Function mGetDataEncKey(DataEncType As EnmDataEncType) As String
        Try
            Select Case DataEncType
                Case EnmDataEncType.Aes
                Case EnmDataEncType.Rsa
                    Dim oPigAes As New PigAes
                Case EnmDataEncType.TripleDES
                    Dim oPigTripleDES As New PigTripleDES
                    'oPigTripleDES.MkEncKey() 
                Case Else
                    Throw New Exception("Invalid DataEncType")
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

    Private Function mMkEncKey(EncKeyUseType As EnmEncKeyUseType, DataEncType As EnmDataEncType, ByRef SrvEncKey As Byte(), ByRef CliEncKey As Byte()) As String
        Dim LOG As New PigStepLog("mMkEncKey")
        Try
            Select Case EncKeyUseType
                Case EnmEncKeyUseType.Multifocal

                Case Else
                    Throw New Exception("Not supported yet")
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mMkEncKey", ex)
        End Try
    End Function

End Class
