'**********************************
'* Name: PigSend
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 数据发送和接收|Data sending and receiving
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 6/11/2021
'* 1.1  7/11/2022   Add InitEnc
'* 1.2  8/11/2022   Modify InitEnc
'**********************************v
Public Class PigSend
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.68"

    ''' <summary>
    ''' 加密方式|Encryption mode
    ''' </summary>
    Public Enum EnmEncType
        Original = 0
        SeowEnc = 1
        SeowEncAndPigAes = 2
    End Enum



    Public ReadOnly Property EncType As EnmEncType

    Private Property mSeowEnc As SeowEnc

    Private Property mPigAes As PigAes

    Public Sub New(EncType As EnmEncType)
        MyBase.New(CLS_VERSION)
        Me.EncType = EncType
        Me.mInitEnc(Nothing, Nothing)
    End Sub

    Private mIsReady As Boolean = False
    Public Property IsReady As Boolean
        Get
            Return mIsReady
        End Get
        Friend Set(value As Boolean)
            mIsReady = value
        End Set
    End Property

    Public Function InitEnc(InSeowEnc As SeowEnc) As String
        Return Me.mInitEnc(InSeowEnc)
    End Function

    Public Function InitEnc(InSeowEnc As SeowEnc, InPigAes As PigAes) As String
        Return Me.mInitEnc(InSeowEnc, InPigAes)
    End Function

    Private Function mInitEnc(InSeowEnc As SeowEnc, Optional InPigAes As PigAes = Nothing) As String
        Dim LOG As New PigStepLog("mInitEnc")
        Try
            Dim bolIsSeowEnc As Boolean = False, bolIsPigAes As Boolean = False
            Select Case Me.EncType
                Case EnmEncType.SeowEnc
                    bolIsSeowEnc = True
                Case EnmEncType.SeowEncAndPigAes
                    bolIsPigAes = True
            End Select
            If bolIsSeowEnc = True Then
                LOG.StepName = "Check InSeowEnc"
                If InSeowEnc Is Nothing Then Throw New Exception("InSeowEnc is nothing")
                If InSeowEnc.IsLoadEncKey = False Then Throw New Exception("Is not LoadEncKey")
                Me.mSeowEnc = InSeowEnc
            End If
            If bolIsPigAes = True Then
                LOG.StepName = "Check InPigAes"
                If InPigAes Is Nothing Then Throw New Exception("InPigAes is nothing")
                If InPigAes.IsLoadEncKey = False Then Throw New Exception("Is not LoadEncKey")
                Me.mPigAes = InPigAes
            End If
            Me.IsReady = True
            Return "OK"
        Catch ex As Exception
            Me.IsReady = False
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function SendBytesData(SendBytes As Byte(), ByRef TarBytes As Byte()) As String
        Try
            If Me.EncType = EnmEncType.Original Then
                TarBytes = SendBytes
            Else
                Dim strRet As String = Me.mSendData(SendBytes, TarBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SendBytesData", ex)
        End Try
    End Function

    Public Function SendBytesData(SendBytes As Byte(), ByRef TarBase64 As String) As String
        Try
            If Me.EncType = EnmEncType.Original Then
                TarBase64 = Convert.ToBase64String(SendBytes)
            Else
                Dim abTar(0) As Byte
                Dim strRet As String = Me.mSendData(SendBytes, abTar)
                If strRet <> "OK" Then Throw New Exception(strRet)
                TarBase64 = Convert.ToBase64String(abTar)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SendBytesData", ex)
        End Try
    End Function

    Public Function SendStrData(SendStr As String, TextType As PigText.enmTextType, ByRef TarBase64 As String) As String
        Try
            Dim oPigText As New PigText(SendStr, TextType)
            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
            If Me.EncType = EnmEncType.Original Then
                TarBase64 = oPigText.Base64
            Else
                Dim abTar(0) As Byte
                Dim strRet As String = Me.mSendData(oPigText.TextBytes, abTar)
                If strRet <> "OK" Then Throw New Exception(strRet)
                TarBase64 = Convert.ToBase64String(abTar)
            End If
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SendStrData", ex)
        End Try
    End Function

    Public Function SendStrData(SendStr As String, TextType As PigText.enmTextType, ByRef TarBytes As Byte()) As String
        Try
            Dim oPigText As New PigText(SendStr, TextType)
            If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
            If Me.EncType = EnmEncType.Original Then
                TarBytes = oPigText.TextBytes
            Else
                Dim strRet As String = Me.mSendData(oPigText.TextBytes, TarBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SendStrData", ex)
        End Try
    End Function


    Private Function mSendData(SendBytes As Byte(), ByRef TarBytes As Byte()) As String
        Dim LOG As New PigStepLog("mSendData")
        Try
            Select Case Me.EncType
                Case EnmEncType.SeowEnc
                    LOG.StepName = "mSeowEnc.Encrypt"
                    LOG.Ret = Me.mSeowEnc.Encrypt(SendBytes, TarBytes)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Case EnmEncType.SeowEncAndPigAes
                    Dim abOut(0) As Byte
                    LOG.StepName = "mSeowEnc.Encrypt"
                    LOG.Ret = Me.mSeowEnc.Encrypt(SendBytes, abOut)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                    LOG.StepName = "mPigAes.Encrypt"
                    LOG.Ret = Me.mPigAes.Encrypt(abOut, TarBytes)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Case Else
                    Throw New Exception("Invalid EncType")
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mReceiveData(ReceiveBytes As Byte(), ByRef TarBytes As Byte()) As String
        Dim LOG As New PigStepLog("mSendData")
        Try
            Select Case Me.EncType
                Case EnmEncType.SeowEnc
                    LOG.StepName = "mSeowEnc.Decrypt"
                    LOG.Ret = Me.mSeowEnc.Decrypt(ReceiveBytes, TarBytes)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Case EnmEncType.SeowEncAndPigAes
                    Dim abOut(0) As Byte
                    LOG.StepName = "mSeowEnc.Decrypt"
                    LOG.Ret = Me.mSeowEnc.Decrypt(ReceiveBytes, abOut)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                    LOG.StepName = "mPigAes.Decrypt"
                    LOG.Ret = Me.mPigAes.Decrypt(abOut, TarBytes)
                    If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                Case Else
                    Throw New Exception("Invalid EncType")
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ReceiveBytesData(ReceiveBytes As Byte(), ByRef TarBytes As Byte()) As String
        Try
            If Me.EncType = EnmEncType.Original Then
                TarBytes = ReceiveBytes
            Else
                Dim strRet As String = Me.mReceiveData(ReceiveBytes, TarBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ReceiveBytesData", ex)
        End Try
    End Function

    Public Function ReceiveBytesData(ReceiveBytes As Byte(), ByRef TarBase64 As String) As String
        Try
            If Me.EncType = EnmEncType.Original Then
                TarBase64 = Convert.ToBase64String(ReceiveBytes)
            Else
                Dim abTar(0) As Byte
                Dim strRet As String = Me.mReceiveData(ReceiveBytes, abTar)
                If strRet <> "OK" Then Throw New Exception(strRet)
                TarBase64 = Convert.ToBase64String(abTar)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ReceiveBytesData", ex)
        End Try
    End Function

    Public Function ReceiveStrData(ReceiveBase64 As String, TextType As PigText.enmTextType, ByRef TarStr As String) As String
        Try
            Dim oPigBytes As New PigBytes(ReceiveBase64)
            If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
            If Me.EncType = EnmEncType.Original Then
                Dim oPigText As New PigText(oPigBytes.Main, TextType)
                TarStr = oPigText.Text
                oPigText = Nothing
            Else
                Dim abTar(0) As Byte
                Dim strRet As String = Me.mReceiveData(oPigBytes.Main, abTar)
                If strRet <> "OK" Then Throw New Exception(strRet)
                Dim oPigText As New PigText(abTar, TextType)
                TarStr = oPigText.Text
                oPigText = Nothing
            End If
            oPigBytes = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ReceiveStrData", ex)
        End Try
    End Function

    Public Function ReceiveStrData(ReceiveBytes As Byte(), TextType As PigText.enmTextType, ByRef TarStr As String) As String
        Try
            If Me.EncType = EnmEncType.Original Then
                Dim oPigText As New PigText(ReceiveBytes, TextType)
                TarStr = oPigText.Text
                oPigText = Nothing
            Else
                Dim abTar(0) As Byte
                Dim strRet As String = Me.mReceiveData(ReceiveBytes, abTar)
                If strRet <> "OK" Then Throw New Exception(strRet)
                Dim oPigText As New PigText(abTar, TextType)
                TarStr = oPigText.Text
                oPigText = Nothing
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ReceiveStrData", ex)
        End Try
    End Function

End Class
