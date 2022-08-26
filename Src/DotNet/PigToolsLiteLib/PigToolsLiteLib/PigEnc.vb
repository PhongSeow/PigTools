'**********************************
'* Name: PigLocalEnc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: A localized encryption and decryption class, ciphertext cannot be decrypted to other machines.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 20/8/2022
'**********************************
Friend Class PigEnc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.8"

    ''' <summary>
    ''' 加密程度|Encryption level
    ''' </summary>
    Public Enum EnmEncLevel
        ''' <summary>最低</summary>
        Minimum = 0
        ''' <summary>较低</summary>
        Low = 1
        ''' <summary>较高</summary>
        Higher = 2
        ''' <summary>最高</summary>
        Supreme = 3
        ''' <summary>左标记开始</summary>
    End Enum

    Public ReadOnly Property EncLevel As Boolean
    Private Property mPigRsa As PigRsa
    Private Property mPigAes As PigAes

    Public Sub New(EncLevel As EnmEncLevel)
        MyBase.New(CLS_VERSION)
        Try
#If NETCOREAPP Then
        Me.EncLevel = EncLevel
#ElseIf NET40_OR_GREATER Then
        Me.EncLevel = EncLevel
#Else
            Select Case EncLevel
                Case EnmEncLevel.Low, EnmEncLevel.Supreme
                    Throw New Exception("This level of encryption only supports .net 4.0 or higher frameworks.")
                Case EnmEncLevel.Minimum, EnmEncLevel.Higher
                    Me.EncLevel = EncLevel
                Case Else
                    Throw New Exception("Unknown encryption level.")
            End Select
#End If
            Me.mPigAes = New PigAes
        Catch ex As Exception
            Me.EncLevel = EnmEncLevel.Minimum
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub
End Class
