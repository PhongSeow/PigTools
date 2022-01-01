'**********************************
'* Name: 豚豚配置|PigConfig
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 配置项|Configuration item
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 18/12/2021
'* 1.1    20/12/2020   Add ConfValue,ContentType
'* 1.2    22/12/2020   Modify ConfValue,mNew
'**********************************
Public Class PigConfig
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.6"

    Public Enum EnmContentType
        ''' <summary>
        ''' Text
        ''' </summary>
        Text = 0
        ''' <summary>
        ''' Encrypted text
        ''' </summary>
        EncText = 1
        ''' <summary>
        ''' Numerical type
        ''' </summary>
        Numerical = 2
        ''' <summary>
        ''' Date and time type
        ''' </summary>
        DateTime = 3
        ''' <summary>
        ''' Boolean value
        ''' </summary>
        BooleanValue = 4

    End Enum


    Public Sub New(ConfName As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent)
    End Sub

    Public Sub New(ConfName As String, ConfValue As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent, ConfValue)
    End Sub

    Public Sub New(ConfName As String, ConfValue As String, ConfDesc As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent, ConfValue, ConfDesc)
    End Sub

    Private Sub mNew(ConfName As String, Parent As PigConfigSession, Optional ConfValue As String = "", Optional ConfDesc As String = "")
        Try
            If ConfName = "" Then Throw New Exception("ConfName is a space")
            If Parent Is Nothing Then Throw New Exception("PigConfigSession Is Nothing")
            With Me
                .ConfName = ConfName
                .Parent = Parent
                .ConfValue = ConfValue
                .ConfDesc = ConfDesc
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 父对象|Parent object
    ''' </summary>
    Private moParent As PigConfigSession
    Public Property Parent As PigConfigSession
        Get
            Return moParent
        End Get
        Friend Set(value As PigConfigSession)
            moParent = value
        End Set
    End Property


    ''' <summary>
    ''' 配置名
    ''' </summary>
    Private mstrConfName As String
    Public Property ConfName As String
        Get
            Return mstrConfName
        End Get
        Friend Set(value As String)
            mstrConfName = value
        End Set
    End Property

    ''' <summary>
    ''' 配置描述
    ''' </summary>
    Private mstrConfDesc As String
    Public Property ConfDesc As String
        Get
            Return mstrConfDesc
        End Get
        Set(value As String)
            mstrConfDesc = value
        End Set
    End Property

    ''' <summary>
    ''' 配置值
    ''' </summary>
    Private mstrConfValue As String
    Private mstrUnEncConfValue As String = ""
    Public Property ConfValue As String
        Get
            Dim LOG As New PigStepLog("ConfValue")
            Try
                If Me.ContentType = EnmContentType.EncText Then
                    If mstrUnEncConfValue = "" Then
                        LOG.StepName = "fGetUnEncStr"
                        LOG.Ret = Me.Parent.Parent.fGetUnEncStr(mstrUnEncConfValue, mstrConfValue)
                        If LOG.Ret <> "OK" Then
                            mstrUnEncConfValue = ""
                            If Me.Parent IsNot Nothing Then LOG.AddStepNameInf(Me.Parent.SessionName)
                            LOG.AddStepNameInf(Me.ConfName)
                            Throw New Exception(LOG.Ret)
                        End If
                    End If
                    Return mstrUnEncConfValue
                Else
                    Return mstrConfValue
                End If
            Catch ex As Exception
                If Me.Parent IsNot Nothing Then
                    If Me.Parent.Parent IsNot Nothing Then
                        Me.Parent.Parent.PrintDebugLog(Me.MyClassName, Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex))
                    End If
                End If
                Return ""
            End Try
        End Get
        Set(value As String)
            mstrConfValue = value
            If Mid(mstrConfValue, 1, 5) = "{Enc}" Then
                mstrUnEncConfValue = ""
                Me.ContentType = EnmContentType.EncText
            ElseIf IsDate(mstrConfValue) = True Then
                Me.ContentType = EnmContentType.DateTime
            ElseIf IsNumeric(mstrConfValue) Then
                Me.ContentType = EnmContentType.Numerical
            Else
                Select Case UCase(mstrConfValue)
                    Case "TRUE", "FALSE"
                        Me.ContentType = EnmContentType.BooleanValue
                    Case Else
                        Me.ContentType = EnmContentType.Text
                End Select
            End If
        End Set
    End Property

    Friend ReadOnly Property fConfValue As String
        Get
            Return mstrConfValue
        End Get
    End Property


    Private mintContentType As EnmContentType
    Public Property ContentType As EnmContentType
        Get
            Return mintContentType
        End Get
        Friend Set(value As EnmContentType)
            mintContentType = value
        End Set
    End Property


End Class

