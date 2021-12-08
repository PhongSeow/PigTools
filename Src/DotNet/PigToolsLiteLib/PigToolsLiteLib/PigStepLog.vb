'**********************************
'* Name: PigStepLog
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigStepLog is for logging and error handling in the process.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.6
'* Create Time: 8/12/2019
'************************************
Public Class PigStepLog
    Public ReadOnly Property SubName As String
    Public ReadOnly Property IsLogUseTime As Boolean
    Public Ret As String = ""
    Private moUseTime As UseTime

    Private Sub mNew()
        If Me.IsLogUseTime = True Then

            moUseTime = New UseTime
            moUseTime.GoBegin()
        End If
    End Sub
    Public Sub New(SubName As String)
        Me.SubName = SubName
        Me.IsLogUseTime = False
        Me.mNew()
    End Sub

    Public Sub New(SubName As String, IsLogUseTime As Boolean)
        Me.SubName = SubName
        Me.IsLogUseTime = IsLogUseTime
        Me.mNew()
    End Sub

    Private mstrStepName As String = ""
    Public Property StepName As String
        Get
            Return mstrStepName
        End Get
        Set(value As String)
            mstrStepName = value
        End Set
    End Property

    Public Sub AddStepNameInf(AddInf As String)
        Me.StepName &= "(" & AddInf & ")"
    End Sub

    Public ReadOnly Property BeginTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.BeginTime
            Else
                Return DateTime.MinValue
            End If
        End Get
    End Property

    Public ReadOnly Property EndTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.EndTime
            Else
                Return DateTime.MaxValue
            End If
        End Get
    End Property

    Public ReadOnly Property AllDiffSeconds As Decimal
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.AllDiffSeconds
            Else
                Return -1
            End If
        End Get
    End Property


    Public Sub ToEnd()
        If Me.IsLogUseTime = True Then
            moUseTime.ToEnd()
        End If
    End Sub
End Class
