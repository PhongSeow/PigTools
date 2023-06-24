'**********************************
'* Name: TextStreamAsc
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Reference Scripting.TextStream through ActiveX
'* Home Url: https://en.seowphong.com
'* Version: 1.2
'* Create Time: 13/3/2022
'* 1.1 13/3/2021   Modify New
'* 1.2 24/6/2023   Modify Obj
'**********************************
Public Class TextStreamAsc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.1"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Friend Property Obj As Object


    Public Sub Close()
        Try
            Me.Obj.Close
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Close", ex)
        End Try
    End Sub

    Public ReadOnly Property AtEndOfStream() As Boolean
        Get
            Try
                Return Me.Obj.AtEndOfStream
            Catch ex As Exception
                Me.SetSubErrInf("AtEndOfStream", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function ReadAll() As String
        Try
            Return Me.Obj.ReadAll
        Catch ex As Exception
            Me.SetSubErrInf("ReadAll", ex)
            Return ""
        End Try
    End Function

    Public Function ReadLine() As String
        Try
            Return Me.Obj.ReadLine
        Catch ex As Exception
            Me.SetSubErrInf("ReadLine", ex)
            Return ""
        End Try
    End Function

    Public Sub Write(Text As String)
        Try
            Me.Obj.Write(Text)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Write", ex)
        End Try
    End Sub

    Public Sub WriteBlankLines(Lines As Long)
        Try
            Me.Obj.WriteBlankLines
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("WriteBlankLines", ex)
        End Try
    End Sub

    Public Sub WriteLine(Text As String)
        Try
            Me.Obj.WriteLine（Text)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("WriteLine", ex)
        End Try
    End Sub

End Class
