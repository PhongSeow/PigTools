'**********************************
'* Name: PigCompressor
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Compression processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 11/12/2019
'* 1.0.2  10/17/2020  optimization
'**********************************

Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Text

Public Class PigCompressor
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.2"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function Compress(SrcStr As String) As Byte()
        Dim abAny As Byte() = Encoding.UTF8.GetBytes(SrcStr)
        Return Me.Compress(abAny)
    End Function

    Public Function Compress(SrcBytes As Byte()) As Byte()
        Try
            Using msAny As MemoryStream = New MemoryStream
                Using gzsAny As GZipStream = New GZipStream(msAny, CompressionMode.Compress)
                    gzsAny.Write(SrcBytes, 0, SrcBytes.Length)
                    gzsAny.Flush()
                End Using
                Return msAny.ToArray
            End Using
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Compress", ex)
            Return Nothing
        End Try
    End Function

    Public Function Depress(CompressBytes As Byte()) As Byte()
#If NET40_OR_GREATER Then
        Try
            Dim buffer As Byte()
            Using msAny As MemoryStream = New MemoryStream(CompressBytes)
                Using msAny2 As MemoryStream = New MemoryStream
                    Using gzsAny As GZipStream = New GZipStream(msAny, CompressionMode.Decompress)
                        gzsAny.CopyTo(msAny2)
                        gzsAny.Flush()
                    End Using
                    buffer = msAny2.ToArray
                End Using
            End Using
            Me.ClearErr()
            Return buffer
        Catch ex As Exception
            Me.SetSubErrInf("Depress", ex)
            Return Nothing
        End Try
#Else
        Try
            Throw New Exception("Need to run in .net 4.0 or higher framework")
        Catch ex As Exception
            Me.SetSubErrInf("Depress", ex)
            Return Nothing
        End Try
#End If

    End Function

End Class
