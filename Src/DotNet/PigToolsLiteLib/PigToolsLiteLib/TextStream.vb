'**********************************
'* Name: TextStream
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Amount to Scripting.TextStream of VB6
'* Home Url: https://en.seowphong.com
'* Version: 1.5
'* Create Time: 30/12/2020
'* 1.0.2    15/1/2021   Err.Raise change to Throw New Exception
'* 1.0.3    23/1/2021   pTextStream rename to TextStream
'* 1.0.4    23/1/2021   Modify Init
'* 1.0.5    25/7/2021   Modify Init
'* 1.1      24/6/2023   Modify IOMode
'* 1.2      4/12/2023   Add TextType
'* 1.3      20/12/2023  Modify New,Init
'* 1.5      23/1/2024  Modify Init
'**********************************
Imports System.IO
Imports System.Text
Public Class TextStream
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "5" & "." & "8"

    Private mEnmIOMode As PigFileSystem.IOMode

    Public ReadOnly Property TextType As PigText.enmTextType = PigText.enmTextType.UTF8

    Private srMain As StreamReader
    Private swMain As StreamWriter

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.TextType = PigText.enmTextType.UnknowOrBin
    End Sub

    Public Sub New(TextType As PigText.enmTextType)
        MyBase.New(CLS_VERSION)
        Me.TextType = TextType
    End Sub

    Private ReadOnly Property mEncTextType As Encoding
        Get
            Select Case Me.TextType
                Case PigText.enmTextType.Ascii
                    Return Encoding.ASCII
                Case PigText.enmTextType.GB2312
                    Return Encoding.GetEncoding("GB2312")
                Case PigText.enmTextType.Unicode
                    Return Encoding.Unicode
                Case PigText.enmTextType.UTF8
                    Return Encoding.UTF8
                Case Else
                    Return Encoding.UTF8
            End Select
        End Get
    End Property


    Friend Sub Init(FilePath As String, IOMode As PigFileSystem.IOMode, Optional Create As Boolean = False)
        Const SUB_NAME As String = "Init"
        Dim strStepName As String = ""
        Try
            mEnmIOMode = IOMode
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                    If Create = True Then
                        If IO.File.Exists(FilePath) = False Then
                            Select Case Me.TextType
                                Case PigText.enmTextType.UnknowOrBin
                                    swMain = New StreamWriter(FilePath, False)
                                Case Else
                                    swMain = New StreamWriter(FilePath, False, Me.mEncTextType)
                            End Select
                            swMain.Close()
                        End If
                    End If
                    Select Case Me.TextType
                        Case PigText.enmTextType.UnknowOrBin
                            srMain = New StreamReader(FilePath)
                        Case Else
                            srMain = New StreamReader(FilePath, Me.mEncTextType)
                    End Select
                Case PigFileSystem.IOMode.ForWriting
                    Select Case Me.TextType
                        Case PigText.enmTextType.UnknowOrBin
                            swMain = New StreamWriter(FilePath, False)
                        Case Else
                            swMain = New StreamWriter(FilePath, False, Me.mEncTextType)
                    End Select
                Case PigFileSystem.IOMode.ForAppending
                    Select Case Me.TextType
                        Case PigText.enmTextType.UnknowOrBin
                            swMain = New StreamWriter(FilePath, True)
                        Case Else
                            swMain = New StreamWriter(FilePath, True, Me.mEncTextType)
                    End Select
                Case Else
                    Throw New Exception("Invalid IOMode " & IOMode.ToString)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            strStepName &= "(" & FilePath & ")"
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
        End Try
    End Sub

    Public ReadOnly Property AtEndOfStream() As Boolean
        Get
            Try
                Select Case mEnmIOMode
                    Case PigFileSystem.IOMode.ForReading
                        Return srMain.EndOfStream
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("AtEndOfStream", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Sub Close()
        Try
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                    srMain.Close()
                Case Else
                    swMain.Close()
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Close", ex)
        End Try
    End Sub

    Public Function ReadAll() As String
        Try
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                    Return srMain.ReadToEnd()
                Case Else
                    Return ""
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("ReadAll", ex)
            Return ""
        End Try
    End Function

    Public Function ReadLine() As String
        Try
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                    Return srMain.ReadLine
                Case Else
                    Return ""
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("ReadLine", ex)
            Return ""
        End Try
    End Function


    Public Sub WriteLine(Text As String)
        Try
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                Case Else
                    swMain.WriteLine(Text)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("WriteLine", ex)
        End Try
    End Sub

    Public Sub WriteBlankLines(Lines As Long)
        For i = 1 To Lines
            Me.WriteLine("")
        Next
    End Sub
    Public Sub Write(Text As String)
        Try
            Select Case mEnmIOMode
                Case PigFileSystem.IOMode.ForReading
                Case Else
                    swMain.Write(Text)
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Write", ex)
        End Try
    End Sub

End Class
