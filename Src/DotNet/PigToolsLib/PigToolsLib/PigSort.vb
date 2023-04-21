'**********************************
'* Name: PigFile
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Sort processing class|排序处理类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 6/9/2022
'* 1.1  10/9/2022   Modify New, add AddValue,MaxSortNo,MinSortNo,GetValue
'**********************************
''' <summary>
''' Sort processing class|排序处理类
''' </summary>
Public Class PigSort
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"
    Public Enum EnmSortWhat
        SortString = 1
        SortLong = 2
        SortDate = 3
        SortDecimal = 4
        SortByte = 5
    End Enum

    Public ReadOnly Property SortString As List(Of String)
    Public ReadOnly Property SortLong As List(Of Long)
    Public ReadOnly Property SortDate As List(Of Date)
    Public ReadOnly Property SortDecimal As List(Of Decimal)
    Public ReadOnly Property SortByte As List(Of Byte)

    Public ReadOnly Property SortWhat As EnmSortWhat

    Public Sub New(SortWhat As EnmSortWhat)
        MyBase.New(CLS_VERSION)
        Try
            Select Case SortWhat
                Case EnmSortWhat.SortByte
                    Me.SortByte = New List(Of Byte)
                Case EnmSortWhat.SortDate
                    Me.SortDate = New List(Of Date)
                Case EnmSortWhat.SortDecimal
                    Me.SortDecimal = New List(Of Decimal)
                Case EnmSortWhat.SortLong
                    Me.SortLong = New List(Of Long)
                Case EnmSortWhat.SortString
                    Me.SortString = New List(Of String)
                Case Else
                    Throw New Exception("Invalid SortWhat is " & SortWhat.ToString)
            End Select
            Me.SortWhat = SortWhat
        Catch ex As Exception

        End Try
    End Sub


    Public Sub AddValue(Value As String)
        Try
            If Me.SortWhat = EnmSortWhat.SortString Then
                Me.SortString.Add(Value)
            Else
                Throw New Exception("SortWhat mismatch")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("AddValue", ex)
        End Try
    End Sub
    Public Sub AddValue(Value As Long)
        Try
            If Me.SortWhat = EnmSortWhat.SortLong Then
                Me.SortLong.Add(Value)
            Else
                Throw New Exception("SortWhat mismatch")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("AddValue", ex)
        End Try
    End Sub

    Public Sub AddValue(Value As Date)
        Try
            If Me.SortWhat = EnmSortWhat.SortDate Then
                Me.SortDate.Add(Value)
            Else
                Throw New Exception("SortWhat mismatch")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("AddValue", ex)
        End Try
    End Sub


    Public Sub AddDecValue(Value As Decimal)
        Try
            If Me.SortWhat = EnmSortWhat.SortDecimal Then
                Me.SortDecimal.Add(Value)
            Else
                Throw New Exception("SortWhat mismatch")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("AddDecValue", ex)
        End Try
    End Sub

    Public Sub AddByteValue(Value As Byte)
        Try
            If Me.SortWhat = EnmSortWhat.SortByte Then
                Me.SortByte.Add(Value)
            Else
                Throw New Exception("SortWhat mismatch")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("AddByteValue", ex)
        End Try
    End Sub

    Public Sub Sort()
        Try
            Select Case Me.SortWhat
                Case EnmSortWhat.SortByte
                    Me.SortByte.Sort()
                Case EnmSortWhat.SortDate
                    Me.SortDate.Sort()
                Case EnmSortWhat.SortDecimal
                    Me.SortDecimal.Sort()
                Case EnmSortWhat.SortLong
                    Me.SortLong.Sort()
                Case EnmSortWhat.SortString
                    Me.SortString.Sort()
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("Sort", ex)
        End Try
    End Sub

    Public ReadOnly Property MaxSortNo As Long
        Get
            Try
                Select Case Me.SortWhat
                    Case EnmSortWhat.SortByte
                        Return Me.SortByte.Count - 1
                    Case EnmSortWhat.SortDate
                        Return Me.SortDate.Count - 1
                    Case EnmSortWhat.SortDecimal
                        Return Me.SortDecimal.Count - 1
                    Case EnmSortWhat.SortLong
                        Return Me.SortLong.Count - 1
                    Case EnmSortWhat.SortString
                        Return Me.SortString.Count - 1
                    Case Else
                        Return -1
                End Select
            Catch ex As Exception
                Me.SetSubErrInf("Sort", ex)
                Return -1
            End Try
        End Get
    End Property

    Public Function GetMinValue() As Object
        Return Me.GetValue(0)
    End Function

    Public Function GetMaxValue() As Object
        Return Me.GetValue(Me.MaxSortNo)
    End Function
    Public Function GetValue(SortNo As Long) As Object
        Try
            If SortNo < 0 Then
                SortNo = 0
            ElseIf SortNo > Me.MaxSortNo Then
                SortNo = Me.MaxSortNo
            End If
            Select Case SortWhat
                Case EnmSortWhat.SortByte
                    GetValue = Me.SortByte.Item(SortNo)
                Case EnmSortWhat.SortDate
                    GetValue = Me.SortDate.Item(SortNo)
                Case EnmSortWhat.SortDecimal
                    GetValue = Me.SortDecimal.Item(SortNo)
                Case EnmSortWhat.SortLong
                    GetValue = Me.SortLong.Item(SortNo)
                Case EnmSortWhat.SortString
                    GetValue = Me.SortString.Item(SortNo)
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("GetValue", ex)
            Return Nothing
        End Try
    End Function
End Class
