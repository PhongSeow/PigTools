'*******************************************************
'* Name: PigJSonForCacheQuery
'* Author: Seow Phong
'* Describe: It is specially designed for JSON returned by CmdSQLSrvSp.CacheQuery and CmdSQLSrvText.CacheQuery of PigSQLSrvLib to simplify writing format.
'* Home Url: http://www.seowphong.com
'* Version: 1.1
'* Create Time: 8/8/2019
'* 1.1      10/10/2021  Add ToNextRow,ToPrevRow,ToNextRecordset,ToPrevRecordset and modify CurrentRow,CurrentRecordset,IsEOF
'*******************************************************

Public Class PigJSonForCacheQuery
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.6"
    Private moPigJSon As PigJSon
    Public Sub New(JSonStr As String)
        MyBase.New(CLS_VERSION)
        Dim strStepName As String = ""
        Try
            strStepName = "New PigJSon"
            moPigJSon = New PigJSon(JSonStr)
            If moPigJSon.LastErr <> "" Then
                Me.PrintDebugLog("New", strStepName, moPigJSon.LastErr)
                Throw New Exception(moPigJSon.LastErr)
            End If
            Me.CurrentRecordset = 1
            Me.CurrentRow = 1
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", strStepName, ex)
        End Try
    End Sub

    Private Function mJSonKeyHead(Optional IsColumn As Boolean = True) As String
        Try
            mJSonKeyHead = "RS[" & Me.CurrentRecordset - 1 & "]."
            If IsColumn = True Then
                mJSonKeyHead &= "ROW[" & Me.CurrentRow - 1 & "]."
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mJSonKeyHead", ex)
            Return ""
        End Try
    End Function

    ''' <summary>Gets the boolean value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetBoolValue(JSonKey As String) As Boolean
        Try
            Return moPigJSon.GetBoolValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetBoolValue", ex)
            Return ""
        End Try
    End Function

    ''' <summary>Gets the date value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Overloads Function GetDateValue(JSonKey As String) As DateTime
        Try
            Return moPigJSon.GetDateValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", ex)
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>Gets the date value of JSON, has IsLocalTime option.</summary>
    ''' <param name="JSonKey">JSON key</param>
    ''' <param name="IsLocalTime">Is it local time</param>
    Public Overloads Function GetDateValue(JSonKey As String, IsLocalTime As Boolean) As DateTime
        Try
            Return moPigJSon.GetDateValue(Me.mJSonKeyHead & JSonKey, IsLocalTime)
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", ex)
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>Gets the long value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetDecValue(JSonKey As String) As Decimal
        Try
            Return moPigJSon.GetDecValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetDecValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Gets the integer value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetIntValue(JSonKey As String) As Integer
        Try
            Return moPigJSon.GetIntValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetIntValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Gets the long value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetLngValue(JSonKey As String) As Long
        Try
            Return moPigJSon.GetLngValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetLngValue", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Gets the string value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetStrValue(JSonKey As String) As String
        Try
            Return moPigJSon.GetStrValue(Me.mJSonKeyHead & JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetStrValue", ex)
            Return ""
        End Try
    End Function


    Public ReadOnly Property TotalRecordset() As Integer
        Get
            Try
                Return moPigJSon.GetIntValue("TotalRS")
            Catch ex As Exception
                Me.SetSubErrInf("TotalRecordset", ex)
                Return -1
            End Try
        End Get
    End Property

    Private mintTotalRows As Integer = 0
    Public ReadOnly Property TotalRows() As Integer
        Get
            Try
                Return moPigJSon.GetIntValue(Me.mJSonKeyHead（False) & "TotalRows")
            Catch ex As Exception
                Me.SetSubErrInf("TotalRows", ex)
                Return -1
            End Try
        End Get
    End Property

    Private mintCurrentRecordset As Integer
    Public Property CurrentRecordset() As Integer
        Get
            Return mintCurrentRecordset
        End Get
        Set(ByVal value As Integer)
            Select Case value
                Case 1 To Me.TotalRecordset
                    mintCurrentRecordset = value
                Case < 1
                    mintCurrentRecordset = 1
                Case > Me.TotalRecordset
                    mintCurrentRow = Me.TotalRecordset
            End Select
        End Set
    End Property

    Private mintCurrentRow As Integer
    Public Property CurrentRow() As Integer
        Get
            Return mintCurrentRow
        End Get
        Friend Set(ByVal value As Integer)
            Select Case value
                Case 1 To Me.TotalRows
                    mintCurrentRow = value
                Case < 1
                    mintCurrentRow = 1
                Case > Me.TotalRows
                    mintCurrentRow = Me.TotalRows
            End Select
        End Set
    End Property

    Public ReadOnly Property IsEOF() As Boolean
        Get
            Try
                Return moPigJSon.GetBoolValue(Me.mJSonKeyHead（False) & "IsEOF")
            Catch ex As Exception
                Me.SetSubErrInf("IsEOF", ex)
                Return True
            End Try
        End Get
    End Property

    Public Sub ToNextRow()
        Me.CurrentRow += 1
    End Sub

    Public Sub ToPrevRow()
        Me.CurrentRow -= 1
    End Sub

    Public Sub ToNextRecordset()
        Me.CurrentRecordset += 1
    End Sub

    Public Sub ToPrevRecordset()
        Me.CurrentRecordset -= 1
    End Sub

End Class
