﻿'*******************************************************
'* Name: PigJSonLite
'* Author: Seow Phong
'* Describe: Simple JSON class.
'* Example JSon = "{""RS"":[{""ROW"":[{""name"":""sp_MSalreadyhavegeneration"",""id"":""-1073624922"",""xtype"":""P"",""uid"":""4"",""info"":""0"",""status"":""0"",""base_schema_ver"":""0"",""replinfo"":""0"",""parent_obj"":""0"",""crdate"":""2012-02-10 20:56:30.180"",""ftcatid"":""0"",""schema_ver"":""0"",""stats_schema_ver"":""0"",""type"":""P"",""userstat"":""0"",""sysstat"":""4"",""indexdel"":""0"",""refdate"":""2012-02-10 20:56:30.180"",""version"":""0"",""deltrig"":""0"",""instrig"":""0"",""updtrig"":""0"",""seltrig"":""0"",""category"":""2"",""cache"":""0""}],""TotalRows"":""1"",""IsEOF"":""True""}],""TotalRS"":""1""}"
'* Home Url: http://www.seowphong.com
'* Version: 1.3
'* Create Time: 8/8/2019
'* 1.0.2    10/8/2020   Code changed from VB6 to VB.NET
'* 1.0.3    12/8/2020   Some Function debugging 
'* 1.0.4    16/8/2020   mAddJSonStr debugging
'* 1.0.5    18/8/2020   AddArrayEle debugging
'* 1.0.6    18/8/2020   Add overloaded function AddArrayEle
'* 1.0.7    18/9/2020   Fix AddEle bug 
'* 1.0.8    19/9/2020   Fix AddArrayEle,mAddEle bug and add AddOneArrayEle
'* 1.0.9    1/10/2020   Fix AddArrayEleValue,add AddArrayEleBegin
'* 1.0.10   10/3/2020   Use PigBaseMini，and add IsGetValueErrRetNothing
'* 1.0.11   4/4/2021   Modify AddArrayEleBegin,mSrc2JSonStr,mJSonStr2Src,mLng2Date
'* 1.0.12   5/7/2021   Remove parsing function
'* 1.0.13   6/7/2021   Modify mLng2Date,AddEle,mDate2Lng
'* 1.0.14   27/8/2021  Modify mLng2Date for NETCOREAPP3_1_OR_GREATER
'* 1.1      14/9/2021  Modify xpJSonEleType,mAddJSonStr, and add AddOneObjectEle
'* 1.2      18/11/2024 Add Object2JSon,JSon2Object,ParseJSON,GetJsonArray,GetJsonNode,mGetStrValue,GetBoolValue,GetDateValue,GetDecValue,GetIntValue,GetLngValue
'* 1.3      21/11/2024 Modify mGetStrValue
'*******************************************************
Imports System.Text
#If NET6_0_OR_GREATER Then
Imports System.Text.Json
Imports System.Text.Json.Nodes
#End If
''' <summary>
''' Simple JSON processing class|简单JSon处理类
''' </summary>
Public Class PigJSonLite
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "16"

    Private ReadOnly Property mPigFunc As New PigFunc
#If NET6_0_OR_GREATER Then
    Private mJsonNode As JsonNode
    Public Property JSonNode As JsonNode
        Get
            Return mJsonNode
        End Get
        Friend Set(value As JsonNode)
            mJsonNode = value
        End Set
    End Property
#End If

    ''' <summary>The type of the JSON element</summary>
    Public Enum xpJSonEleType
        ''' <summary>First element</summary>
        FristEle = 0
        ''' <summary>First array element</summary>
        FristArrayEle = 5
        ''' <summary>Not the first element</summary>
        NotFristEle = 10
        ''' <summary>Element, need to be escaped</summary>
        EleValue = 20
        ''' <summary>The value of the array element, which does not need to be escaped</summary>
        ArrayValue = 100
        ''' <summary>The value of the object element, which does not need to be escaped</summary>
        ObjectValue = 110
    End Enum

    ''' <summary>Separator or control symbol in JSON</summary>
    Public Enum xpSymbolType
        ''' <summary>The end flag of the element</summary>
        EleEndFlag = 0
        '''' <summary>The begin flag of the array</summary>
        'ArrayBeginFlag = 5
        ''' <summary>The end flag of the array</summary>
        ArrayEndFlag = 10
        '''' <summary>The separator for the array</summary>
        'ArraySeparator = 20
    End Enum

    Private mbolIsParse As Boolean
    Private msbMain As System.Text.StringBuilder
    Private mstrLastErr As String
    Private mstrClsName = Me.GetType.Name.ToString()

    ''' <summary>
    ''' If Get value Error, return Nothing 
    ''' </summary>
    Private mbolIsGetValueErrRetNothing As Boolean = False
    Public Property IsGetValueErrRetNothing() As Boolean
        Get
            Return mbolIsGetValueErrRetNothing
        End Get
        Friend Set(ByVal value As Boolean)
            mbolIsGetValueErrRetNothing = value
        End Set
    End Property


    ''' <summary>Full JSON string, the assembled JSON string is in compact format and can be displayed in a third-party format.</summary>
    Public ReadOnly Property MainJSonStr As String
        Get
            MainJSonStr = msbMain.ToString
        End Get
    End Property


    ''' <summary>Whether the JSON has been resolved or not, the value of JSON can be read through the key.</summary>
    Public ReadOnly Property IsParse As Boolean
        Get
            IsParse = mbolIsParse
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of microseconds of Greenwich mean time for the current time|获取当前时间的格林威治时间微秒数
    ''' </summary>
    ''' <param name="DateValue"></param>
    ''' <returns></returns>
    Private Function mDate2Lng(DateValue As DateTime) As Long
        Dim dteStart As New DateTime(1970, 1, 1)
        Dim mtsTimeDiff As TimeSpan = DateValue - dteStart
        Try
            Return mtsTimeDiff.TotalMilliseconds
        Catch ex As Exception
            Me.SetSubErrInf("mDate2Lng", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Add a non array element (long value)</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="IntValue">The long or integer value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, IntValue As Long, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet As String = Me.mAddEle(EleKey, IntValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.IntValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (string value)</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="StrValue">The string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, StrValue As String, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet As String = Me.mAddEle(EleKey, StrValue, IsFirstEle, True)
            If strRet <> "" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.StrValue", ex)
        End Try
    End Sub


    ''' <summary>Add a non array element (boolean value)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="BoolValue">The boolean value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, BoolValue As Boolean, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet = Me.mAddEle(EleKey, BoolValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.BoolValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (date value)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="DecValue">The long value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, DecValue As Decimal, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet = Me.mAddEle(EleKey, DecValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.DateValue", ex)
        End Try
    End Sub


    ''' <summary>Add a non array element (date value, need to specify whether it is a local time)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="DateValue">The date value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, DateValue As DateTime, Optional IsFirstEle As Boolean = False)
        Try
            Dim lngDate As Long = Me.mDate2Lng(DateValue)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Dim strRet = Me.mAddEle(EleKey, lngDate.ToString, IsFirstEle, False)
            If strRet <> "" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.DateValue.IsLocalTime", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="LngValue">The number of milliseconds since 1970-1-1</param>
    ''' <param name="IsLocalTime">Convert to local time</param>
    ''' <returns></returns>
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Private Function mLng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            Dim intHourAdd As Integer = 0
            If IsLocalTime = False Then
                Dim oTimeZoneInfo As System.TimeZoneInfo
                oTimeZoneInfo = System.TimeZoneInfo.Local
                intHourAdd = oTimeZoneInfo.GetUtcOffset(Now).Hours
            End If

            Return dteStart.AddSeconds(LngValue + intHourAdd * 3600)
            Me.ClearErr()
        Catch ex As Exception
            Return dteStart
            Me.SetSubErrInf("mLng2Date", ex)
        End Try
    End Function
#Else
    Private Function mLng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            If IsLocalTime = False Then
                mLng2Date = dteStart.AddMilliseconds(LngValue - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000)
            Else
                mLng2Date = dteStart.AddMilliseconds(LngValue)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mLng2Date", ex)
            Return DateTime.MinValue
        End Try
    End Function
#End If


    ''' <summary>Add a JSON symbol</summary>
    ''' <param name="SymbolType">Symbol type</param>
    Public Sub AddSymbol(SymbolType As xpSymbolType)
        Try
            Select Case SymbolType
                'Case xpSymbolType.ArrayBeginFlag
                '    msbMain.Append("[")
                Case xpSymbolType.ArrayEndFlag
                    msbMain.Append("]")
                'Case xpSymbolType.ArraySeparator
                '    msbMain.Append(",")
                Case xpSymbolType.EleEndFlag
                    msbMain.Append("}")
                Case Else
                    Throw New Exception("Invalid SymbolType")
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddSymbol", ex)
        End Try
    End Sub


    ''' <summary>Add a non array JSON element</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="EleValue">The string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Private Function mAddEle(EleKey As String, EleValue As String, IsFirstEle As Boolean, IsChgCtlStr As Boolean) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "Add EleKey"
            If IsFirstEle = True Then
                If EleKey <> "" Then
                    strRet = mAddJSonStr(msbMain, xpJSonEleType.FristEle, EleKey, "", False)
                Else
                    strRet = "OK"
                End If
            Else
                strRet = mAddJSonStr(msbMain, xpJSonEleType.NotFristEle, EleKey, "", False)
            End If
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.EleValue, "", EleValue, IsChgCtlStr)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAddEle", strStepName, ex)
        End Try
    End Function

    ''' <summary>Add a array JSON element begin</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddArrayEleBegin(EleKey As String, IsFirstEle As Boolean)
        Dim strStepName As String = "", strRet As String = ""
        Try
            mSrc2JSonStr(EleKey)
            With msbMain
                If IsFirstEle = False Then
                    .Append(",")
                Else
                    .Append("{")
                End If
                .Append("""")
                .Append(EleKey)
                .Append(""":[")
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEle", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add a array JSON element begin</summary>
    ''' <param name="EleKey">The key of the element</param>
    Public Overloads Sub AddArrayEleBegin(EleKey As String)
        Dim strStepName As String = "", strRet As String = ""
        Try
            mSrc2JSonStr(EleKey)
            With msbMain
                .Append(",")
                .Append("""")
                .Append(EleKey)
                .Append(""":[")
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEle", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add a array JSON element</summary>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddArrayEleValue(ArrayEleValue As String, IsFirstEle As Boolean)
        Dim strStepName As String = "", strRet As String = ""
        Try
            If IsFirstEle = False Then msbMain.Append(",")
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", ArrayEleValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add one array JSON element</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddOneArrayEle(EleKey As String, ArrayEleValue As String, Optional IsFirstEle As Boolean = False)
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "Check EleKey"
            If EleKey = "" Then Throw New Exception("Need EleKey")
            If IsFirstEle = True Then
                strRet = mAddJSonStr(msbMain, xpJSonEleType.FristEle, EleKey, "")
            Else
                strRet = mAddJSonStr(msbMain, xpJSonEleType.NotFristEle, EleKey, "")
            End If
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", "[" & ArrayEleValue & "]")
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub


    ''' <summary>Add a array JSON element</summary>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    Public Overloads Sub AddArrayEleValue(ArrayEleValue As String)
        Dim strStepName As String = "", strRet As String = ""
        Try
            msbMain.Append(",")
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", ArrayEleValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Reset so that JSON can be assembled.</summary>
    Public Sub Reset()
        Try
            If Not msbMain Is Nothing Then msbMain = Nothing
            msbMain = New System.Text.StringBuilder("")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Reset", ex)
        End Try
    End Sub

    Private Function mAddJSonStr(ByRef sbJSonStr As System.Text.StringBuilder, JSonEleType As xpJSonEleType, ColName As String, ColValue As String, Optional IsChgCtlStr As Boolean = False) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
10:         Select Case JSonEleType
                Case xpJSonEleType.FristEle, xpJSonEleType.NotFristEle, xpJSonEleType.FristArrayEle
                    Select Case JSonEleType
                        Case xpJSonEleType.FristEle
                            sbJSonStr.Append("{")
                        Case xpJSonEleType.FristArrayEle
                            sbJSonStr.Append("[")
                        Case xpJSonEleType.NotFristEle
                            sbJSonStr.Append(",")
                    End Select
                    If ColName <> "" Then
                        mSrc2JSonStr(ColName)
                        sbJSonStr.Append("""")
                        sbJSonStr.Append(ColName)
                        sbJSonStr.Append(""":")
                    End If
                Case xpJSonEleType.EleValue, xpJSonEleType.ArrayValue, xpJSonEleType.ObjectValue
                    Select Case JSonEleType
                        Case xpJSonEleType.EleValue
                            If IsChgCtlStr = True Then mSrc2CtlStr(ColValue)
                            mSrc2JSonStr(ColValue)
                            sbJSonStr.Append("""")
                            sbJSonStr.Append(ColValue)
                            sbJSonStr.Append("""")
                        Case xpJSonEleType.ArrayValue, xpJSonEleType.ObjectValue
                            sbJSonStr.Append(ColValue)
                    End Select
                Case Else
                    strRet = "Invalid jsoneletype:" & JSonEleType.ToString
                    Throw New Exception(strRet)
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAddJSonStr", strStepName, ex)
        End Try
    End Function

    Private Sub mSrc2CtlStr(ByRef SrcStr As String)
        Try
            If SrcStr.IndexOf(vbCrLf) > 0 Then SrcStr = Replace(SrcStr, vbCrLf, "\r\n")
            If SrcStr.IndexOf(vbCr) > 0 Then SrcStr = Replace(SrcStr, vbCr, "\r")
            If SrcStr.IndexOf(vbTab) > 0 Then SrcStr = Replace(SrcStr, vbTab, "\t")
            If SrcStr.IndexOf(vbBack) > 0 Then SrcStr = Replace(SrcStr, vbBack, "\b")
            If SrcStr.IndexOf(vbFormFeed) > 0 Then SrcStr = Replace(SrcStr, vbFormFeed, "\f")
            If SrcStr.IndexOf(vbVerticalTab) > 0 Then SrcStr = Replace(SrcStr, vbVerticalTab, "\v")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mSrc2CtlStr", ex)
        End Try
    End Sub

    Private Sub mCtlStr2Src(ByRef CtlStr As String)
        Try
            If CtlStr.IndexOf("\r\n") > 0 Then CtlStr = Replace(CtlStr, "\r\n", vbCrLf)
            If CtlStr.IndexOf("\r") > 0 Then CtlStr = Replace(CtlStr, "\r", vbCr)
            If CtlStr.IndexOf("\t") > 0 Then CtlStr = Replace(CtlStr, "\t", vbTab)
            If CtlStr.IndexOf("\b") > 0 Then CtlStr = Replace(CtlStr, "\b", vbBack)
            If CtlStr.IndexOf(vbFormFeed) > 0 Then CtlStr = Replace(CtlStr, "\f", vbFormFeed)
            If CtlStr.IndexOf(vbVerticalTab) > 0 Then CtlStr = Replace(CtlStr, "\v", vbVerticalTab)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mCtlStr2Src", ex)
        End Try
    End Sub

    Private Sub mSrc2JSonStr(ByRef SrcStr As String)
        Try
            If InStr(SrcStr, "\") <> 0 Then
                SrcStr = Replace(SrcStr, "\", "\\")
            End If
            If InStr(SrcStr, """") <> 0 Then
                SrcStr = Replace(SrcStr, """", "\""")
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mSrc2JSonStr", ex)
        End Try
    End Sub

    Private Sub mJSonStr2Src(ByRef JSonStr As String)
        Try
            If InStr(JSonStr, "\""") <> 0 Then
                JSonStr = Replace(JSonStr, "\""", """")
            End If
            If InStr(JSonStr, "\\") <> 0 Then
                JSonStr = Replace(JSonStr, "\\", "\")
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mJSonStr2Src", ex)
        End Try
    End Sub


    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.Reset()
    End Sub

    ''' <summary>Add one object JSON element</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="ObjectEleValue">The object string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddOneObjectEle(EleKey As String, ObjectEleValue As String, Optional IsFirstEle As Boolean = False)
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "Check EleKey"
            If EleKey = "" Then Throw New Exception("Need EleKey")
            If IsFirstEle = True Then
                strRet = mAddJSonStr(msbMain, xpJSonEleType.FristEle, EleKey, "")
            Else
                strRet = mAddJSonStr(msbMain, xpJSonEleType.NotFristEle, EleKey, "")
            End If
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ObjectValue, "", ObjectEleValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception

            Me.SetSubErrInf("AddObjectEleValue", strStepName, ex)
        End Try
    End Sub


#If NET6_0_OR_GREATER Then

    ''' <summary>
    ''' Get the float value of the JSon element|获取JSon元素的浮点值
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].ROW[0].indexdel|例如：RS[0].ROW[0].indexdel</param>
    ''' <returns></returns>
    Public Function GetDecValue(JSonKey As String) As Decimal
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return CDec(strValue)
        Catch ex As Exception
            Me.SetSubErrInf("GetIntValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function

    ''' <summary>
    ''' Get the date value of the JSon element|获取JSon元素的日期值
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].ROW[0].refdate|例如：RS[0].ROW[0].refdate</param>
    ''' <param name="IsLocalTime">Is it local time|是否本地时间</param>
    ''' <returns></returns>
    Public Function GetDateValue(JSonKey As String, Optional IsLocalTime As Boolean = True) As DateTime
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            If IsDate(strValue) = True Then
                GetDateValue = CDate(strValue)
            ElseIf IsNumeric(strValue) Then
                GetDateValue = Me.mLng2Date(CLng(strValue), IsLocalTime)
                If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            Else
                Throw New Exception("Not date string")
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return DateTime.MinValue
            End If
        End Try
    End Function


    ''' <summary>
    ''' Get the long integer value of the JSon element|获取JSon元素的长整型值
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].TotalRows|例如：RS[0].TotalRows</param>
    ''' <returns></returns>
    Public Function GetLngValue(JSonKey As String) As Long
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return CLng(strValue)
        Catch ex As Exception
            Me.SetSubErrInf("GetLngValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function

    ''' <summary>
    ''' 获取JSon元素的整型值|Get the integer value of the JSon element
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].TotalRows|例如：RS[0].TotalRows</param>
    ''' <returns></returns>
    Public Function GetIntValue(JSonKey As String) As Integer
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return CInt(strValue)
        Catch ex As Exception
            Me.SetSubErrInf("GetIntValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function


    ''' <summary>
    ''' 获取JSon元素的布尔值|Get the boolean value of the JSon element
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].IsEOF|例如：RS[0].IsEOF</param>
    ''' <returns></returns>
    Public Function GetBoolValue(JSonKey As String) As Boolean
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return CBool(strValue)
        Catch ex As Exception
            Me.SetSubErrInf("GetBoolValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return False
            End If
        End Try
    End Function

    ''' <summary>
    ''' 获取JSon元素的字符串值|Get the string value of the JSon element
    ''' </summary>
    ''' <param name="JSonKey">For example:RS[0].ROW[0].name|例如：RS[0].ROW[0].name</param>
    ''' <returns></returns>
    Public Function GetStrValue(JSonKey As String) As String
        Try
            Dim strRet As String = ""
            Dim strValue As String = ""
            strRet = Me.mGetStrValue(JSonKey, strValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Return strValue
        Catch ex As Exception
            Me.SetSubErrInf("GetStrValue", ex)
            Return ""
        End Try
    End Function

    Private Function mGetStrValue(JSonKey As String, ByRef OutStrValue As String) As String
        Dim LOG As New PigStepLog("mGetStrValue")
        Dim strJSonKey As String = ""
        Dim strSubJSonKey As String = ""
        Try
            OutStrValue = ""
            If Me.JSonNode IsNot Nothing Then
                If InStr(JSonKey, ".") = 0 Then
                    OutStrValue = Me.JSonNode(JSonKey).ToString()
                Else
                    strJSonKey = JSonKey
                    Dim oSubNode = Me.JSonNode
                    Do While True
                        strSubJSonKey = Me.mPigFunc.GetStr(strJSonKey, "", ".", True)
                        If strSubJSonKey = "" Then
                            strSubJSonKey = strJSonKey
                            If InStr(strSubJSonKey, "[") > 0 Then
                                Dim strIndex As String = Me.mPigFunc.GetStr(strSubJSonKey, "[", "]")
                                If IsNumeric(strIndex) Then
                                    Dim intIndex As Integer = CInt(strIndex)
                                    OutStrValue = oSubNode(strSubJSonKey)(intIndex).ToString
                                End If
                            Else
                                OutStrValue = oSubNode(strSubJSonKey).ToString
                            End If
                            Exit Do
                        End If
                        If InStr(strSubJSonKey, "[") > 0 Then
                            Dim strIndex As String = Me.mPigFunc.GetStr(strSubJSonKey, "[", "]")
                            If IsNumeric(strIndex) Then
                                Dim intIndex As Integer = CInt(strIndex)
                                oSubNode = oSubNode(strSubJSonKey)(intIndex)
                            End If
                        Else
                            oSubNode = oSubNode(strSubJSonKey)
                        End If
                    Loop

                End If
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf("JSonKey=" & strJSonKey)
            LOG.AddStepNameInf("SubJSonKey=" & strSubJSonKey)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 获取一个JSon数组|Get a JSon array
    ''' </summary>
    ''' <param name="JSonKey"></param>
    ''' <param name="ArrayIndex"></param>
    ''' <returns></returns>
    Public Function GetJsonArray(JSonKey As String, ArrayIndex As Integer) As JsonArray
        Try
            Return Me.JsonNode(ArrayIndex)(JSonKey)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("GetJsonArray", ex)
        End Try
    End Function

    ''' <summary>
    ''' 获取一个JSon节点|Get a JSon node
    ''' </summary>
    ''' <param name="JSonKey"></param>
    ''' <returns></returns>
    Public Function GetJsonNode(JSonKey As String) As JsonNode
        Try
            Return Me.JsonNode(JSonKey)
        Catch ex As Exception
            Me.SetSubErrInf("GetJsonNode", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 解析JSon字符串|Parse JSon string
    ''' </summary>
    ''' <param name="JSonStr"></param>
    ''' <returns></returns>
    Public Function ParseJSON(JSonStr As String) As String
        Try
            Me.JsonNode = JsonNode.Parse(JSonStr)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ParseJSON", ex)
        End Try
    End Function

    ''' <summary>
    ''' JSon serialization|JSon序列化
    ''' </summary>
    ''' <param name="InObj">Input Object|输入对象</param>
    ''' <param name="OutJSon">|输出的JSon</param>
    ''' <param name="IsIndented">Is it indented|是否缩进</param>
    ''' <returns></returns>
    Public Function Object2JSon(InObj As Object, ByRef OutJSon As String, Optional IsIndented As Boolean = False) As String
        Try
            Dim Options = New JsonSerializerOptions()
            Options.WriteIndented = True
            OutJSon = JsonSerializer.Serialize(InObj, Options)
            Return "OK"
        Catch ex As Exception
            OutJSon = ""
            Return Me.GetSubErrInf("JSon2Object", ex)
        End Try
    End Function

    ''' <summary>
    ''' JSon deserialization|JSon反序列化
    ''' </summary>
    ''' <param name="InJSon">Enter JSon string|输入JSon字符串</param>
    ''' <param name="OutObject">Output Object|输出对象</param>
    ''' <returns></returns>
    Public Function JSon2Object(Of T)(InJSon As String, ByRef OutObject As T) As String
        Try
            Dim Options = New JsonSerializerOptions()
            OutObject = JsonSerializer.Deserialize(Of T)(InJSon, Options)
            Return "OK"
        Catch ex As Exception
            OutObject = Nothing
            Return Me.GetSubErrInf("JSon2Object", ex)
        End Try
    End Function

#End If

End Class
