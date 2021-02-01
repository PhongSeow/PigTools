'**********************************
'* Name: PigXml
'* Author: Seow Phong.
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Processing XML string splicing and parsing. 处理XML字符串拼接及解析
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.10
'* Create Time: 8/11/2019
'1.0.2  2019-11-10  修改bug
'1.0.3  2020-5-26  修改bug
'1.0.5  2020-6-5  增加 CData处理和流处理
'1.0.6  2020-6-7  增加 CData处理和流处理优化
'1.0.8  2020-6-19  修改 FillByXmlReader
'1.0.9  2020-6-20  修改 Bug
'1.0.10 2020-7-6    增加 XmlGetInt
'*******************************************************

Imports System.Xml
Public Class PigXml
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.10"
    Private mstrMainXml As String
    Private mslMain As SortedList

    ''' <summary>增加一个XML标记的部分</summary>
    Private Enum xpXMLAddWhere
        ''' <summary>左标记</summary>
        Left = 0
        ''' <summary>右标记</summary>
        Right = 1
        ''' <summary>左右标记</summary>
        Both = 2
        ''' <summary>只有值</summary>
        ValueOnly = 3
    End Enum

    Private msbMain As System.Text.StringBuilder    '主体的XML
    Private mbolIsCrLf As Boolean


    Public Sub New(IsCrLf As Boolean)
        MyBase.New(CLS_VERSION)
        mbolIsCrLf = IsCrLf
        Me.Clear()
    End Sub

    Public Sub SetMainXml(InXml As String)
        mstrMainXml = InXml
    End Sub

    ''' <summary>通过XmlReader填充元素</summary>
    ''' <param name="oXmlReader">XmlReader对象</param>
    ''' <param name="DateKeys">日期型的键值，有则转为日期型</param>
    Public Function FillByXmlReader(ByRef oXmlReader As XmlReader, Optional DateKeys As String = "") As String
        Dim strStepName As String = ""
        Try
            strStepName = "mslMain.Clear"
            mslMain.Clear()
            strStepName = "读取XML流"
            While oXmlReader.Read
                Select Case oXmlReader.NodeType
                    Case XmlNodeType.Element
                        Dim strEleName As String = oXmlReader.Name
                        Dim intEleIdx As Integer = mslMain.IndexOfKey(strEleName)
                        strStepName = strEleName & ".Read"
                        oXmlReader.Read() '读取内容
                        Dim strValue As String = oXmlReader.Value
                        Select Case oXmlReader.NodeType
                            Case XmlNodeType.CDATA
                                If intEleIdx = -1 Then
                                    mslMain.Add(strEleName, strValue)
                                End If
                            Case Else
                                If InStr(strValue, ".") > 0 Then
                                    Dim decValue As Decimal = CDec(strValue)
                                    If intEleIdx = -1 Then
                                        mslMain.Add(strEleName, decValue)
                                    End If
                                Else
                                    Dim lngValue As Long = CLng(strValue)
                                    If InStr(DateKeys, "<" & strEleName & ">") > 0 Then
                                        Dim dteValue As DateTime = Me.mLng2Date(lngValue)
                                        If intEleIdx = -1 Then
                                            mslMain.Add(strEleName, dteValue)
                                        End If
                                    Else
                                        If intEleIdx = -1 Then
                                            mslMain.Add(strEleName, lngValue)
                                        End If
                                    End If
                                End If
                        End Select
                End Select
            End While
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("FillByXmlReader", strStepName, ex)
        End Try
    End Function


    ''' <summary>获取一个元素字符串值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetStr(XMLSign As String) As String
        Return Me.mXmlGetStr(mstrMainXml, XMLSign, False)
    End Function

    ''' <summary>获取一个元素长整值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetLong(XMLSign As String) As Long
        Try
            XmlGetLong = CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetLong", ex)
            Return 0
        End Try
    End Function

    ''' <summary>获取一个元素整值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetInt(XMLSign As String) As Integer
        Try
            XmlGetInt = CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetLong", ex)
            Return 0
        End Try
    End Function

    ''' <summary>获取一个元素浮点值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDec(XMLSign As String) As Decimal
        Try
            XmlGetDec = CDec(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDec", ex)
            Return 0
        End Try
    End Function

    ''' <summary>整型转日期型</summary>
    ''' <param name="LngValue">整型值，即1970-1-1以来的秒数</param>
    ''' <param name="IsLocalTime">是否本地时间，否则为格林威治时间</param>
    Private Function mLng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            Dim intHourAdd As Integer = 0
            If IsLocalTime = True Then intHourAdd = System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours

            Return dteStart.AddSeconds(LngValue + intHourAdd * 3600)
            Me.ClearErr()
        Catch ex As Exception
            Return dteStart
            Me.SetSubErrInf("mLng2Date", ex)
        End Try
    End Function

    ''' <summary>获取一个元素日期值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDate(XMLSign As String) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            XmlGetDate = Me.mLng2Date(CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False)))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDate", ex)
            Return dteStart
        End Try
    End Function

    ''' <summary>获取一个元素日期值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDate(XMLSign As String, IsLocalTime As Boolean) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            XmlGetDate = Me.mLng2Date(CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False)), IsLocalTime)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDate", ex)
            Return dteStart
        End Try
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">源XML串，读取后会截去</param>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Private Overloads Function mXmlGetStr(ByRef SrcXmlStr As String, XMLSign As String, IsCData As Boolean) As String
        Try
            Dim strBegin As String, strEnd As String
            strBegin = "<" & XMLSign & ">"
            strEnd = "</" & XMLSign & ">"
            If IsCData = True Then
                strBegin &= "<![CDATA["
                strEnd = "]]>" & strEnd
            End If
            mXmlGetStr = Me.GetStr(SrcXmlStr, strBegin, strEnd, True)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mXmlGetStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function XmlGetStr(XMLSign As String, IsCData As Boolean) As String
        Return Me.mXmlGetStr(mstrMainXml, XMLSign, IsCData)
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">XML源串</param>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetStr(ByRef SrcXmlStr As String, XMLSign As String) As String
        Return Me.mXmlGetStr(SrcXmlStr, XMLSign, False)
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">XML源串</param>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function XmlGetStr(ByRef SrcXmlStr As String, XMLSign As String, IsCData As Boolean) As String
        Return Me.mXmlGetStr(SrcXmlStr, XMLSign, IsCData)
    End Function



    ''' <remarks>截取字符串</remarks>
    Public Function GetStr(ByRef SrcStr As String, BeginKey As String, EndKey As String, Optional IsCut As Boolean = True) As String
        Dim intBegin As Integer
        Dim intEnd As Integer
        Dim intBeginLen As Integer
        Dim intEndLen As Integer
        Try
            intBeginLen = Len(BeginKey)
            intBegin = InStr(SrcStr, BeginKey)
            intEndLen = Len(EndKey)
            If intEndLen = 0 Then
                intEnd = Len(SrcStr) + 1
            Else
                intEnd = InStr(intBegin + intBeginLen + 1, SrcStr, EndKey)
                If intEnd = 0 Then Err.Raise(-1, , "intEnd为0")
            End If
            If intEnd <= intBegin Then Err.Raise(-1, , "intEnd <= intBegin")
            If intBegin = 0 Then Err.Raise(-1, , "intBegin为0")
            GetStr = Mid(SrcStr, intBegin + intBeginLen, (intEnd - intBegin - intBeginLen))
            If IsCut = True Then
                SrcStr = Left(SrcStr, intBegin - 1) & Mid(SrcStr, intEnd + intEndLen)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetStr", ex)
            Return ""
        End Try
    End Function


    ''' <summary>增加一个完整元素</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, mbolIsCrLf)
            If AddEle <> "OK" Then Err.Raise(-1, , AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, IsCData As Boolean) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, mbolIsCrLf, , IsCData)
            If AddEle <> "OK" Then Err.Raise(-1, , AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素带缩进</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    ''' <remarks></remarks>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, LeftTab As Integer) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, mbolIsCrLf, LeftTab)
            If AddEle <> "OK" Then Err.Raise(-1, , AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素带缩进</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    ''' <param name="IsCData">是否CDATA</param>
    ''' <remarks></remarks>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, LeftTab As Integer, IsCData As Boolean) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, mbolIsCrLf, LeftTab, IsCData)
            If AddEle <> "OK" Then Err.Raise(-1, , AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个XML值</summary>
    Public Function AddEleValue(XMLValue As String) As String
        Try
            AddEleValue = Me.mXMLAddStr(msbMain, "", XMLValue, xpXMLAddWhere.ValueOnly)
            If AddEleValue <> "OK" Then Err.Raise(-1, , AddEleValue)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleValue", ex)
        End Try
    End Function

    ''' <summary>增加一个XML左标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function AddEleLeftSign(XMLSign As String) As String
        Try
            AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Left, mbolIsCrLf)
            If AddEleLeftSign <> "OK" Then Err.Raise(-1, , AddEleLeftSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftSign", ex)
        End Try
    End Function

    ''' <summary>增加一个XML左标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    Public Overloads Function AddEleLeftSign(XMLSign As String, LeftTab As Integer) As String
        Try
            AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Left, mbolIsCrLf, LeftTab)
            If AddEleLeftSign <> "OK" Then Err.Raise(-1, , AddEleLeftSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftSign", ex)
        End Try
    End Function

    ''' <summary>增加一个XML右标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function AddEleRightSign(XMLSign As String) As String
        Try
            AddEleRightSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Right, mbolIsCrLf)
            If AddEleRightSign <> "OK" Then Err.Raise(-1, , AddEleRightSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleRightSign", ex)
        End Try
    End Function

    ''' <summary>增加一个XML右标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    Public Overloads Function AddEleRightSign(XMLSign As String, LeftTab As Integer) As String
        Try
            AddEleRightSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Right, mbolIsCrLf, LeftTab)
            If AddEleRightSign <> "OK" Then Err.Raise(-1, , AddEleRightSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleRightSign", ex)
        End Try
    End Function

    Private Function mXMLAddStr(ByRef sbAny As System.Text.StringBuilder, ByRef XMLSign As String, ByRef XMLValue As String,
                         Optional ByVal AddWhere As xpXMLAddWhere = xpXMLAddWhere.Both,
                         Optional ByVal IsCrlf As Boolean = False,
                         Optional ByVal LeftTab As Integer = 0,
                         Optional ByVal IsCData As Boolean = False) As String
        Try
            Select Case AddWhere
                Case xpXMLAddWhere.Left, xpXMLAddWhere.Both
10:                 If LeftTab > 0 Then
                        For i = 1 To LeftTab
                            sbAny.Append(vbTab)
                        Next
                    End If
            End Select
            Select Case AddWhere
                Case xpXMLAddWhere.Left
20:                 sbAny.Append("<")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                    If IsCData = True Then
                        sbAny.Append("<![CDATA[")
                    End If
                Case xpXMLAddWhere.Right
                    If IsCData = True Then
                        sbAny.Append("]]>")
                    End If
30:                 sbAny.Append("</")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                Case xpXMLAddWhere.Both
40:                 sbAny.Append("<")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                    If IsCData = True Then
                        sbAny.Append("<![CDATA[")
                    End If
                    sbAny.Append(XMLValue)
                    If IsCData = True Then
                        sbAny.Append("]]>")
                    End If
                    sbAny.Append("</")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                Case xpXMLAddWhere.ValueOnly
                    sbAny.Append(XMLValue)
            End Select
50:         If IsCrlf = True Then sbAny.Append(vbCrLf)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mXMLAddStr", ex)
        End Try
    End Function

    Public Function Clear() As String
        '初始化内部XML
        Try
            msbMain = Nothing
            msbMain = New System.Text.StringBuilder("")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function


    Public ReadOnly Property MainXmlStr As String
        Get
            MainXmlStr = msbMain.ToString
        End Get
    End Property


End Class
