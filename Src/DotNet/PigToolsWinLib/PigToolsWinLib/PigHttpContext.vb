﻿'**********************************
'* Name: PigHttpContext
'* Author: Seow Phong
'* License: Copyright (c) 2019-2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 给 HttpContext 加壳，实现一系统功能
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.9
'* Create Time: 31/8/2019
'* 1.0.2  2020-1-29   改用fGEBaseMini
'* 1.0.3  2020-1-31   BinaryRead，BinaryWrite
'* 1.0.5  2020-2-6   增加 WebSiteUrl
'* 1.0.6  2020-2-8    删除 json
'* 1.0.7  2020-3-11  增加 IsLocalHost
'* 1.0.8  4/3/2021  Add to PigToolsWinLib
'* 1.0.9  26/7/2021  Modify BinaryWrite,BinaryRead
'* 1.1    28/28/2021  Modify for not NETFRAMEWORK
'* 1.2    15/12/2021  Use LOG
'* 1.3    26/7/2022  Modify Imports
'* 1.5    19/8/2022  Use PigBaseLocal
'* 1.6    6/12/2022  Add GetHeaders
'* 1.7    8/12/2022  Add RequestItemLongValue,RequestItemBoolValue,RequestItemDateValue,RequestItemDecValue,IsEscapeHTML, modify related interfaces 
'* 1.8    1/6/2023  Use PigBaseLocal, Imports PigToolsLiteLib
'* 1.9  28/7/2024   Modify PigStepLog to StruStepLog
'************************************
#If NETFRAMEWORK Then
Imports System.Web
#Else
Imports Microsoft.AspNetCore.Http
#End If
Imports System.Collections.Specialized
Imports PigToolsLiteLib

''' <summary>
''' HTTP context processing class|HTTP上下文处理类
''' </summary>
Public Class PigHttpContext
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1." & "9" & "." & "2"

    Public Enum enmWhatHtmlEle '什么HTML元素
        Table = 1 '表格
        TabHead = 2 '表格的表头
        TabRow = 3 '表格的行
        TabCell = 4 '表格的表格单元
    End Enum

    Public HcMain As HttpContext
    Private mstrClientIp As String   '客户端IP，自动识别


    Private mstrRefererUrl As String '引用本页面的上一个URL
    Private mstrPageFullUrl As String '本页面的完整URL
    Private mstrUserAgent As String '客户端用户信息


    Private moPigFunc As New PigFunc

    Public Sub New(ByRef HcMain As HttpContext)
        'InitGEBase：可以由一个全局的 GEBase 来初始化参数
        MyBase.New(CLS_VERSION)
        Me.HcMain = HcMain
    End Sub

    Property mIsEscapeHTML As Boolean = True
    ''' <summary>
    ''' Whether to escape HTML|是否转义HTML
    ''' </summary>
    ''' <returns></returns>
    Public Property IsEscapeHTML As Boolean
        Get
            Return mIsEscapeHTML
        End Get
        Set(value As Boolean)
            mIsEscapeHTML = value
        End Set
    End Property

    Public ReadOnly Property IsPost As Boolean
        Get
            If UCase(Me.ServerVariables("REQUEST_METHOD")) = "POST" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsMobile As Boolean
        Get
            If InStr(Me.UserAgent, "Mobile") > 0 Then
                IsMobile = True
            Else
                IsMobile = False
            End If
        End Get
    End Property

    Public ReadOnly Property IsWeiXin As Boolean
        Get
            If InStr(Me.UserAgent, "MicroMessenger") > 0 Then
                IsWeiXin = True
            Else
                IsWeiXin = False
            End If
        End Get
    End Property

    Public ReadOnly Property IsAndroid As Boolean
        Get
            If InStr(Me.UserAgent, "Android") > 0 Then
                IsAndroid = True
            Else
                IsAndroid = False
            End If
        End Get
    End Property

    Public ReadOnly Property IsiPhone As Boolean
        Get
            If InStr(Me.UserAgent, "iPhone") > 0 Then
                IsiPhone = True
            Else
                IsiPhone = False
            End If
        End Get
    End Property

    Public ReadOnly Property IsMingxing As Boolean
        Get
            If InStr(Me.UserAgent, "MinxingMessenger") > 0 Then
                IsMingxing = True
            Else
                IsMingxing = False
            End If
        End Get
    End Property

    ''' <summary>
    ''' 客户端用户信息
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property UserAgent() As String
        Get
            Try
                If mstrUserAgent = "" Then mstrUserAgent = Me.ServerVariables("HTTP_USER_AGENT")
                UserAgent = mstrUserAgent
            Catch ex As Exception
                UserAgent = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property WebSiteUrl() As String
        Get
            Try
                Return HcMain.Request.Url.Authority
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property



    Public ReadOnly Property PageFullUrl() As String
        '本页面的完整URL
        Get
            Try
                Return HcMain.Request.Url.AbsoluteUri
            Catch ex As Exception
                Return ""
            End Try
            'Try
            '    If mstrPageFullUrl = "" Then
            '        If UCase(Me.ServerVariables("HTTPS")) = "OFF" Then
            '            mstrPageFullUrl = "http://"
            '        Else
            '            mstrPageFullUrl = "https://"
            '        End If
            '        mstrPageFullUrl &= Me.ServerVariables("HTTP_HOST")
            '        mstrPageFullUrl &= Me.ServerVariables("URL")
            '    End If
            '    PageFullUrl = mstrPageFullUrl
            'Catch ex As Exception
            '    PageFullUrl = ""
            'End Try
        End Get
    End Property

    Public ReadOnly Property ReqGetQueryString() As String
        '本页面的完整URL
        Get
            Try
                ReqGetQueryString = HcMain.Request.QueryString.ToString
            Catch ex As Exception
                ReqGetQueryString = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ReqPostString() As String
        '本页面的完整URL
        Get
            Try
                Dim i As Integer
                ReqPostString = ""
                For i = 0 To HcMain.Request.Form.Count
                    If ReqPostString = "" Then ReqPostString &= "&"
                    ReqPostString &= HcMain.Request.Form.Keys(0).ToString & "="
                    ReqPostString &= HcMain.Request.Form.Item(0).ToString
                Next
            Catch ex As Exception
                ReqPostString = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property RefererUrl() As String
        '引用URL，发起访问的上一个URL
        Get
            Try
                If mstrRefererUrl = "" Then mstrRefererUrl = Me.ServerVariables("HTTP_REFERER")
                RefererUrl = mstrRefererUrl
            Catch ex As Exception
                RefererUrl = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property IsLocalHost() As Boolean
        Get
            Try
                Dim strClietIp As String = ""
                If strClietIp = "" Then strClietIp = Me.ServerVariables("REMOTE_HOST")
                If strClietIp = "" Then strClietIp = Me.ServerVariables("REMOTE_ADDR")
                Select Case strClietIp
                    Case "::1", "127.0.0.1"
                        Return True
                    Case Else
                        Return False
                End Select
            Catch ex As Exception
                Return False
            End Try
        End Get
    End Property

    Public ReadOnly Property ClientIp() As String
        'HTTP服务器变量

        Get
            Try
                If mstrClientIp = "" Then
                    mstrClientIp = Me.ServerVariables("HTTP_X_FORWARDED_FOR")
                    If mstrClientIp = "" Then mstrClientIp = Me.ServerVariables("REMOTE_HOST")
                    If mstrClientIp = "" Then mstrClientIp = Me.ServerVariables("REMOTE_ADDR")
                    If mstrClientIp = "::1" Then mstrClientIp = "127.0.0.1"
                End If
                ClientIp = mstrClientIp
            Catch ex As Exception
                ClientIp = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property HeaderValue(Key As String) As String
        'HTTP服务器变量
        Get
            Try
                HeaderValue = HcMain.Request.Headers.Get(Key)
                If Me.IsEscapeHTML = True Then HeaderValue = Me.moPigFunc.ConvertHtmlStr(HeaderValue, PigFunc.EmnHowToConvHtml.DisableHTMLOnlySymbol)
            Catch ex As Exception
                HeaderValue = ""
            End Try
        End Get
    End Property

    Public ReadOnly Property ServerVariables(VarName As String) As String
        'HTTP服务器变量
        Get
            Try
                ServerVariables = HcMain.Request.ServerVariables(VarName).ToString
            Catch ex As Exception
                ServerVariables = ""
            End Try
        End Get
    End Property

    Public Property Session(SessionName As String) As String
        Get
            Try
                Session = HcMain.Session(SessionName).ToString
            Catch ex As Exception
                Session = ""
                Me.SetSubErrInf("Session_Get", ex)
            End Try
        End Get
        Set(value As String)
            Try
                HcMain.Session(SessionName) = value
            Catch ex As Exception
                Me.SetSubErrInf("Session_Set", ex)
            End Try
        End Set
    End Property

    Public ReadOnly Property RequestItemDecValue(ItemName As String) As Decimal
        Get
            Dim strValue As String = Me.RequestItem(ItemName)
            Return Me.moPigFunc.GEDec(strValue)
        End Get
    End Property

    Public ReadOnly Property RequestItemDateValue(ItemName As String) As Date
        Get
            Dim strValue As String = Me.RequestItem(ItemName)
            Return Me.moPigFunc.GECDate(strValue)
        End Get
    End Property

    Public ReadOnly Property RequestItemBoolValue(ItemName As String) As Boolean
        Get
            Dim strValue As String = Me.RequestItem(ItemName)
            Return Me.moPigFunc.GECBool(strValue)
        End Get
    End Property

    Public ReadOnly Property RequestItemLongValue(ItemName As String) As Long
        Get
            Dim strValue As String = Me.RequestItem(ItemName)
            Return Me.moPigFunc.GECLng(strValue)
        End Get
    End Property

    Public ReadOnly Property RequestItem(ItemName As String) As String
        Get
            Try
                RequestItem = HcMain.Request.Item(ItemName).ToString
                If Me.IsEscapeHTML = True Then RequestItem = Me.moPigFunc.ConvertHtmlStr(RequestItem, PigFunc.EmnHowToConvHtml.DisableHTML)
            Catch ex As Exception
                RequestItem = ""
            End Try
        End Get
    End Property

    Public Overloads Function WriteFile(FilePath As String) As String
        WriteFile = Me.WriteFileMain(FilePath, "", True)
    End Function

    Public Overloads Function WriteFile(FilePath As String, IsWrite2Memory As Boolean) As String
        WriteFile = Me.WriteFileMain(FilePath, "", IsWrite2Memory)
    End Function

    Public Overloads Function WriteFile(FilePath As String, ContentType As String) As String
        WriteFile = Me.WriteFileMain(FilePath, ContentType, True)
    End Function

    Public Overloads Function WriteFile(FilePath As String, ContentType As String, IsWrite2Memory As Boolean) As String
        WriteFile = Me.WriteFileMain(FilePath, ContentType, IsWrite2Memory)
    End Function



    ''' <summary>
    ''' 向HTTP客户端写文件
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="ContentType"></param>
    ''' <param name="IsWrite2Memory"></param>
    ''' <returns></returns>
    Private Function WriteFileMain(FilePath As String, Optional ContentType As String = "", Optional IsWrite2Memory As Boolean = True) As String
        Dim LOG As New StruStepLog : LOG.SubName = "WriteFileMain"
        Try
            If ContentType = "" Then
                Dim strExtName As String = LCase(moPigFunc.GetFilePart(FilePath, PigFunc.EnmFilePart.ExtName))
                WriteFileMain = "Automatically identify the corresponding contenttype"
                Select Case strExtName
                    Case "css"
                        ContentType = "text/css"
                    Case "doc"
                        ContentType = "application/msword"
                    Case "docx"
                        ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    Case "gif"
                        ContentType = "image/gif"
                    Case "html", "htm"
                        ContentType = "text/html"
                    Case "jpg", "jpeg"
                        ContentType = "image/jpeg"
                    Case "ppt"
                        ContentType = "application/vnd.ms-powerpoint"
                    Case "ppsx"
                        ContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
                    Case "png"
                        ContentType = "image/png"
                    Case "rar"
                        ContentType = "application/octet-stream"
                    Case "xlsx"
                        ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    Case "xls", "xlm"
                        ContentType = "application/vnd.ms-excel"
                    Case "xml"
                        ContentType = "text/xml"
                    Case "zip"
                        ContentType = "application/x-zip-compressed"
                    Case Else
                        LOG.Ret = "Unrecognized file extension:" & strExtName
                        Throw New Exception(LOG.Ret)
                End Select
            End If
            Dim oFile As New System.IO.FileInfo(FilePath)
            With HcMain
                .Response.WriteFile(FilePath, IsWrite2Memory)
                .Response.AddHeader("Content-Disposition", "attachment; filename=" + .Server.UrlEncode(moPigFunc.GetFilePart(FilePath, PigFunc.EnmFilePart.FileTitle)))
                .Response.AddHeader("Content-Length", oFile.Length.ToString())
                .Response.ContentType = ContentType
                .Response.WriteFile(FilePath)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function BinaryWrite(ByRef ResBytes As Byte()) As String
        Try
            If Me.IsPost = False Then Err.Raise(-1, , "Only post mode can be used")
            With Me.HcMain.Response
                .BinaryWrite(ResBytes)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("BinaryWrite", ex)
        End Try
    End Function

    Public Function BinaryRead(ByRef ReqBytes As Byte()) As String
        Try
            If Me.IsPost = False Then Err.Raise(-1, , "Only post mode can be used")
            With Me.HcMain.Request
                ReqBytes = .BinaryRead(.TotalBytes)
            End With
            Return "OK"
        Catch ex As Exception
            ReqBytes = Nothing
            Return Me.GetSubErrInf("BinaryRead", ex)
        End Try
    End Function


    ''' <summary>
    ''' 向HTTP客户端写页面
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    Private Function WritePageMain(FilePath As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "WritePageMain"
        Try
            With Me.HcMain
                .Response.Write(My.Computer.FileSystem.ReadAllText(FilePath))    '默认为UTF8
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ResponsJsAlert(AlertDesc As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsJsAlert"
        Try
            Dim strHtml As String = "<script>window.alert(""" & AlertDesc & """);</script>"
            Me.HcMain.Response.Write(strHtml)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Overloads Function ResponsTableBegin(Optional Style As String = "width:100%;") As String
        ResponsTableBegin = ResponsTableBeginMain(True, , , Style)
    End Function


    Public Function ResponseWrite(WriteStr As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "ResponseWrite"
        Try
            LOG.StepName = "HcMain.Response.Write"
            With Me.HcMain.Response
                .Write(WriteStr)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ResponsHtmlEnd(WhatTabEle As enmWhatHtmlEle) As String
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsHtmlEnd"
        Try
            With Me.HcMain.Response
                Select Case WhatTabEle
                    Case enmWhatHtmlEle.Table
                        .Write("</table>")
                    Case enmWhatHtmlEle.TabHead
                        .Write("</th>")
                    Case enmWhatHtmlEle.TabRow
                        .Write("</tr>")
                    Case enmWhatHtmlEle.TabCell
                        .Write("</td>")
                End Select
            End With
            ResponsHtmlEnd = "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Overloads Function ResponsTabRowBegin() As String
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsTabRowBegin"
        Try
            With Me.HcMain.Response
                .Write("<tr>")
            End With
            ResponsTabRowBegin = "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Overloads Function ResponsTableBegin() As String
        ResponsTableBegin = Me.ResponsTableBeginMain(True)
    End Function


    Public Overloads Function ResponsTabCellBegin() As String
        ResponsTabCellBegin = Me.ResponsTabCellBeginMain(True)
    End Function

    Public Overloads Function ResponsTabCellBegin(Colspan As Integer) As String
        ResponsTabCellBegin = Me.ResponsTabCellBeginMain(True)
    End Function

    Public Overloads Function ResponsTabCellBegin(IsRowspan As Boolean, Span As Integer) As String
        If IsRowspan = True Then
            ResponsTabCellBegin = Me.ResponsTabCellBeginMain(True, , Span)
        Else
            ResponsTabCellBegin = Me.ResponsTabCellBeginMain(True, Span)
        End If
    End Function


    Public Overloads Function ResponsTabHeadBegin() As String
        ResponsTabHeadBegin = Me.ResponsTabHeadBeginMain(True)
    End Function

    Public Overloads Function ResponsTabHeadBegin(Colspan As Integer) As String
        ResponsTabHeadBegin = Me.ResponsTabHeadBeginMain(True, Colspan)
    End Function

    Public Overloads Function ResponsTabHeadBegin(IsRowspan As Boolean, Span As Integer) As String
        If IsRowspan = True Then
            ResponsTabHeadBegin = Me.ResponsTabHeadBeginMain(True, , Span)
        Else
            ResponsTabHeadBegin = Me.ResponsTabHeadBeginMain(True, Span)
        End If
    End Function

    Private Function ResponsTabCellBeginMain(IsCrLf As Boolean, Optional Colspan As Integer = 1, Optional Rowspan As Integer = 1) As String
        '向HTTP客户端写表格单元开始
        'Colspan：规定表头单元格可横跨的列数
        'Rowspan：规定表头单元格可横跨的行数。
        'Scope：规定表头单元格是否是行、列、行组或列组的头部。
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsTabCellBeginMain"
        Try
            With Me.HcMain.Response
                .Write("<td")
                If Colspan > 1 Then .Write(" colspan=""" & Colspan.ToString & """")
                If Colspan > 1 Then .Write(" rowspan=""" & Rowspan.ToString & """")
                .Write(">")
                If IsCrLf = True Then .Write(vbCrLf)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Private Function ResponsTabHeadBeginMain(IsCrLf As Boolean, Optional Colspan As Integer = 1, Optional Rowspan As Integer = 1, Optional Scope As String = "") As String
        '向HTTP客户端写表格头开始
        'Colspan：规定表头单元格可横跨的列数
        'Rowspan：规定表头单元格可横跨的行数。
        'Scope：规定表头单元格是否是行、列、行组或列组的头部。
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsTabHeadBeginMain"
        Try
            With Me.HcMain.Response
                .Write("<th")
                If Colspan > 1 Then .Write(" colspan=""" & Colspan.ToString & """")
                If Colspan > 1 Then .Write(" rowspan=""" & Rowspan.ToString & """")
                Select Case LCase(Scope)
                    Case "col", "colgroup", "row", "rowgroup"
                        .Write(" scope=""" & Scope & """")
                    Case ""
                    Case Else
                        Err.Raise(-1, , "Scope=" & Scope & " invalid")
                End Select
                .Write(">")
                If IsCrLf = True Then .Write(vbCrLf)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' 向HTTP客户端写表格开始
    ''' </summary>
    ''' <param name="IsCrLf"></param>
    ''' <param name="Border"></param>
    ''' <param name="ClassName"></param>
    ''' <param name="Style"></param>
    ''' <returns></returns>
    Private Function ResponsTableBeginMain(IsCrLf As Boolean, Optional Border As Integer = 1, Optional ClassName As String = "", Optional Style As String = "width:100%;") As String
        Dim LOG As New StruStepLog : LOG.SubName = "ResponsTableBeginMain"
        Try
            With Me.HcMain.Response
                .Write("<table border=""" & Border.ToString & """")
                If ClassName <> "" Then .Write(" class=""" & ClassName & """")
                If Style <> "" Then .Write(" style=""" & Style & """")
                .Write(">")
                If IsCrLf = True Then .Write(vbCrLf)
            End With
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Overloads Function WritePage(FilePath As String) As String
        WritePage = Me.WritePageMain(FilePath)
    End Function

    Public Function GetHeaders(ByRef OutHeaders As String) As String
        Dim LOG As New StruStepLog : LOG.SubName = "GetHeaders"
        Try
            Dim strKey As String, strValue As String
            LOG.StepName = "NameValueCollection"
            Dim nvcMain As NameValueCollection = HcMain.Request.Headers
            Dim astrKey As String(), astrValue As String()
            LOG.StepName = "AllKeys"
            astrKey = nvcMain.AllKeys
            OutHeaders = "Headers.Count=" & nvcMain.Count & Me.OsCrLf
            LOG.StepName = "For Keys"
            For i = 0 To astrKey.GetUpperBound(0)
                strKey = astrKey(i)
                LOG.StepName = strKey & "GetValues"
                astrValue = nvcMain.GetValues(i)
                LOG.StepName = "For Values"
                For j = 0 To astrValue.GetUpperBound(0)
                    strValue = astrValue(j)
                    OutHeaders &= strKey & "=" & strValue & Me.OsCrLf
                Next
            Next
            Return "OK"
        Catch ex As Exception
            OutHeaders = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
