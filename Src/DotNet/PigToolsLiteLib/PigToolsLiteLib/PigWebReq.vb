'**********************************
'* Name: PigWebReq
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Http Web Request operation
'* Home Url: https://en.seowphong.com
'* Version: 1.5
'* Create Time: 5/2/2021
'*1.0.2  25/2/2021   Add Me.ClearErr()
'*1.0.3  9/3/2021  Modify GetText,GetTextAuth,PostText,PostTextAuth
'*1.1  27/10/2021  Add DownloadFile, modify MainNew
'*1.2  20/3/2024   Add GetHeadValue,AddHeader,GetAllHeads
'*1.3  25/3/2024   Add PostRaw
'*1.5  30/5/2024   Rewrite the internal code of a class
'**********************************
Imports System.Net
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.Logging

''' <summary>
''' WEB request processing class|WEB请求处理类
''' </summary>
Public Class PigWebReq
    Inherits PigBaseMini
    Const CLS_VERSION As String = "1" & "." & "5" & "." & "168"
    Private ReadOnly Property mUrl As String = ""
    Private ReadOnly Property mPara As String = ""
    Private Property mUri As System.Uri
    Private Property mHttpWebRequest As HttpWebRequest
    Private Property mHttpWebResponse As HttpWebResponse
    Private ReadOnly Property mUserAgent As String = ""
    Public Property UseTimeItem As New UseTime


    Public Sub New(Url As String, Para As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.mUrl = Url
        Me.mPara = Para
        Me.mUserAgent = UserAgent
        Me.InitHttpWebRequest()
    End Sub

    Public Sub New(Url As String)
        MyBase.New(CLS_VERSION)
        Me.mUrl = Url
        Me.mUserAgent = Me.MyClassName & "/" & Me.CLS_VERSION
        Me.InitHttpWebRequest()
    End Sub

    Public Sub New(Url As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.mUrl = Url
        Me.mUserAgent = UserAgent
        Me.InitHttpWebRequest()
    End Sub

    Private mResString As String
    Public Property ResString As String
        Get
            ResString = mResString
        End Get
        Friend Set(value As String)
            mResString = value
        End Set
    End Property


    ''' <summary>
    ''' Initialize HttpWebRequest object|初始化HttpWebRequest对象
    ''' </summary>
    ''' <returns></returns>
    Public Function InitHttpWebRequest() As String
        Dim LOG As New PigStepLog("InitHttpWebRequest")
        Try
            If Me.mPara = "" Then
10:             Me.mUri = New System.Uri(Me.mUrl)
            Else
20:             Me.mUri = New System.Uri(Me.mUrl & "?" & Me.mPara)
            End If
            LOG.StepName = "HttpWebRequest.Create:"
            Me.mHttpWebRequest = System.Net.HttpWebRequest.Create(mUri)
            Me.mHttpWebRequest.UserAgent = Me.mUserAgent
            Me.mHttpWebResponse = Nothing
            Me.ResString = ""
            Return "OK"
        Catch ex As Exception
            Me.mUri = Nothing
            Me.mHttpWebRequest = Nothing
            LOG.AddStepNameInf("Url=" & Me.mUrl)
            If Me.mPara <> "" Then LOG.AddStepNameInf("Para=" & Me.mPara)
            If Me.mUserAgent <> "" Then LOG.AddStepNameInf("UserAgent=" & Me.mUserAgent)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function




    ''' <summary>
    ''' 以 GET 方式发送请求|Send request in GET method
    ''' </summary>
    ''' <returns></returns>
    Public Function GetText() As String
        Dim LOG As New PigStepLog("GetText")
        Try
            Me.UseTimeItem.GoBegin()
            mHttpWebRequest.Method = "GET"
            LOG.StepName = "GetResponse"
            Me.mHttpWebResponse = mHttpWebRequest.GetResponse
            LOG.StepName = "GetResponseStream"
            Dim msrRes As New StreamReader(Me.mHttpWebResponse.GetResponseStream)
            LOG.StepName = "ReadToEnd"
            Me.ResString = msrRes.ReadToEnd
            msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.ResString = ""
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function GetTextAuth(AccessToken As String) As String
        Dim LOG As New PigStepLog("GetTextAuth")
        Try
            Me.UseTimeItem.GoBegin()
            LOG.StepName = "Set Properties"
            With mHttpWebRequest
                .Headers("Authorization") = "Bearer " & AccessToken
                .Method = "GET"
                .PreAuthenticate = True
            End With
            LOG.StepName = "GetResponse"
            Me.mHttpWebResponse = mHttpWebRequest.GetResponse
            LOG.StepName = "GetResponseStream"
            Dim msrRes As New StreamReader(mHttpWebRequest.GetResponse().GetResponseStream)
            LOG.StepName = "ReadToEnd"
            Me.ResString = msrRes.ReadToEnd
            msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.ResString = ""
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

    ''' <summary>
    ''' Add a header|添加一个头
    ''' </summary>
    ''' <param name="HeadKey"></param>
    ''' <param name="HeadValue"></param>
    ''' <returns></returns>
    Public Function AddHeader(HeadKey As String, HeadValue As String) As String
        Dim LOG As New PigStepLog("AddHeader")
        Try
            If mHttpWebRequest Is Nothing Then
                LOG.StepName = "InitHttpWebRequest"
                LOG.Ret = Me.InitHttpWebRequest()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            mHttpWebRequest.Headers(HeadKey) = HeadValue
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(HeadKey)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function PostTextAuth(Para As String, AccessToken As String) As String
        Dim LOG As New PigStepLog("PostTextAuth")
        Try
            Me.UseTimeItem.GoBegin()
            LOG.StepName = "Set Properties"
            mHttpWebRequest.Method = "POST"
            mHttpWebRequest.ContentType = "application/x-www-form-urlencoded"
            mHttpWebRequest.Headers("AUTHORIZATION") = "Bearer " & AccessToken
            mHttpWebRequest.PreAuthenticate = True
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
            mHttpWebRequest.ContentLength = bys.Length
            LOG.StepName = "GetRequestStream"
            Dim newStream As Stream = mHttpWebRequest.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
            newStream.Close()
            LOG.StepName = "GetResponse"
            Me.mHttpWebResponse = mHttpWebRequest.GetResponse
            LOG.StepName = "GetResponseStream"
            Dim sr As StreamReader = New StreamReader(mHttpWebRequest.GetResponse().GetResponseStream)
            LOG.StepName = "ReadToEnd"
            Me.ResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.ResString = ""
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

    ''' <summary>
    ''' Get the value of the specified header|获取指定头的值
    ''' </summary>
    ''' <param name="HeadName"></param>
    ''' <returns></returns>
    Public Function GetHeadValue(HeadName As String) As String
        Try
            If Me.mHttpWebResponse Is Nothing Then Throw New Exception("HTTP method not executed")
            Return Me.mHttpWebResponse.Headers(HeadName)
        Catch ex As Exception
            Me.SetSubErrInf("GetHeadValue", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get all headers|获取所有头
    ''' </summary>
    ''' <returns></returns>
    Public Function GetAllHeads() As String
        Try
            If Me.mHttpWebResponse Is Nothing Then Throw New Exception("HTTP method not executed")
            Return Me.mHttpWebResponse.Headers.ToString
        Catch ex As Exception
            Me.SetSubErrInf("GetAllHeads", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Send POST request in raw JSON format|以原始JSON格式发送POST请求
    ''' </summary>
    ''' <param name="JSon"></param>
    ''' <returns></returns>
    Public Function PostRaw(JSon As String) As String
        Dim LOG As New PigStepLog("PostRaw")
        Try
            Me.UseTimeItem.GoBegin()
            LOG.StepName = "Set Properties"
            mHttpWebRequest.Method = "POST"
            mHttpWebRequest.ContentType = "application/json"
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(JSon)
            mHttpWebRequest.ContentLength = bys.Length
            LOG.StepName = "GetRequestStream"
            Dim newStream As Stream = mHttpWebRequest.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
            newStream.Close()
            LOG.StepName = "GetResponse"
            Me.mHttpWebResponse = Me.mHttpWebRequest.GetResponse
            LOG.StepName = "GetResponseStream"
            Dim srMain As StreamReader = New StreamReader(Me.mHttpWebResponse.GetResponseStream)
            LOG.StepName = "ReadToEnd"
            Me.ResString = srMain.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.ResString = ""
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function


    ''' <summary>
    ''' Send request in POST method|以 POST 方式发送请求
    ''' </summary>
    ''' <param name="Para"></param>
    ''' <returns></returns>
    Public Function PostText(Para As String) As String
        Dim LOG As New PigStepLog("PostText")
        Try
            Me.UseTimeItem.GoBegin()
            LOG.StepName = "Set Properties"
            mHttpWebRequest.Method = "POST"
            mHttpWebRequest.ContentType = "application/x-www-form-urlencoded"
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
            mHttpWebRequest.ContentLength = bys.Length
            LOG.StepName = "GetRequestStream"
            Dim newStream As Stream = mHttpWebRequest.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
            newStream.Close()
            LOG.StepName = "GetResponse"
            Me.mHttpWebResponse = mHttpWebRequest.GetResponse
            LOG.StepName = "GetResponseStream"
            Dim srMain As StreamReader = New StreamReader(Me.mHttpWebResponse.GetResponseStream())
            LOG.StepName = "ReadToEnd"
            Me.ResString = srMain.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.ResString = ""
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

    ''' <summary>
    ''' Download file|下载文件
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="MaxSize"></param>
    ''' <returns></returns>
    Public Function DownloadFile(FilePath As String, Optional MaxSize As Integer = 1073741824) As String
        Dim LOG As New PigStepLog("DownloadFile")
        Try
            Me.UseTimeItem.GoBegin()
            If Me.mHttpWebRequest Is Nothing Then
                LOG.StepName = "InitHttpWebRequest"
                LOG.Ret = Me.InitHttpWebRequest()
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "AddRange"
            mHttpWebRequest.AddRange(0, MaxSize)
            LOG.StepName = "GetResponseStream"
            Dim oStream As Stream = mHttpWebRequest.GetResponse().GetResponseStream
            LOG.StepName = "New FileStream"
            Dim fsMain As New FileStream(FilePath, FileMode.OpenOrCreate)
            Dim abData(-1) As Byte
            Dim intCount As Integer, intPos As Integer = 0, intOnceSize As Integer = 1024000
            Do While True
                ReDim abData(intOnceSize - 1)
                intCount = oStream.Read(abData, 0, intOnceSize)
                If intCount <= 0 Then Exit Do
                If intCount < intOnceSize Then ReDim Preserve abData(intCount - 1)
                fsMain.Write(abData, 0, intCount)
                intPos += intCount
            Loop
            LOG.StepName = "Close"
            oStream.Close()
            fsMain.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            LOG.AddStepNameInf(FilePath)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

End Class
