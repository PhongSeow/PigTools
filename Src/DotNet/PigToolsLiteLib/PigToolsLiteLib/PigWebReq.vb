'**********************************
'* Name: PigWebReq
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Http Web Request operation
'* Home Url: https://en.seowphong.com
'* Version: 1.3
'* Create Time: 5/2/2021
'*1.0.2  25/2/2021   Add Me.ClearErr()
'*1.0.3  9/3/2021  Modify GetText,GetTextAuth,PostText,PostTextAuth
'*1.1  27/10/2021  Add DownloadFile, modify MainNew
'*1.2  20/3/2024   Add GetHeadValue,AddHeader,GetAllHeads
'*1.3  25/3/2024   Add PostRaw
'**********************************
Imports System.Net
Imports System.IO
Imports System.Text

''' <summary>
''' WEB request processing class|WEB请求处理类
''' </summary>
Public Class PigWebReq
    Inherits PigBaseMini
    Const CLS_VERSION As String = "1" & "." & "3" & "." & "6"
    Private mstrUrl As String
    Private mstrPara As String
    Private muriMain As System.Uri
    Private mhwrMain As HttpWebRequest
    Private msrRes As StreamReader
    Private mstrResString As String
    Private mstrUserAgent As String
    Public UseTimeItem As New UseTime

    Public Sub New(Url As String, Para As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, Para, UserAgent)
    End Sub

    Public Sub New(Url As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, "", "")
    End Sub

    Public Sub New(Url As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, "", UserAgent)
    End Sub

    Public ReadOnly Property ResString As String
        Get
            ResString = mstrResString
        End Get
    End Property

    Public ReadOnly Property UserAgent As String
        Get
            UserAgent = mstrUserAgent
        End Get
    End Property

    Private Sub MainNew(Url As String, Para As String, UserAgent As String)
        Dim strStepName As String = ""
        Try
            If Len(Para) = 0 Then
                strStepName = "Set Uri" & Url & ":"
10:             muriMain = New System.Uri(Url)
            Else
                strStepName = "Set Uri" & Url & "?" & Para & ":"
20:             muriMain = New System.Uri(Url & "?" & Para)
            End If
            If Len(UserAgent) > 0 Then
                mstrUserAgent = UserAgent
            Else
                mstrUserAgent = "PigWebReq"
            End If
            strStepName = "HttpWebRequest.Create:"
30:         mhwrMain = System.Net.HttpWebRequest.Create(muriMain)
            mhwrMain.UserAgent = mstrUserAgent
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("MainNew", strStepName, ex)
        End Try
    End Sub

    Public Function GetText() As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        GetText = ""
        Try
10:         mhwrMain.Method = "GET"
            strStepName = "New StreamReader"
20:         msrRes = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            strStepName = "ReadToEnd"
30:         mstrResString = msrRes.ReadToEnd
40:         msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("GetText", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function GetTextAuth(AccessToken As String) As String
        'AccessToken：在接入端中设置
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        GetTextAuth = ""
        Try
            'mhwrMain.UseDefaultCredentials = False
            With mhwrMain
                .Headers("Authorization") = "Bearer " & AccessToken
                '                .Headers("value") = "Bearer " & AccessToken
                .Method = "GET"
                .PreAuthenticate = True
            End With
            strStepName = "New StreamReader"
20:         msrRes = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            strStepName = "ReadToEnd"
30:         mstrResString = msrRes.ReadToEnd
40:         msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("GetTextAuth", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function AddHeader(HeadKey As String, HeadValue As String) As String
        Try
            mhwrMain.Headers(HeadKey) = HeadValue
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

    Public Function PostTextAuth(Para As String, AccessToken As String) As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        PostTextAuth = ""
        Try
10:         mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/x-www-form-urlencoded"
            mhwrMain.Headers("AUTHORIZATION") = "Bearer " & AccessToken
            mhwrMain.PreAuthenticate = True
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
20:         mhwrMain.ContentLength = bys.Length
            strStepName = "mhwrMain.GetRequestStream():"
            Dim newStream As Stream = mhwrMain.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
30:         newStream.Close()
            strStepName = "mhwrMain.GetResponse().GetResponseStream():"
            Dim sr As StreamReader = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
40:         mstrResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("PostTextAuth", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function GetHeadValue(HeadName As String) As String
        Try
            Return mhwrMain.GetResponse().Headers(HeadName)
        Catch ex As Exception
            Me.SetSubErrInf("GetHeadValue", ex)
            Return ""
        End Try
    End Function

    Public Function GetAllHeads() As String
        Try
            Return mhwrMain.GetResponse().Headers.ToString
        Catch ex As Exception
            Me.SetSubErrInf("GetAllHeads", ex)
            Return ""
        End Try
    End Function

    Public Function PostRaw(JSon As String) As String
        Dim LOG As New PigStepLog("PostRaw")
        Me.UseTimeItem.GoBegin()
        Try
            LOG.StepName = "Init"
            mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/json"
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(JSon)
            mhwrMain.ContentLength = bys.Length
            LOG.StepName = "GetRequestStream():"
            Dim newStream As Stream = mhwrMain.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
            newStream.Close()
            LOG.StepName = "New StreamReader"
            Dim sr As StreamReader = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            LOG.StepName = "ReadToEnd"
            mstrResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    ''' <summary>
    ''' 以 x-www-form-urlencoded 格式发送POST请求|Send POST request in x-www-form-urlencoded format
    ''' </summary>
    ''' <param name="Para">格式：p1=v1&p2=v2...|Format:p1=v1&p2=v2...</param>
    ''' <returns></returns>
    Public Function PostText(Para As String) As String
        Dim LOG As New PigStepLog("PostText")
        Me.UseTimeItem.GoBegin()
        Try
            LOG.StepName = "Init"
            mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/x-www-form-urlencoded"
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
            mhwrMain.ContentLength = bys.Length
            LOG.StepName = "GetRequestStream"
            Dim newStream As Stream = mhwrMain.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
            newStream.Close()
            LOG.StepName = "New StreamReader"
            Dim sr As StreamReader = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            mstrResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function DownloadFile(FilePath As String, Optional MaxSize As Integer = 1073741824) As String
        Dim LOG As New PigStepLog("DownloadFile")
        Try
            LOG.StepName = "New StreamReader"
            mhwrMain.AddRange(0, MaxSize)
            LOG.StepName = "GetResponseStream"
            Dim oStream As Stream = mhwrMain.GetResponse().GetResponseStream
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
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
