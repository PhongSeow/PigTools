'**********************************
'* Name: PigWebReq
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Http Web Request operation
'* Home Url: https://en.seowphong.com
'* Version: 1.1
'* Create Time: 5/2/2021
'*1.0.2  25/2/2021   Add Me.ClearErr()
'*1.0.3  9/3/2021   Modify GetText,GetTextAuth,PostText,PostTextAuth
'*1.1  27/10/2021   Add DownloadFile, modify MainNew
'**********************************
Imports System.Net
Imports System.IO
Imports System.Text

''' <summary>
''' WEB request processing class|WEB请求处理类
''' </summary>
Public Class PigWebReq
    Inherits PigBaseMini
    Const CLS_VERSION As String = "1.1.8"
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

    Public Function PostText(Para As String) As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        PostText = ""
        Try
10:         mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/x-www-form-urlencoded"
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
            Me.SetSubErrInf("PostText", strStepName, ex)
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
            '//流对象使用完后自动关闭
            'Using (Stream stream = hwr.GetResponse().GetResponseStream())
            '{
            '//文件流，流信息读到文件流中，读完关闭
            'Using (FileStream fs = File.Create(@"C:\Users\Evan\Desktop\test\755.jpg"))
            '{
            '//建立字节组，并设置它的大小是多少字节
            'Byte[] bytes = New Byte[102400];
            'Int n = 1;
            'While (n > 0)
            '{
            '//一次从流中读多少字节，并把值赋给Ｎ，当读完后，Ｎ为０,并退出循环
            'n = Stream.Read(bytes, 0, 10240);
            'fs.Write(bytes, 0, n);　//将指定字节的流信息写入文件流中
            Me.UseTimeItem.ToEnd()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
