'**********************************
'* Name: IISMain
'* Author: 萧鹏
'* Describe: IIS线程主程序
'* Version: 1.0.5
'* Create Time: 10/3/2021
'* 1.0.2  11/3/2021   初始化
'* 1.0.3  31/3/2021   修改 Login,ReLogin
'* 1.0.5  4/4/2021   修改 ReLogin
'**********************************
Imports System.Web
Imports System.Web.Services
Imports PigToolsLib
Imports ObjAdoDBLib
Imports PigWxWorkLib

Public Class IISMain
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.5"
    Private moPigFunc As New PigFunc
    Public Sub New()
        MyBase.New(CLS_VERSION)
        If goIISMain Is Nothing Then
            goIISMain = New fIISMain
        End If
        goIISMain.RefConf()
        If goIISMain.IsDebug = True Then
            Me.SetDebug(goIISMain.LogFilePath)
        End If
        Me.SetDebug("E:\IISWeb\Site8271\WxWork2LxChkLog\" & Me.AppTitle & ".log")
    End Sub

    ''' <summary>
    ''' 刷新登录凭证
    ''' </summary>
    Friend Sub RefAccessToken()
        Const SUB_NAME As String = "RefAccessToken"
        Dim strStepName As String = ""
        Try
            With goIISMain
                If .IsAccessTokenReady = False Then
                    strStepName = ".GetItem(AccessToken)"
                    With .PigKeyValueList
                        Dim oPigKeyValue As PigKeyValue = .GetItem("WxWorkAccessToken")
                        If oPigKeyValue Is Nothing Then
                            If goIISMain.IsDBConnReady = False Then
                                strStepName = "RefDBConn"
                                goIISMain.RefDBConn()
                                If goIISMain.LastErr <> "" Then Throw New Exception(goIISMain.LastErr)
                            End If
                            Dim strSQL As String = "EXEC dbo.gespKeyValueQry @OptCode=2,@KeyName='AccessToken'"
                            strStepName = "Execute"
                            Dim rs As Recordset, strKeyValue As String
                            rs = goIISMain.Connection.Execute(strSQL)
                            If goIISMain.Connection.LastErr <> "" Then Throw New Exception(goIISMain.Connection.LastErr)
                            If rs.EOF = True Then
                                strKeyValue = ""
                            Else
                                strKeyValue = rs.Fields.Item("KeyValue").Value
                            End If
                            Dim oWorkApp As WorkApp, bolIsReNew As Boolean = False, dteAccessTokenExpiresTime As DateTime, oPigJSon As fPigJSon, strTxRes As String = "", strJSonStr As String = ""
                            If strKeyValue = "" Then
                                bolIsReNew = True
                            Else
                                strStepName = "New fPigJSon(KeyValue)"
                                oPigJSon = New fPigJSon(strKeyValue)
                                If oPigJSon.LastErr = "" Then
                                    strStepName = "New WorkApp并使用缓存AccessToken"
                                    oWorkApp = New WorkApp(goIISMain.CorpId, goIISMain.CorpSecret, oPigJSon.GetStrValue("AccessToken"), oPigJSon.GetDateValue("ExpiresTime"))
                                    dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                    If oWorkApp.IsAccessTokenReady = False Then
                                        bolIsReNew = True
                                    Else
                                        goIISMain.AccessToken = oWorkApp.AccessToken
                                        goIISMain.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                    End If
                                Else
                                    bolIsReNew = True
                                End If
                            End If
                            If bolIsReNew = True Then
                                strStepName = "ReNew WorkApp"
                                oWorkApp = New WorkApp(goIISMain.CorpId, goIISMain.CorpSecret)
                                If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
                                strStepName = "ReNew RefAccessToken"
                                If Me.IsDebug = True Then Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, ""))
                                oWorkApp.RefAccessToken(True)
                                If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
                                dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                goIISMain.AccessToken = oWorkApp.AccessToken
                                goIISMain.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                strStepName = "Get MainJSonStr"
                                oPigJSon = New fPigJSon
                                With oPigJSon
                                    .AddEle("AccessToken", oWorkApp.AccessToken, True)
                                    .AddEle("ExpiresTime", oWorkApp.AccessTokenExpiresTime)
                                    .AddSymbol(fPigJSon.xpSymbolType.EleEndFlag)
                                    strJSonStr = .MainJSonStr
                                    If .LastErr <> "" Then Throw New Exception(.LastErr)
                                End With
                                strSQL = "EXEC dbo.gespKeyValueMgr @OptCode=1,@KeyName='AccessToken',@KeyValue=" & SQLStr(strJSonStr) & ",@ExpTime=" & SQLDate(oWorkApp.AccessTokenExpiresTime)
                                strStepName = "Execute"
                                rs = goIISMain.Connection.Execute(strSQL)
                                If goIISMain.Connection.LastErr <> "" Then
                                    Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, goIISMain.Connection.LastErr))
                                    If Me.IsDebug = True Then
                                        Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, strSQL))
                                    End If
                                ElseIf rs.EOF = False Then
                                    strTxRes = rs.Fields.Item("TxRes").Value
                                    If strTxRes <> "OK" Then Throw New Exception(strTxRes)
                                End If
                            End If
                            strStepName = "New PigKeyValue"
                            oPigKeyValue = New PigKeyValue("AccessToken", dteAccessTokenExpiresTime, strJSonStr)
                            strStepName = "AddNewItem"
                            .AddNewItem(oPigKeyValue)
                            If .LastErr <> "" Then Throw New Exception(.LastErr)
                        End If
                    End With
                End If
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(SUB_NAME, strStepName, ex)
            Me.mPrintErrLogInf(Me.LastErr)
        End Try
    End Sub

    Private Function mGetWorkApp() As WorkApp
        Dim strStepName As String = ""
        Try
            With goIISMain
                If .IsAccessTokenReady = False Then
                    Throw New Exception("AccessToken未就绪")
                End If
                strStepName = "New WorkApp"
                Dim oWorkApp As New WorkApp(.CorpId, .CorpSecret, .AccessToken, .AccessTokenExpiresTime)
                If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
                mGetWorkApp = oWorkApp
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mGetWorkApp", strStepName, ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 登录入口
    ''' </summary>
    ''' <param name="HcMain"></param>
    Public Sub Login(ByRef HcMain As HttpContext)
        Dim strStepName As String = "", bolIsNoLog As Boolean = False
        Try
            strStepName = "RefDBConn"
            goIISMain.RefDBConn()
            If goIISMain.LastErr <> "" Then Throw New Exception(goIISMain.LastErr)
            strStepName = "RefAccessToken"
            Me.RefAccessToken()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            strStepName = "mGetWorkApp"
            Dim oWorkApp As WorkApp = Me.mGetWorkApp
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            With oWorkApp
                strStepName = "RequestItem(State)"
                Dim oPigHttpContext As New PigHttpContext(HcMain)
                Dim strState As String = oPigHttpContext.RequestItem("State")
                If strState = "" Then strState = "WxWorkLx2LoginTest"
                oPigHttpContext = Nothing
                strStepName = "GetOauth2Url"
                Me.PrintDebugLog("RediretUri=" & goIISMain.JSonConf.GetStrValue("RediretUri"))
                Dim strUrl As String = .GetOauth2Url(goIISMain.JSonConf.GetStrValue("RediretUri"), strState)
                If Me.IsDebug = True Then Me.mPrintErrLogInf(strUrl)
                If .LastErr <> "" Then
                    Throw New Exception(.LastErr)
                Else
                    strStepName = "Redirect"
                    bolIsNoLog = True
                    Me.PrintDebugLog("Redirect=" & strUrl)
                    HcMain.Response.Redirect(strUrl)
                End If
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Login", strStepName, ex)
            If bolIsNoLog = False Then Me.mPrintErrLogInf(Me.LastErr)
        End Try
    End Sub

    ''' <summary>
    ''' 登录后跳转腾讯回调入口
    ''' </summary>
    ''' <param name="HcMain"></param>
    Public Sub ReLogin(ByRef HcMain As HttpContext)
        Const SUB_NAME As String = "ReLogin"
        Dim strStepName As String = "", bolIsNoLog As Boolean = False
        Try
            strStepName = "RefDBConn"
            goIISMain.RefDBConn()
            If goIISMain.LastErr <> "" Then Throw New Exception(goIISMain.LastErr)
            strStepName = "RefAccessToken"
            Me.RefAccessToken()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            strStepName = "mGetWorkApp"
            Dim oWorkApp As WorkApp = Me.mGetWorkApp
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            With oWorkApp
                strStepName = "GetWorkMemberFromOauth2Redirect"
                Dim oWorkMember As WorkMember = .GetWorkMemberFromOauth2Redirect(HcMain)
                If .LastErr <> "" Then Throw New Exception(.LastErr)
                Dim strSQL As String = ""
                If goIISMain.IsGD = True Then
                    strSQL = "EXEC dbo.gespUserQry @OptCode=1,@MobileNo=" & SQLStr(oWorkMember.UserId)
                Else
                    strSQL = "EXEC dbo.gespUserQry @OptCode=2,@ccb_empid=" & SQLStr(oWorkMember.UserId)
                End If
                strStepName = "Execute(" & strSQL & ")"
                Dim rs As Recordset
                rs = goIISMain.Connection.Execute(strSQL)
                If goIISMain.Connection.LastErr <> "" Then Throw New Exception(goIISMain.Connection.LastErr)
                If rs.EOF = True Then
                    Throw New Exception(oWorkMember.UserId & "无法获取UAAS")
                Else
                    Dim oPigHttpContext As New PigHttpContext(HcMain)
                    Dim strState As String = oPigHttpContext.RequestItem("State")
                    oPigHttpContext = Nothing
                    Dim strReUrl As String = ""
                    Select Case strState
                        Case "WxWorkLx2LoginTest"
                            strReUrl = "http://f5.jhvpn.cn:8813/LxWork/WxWorkLx2LoginTest/Login.ashx"
                        Case Else
                            Throw New Exception("无效AppID")
                    End Select
                    strStepName = "Redirect"
                    bolIsNoLog = True
                    Dim strUaas As String = rs.Fields.Item("Uaas").StrValue
                    Dim strKeyName As String = strUaas & "_UAT"
                    With goIISMain.PigKeyValueList
                        Dim oPigKeyValue As PigKeyValue = .GetItem(strKeyName)
                        If oPigKeyValue Is Nothing Then
                            If goIISMain.IsDBConnReady = False Then
                                strStepName = "RefDBConn"
                                goIISMain.RefDBConn()
                                If goIISMain.LastErr <> "" Then Throw New Exception(goIISMain.LastErr)
                            End If
                            strSQL = "EXEC dbo.gespKeyValueQry @OptCode=2,@KeyName=" & SQLStr(strKeyName)
                            strStepName = "Execute(" & strSQL & ")"
                            Dim strKeyValue As String
                            rs = goIISMain.Connection.Execute(strSQL)
                            If goIISMain.Connection.LastErr <> "" Then Throw New Exception(goIISMain.Connection.LastErr)
                            If rs.EOF = True Then
                                strKeyValue = ""
                            Else
                                strKeyValue = rs.Fields.Item("KeyValue").Value
                            End If
                            Dim bolIsReNew As Boolean = False, dteAccessTokenExpiresTime As DateTime, oPigJSon As fPigJSon, strTxRes As String = "", strJSonStr As String = ""
                            If strKeyValue = "" Then
                                bolIsReNew = True
                            Else
                                strStepName = "New fPigJSon(KeyValue)"
                                oPigJSon = New fPigJSon(strKeyValue)
                                If oPigJSon.LastErr = "" Then
                                    strSQL = "EXEC dbo.gespUserQry @OptCode=3,@Uaas=" & SQLStr(strUaas) & ",@ByAppID=" & SQLStr(strState)
                                    strStepName = "Execute(" & strSQL & ")"
                                    rs = goIISMain.Connection.Execute(strSQL)
                                    If goIISMain.Connection.LastErr <> "" Then Throw New Exception(goIISMain.Connection.LastErr)
                                    If rs.EOF = True Then Throw New Exception(strUaas & "无法获取详细信息")
                                    If rs.Fields.Item("TxRes").StrValue <> "OK" Then Throw New Exception(rs.Fields.Item("TxRes").StrValue)
                                    Dim rs2 As Recordset = rs.NextRecordset

                                    dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                    If oWorkApp.IsAccessTokenReady = False Then
                                        bolIsReNew = True
                                    Else
                                        goIISMain.AccessToken = oWorkApp.AccessToken
                                        goIISMain.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                    End If
                                Else
                                    bolIsReNew = True
                                End If
                            End If
                            If bolIsReNew = True Then
                                strStepName = "ReNew WorkApp"
                                oWorkApp = New WorkApp(goIISMain.CorpId, goIISMain.CorpSecret)
                                If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
                                strStepName = "ReNew RefAccessToken"
                                If Me.IsDebug = True Then Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, ""))
                                oWorkApp.RefAccessToken(True)
                                If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
                                dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                goIISMain.AccessToken = oWorkApp.AccessToken
                                goIISMain.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
                                strStepName = "Get MainJSonStr"
                                oPigJSon = New fPigJSon
                                With oPigJSon
                                    .AddEle("AccessToken", oWorkApp.AccessToken, True)
                                    .AddEle("ExpiresTime", oWorkApp.AccessTokenExpiresTime)
                                    .AddSymbol(fPigJSon.xpSymbolType.EleEndFlag)
                                    strJSonStr = .MainJSonStr
                                    If .LastErr <> "" Then Throw New Exception(.LastErr)
                                End With
                                strSQL = "EXEC dbo.gespKeyValueMgr @OptCode=1,@KeyName='AccessToken',@KeyValue=" & SQLStr(strJSonStr) & ",@ExpTime=" & SQLDate(oWorkApp.AccessTokenExpiresTime)
                                strStepName = "Execute"
                                rs = goIISMain.Connection.Execute(strSQL)
                                If goIISMain.Connection.LastErr <> "" Then
                                    Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, goIISMain.Connection.LastErr))
                                    If Me.IsDebug = True Then
                                        Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, strSQL))
                                    End If
                                ElseIf rs.EOF = False Then
                                    strTxRes = rs.Fields.Item("TxRes").Value
                                    If strTxRes <> "OK" Then Throw New Exception(strTxRes)
                                End If
                            End If
                            strStepName = "New PigKeyValue"
                            oPigKeyValue = New PigKeyValue("AccessToken", dteAccessTokenExpiresTime, strJSonStr)
                            strStepName = "AddNewItem"
                            .AddNewItem(oPigKeyValue)
                            If .LastErr <> "" Then Throw New Exception(.LastErr)
                        End If
                    End With
                    Dim strPara As String = "timestamp=1617258363&nonce=990996"
                    strPara &= "&login_name=" & rs.Fields.Item("Uaas").StrValue
                            strPara &= "&mx_sso_token=lAkAADqpI9Qg-36diC9r2eJWTCYuepKOigexHBejxUwUxaZO"
                            strPara &= "&signed=FTeqymu0tDqw40hDW1OgC"
                            strPara &= "&IsWxWork=1"
                            Me.PrintDebugLog("Redirect=" & strReUrl & "?" & strPara)
                            HcMain.Response.Redirect(strReUrl & "?" & strPara)
                        End If
                    End With
                    Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("ReLogin", strStepName, ex)
            HcMain.Response.Write("<p><h1>出错</h1></p>")
            If bolIsNoLog = False Then Me.mPrintErrLogInf(Me.LastErr)
        End Try
    End Sub
    Private Function mPrintErrLogInf(ErrInf As String) As String
        Try
            moPigFunc.OptLogInf(ErrInf, goIISMain.LogFilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mPrintErrLogInf", ex)
        End Try
    End Function

End Class
