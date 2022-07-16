'**********************************
'* Name: fIISMain
'* Author: 萧鹏
'* Describe: IIS常驻进程主程序
'* Version: 1.0.4
'* Create Time: 10/3/2021
'* 1.0.2  11/3/2021   初始化
'* 1.0.3  4/4/2021   修改
'* 1.0.4  6/4/2021   修改GetKeyValue,SaveKeyValue
'**********************************
Imports PigToolsLib
Imports ObjAdoDBLib
Imports ObjScriptingLib

Friend Class fIISMain
	Inherits fPigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"
	Private moPigFunc As New PigFunc
	Private moFS As New FileSystemObject
	Public Connection As Connection
	Public JSonConf As fPigJSon
	Public PigKeyValueList As New PigKeyValueList

	Public Sub New()
		MyBase.New(CLS_VERSION)
		Me.mNew()
	End Sub

	Public Sub New(ConfFilePath As String, LogFilePath As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(ConfFilePath, LogFilePath)
	End Sub

	Private Function mPrintErrLogInf(ErrInf As String) As String
		Try
			moPigFunc.OptLogInf(ErrInf, Me.LogFilePath)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf("PrintErrLogInf", ex)
		End Try
	End Function

	Private Sub mNew(Optional ConfFilePath As String = "", Optional LogDirPath As String = "")
		Try
			Dim strAppPath As String = Me.AppPath
			If Right(strAppPath, 1) = "\" Then strAppPath = Left(strAppPath, strAppPath.Length - 1)
			If ConfFilePath = "" Then
				Me.ConfFilePath = strAppPath & "\Bin\" & Me.AppTitle & ".conf"
			Else
				Me.ConfFilePath = ConfFilePath
			End If
			If LogDirPath = "" Then
				Me.LogDirPath = moPigFunc.GetFilePart(strAppPath, PigFunc.enmFilePart.Path) & "\" & moPigFunc.GetFilePart(strAppPath, PigFunc.enmFilePart.FileTitle) & "Log"
			Else
				Me.LogDirPath = LogDirPath
			End If

			If moFS.FolderExists(moPigFunc.GetFilePart(Me.ConfFilePath, PigFunc.enmFilePart.Path)) = False Then
				Throw New Exception("配置文件目录" & Me.ConfFilePath & "不存在")
			End If
			If moFS.FolderExists(Me.LogDirPath) = False Then
				Throw New Exception("日志目录" & Me.LogDirPath & "不存在")
			End If
			Me.RefLogFilePath()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", ex)
			Me.mPrintErrLogInf(Me.LastErr)
		End Try
	End Sub

	Private Sub RefLogFilePath()
		Try
			Me.LogFilePath = Me.LogDirPath & "\" & Me.AppTitle & Format(Now, "yyyyMMdd") & ".log"
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RefLogFilePath", ex)
		End Try
	End Sub

	''' <summary>
	''' 日志目录路径
	''' </summary>
	Private mstrLogDirPath As String
	Public Property LogDirPath() As String
		Get
			Return mstrLogDirPath
		End Get
		Friend Set(ByVal value As String)
			mstrLogDirPath = value
		End Set
	End Property

	''' <summary>
	''' 日志文件路径
	''' </summary>
	Private mstrLogFilePath As String
	Public Property LogFilePath() As String
		Get
			Return mstrLogFilePath
		End Get
		Friend Set(ByVal value As String)
			mstrLogFilePath = value
		End Set
	End Property

	''' <summary>
	''' 配置文件路径
	''' </summary>
	Private mstrConfFilePath As String
	Public Property ConfFilePath() As String
		Get
			Return mstrConfFilePath
		End Get
		Friend Set(ByVal value As String)
			mstrConfFilePath = value
		End Set
	End Property

	''' <summary>
	''' 最近刷新配置时间
	''' </summary>
	Private moLastRefConfTime As DateTime = DateTime.MinValue
	Public Property LastRefConfTime() As DateTime
		Get
			Return moLastRefConfTime
		End Get
		Friend Set(ByVal value As DateTime)
			moLastRefConfTime = value
		End Set
	End Property

	''' <summary>
	''' 配置文件更新时间
	''' </summary>
	Private moConfFileUpdTime As DateTime = DateTime.MinValue
	Public Property ConfFileUpdTime() As DateTime
		Get
			Return moConfFileUpdTime
		End Get
		Friend Set(ByVal value As DateTime)
			moConfFileUpdTime = value
		End Set
	End Property

	''' <summary>
	''' 刷新配置
	''' </summary>
	Public Sub RefConf()
		Dim strStepName As String = ""
		Try
			If Math.Abs(DateDiff(DateInterval.Minute, Me.LastRefConfTime, Now)) > 1 Then
				Me.RefLogFilePath()
				If moFS.FileExists(Me.ConfFilePath) = False Then
					Dim oPigJSon As New fPigJSon
					With oPigJSon
						.AddEle("CorpId", "", True)
						.AddEle("CorpSecret", "")
						.AddEle("AgentId", "")
						.AddEle("RediretUri", "")
						.AddEle("State", "")
						.AddEle("WxWorkSrv", "")
						.AddEle("WxWorkDB", "")
						.AddEle("WxWorkUser", "")
						.AddEle("WxWorkPwd", "")
						.AddEle("IsDebug", False)
						.AddEle("IsGD", True)
						.AddSymbol(fPigJSon.xpSymbolType.EleEndFlag)
					End With
					Dim tsOut As TextStream
					strStepName = "OpenTextFile(" & Me.ConfFilePath & ",ForWriting)"
					tsOut = moFS.OpenTextFile(Me.ConfFilePath, FileSystemObject.IOMode.ForWriting, True)
					If moFS.LastErr <> "" Then Throw New Exception(moFS.LastErr)
					strStepName = "ReadAll"
					tsOut.Write(oPigJSon.MainJSonStr)
					If tsOut.LastErr <> "" Then Throw New Exception(tsOut.LastErr)
					tsOut.Close()
				End If
				strStepName = "GetFile(" & Me.ConfFilePath & ")"
				Dim oFile As File = moFS.GetFile(Me.ConfFilePath)
				If oFile.DateLastModified <> Me.ConfFileUpdTime Then
					If moFS.LastErr <> "" Then Throw New Exception(moFS.LastErr)
					Dim tsIn As TextStream
					strStepName = "OpenTextFile(" & Me.ConfFilePath & ",ForReading)"
					tsIn = moFS.OpenTextFile(Me.ConfFilePath, FileSystemObject.IOMode.ForReading, True)
					If moFS.LastErr <> "" Then Throw New Exception(moFS.LastErr)
					strStepName = "ReadAll"
					Dim strData As String = tsIn.ReadAll
					If tsIn.LastErr <> "" Then Throw New Exception(tsIn.LastErr)
					tsIn.Close()
					strStepName = "New fPigJSon.Conf"
					JSonConf = New fPigJSon(strData)
					If JSonConf.LastErr <> "" Then Throw New Exception(JSonConf.LastErr)
					If JSonConf.GetBoolValue("IsDebug") = True Then
						Me.SetDebug(Me.LogFilePath)
					End If
					Me.LastRefConfTime = Now
				End If
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RefConf", strStepName, ex)
			Me.mPrintErrLogInf(Me.LastErr)
		End Try
	End Sub

	''' <summary>
	''' 数据库连接是否可用
	''' </summary>
	Public ReadOnly Property IsDBConnReady() As Boolean
		Get
			If Connection Is Nothing Then
				Return False
			ElseIf Connection.State = 1 Then
				Return True
			Else
				Return False
			End If
		End Get
	End Property

	''' <summary>
	''' 登录凭证是否可用
	''' </summary>
	Public ReadOnly Property IsAccessTokenReady() As Boolean
		Get
			If Me.AccessToken = "" Then
				Return False
			ElseIf Me.AccessTokenExpiresTime < Now Then
				Return False
			Else
				Return True
			End If
		End Get
	End Property

	''' <summary>
	''' 登录凭证
	''' </summary>
	Private mstrAccessToken As String
	Public Property AccessToken() As String
		Get
			Return mstrAccessToken
		End Get
		Friend Set(ByVal value As String)
			mstrAccessToken = value
		End Set
	End Property

	''' <summary>
	''' 登录凭证超时时间
	''' </summary>
	Private moAccessTokenExpiresTime As DateTime
	Public Property AccessTokenExpiresTime() As DateTime
		Get
			Return moAccessTokenExpiresTime
		End Get
		Friend Set(ByVal value As DateTime)
			moAccessTokenExpiresTime = value
		End Set
	End Property



	''' <summary>
	''' 刷新数据库连接
	''' </summary>
	Public Sub RefDBConn()
		Dim strStepName As String = ""
		Try
			If Me.IsDBConnReady = False Then
				strStepName = "RefConf"
				Me.RefConf()
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
				Dim strConn As String
				With JSonConf
					strConn = "Provider=Sqloledb;Data Source=" & .GetStrValue("WxWorkSrv") & ";Database=" & .GetStrValue("WxWorkDB") & ";User ID=" & .GetStrValue("WxWorkUser") & ";Password=" & .GetStrValue("WxWorkPwd")
				End With
				strStepName = "Open Connection"
				Connection = New Connection
				With Connection
					.ConnectionTimeout = 5
					.CommandTimeout = 60
					.ConnectionString = strConn
					.Open()
					If .LastErr <> "" Then Throw New Exception(.LastErr)
				End With
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RefDBConn", strStepName, ex)
			Me.mPrintErrLogInf(Me.LastErr)
		End Try
	End Sub


	''' <summary>
	''' 公司标识
	''' </summary>
	Public ReadOnly Property CorpId() As String
		Get
			Try
				Return JSonConf.GetStrValue("CorpId")
			Catch ex As Exception
				Return ""
			End Try
		End Get
	End Property

	''' <summary>
	''' 是否广东分行企业微信
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property IsGD() As Boolean
		Get
			Try
				Return JSonConf.GetBoolValue("IsGD")
			Catch ex As Exception
				Return False
			End Try
		End Get
	End Property

	''' <summary>
	''' 应用的凭证密钥
	''' </summary>
	Public ReadOnly Property CorpSecret() As String
		Get
			Try
				Return JSonConf.GetStrValue("CorpSecret")
			Catch ex As Exception
				Return ""
			End Try
		End Get
	End Property

	Public Sub ClearAccessToken()
		Try
			Me.AccessToken = ""
			Me.AccessTokenExpiresTime = DateTime.MinValue
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("ClearAccessToken", ex)
		End Try
	End Sub

	Public Function GetKeyValue(KeyName As String) As PigKeyValue
		Dim strStepName As String = ""
		Try
			strStepName = ".GetItem(" & KeyName & ")"
			With Me.PigKeyValueList
				Dim oPigKeyValue As PigKeyValue = .GetItem(KeyName)
				If oPigKeyValue Is Nothing Then
					If Me.IsDBConnReady = False Then
						strStepName = "RefDBConn"
						Me.RefDBConn()
						If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
					End If
					Dim strSQL As String = "EXEC dbo.gespKeyValueQry @OptCode=2,@KeyName=" & SQLStr(KeyName)
					strStepName = "Execute(" & strSQL & ")"
					Dim rs As Recordset, strKeyValue As String
					rs = Me.Connection.Execute(strSQL)
					If Me.Connection.LastErr <> "" Then Throw New Exception(Me.Connection.LastErr)
					If rs.EOF = True Then
						strKeyValue = ""
					Else
						strKeyValue = rs.Fields.Item("KeyValue").StrValue
					End If
					If rs.Fields.Item("IsBinValue").BooleanValue = True Then
						Dim oPigText As New PigText(strKeyValue, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
						strKeyValue = oPigText.Text
						oPigText = Nothing
					End If
					Dim oWorkApp As WorkApp, bolIsReNew As Boolean = False, dteAccessTokenExpiresTime As DateTime, oPigJSon As fPigJSon, strTxRes As String = "", strJSonStr As String = ""
					If strKeyValue = "" Then
						bolIsReNew = True
					Else
						strStepName = "New fPigJSon(KeyValue)"
						oPigJSon = New fPigJSon(strKeyValue)
						If oPigJSon.LastErr = "" Then
							strStepName = "New WorkApp并使用缓存AccessToken"
							oWorkApp = New WorkApp(Me.CorpId, Me.CorpSecret, oPigJSon.GetStrValue("AccessToken"), oPigJSon.GetDateValue("ExpiresTime"))
							dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
							If oWorkApp.IsAccessTokenReady = False Then
								bolIsReNew = True
							Else
								Me.AccessToken = oWorkApp.AccessToken
								Me.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
							End If
						Else
							bolIsReNew = True
						End If
					End If
					If bolIsReNew = True Then
						strStepName = "ReNew WorkApp"
						oWorkApp = New WorkApp(Me.CorpId, Me.CorpSecret)
						If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
						strStepName = "ReNew RefAccessToken"
						If Me.IsDebug = True Then Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, ""))
						oWorkApp.RefAccessToken(True)
						If oWorkApp.LastErr <> "" Then Throw New Exception(oWorkApp.LastErr)
						dteAccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
						Me.AccessToken = oWorkApp.AccessToken
						Me.AccessTokenExpiresTime = oWorkApp.AccessTokenExpiresTime
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
						rs = Me.Connection.Execute(strSQL)
						If Me.Connection.LastErr <> "" Then
							Me.mPrintErrLogInf(Me.GetSubStepDebugInf(SUB_NAME, strStepName, Me.Connection.LastErr))
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
		Catch ex As Exception

		End Try
	End Function

	Public Sub SaveKeyValue(KeyName As String)
		Try

		Catch ex As Exception

		End Try
	End Sub


End Class
