'**********************************
'* Name: PigFolders
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigFolder 的集合类|Collection class of PigFolder
'* Home Url: https://en.seowphong.com
'* Version: 1.0
'* Create Time: 11/6/2023
'* 1.1  22/11/2023   Add FilesSize,mGetAllFastPigMD5,AllFiles
'************************************
Public Class PigFolders
	Inherits PigBaseMini
	Implements IEnumerable(Of PigFolder)
	Private Const CLS_VERSION As String = "1.1.28"
	Private ReadOnly moList As New List(Of PigFolder)
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property Count() As Integer
		Get
			Try
				Return moList.Count
			Catch ex As Exception
				Me.SetSubErrInf("Count", ex)
				Return -1
			End Try
		End Get
	End Property
	Public Function GetEnumerator() As IEnumerator(Of PigFolder) Implements IEnumerable(Of PigFolder).GetEnumerator
		Return moList.GetEnumerator()
	End Function
	Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
	Public ReadOnly Property Item(Index As Integer) As PigFolder
		Get
			Try
				Return moList.Item(Index)
			Catch ex As Exception
				Me.SetSubErrInf("Item.Index", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(FolderPath As String) As PigFolder
		Get
			Try
				Item = Nothing
				For Each oPigFolder As PigFolder In moList
					If oPigFolder.FolderPath = FolderPath Then
						Item = oPigFolder
						Exit For
					End If
				Next
			Catch ex As Exception
				Me.SetSubErrInf("Item.Key", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function IsItemExists(FolderPath) As Boolean
		Try
			IsItemExists = False
			For Each oPigFolder As PigFolder In moList
				If oPigFolder.FolderPath = FolderPath Then
					IsItemExists = True
					Exit For
				End If
			Next
		Catch ex As Exception
			Me.SetSubErrInf("IsItemExists", ex)
			Return False
		End Try
	End Function
	Private Sub mAdd(NewItem As PigFolder)
		Try
			If Me.IsItemExists(NewItem.FolderPath) = True Then Throw New Exception(NewItem.FolderPath & "Already exists")
			moList.Add(NewItem)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mAdd", ex)
		End Try
	End Sub
	Public Sub Add(NewItem As PigFolder)
		Me.mAdd(NewItem)
	End Sub
	Public Function AddOrGet(FolderPath As String) As PigFolder
		Dim LOG As New PigStepLog("AddOrGet")
		Try
			If Me.IsItemExists(FolderPath) = True Then
				Return Me.Item(FolderPath)
			Else
				Return Me.Add(FolderPath)
			End If
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Public Function Add(FolderPath As String) As PigFolder
		Dim LOG As New PigStepLog("Add")
		Try
			LOG.StepName = "New PigFolder"
			Dim oPigFolder As New PigFolder(FolderPath)
			If oPigFolder.LastErr <> "" Then Throw New Exception(oPigFolder.LastErr)
			LOG.StepName = "mAdd"
			Me.mAdd(oPigFolder)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Add = oPigFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function
	Private Sub Remove(FolderPath)
		Dim strStepName As String = ""
		Try
			strStepName = "For Each"
			For Each oPigFolder As PigFolder In moList
				If oPigFolder.FolderPath = FolderPath Then
					strStepName = "Remove " & FolderPath
					moList.Remove(oPigFolder)
					Exit For
				End If
			Next
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Remove.Key", strStepName, ex)
		End Try
	End Sub
	Public Sub Remove(Index As Integer)
		Dim strStepName As String = ""
		Try
			strStepName = "Index=" & Index.ToString
			moList.RemoveAt(Index)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Remove.Index", strStepName, ex)
		End Try
	End Sub
	Public Sub Clear()
		Try
			moList.Clear()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Clear", ex)
		End Try
	End Sub

	''' <summary>
	''' The file space size of all directories|所有目录的文件空间大小
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property AllFilesSize() As Long
		Get
			Try
				AllFilesSize = 0
				For Each oPigFolder As PigFolder In moList
					AllFilesSize += oPigFolder.FilesSize
				Next
			Catch ex As Exception
				Me.SetSubErrInf("AllFilesSize", ex)
				Return -1
			End Try
		End Get
	End Property

	Public ReadOnly Property AllFiles() As Long
		Get
			Try
				AllFiles = 0
				For Each oPigFolder As PigFolder In moList
					oPigFolder.RefPigFiles()
					AllFiles += oPigFolder.PigFiles.Count
				Next
			Catch ex As Exception
				Me.SetSubErrInf("AllFiles", ex)
				Return -1
			End Try
		End Get
	End Property

	Public Function GetAllFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As PigFolder.EnmGetFastPigMD5Type, ScanSize As Integer) As String
		Return Me.mGetAllFastPigMD5(FastPigMD5, GetFastPigMD5Type, ScanSize)
	End Function

	Public Function GetAllFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As PigFolder.EnmGetFastPigMD5Type) As String
		Return Me.mGetAllFastPigMD5(FastPigMD5, GetFastPigMD5Type)
	End Function

	Public Function GetAllFastPigMD5(ByRef FastPigMD5 As String, GetFastPigMD5Type As PigFolder.EnmGetFastPigMD5Type, ScanSize As Integer) As String
		Try
			Dim strRet As String = ""
			Dim oPigMD5 As PigMD5 = Nothing
			strRet = Me.mGetAllFastPigMD5(oPigMD5, GetFastPigMD5Type, ScanSize)
			If strRet <> "OK" Then Throw New Exception(strRet)
			FastPigMD5 = oPigMD5.PigMD5
			oPigMD5 = Nothing
			Return "OK"
		Catch ex As Exception
			FastPigMD5 = ""
			Return Me.GetSubErrInf("GetAllFastPigMD5", ex)
		End Try
	End Function

	Public Function GetAllFastPigMD5(ByRef FastPigMD5 As String, GetFastPigMD5Type As PigFolder.EnmGetFastPigMD5Type) As String
		Try
			Dim strRet As String = ""
			Dim oPigMD5 As PigMD5 = Nothing
			strRet = Me.mGetAllFastPigMD5(oPigMD5, GetFastPigMD5Type)
			If strRet <> "OK" Then Throw New Exception(strRet)
			FastPigMD5 = oPigMD5.PigMD5
			oPigMD5 = Nothing
			Return "OK"
		Catch ex As Exception
			FastPigMD5 = ""
			Return Me.GetSubErrInf("GetAllFastPigMD5", ex)
		End Try
	End Function

	Private Function mGetAllFastPigMD5(ByRef FastPigMD5 As PigMD5, GetFastPigMD5Type As PigFolder.EnmGetFastPigMD5Type, Optional ScanSize As Integer = 20480) As String
		Dim LOG As New PigStepLog("mGetAllFastPigMD5")
		Try
			Dim pbMain As New PigBytes
			LOG.StepName = "SetValue"
			LOG.Ret = pbMain.SetValue(Me.Count)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf("Count")
				Throw New Exception(LOG.Ret)
			End If
			Dim oList As New List(Of String)
			For Each oPigFolder As PigFolder In moList
				oList.Add(oPigFolder.FolderPath)
			Next
			oList.Sort()
			For i = 0 To oList.Count - 1
				With Me.Item(oList.Item(i))
					Dim oPigMD5 As PigMD5 = Nothing
					LOG.Ret = .GetFastPigMD5(oPigMD5, GetFastPigMD5Type, ScanSize)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf("GetFastPigMD5")
						LOG.AddStepNameInf(.FolderPath)
						Throw New Exception(LOG.Ret)
					End If
					LOG.Ret = pbMain.SetValue(oPigMD5.PigMD5Bytes)
					If LOG.Ret <> "OK" Then
						LOG.AddStepNameInf("PigMD5")
						LOG.AddStepNameInf(.FolderPath)
						Throw New Exception(LOG.Ret)
					End If
				End With
			Next
			LOG.StepName = "New PigMD5"
			FastPigMD5 = New PigMD5(pbMain.Main)
			If FastPigMD5.LastErr <> "" Then
				LOG.Ret = FastPigMD5.LastErr
				Throw New Exception(LOG.Ret)
			End If
			pbMain = Nothing
			Return "OK"
		Catch ex As Exception
			FastPigMD5 = Nothing
			Return Me.GetSubErrInf("GetAllFastPigMD5", ex)
		End Try
	End Function

End Class