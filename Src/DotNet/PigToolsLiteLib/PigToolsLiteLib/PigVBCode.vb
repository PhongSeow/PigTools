'**********************************
'* Name: PigVBCode
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: And generate VB code|且于生成VB的代码
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.5
'* Create Time: 16/6/2021
'* 1.1  1/7/2022    Modify MkCollectionClass
'* 1.2  6/7/2022    Modify MkCollectionClass
'* 1.3  2/11/2022   Add mMkBytes2Func,MkBytes2Func
'* 1.5  8/11/2022   Modify MkStr2Func,MkBytes2Func
'**********************************
Public Class PigVBCode
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.5.8"

    Private Enum EnmMkBytes2FuncRetType
        RetString = 0
        RetBase64 = 1
        RetBytes = 2
    End Enum

    Private ReadOnly Property mPigFunc As New PigFunc


    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <summary>
    ''' 生成集合类代码|Generate collection class code
    ''' </summary>
    ''' <param name="MemberClassName">成员类名称|Member class name</param>
    ''' <param name="MemberClassKeyName">成员类键值，类型默认字符型|Member class key value, default character type</param>
    ''' <returns></returns>
    Public Function MkCollectionClass(ByRef OutVBCode As String, MemberClassName As String, MemberClassKeyName As String) As String
        Try
            OutVBCode = "Imports PigToolsLiteLib" & vbCrLf
            OutVBCode &= "Public Class " & MemberClassName & "s" & vbCrLf
            OutVBCode &= vbTab & "Inherits PigBaseLocal" & vbCrLf
            OutVBCode &= vbTab & "Implements IEnumerable(Of " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & "Private Const CLS_VERSION As String = ""1.0.0""" & vbCrLf
            OutVBCode &= vbTab & "Private ReadOnly moList As New List(Of " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & "Public Sub New()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "MyBase.New(CLS_VERSION)" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf
            OutVBCode &= vbTab & "Public ReadOnly Property Count() As Integer" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return moList.Count" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Count"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return -1" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf
            OutVBCode &= vbTab & "Public Function GetEnumerator() As IEnumerator(Of " & MemberClassName & ") Implements IEnumerable(Of " & MemberClassName & ").GetEnumerator" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Return moList.GetEnumerator()" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf
            OutVBCode &= vbTab & "Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Return Me.GetEnumerator()" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf
            OutVBCode &= vbTab & "Public ReadOnly Property Item(Index As Integer) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return moList.Item(Index)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Item.Index"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf

            OutVBCode &= vbTab & "Public ReadOnly Property Item(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Item = Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Item = o" & MemberClassName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Item.Key"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf

            OutVBCode &= vbTab & "Public Function IsItemExists(" & MemberClassKeyName & ") As Boolean" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "IsItemExists = False" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "IsItemExists = True" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""IsItemExists"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return False" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Private Sub mAdd(NewItem As " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.IsItemExists(NewItem." & MemberClassKeyName & ") = True Then Throw New Exception(NewItem." & MemberClassKeyName & " & ""Already exists"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.Add(NewItem)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""mAdd"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Add(NewItem As " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Me.mAdd(NewItem)" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Function AddOrGet(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim LOG As New PigStepLog(""AddOrGet"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.IsItemExists(" & MemberClassKeyName & ") = True Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Me.Item(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Else" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Me.Add(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Public Function Add(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim LOG As New PigStepLog(""Add"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "LOG.StepName = ""New " & MemberClassName & """" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Dim o" & MemberClassName & " As New " & MemberClassName & "(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If o" & MemberClassName & ".LastErr <> """" Then Throw New Exception(o" & MemberClassName & ".LastErr)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "LOG.StepName = ""mAdd""" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.mAdd(o" & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.LastErr <> """" Then Throw New Exception(Me.LastErr)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Add = o" & MemberClassName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Private Sub Remove(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim strStepName As String = """"" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "strStepName = ""For Each""" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "strStepName = ""Remove "" & " & MemberClassKeyName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "moList.Remove(o" & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Remove.Key"", strStepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Remove(Index As Integer)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim strStepName As String = """"" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "strStepName = ""Index="" & Index.ToString" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.RemoveAt(Index)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Remove.Index"", strStepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Clear()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.Clear()" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Clear"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= "End Class"

            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function



    ''' <summary>
    ''' 将一个常量的字节数组变成一个函数的代码，达到混淆目的，通过调用该函数还原字节数组的内容|Turn a constant byte array into a function code to achieve the purpose of confusion, and restore the contents of the byte array by calling this function
    ''' </summary>
    ''' <param name="InBytes">输入的字节数组|Byte array entered</param>
    ''' <param name="FuncName">函数名|Function Name</param>
    ''' <param name="FuncCode">输出的函数代码|Output function code</param>
    ''' <returns></returns>
    Public Function MkBytes2Func(InBytes As Byte(), FuncName As String, ByRef FuncCode As String) As String
        Dim LOG As New PigStepLog("MkBytes2Func")
        Try
            LOG.StepName = "mMkBytes2Func"
            LOG.Ret = Me.mMkBytes2Func(InBytes, FuncName, EnmMkBytes2FuncRetType.RetBytes, FuncCode)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            FuncCode = "''' <summary>" & Convert.ToBase64String(InBytes) & "</summary>" & Me.OsCrLf & FuncCode
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 将一个字符串变成一个函数的代码，达到混淆目的，通过调用该函数还原字符串的内容|Turn a string into the code of a function to achieve the purpose of confusion, and restore the content of the string by calling the function
    ''' </summary>
    ''' <param name="InpStr">输入的字符串|String entered</param>
    ''' <param name="TextType">文本类型|Text Type</param>
    ''' <param name="FuncName">函数名|Function Name</param>
    ''' <param name="FuncCode">输出的函数代码|Output function code</param>
    ''' <returns></returns>
    Public Function MkStr2Func(InpStr As String, TextType As PigText.enmTextType, FuncName As String, ByRef FuncCode As String) As String
        Dim LOG As New PigStepLog("MkStr2Func")
        Try
            LOG.StepName = "New PigText"
            Dim ptIn As New PigText(InpStr, TextType)
            Dim abOut(0) As Byte
            LOG.StepName = "mMkBytes2Func"
            LOG.Ret = Me.mMkBytes2Func(ptIn.TextBytes, FuncName, EnmMkBytes2FuncRetType.RetString, FuncCode, TextType)
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            FuncCode = "''' <summary>" & InpStr & "</summary>" & Me.OsCrLf & FuncCode
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mMkBytes2Func(InBytes As Byte(), FuncName As String, MkBytes2FuncRetType As EnmMkBytes2FuncRetType, ByRef FuncCode As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UnknowOrBin) As String
        Dim LOG As New PigStepLog("mMkBytes2Func")
        Try
            Dim strRetType As String, strRetErr As String, strRetOK As String
            Select Case MkBytes2FuncRetType
                Case EnmMkBytes2FuncRetType.RetBytes
                    strRetType = "Byte()"
                    strRetErr = "Return Nothing"
                    strRetOK = "Return abMain"
                Case EnmMkBytes2FuncRetType.RetString, EnmMkBytes2FuncRetType.RetBase64
                    strRetType = "String"
                    strRetErr = "Return """""
                    If MkBytes2FuncRetType = EnmMkBytes2FuncRetType.RetBase64 Then
                        strRetOK = "Return Convert.ToBase64String(abMain)"
                    Else
                        strRetOK = ""
                    End If
                Case Else
                    Throw New Exception("Invalid MkBytes2FuncRetType")
            End Select
            Dim strData As String = ""
            Dim intLen As Integer = InBytes.Length, i As Integer, intPos As Integer
            Dim abOut(1, intLen - 1) As Integer
            For i = 0 To intLen - 1
                abOut(0, i) = -1
                abOut(1, i) = -1
            Next
            Do While True
                intPos = Me.mPigFunc.GetRandNum(0, intLen - 1)
                Dim bolIsFind As Boolean = False
                For i = 0 To intLen - 1
                    If abOut(1, i) = -1 Then
                        bolIsFind = True
                        Exit For
                    End If
                Next
                If bolIsFind = False Then Exit Do
                bolIsFind = False
                For i = 0 To intLen - 1
                    If abOut(1, i) = intPos Then
                        bolIsFind = True
                        Exit For
                    End If
                Next
                If bolIsFind = False Then
                    For i = 0 To intLen - 1
                        If abOut(0, i) = -1 Then
                            abOut(0, i) = InBytes(intPos)
                            abOut(1, i) = intPos
                            Exit For
                        End If
                    Next
                End If
            Loop
            Dim intLastPos As Integer = 0
            For i = 0 To intLen - 1
                intPos = abOut(1, i)
                Dim strAdd As String = ""
                If intPos > intLastPos Then
                    strAdd = "i += " & (intPos - intLastPos).ToString
                Else
                    strAdd = "i -= " & (intLastPos - intPos).ToString
                End If
                strAdd &= " : abMain(i) = " & abOut(0, i).ToString
                Me.mPigFunc.AddMultiLineText(strData, strAdd, 2)
                intLastPos = intPos
            Next
            FuncCode = ""
            Dim strLine As String = ""
            With Me.mPigFunc
                strLine = "Private Function " & FuncName & "(Optional FuncRes As String = ""OK"") As " & strRetType
                .AddMultiLineText(FuncCode, strLine)
                .AddMultiLineText(FuncCode, "Try", 1)
                .AddMultiLineText(FuncCode, "Dim abMain(" & (intLen - 1).ToString & ") As Byte", 2)
                .AddMultiLineText(FuncCode, "Dim i As Integer = 0", 2)
                FuncCode &= strData
                If MkBytes2FuncRetType = EnmMkBytes2FuncRetType.RetString Then
                    Dim strTextType As String
                    Select Case TextType
                        Case PigText.enmTextType.Ascii
                            strTextType = "PigText.enmTextType.Ascii"
                        Case PigText.enmTextType.Unicode
                            strTextType = "PigText.enmTextType.Unicode"
                        Case PigText.enmTextType.UTF8
                            strTextType = "PigText.enmTextType.UTF8"
                        Case Else
                            strTextType = "PigText.enmTextType.UTF8"
                    End Select
                    .AddMultiLineText(strRetOK, "Dim ptOut As New PigText(abMain, " & strTextType & ")")
                    .AddMultiLineText(strRetOK, FuncName & " = ptOut.Text", 2)
                    .AddMultiLineText(strRetOK, "ptOut = Nothing", 2)
                    .AddMultiLineText(FuncCode, strRetOK, 2)
                Else
                    .AddMultiLineText(FuncCode, strRetOK, 2)
                End If
                .AddMultiLineText(FuncCode, "Catch ex As Exception", 1)
                .AddMultiLineText(FuncCode, "FuncRes = ex.Message.ToString", 2)
                .AddMultiLineText(FuncCode, strRetErr, 2)
                .AddMultiLineText(FuncCode, "End Try", 1)
                .AddMultiLineText(FuncCode, "End Function")
            End With
            Return "OK"
        Catch ex As Exception
            FuncCode = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


End Class
