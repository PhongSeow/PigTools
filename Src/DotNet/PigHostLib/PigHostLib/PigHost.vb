'**********************************
'* Name: 豚豚主机|PigHost
'* Author: Seow Phong
'* License: Copyright (c) 2021-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 主机信息处理|Host information processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 8/10/2021
'* 1.1    12/10/2021   Add New
'* 1.2    18/10/2021   Modify New
'* 1.3    23/7/2022   Modify New
'* 1.5    24/7/2022   Rename MyHost,
'* 1.6    26/7/2022   Modify Imports
'* 1.7    29/7/2022   Modify Imports
'* 1.8    4/10/2022   Rename PigHost, add mNew, modify New
'* 1.9    15/10/2022  Rename PigHost, add mNew, modify New
'**********************************
Imports PigToolsLiteLib
#If NETFRAMEWORK Then
Imports PigSQLSrvLib
#Else
Imports PigSQLSrvCoreLib
#End If
Imports PigCmdLib

Public Class PigHost
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.8.2"

    ''' <summary>
    ''' 主机名|host name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property HostID As String
    ''' <summary>
    ''' 操作系统名称|Operating system name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property OSCaption As String
    Public ReadOnly Property UUID As String

    Private Property mPigFunc As New PigFunc
    Private Property mPigSysCmd As New PigSysCmd

    Public Sub New(HostID)
        MyBase.New(CLS_VERSION)
        Me.HostID = HostID
    End Sub


    'Automatically generate statements - Begin
    Private mUpdateCheck As New UpdateCheck
    Public ReadOnly Property LastUpdateTime() As Date
        Get
            Return mUpdateCheck.LastUpdateTime
        End Get
    End Property
    Public ReadOnly Property IsUpdate(PropertyName As String) As Boolean
        Get
            Return mUpdateCheck.IsUpdated(PropertyName)
        End Get
    End Property
    Public ReadOnly Property HasUpdated() As Boolean
        Get
            Return mUpdateCheck.HasUpdated
        End Get
    End Property
    Public Sub UpdateCheckClear()
        mUpdateCheck.Clear()
    End Sub
    Private mHostIp As String
    Public Property HostIp() As String
        Get
            Return mHostIp
        End Get
        Friend Set(value As String)
            If value <> mHostIp Then
                Me.mUpdateCheck.Add("HostIp")
                mHostIp = value
            End If
        End Set
    End Property
    Private mHostName As String
    Public Property HostName() As String
        Get
            Return mHostName
        End Get
        Friend Set(value As String)
            If value <> mHostName Then
                Me.mUpdateCheck.Add("HostName")
                mHostName = value
            End If
        End Set
    End Property
    Private mIsUse As Boolean
    Public Property IsUse() As Boolean
        Get
            Return mIsUse
        End Get
        Friend Set(value As Boolean)
            If value <> mIsUse Then
                Me.mUpdateCheck.Add("IsUse")
                mIsUse = value
            End If
        End Set
    End Property
    Private mCreateTime As DateTime = #1/1/1900#
    Public Property CreateTime() As DateTime
        Get
            Return mCreateTime
        End Get
        Friend Set(value As DateTime)
            If value <> mCreateTime Then
                Me.mUpdateCheck.Add("CreateTime")
                mCreateTime = value
            End If
        End Set
    End Property
    Private mUpdateTime As DateTime = #1/1/1900#
    Public Property UpdateTime() As DateTime
        Get
            Return mUpdateTime
        End Get
        Friend Set(value As DateTime)
            If value <> mUpdateTime Then
                Me.mUpdateCheck.Add("UpdateTime")
                mUpdateTime = value
            End If
        End Set
    End Property
    Private mRemark As String
    Public Property Remark() As String
        Get
            Return mRemark
        End Get
        Friend Set(value As String)
            If value <> mRemark Then
                Me.mUpdateCheck.Add("Remark")
                mRemark = value
            End If
        End Set
    End Property
    Private mConfInf As String
    Public Property ConfInf() As String
        Get
            Return mConfInf
        End Get
        Friend Set(value As String)
            If value <> mConfInf Then
                Me.mUpdateCheck.Add("ConfInf")
                mConfInf = value
            End If
        End Set
    End Property
    Private mStatusInf As String
    Public Property StatusInf() As String
        Get
            Return mStatusInf
        End Get
        Friend Set(value As String)
            If value <> mStatusInf Then
                Me.mUpdateCheck.Add("StatusInf")
                mStatusInf = value
            End If
        End Set
    End Property


    Friend Function fFillByRs(ByRef InRs As Recordset, Optional ByRef UpdateCnt As Integer = 0) As String
        Try
            If InRs.EOF = False Then
                With InRs.Fields
                    If .IsItemExists("HostIp") = True Then
                        If Me.HostIp <> .Item("HostIp").StrValue Then
                            Me.HostIp = .Item("HostIp").StrValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("HostName") = True Then
                        If Me.HostName <> .Item("HostName").StrValue Then
                            Me.HostName = .Item("HostName").StrValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("IsUse") = True Then
                        If Me.IsUse <> .Item("IsUse").BooleanValue Then
                            Me.IsUse = .Item("IsUse").BooleanValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("CreateTime") = True Then
                        If Me.CreateTime <> .Item("CreateTime").DateValue Then
                            Me.CreateTime = .Item("CreateTime").DateValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("UpdateTime") = True Then
                        If Me.UpdateTime <> .Item("UpdateTime").DateValue Then
                            Me.UpdateTime = .Item("UpdateTime").DateValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("Remark") = True Then
                        If Me.Remark <> .Item("Remark").StrValue Then
                            Me.Remark = .Item("Remark").StrValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("ConfInf") = True Then
                        If Me.ConfInf <> .Item("ConfInf").StrValue Then
                            Me.ConfInf = .Item("ConfInf").StrValue
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsItemExists("StatusInf") = True Then
                        If Me.StatusInf <> .Item("StatusInf").StrValue Then
                            Me.StatusInf = .Item("StatusInf").StrValue
                            UpdateCnt += 1
                        End If
                    End If
                    Me.mUpdateCheck.Clear()
                End With
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fFillByRs", ex)
        End Try
    End Function


    Friend Function fFillByXmlRs(ByRef InXmlRs As XmlRS, RSNo As Integer, RowNo As Integer, Optional ByRef UpdateCnt As Integer = 0) As String
        Try
            If RowNo <= InXmlRs.TotalRows(RSNo) Then
                With InXmlRs
                    If .IsColExists(RSNo, "HostIp") = True Then
                        If Me.HostIp <> .StrValue(RSNo, RowNo, "HostIp") Then
                            Me.HostIp = .StrValue(RSNo, RowNo, "HostIp")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "HostName") = True Then
                        If Me.HostName <> .StrValue(RSNo, RowNo, "HostName") Then
                            Me.HostName = .StrValue(RSNo, RowNo, "HostName")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "IsUse") = True Then
                        If Me.IsUse <> .BooleanValue(RSNo, RowNo, "IsUse") Then
                            Me.IsUse = .BooleanValue(RSNo, RowNo, "IsUse")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "CreateTime") = True Then
                        If Me.CreateTime <> .DateValue(RSNo, RowNo, "CreateTime") Then
                            Me.CreateTime = .DateValue(RSNo, RowNo, "CreateTime")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "UpdateTime") = True Then
                        If Me.UpdateTime <> .DateValue(RSNo, RowNo, "UpdateTime") Then
                            Me.UpdateTime = .DateValue(RSNo, RowNo, "UpdateTime")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "Remark") = True Then
                        If Me.Remark <> .StrValue(RSNo, RowNo, "Remark") Then
                            Me.Remark = .StrValue(RSNo, RowNo, "Remark")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "ConfInf") = True Then
                        If Me.ConfInf <> .StrValue(RSNo, RowNo, "ConfInf") Then
                            Me.ConfInf = .StrValue(RSNo, RowNo, "ConfInf")
                            UpdateCnt += 1
                        End If
                    End If
                    If .IsColExists(RSNo, "StatusInf") = True Then
                        If Me.StatusInf <> .StrValue(RSNo, RowNo, "StatusInf") Then
                            Me.StatusInf = .StrValue(RSNo, RowNo, "StatusInf")
                            UpdateCnt += 1
                        End If
                    End If
                    Me.mUpdateCheck.Clear()
                End With
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fFillByXmlRs", ex)
        End Try
    End Function
    'Automatically generate statements - End

End Class
