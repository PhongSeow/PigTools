'**********************************
'* Name: NoScanDirItem
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Directory entries not scanned
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 23/6/2021
'* 1.1  1/3/2023   Use PigBaseLocal
'************************************
''' <summary>
''' Directory entries not scanned|不扫描的目录项
''' </summary>
Public Class NoScanDirItem
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1.1.2"
    Public ReadOnly Property DirPath As String
    Public Sub New(DirPath As String)
        MyBase.New(CLS_VERSION)
        Me.DirPath = DirPath
    End Sub

    Public ReadOnly Property IsMeNoScan(ByVal ChkDirPath As String) As Boolean
        Get
            Try
                Dim strDirPath As String, lngLen As Long
                ChkDirPath = UCase(ChkDirPath)
                strDirPath = UCase(Me.DirPath)
                lngLen = Len(strDirPath)
                If Left(ChkDirPath, lngLen) = strDirPath Then
                    IsMeNoScan = True
                Else
                    IsMeNoScan = False
                End If
            Catch ex As Exception
                Me.SetSubErrInf("IsMeNoScan", ex)
                Return False
            End Try
        End Get
    End Property



End Class
