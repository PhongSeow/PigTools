'**********************************
'* Name: CmdMenu
'* Author: Seow Phong
'* License: Copyright (c) 2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 命令行菜单项|Command line menu items
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 5/3/2024
'* 1.1  6/3/2024  Modify New
'**********************************
Imports System.ComponentModel
Imports PigToolsLiteLib
Friend Class CmdMenu
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "1" & "." & "2"

    Public ReadOnly Property MenuKey As String
    Public ReadOnly Property MenuText As String
    Public ReadOnly Property MenuType As EnmMenuItemType
    Public ReadOnly Property Parent As PigCmdMenu

    Private mHostKeyLetter As String = "A"
    Public Property HostKeyLetter As String
        Get
            Return mHostKeyLetter
        End Get
        Friend Set(value As String)
            mHostKeyLetter = value
        End Set
    End Property

    Public Enum EnmMenuItemType
        MenuItem = 0
        MenuBar = 1
    End Enum

    Public Sub New(Parent As PigCmdMenu, MenuType As EnmMenuItemType, MenuKey As String, MenuText As String)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
        Me.MenuType = MenuType
        Select Case Me.MenuType
            Case EnmMenuItemType.MenuBar
                Me.MenuKey = Me.Parent.mPigFunc.GetPKeyValue("-", False)
                Me.MenuText = "-"
            Case EnmMenuItemType.MenuItem
                Me.MenuKey = MenuKey
                Me.MenuText = MenuText
        End Select
    End Sub

    Friend ReadOnly Property mFullMenuText As String
        Get
            Try
                Dim strLine As String = ""
                strLine = "* "
                strLine &= Me.HostKeyLetter & " - "
                strLine &= Me.MenuText
                Return strLine
            Catch ex As Exception
                Me.SetSubErrInf("mFullMenuText.Get", ex)
                Return ""
            End Try
        End Get
    End Property

End Class
