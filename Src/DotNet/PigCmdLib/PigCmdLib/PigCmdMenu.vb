'**********************************
'* Name: PigCmdMenuApp
'* Author: Seow Phong
'* License: Copyright (c) 2024 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 命令行菜单应用|Command line menu application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3
'* Create Time: 5/3/2024
'* 1.1  8/3/2024  Modify SelectMenu, add mFullMenuTitle,mGetMaxLineCol,mGetMaxPage
'* 1.2  8/3/2024  Modify SelectMenu, add mRefreshMenuProperty,AddMenuItem,AddMenuBarItem,PageItems,CurrPage,RowCol,MaxLineCol
'* 1.3  13/3/2024  Modify mGetMaxLineCol
'**********************************
Imports PigToolsLiteLib
Public Class PigCmdMenu
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "3" & "." & "8"

    Friend ReadOnly Property mPigFunc As New PigFunc
    Public ReadOnly Property MenuTitle As String
    Public ReadOnly Property IsTopMenu As Boolean
    Private ReadOnly Property mCmdMenus As New CmdMenus

    Private mPageItems As Integer = 21
    Public Property PageItems As Integer
        Get
            Return mPageItems
        End Get
        Friend Set(value As Integer)
            If value < 1 Then value = 1
            If value > 21 Then value = 21
            mPageItems = value
            Me.mRefreshMenuProperty()
        End Set
    End Property

    Private mCurrPage As Integer = 1
    Public Property CurrPage As Integer
        Get
            Return mCurrPage
        End Get
        Friend Set(value As Integer)
            If value < 0 Then value = 0
            If value > Me.mGetMaxPage Then value = Me.mGetMaxPage
            mCurrPage = value
            Me.mRefreshMenuProperty()
        End Set
    End Property

    Private mCurrPageBeginIndex As Integer
    Private Property CurrPageBeginIndex() As Integer
        Get
            Return mCurrPageBeginIndex
        End Get
        Set(value As Integer)
            mCurrPageBeginIndex = value
        End Set
    End Property

    Private mCurrPageEndIndex As Integer
    Private Property CurrPageEndIndex() As Integer
        Get
            Return mCurrPageEndIndex
        End Get
        Set(value As Integer)
            mCurrPageEndIndex = value
        End Set
    End Property

    Private mHasPreviousPage As Boolean
    Private Property HasPreviousPage() As Boolean
        Get
            Return mHasPreviousPage
        End Get
        Set(value As Boolean)
            mHasPreviousPage = value
        End Set
    End Property

    Private mHasNextPage As Boolean
    Private Property HasNextPage() As Boolean
        Get
            Return mHasNextPage
        End Get
        Set(value As Boolean)
            mHasNextPage = value
        End Set
    End Property

    Private mRowCol As Integer = 76
    Public Property RowCol As Integer
        Get
            Return mRowCol
        End Get
        Friend Set(value As Integer)
            mRowCol = value
        End Set
    End Property

    Private mMaxLineCol As Integer
    Public Property MaxLineCol As Integer
        Get
            Return mMaxLineCol
        End Get
        Friend Set(value As Integer)
            mMaxLineCol = value
        End Set
    End Property

    ''' <summary>
    ''' 实例化|Instantiate
    ''' </summary>
    ''' <param name="MenuTitle">菜单标题|menu title</param>
    ''' <param name="IsTopMenu">是否顶层菜单|Is it a top-level menu</param>
    ''' </summary>
    Public Sub New(MenuTitle As String, IsTopMenu As Boolean)
        MyBase.New(CLS_VERSION)
        Me.MenuTitle = MenuTitle
        Me.IsTopMenu = IsTopMenu
    End Sub

    Public Function ClearMenuItems() As String
        Dim LOG As New PigStepLog("ClearMenuItems")
        Try
            LOG.StepName = "CmdMenus.Clear"
            Me.mCmdMenus.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function AddMenuItem(MenuKey As String, MenuText As String) As String
        Dim LOG As New PigStepLog("AddMenuItem")
        Try
            If Me.mCmdMenus.IsItemExists(MenuKey) = True Then
                LOG.StepName = "Check "
                LOG.AddStepNameInf(MenuKey)
                LOG.AddStepNameInf(MenuText)
                LOG.Ret = "The Menu Key already exists."
                Throw New Exception(LOG.Ret)
            End If
            LOG.StepName = "CmdMenus.Add"
            Me.mCmdMenus.Add(Me, CmdMenu.EnmMenuItemType.MenuItem, MenuKey, MenuText)
            If Me.mCmdMenus.LastErr <> "" Then
                LOG.Ret = Me.mCmdMenus.LastErr
                Throw New Exception(LOG.Ret)
            End If
            Me.mMaxLineCol = -1
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddMenuBarItem() As String
        Dim LOG As New PigStepLog("AddMenuBarItem")
        Try
            LOG.StepName = "CmdMenus.Add"
            Me.mCmdMenus.Add(Me, CmdMenu.EnmMenuItemType.MenuBar, "", "")
            If Me.mCmdMenus.LastErr <> "" Then
                LOG.Ret = Me.mCmdMenus.LastErr
                Throw New Exception(LOG.Ret)
            End If
            Me.mMaxLineCol = -1
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Private Function mGetMaxPage() As Integer
        Try
            Dim intMaxPage As Integer = Me.mCmdMenus.Count \ Me.mPageItems
            If Me.mCmdMenus.Count Mod Me.mPageItems > 0 Then
                intMaxPage += 1
            End If
            Return intMaxPage
        Catch ex As Exception
            Me.SetSubErrInf("mGetMaxPage", ex)
            Return 0
        End Try
    End Function

    Private ReadOnly Property mFullMenuTitle As String
        Get
            Try
                Dim strLine As String = ""
                strLine = "* "
                strLine &= Me.mPigFunc.GetAlignStr(Me.MenuTitle, PigFunc.EnmAlignment.Center, Me.RowCol - 4)
                strLine &= " *"
                Return strLine
            Catch ex As Exception
                Me.SetSubErrInf("mFullMenuTitle.Get", ex)
                Return ""
            End Try
        End Get
    End Property


    Private Function mGetMaxLineCol() As Integer
        Try
            Dim intMaxLineCol As Integer = Me.mPigFunc.LenA(Me.MenuTitle)
            For Each oCmdMenu As CmdMenu In Me.mCmdMenus
                With oCmdMenu
                    Dim intText As Integer
                    intText = Me.mPigFunc.LenA(.MenuText) + 4
                    If intText > intMaxLineCol Then
                        intMaxLineCol = intText
                    End If
                End With
            Next
            Me.RowCol = intMaxLineCol + 4
            Return intMaxLineCol
        Catch ex As Exception
            Me.SetSubErrInf("mGetMaxLineCol", ex)
            Return -1
        End Try
    End Function

    Private Function mRefreshMenuProperty() As String
        Dim LOG As New PigStepLog("mRefreshMenuProperty")
        Try
            If Me.mMaxLineCol < 0 Then
                Me.mMaxLineCol = Me.mGetMaxLineCol
            End If
            Dim intRows As Integer = 0
            Me.mCurrPageBeginIndex = (Me.CurrPage - 1) * Me.PageItems
            Me.mCurrPageEndIndex = Me.mCurrPageBeginIndex + Me.PageItems - 1
            If Me.mCurrPageEndIndex > Me.mCmdMenus.Count - 1 Then
                Me.mCurrPageEndIndex = Me.mCmdMenus.Count - 1
            End If
            Me.HasNextPage = False
            Me.HasPreviousPage = False
            For i = 0 To Me.mCmdMenus.Count - 1
                With Me.mCmdMenus.Item(i)
                    Select Case i
                        Case 0 To Me.mCurrPageBeginIndex - 1
                            If .MenuType <> CmdMenu.EnmMenuItemType.MenuBar And Me.mHasPreviousPage = False Then
                                Me.mHasPreviousPage = True
                            End If
                        Case > Me.mCurrPageEndIndex
                            If .MenuType <> CmdMenu.EnmMenuItemType.MenuBar And Me.mHasNextPage = False Then
                                Me.mHasNextPage = True
                            End If
                    End Select
                    .HostKeyLetter = ""
                End With
            Next
            Dim bytCurrLetter As Byte = Asc("A")
            For i = Me.mCurrPageBeginIndex To Me.mCurrPageEndIndex
                With Me.mCmdMenus.Item(i)
                    If .MenuType = CmdMenu.EnmMenuItemType.MenuItem Then
                        Dim strLetter As String = Chr(bytCurrLetter)
                        Select Case strLetter
                            Case "Q"
                                bytCurrLetter += 1
                                .HostKeyLetter = Chr(bytCurrLetter)
                                bytCurrLetter += 1
                            Case "N"
                                If Me.HasNextPage = False Then
                                    .HostKeyLetter = strLetter
                                Else
                                    bytCurrLetter += 1
                                    .HostKeyLetter = Chr(bytCurrLetter)
                                End If
                                bytCurrLetter += 1
                            Case "P"
                                If Me.HasPreviousPage = False Then
                                    .HostKeyLetter = strLetter
                                Else
                                    bytCurrLetter += 1
                                    .HostKeyLetter = Chr(bytCurrLetter)
                                End If
                                bytCurrLetter += 1
                            Case Else
                                .HostKeyLetter = strLetter
                                bytCurrLetter += 1
                        End Select
                    End If
                End With
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function SelectMenu(ByRef OutMenuKey As String) As String
        Dim LOG As New PigStepLog("SelectMenu")
        Try
            LOG.StepName = "mRefreshMenuProperty"
            LOG.Ret = Me.mRefreshMenuProperty
            If LOG.Ret <> "OK" Then
                Throw New Exception(LOG.Ret)
            End If
            Dim strBarLine As String = Me.mPigFunc.GetRepeatStr(Me.RowCol, "*")
            OutMenuKey = ""
            Do While True
                Dim strLine As String = ""
                Console.Clear()
                Console.WriteLine(strBarLine)
                Console.WriteLine(Me.mFullMenuTitle)
                Console.WriteLine(strBarLine)
                strLine = "* Q - "
                If Me.IsTopMenu = True Then
                    strLine &= Me.mPigFunc.GetAlignStr("Exit", PigFunc.EnmAlignment.Left, Me.RowCol)
                Else
                    strLine &= Me.mPigFunc.GetAlignStr("Up", PigFunc.EnmAlignment.Left, Me.RowCol)
                End If
                Console.WriteLine(strLine)
                If Me.HasPreviousPage = True Then
                    strLine = "* P - "
                    strLine &= Me.mPigFunc.GetAlignStr("Previous Page", PigFunc.EnmAlignment.Left, Me.RowCol)
                    Console.WriteLine(strLine)
                End If
                If Me.HasNextPage = True Then
                    strLine = "* N - "
                    strLine &= Me.mPigFunc.GetAlignStr("Next Page", PigFunc.EnmAlignment.Left, Me.RowCol)
                    Console.WriteLine(strLine)
                End If
                Console.WriteLine(strBarLine)
                Dim intRows As Integer = 0
                For i = Me.CurrPageBeginIndex To Me.CurrPageEndIndex
                    With Me.mCmdMenus.Item(i)
                        If .MenuType = CmdMenu.EnmMenuItemType.MenuBar Then
                            Console.WriteLine(strBarLine)
                        Else
                            Console.WriteLine((.mFullMenuText))
                        End If
                    End With
                Next
                Console.WriteLine(strBarLine)
                Dim oConsoleKey As ConsoleKey = Console.ReadKey(True).Key
                Select Case oConsoleKey
                    Case ConsoleKey.Q, ConsoleKey.Escape
                        OutMenuKey = ""
                        Exit Do
                    Case Else
                        Dim bolIsFind As Boolean = True
                        Select Case oConsoleKey
                            Case ConsoleKey.N
                                If Me.HasNextPage = True Then
                                    Me.CurrPage += 1
                                    bolIsFind = False
                                End If
                            Case ConsoleKey.P
                                If Me.HasPreviousPage = True Then
                                    Me.CurrPage -= 1
                                    bolIsFind = False
                                End If
                            Case Else
                                bolIsFind = True
                        End Select
                        If bolIsFind = True Then
                            Dim strKey As String = oConsoleKey.ToString
                            For Each oCmdMenu As CmdMenu In Me.mCmdMenus
                                If oCmdMenu.HostKeyLetter = strKey Then
                                    OutMenuKey = oCmdMenu.MenuKey
                                    Exit Do
                                End If
                            Next
                        End If
                End Select
            Loop
            Return "OK"
        Catch ex As Exception
            OutMenuKey = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
