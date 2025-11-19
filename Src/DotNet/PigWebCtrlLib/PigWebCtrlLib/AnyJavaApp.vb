'**********************************
'* Name: AnyJavaApp
'* Author: Seow Phong
'* License: Copyright (c) 2025 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: A universal JAVA application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 19/4/2025
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
Friend Class AnyJavaApp
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "2"

    Public ReadOnly Property AnyJavas As AnyJavas
    Friend ReadOnly Property PigProcApp As New PigProcApp

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.AnyJavas = New AnyJavas
    End Sub

    Public Function GetHostJavaProcList(ByRef OutJavaProcList As String) As String
        Try

        Catch ex As Exception

        End Try
    End Function

End Class
