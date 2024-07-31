'**********************************
'* Name: WebLogicDeploy
'* Author: Seow Phong
'* License: Copyright (c) 2023 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Weblogic deployment application
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 28/2/2023
'************************************
Imports PigCmdLib
Imports PigToolsLiteLib
''' <summary>
''' WebLogic deployment application class|WebLogic部署应用类
''' </summary>
Public Class WebLogicDeploy
    Inherits PigBaseLocal
    Private Const CLS_VERSION As String = "1" & "." & "0" & "." & "10"

    Public Enum EnmModuleType
        War = 0
        Dir = 1
    End Enum
    Public ReadOnly Property DeployName As String
    Public ReadOnly Property Parent As WebLogicDomain

    Private mModuleType As EnmModuleType
    Public Property ModuleType As EnmModuleType
        Get
            Return mModuleType
        End Get
        Friend Set(value As EnmModuleType)
            mModuleType = value
        End Set
    End Property

    Private mSourcePath As String
    Public Property SourcePath As String
        Get
            Return mSourcePath
        End Get
        Friend Set(value As String)
            mSourcePath = value
        End Set
    End Property


    Public Sub New(DeployName As String, Parent As WebLogicDomain)
        MyBase.New(CLS_VERSION)
        Me.DeployName = DeployName
        Me.Parent = Parent
    End Sub
End Class
