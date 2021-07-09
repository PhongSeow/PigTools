VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NoScanDirItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public DirPath As String


Public Property Get IsMeNoScan(ByVal ChkDirPath As String) As Boolean
Dim strDirPath As String, lngLen As Long
    ChkDirPath = UCase(ChkDirPath)
    strDirPath = UCase(Me.DirPath)
    If InStr(Me.DirPath, ":\") > 0 Then
        '绝对目录，全匹配
        lngLen = Len(strDirPath)
        If Left(ChkDirPath, lngLen) = strDirPath Then
            IsMeNoScan = True
        Else
            IsMeNoScan = False
        End If
    Else    '相对目录，模糊匹配
        If InStr(ChkDirPath, strDirPath) > 0 Then
            IsMeNoScan = True
        Else
            IsMeNoScan = False
        End If
    End If
End Property

