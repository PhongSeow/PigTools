Imports GetFileAndDirListLib
Imports PigCmdLib
Imports PigToolsLiteLib


Public Class ConsoleDemo
    Public WithEvents GetFileAndDirListApp As GetFileAndDirListApp
    Public PigConsole As New PigConsole
    Public RootDirPath As String
    Public LogFilePath As String
    Public Ret As String
    Public PigFunc As New PigFunc

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to New")
            Console.WriteLine("Press B to Start")
            Console.WriteLine("Press C to Show Status")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("New")
                    Console.WriteLine("*******************")
                    Me.Ret = Me.PigConsole.GetLine("Enter the starting folder for the scan", Me.RootDirPath)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Me.Ret = Me.PigConsole.GetLine("Enter the log file path", Me.LogFilePath)
                    If Me.Ret <> "OK" Then Console.WriteLine(Me.Ret)
                    Me.GetFileAndDirListApp = New GetFileAndDirListApp(Me.RootDirPath, Me.LogFilePath)
                    Console.WriteLine("RootDirPath=" & Me.RootDirPath)
                    If Me.GetFileAndDirListApp.LastErr <> "" Then
                        Console.WriteLine(Me.GetFileAndDirListApp.LastErr)
                    End If
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("Start")
                    Console.WriteLine("*******************")
                    If Me.GetFileAndDirListApp Is Nothing Then
                        Console.WriteLine("GetFileAndDirListApp Is Nothing")
                    Else
                        Dim bolIsFull As Boolean = False
                        bolIsFull = Me.PigConsole.IsYesOrNo("Is full PigMD5")
                        Console.WriteLine("Start" & Me.GetFileAndDirListApp.DirListPath)
                        Me.GetFileAndDirListApp.Start(bolIsFull)
                        If Me.GetFileAndDirListApp.LastErr <> "" Then
                            Console.WriteLine(Me.GetFileAndDirListApp.LastErr)
                        Else
                            Console.WriteLine("Start scanning.")
                        End If
                    End If
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("Show Status")
                    Console.WriteLine("*******************")
                    If Me.GetFileAndDirListApp Is Nothing Then
                        Console.WriteLine("GetFileAndDirListApp Is Nothing")
                    Else
                        With Me.GetFileAndDirListApp
                            .RefStatus()
                            Console.WriteLine("RunStatus=" & .RunStatus.ToString)
                            Console.WriteLine("StartTime=" & .StartTime)
                            Console.WriteLine("TimeoutTime=" & .TimeoutTime)
                            Console.WriteLine("EndTime=" & .EndTime)
                            Console.WriteLine("UseSeconds=" & .UseSeconds)
                            Console.WriteLine("ScanFolders=" & .ScanFolders)
                            Console.WriteLine("ScanFiles=" & .ScanFiles)
                            Console.WriteLine("IsAbsolutePath=" & .IsAbsolutePath)
                            Console.WriteLine("IsFullPigMD5=" & .IsFullPigMD5)
                            Console.WriteLine("CurrProcThreadID=" & .CurrProcThreadID)
                            Console.WriteLine("RootDirPath=" & .RootDirPath)
                            Console.WriteLine("LogFilePath=" & .LogFilePath)
                        End With
                    End If
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub GetFileAndDirListApp_ScanFail(ErrInf As String) Handles GetFileAndDirListApp.ScanFail
        Console.WriteLine("GetFileAndDirListApp_ScanFail:" & ErrInf)
    End Sub

    Private Sub GetFileAndDirListApp_ScanOK() Handles GetFileAndDirListApp.ScanOK
        Console.WriteLine("GetFileAndDirListApp_ScanOK")
    End Sub
End Class
